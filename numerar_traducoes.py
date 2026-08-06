"""
Numera traduções em ./saida, grava na planilha e atualiza apenas o rodapé dos .docx.
Fluxo:
- varre a pasta base e subpastas em busca de .docx;
- identifica a família pelo nome da subpasta (ou pelo nome do arquivo, se estiver direto na pasta base);
- evita duplicidade somente quando FAMÍLIAS + DOCUMENTO + DATA forem iguais;
- encontra a primeira linha com B:E vazias;
- grava a linha na planilha;
- atualiza o rodapé do Word;
- se o Word falhar depois da planilha gravada, o erro é registrado no log.
"""

from __future__ import annotations

import argparse
import dataclasses
import logging
import os
import re
import sys
import unicodedata
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Optional

try:
    import win32com.client  # type: ignore
except Exception as exc:  # pragma: no cover - hard dependency on Windows/pywin32
    win32com = None
    _win32_import_error = exc
else:
    win32com = win32com.client
    _win32_import_error = None

LOGGER = logging.getLogger("numerar_traducoes")

# =========================
# CONFIGURAÇÕES EDITÁVEIS
# =========================

BASE_DIR_DEFAULT = "./saida"
DEFAULT_FOOTER_PLACEHOLDERS = [r"Tradução"]
TRANSLATION_SUFFIX = "C"
SUPPORTED_EXTENSIONS = {".docx"}

PROJECT_DIR = Path(__file__).resolve().parent

BASE_DIR = PROJECT_DIR / "saida"
GOOGLE_SHEETS_ID = "1ToLtlZ-bKJWiu6wTV6ly-ao-qqrOc9Ixg8g203uoNK0"
GOOGLE_SHEETS_WORKSHEET = "DEISI"
GOOGLE_SHEETS_CREDENTIALS_JSON = PROJECT_DIR / "numero-traducao-82a46e53b009.json"

SHEET_FAMILY_HEADER = "FAMÍLIAS"
SHEET_DOCUMENT_HEADER = "DOCUMENTO"
SHEET_DATE_HEADER = "DATA"
SHEET_FOLHA_HEADER = "FOLHA Nº"
SHEET_TRANSLATION_HEADER = "N. TRADUÇÃO"


# =========================
# MODELOS DE DADOS
# =========================


@dataclasses.dataclass
class SheetRow:
    index: int  # 1-based row number in the sheet
    values: list[str]


class SheetError(RuntimeError):
    pass


class DocumentError(RuntimeError):
    pass


class BaseSheetClient:
    def load_rows(self) -> tuple[list[str], list[SheetRow]]:
        raise NotImplementedError

    def ensure_row_exists(self, row_number: int) -> None:
        raise NotImplementedError

    def update_row_values(self, row_number: int, start_column: str, values: list[str]) -> None:
        raise NotImplementedError


class GSpreadSheetClient(BaseSheetClient):
    def __init__(self, spreadsheet_id: str, worksheet_name: str, credentials_json: Optional[str] = None):
        try:
            import gspread  # type: ignore
        except Exception as exc:  # pragma: no cover - optional dependency
            raise SheetError(
                "gspread não está instalado. Instale gspread para gravar em planilha privada."
            ) from exc

        self._gspread = gspread
        self._spreadsheet_id = spreadsheet_id
        self._worksheet_name = worksheet_name
        self._credentials_json = credentials_json
        self._client = self._build_client()
        self._worksheet = self._open_worksheet()

    def _build_client(self):
        if self._credentials_json:
            return self._gspread.service_account(filename=self._credentials_json)
        return self._gspread.service_account()

    def _open_worksheet(self):
        spreadsheet = self._client.open_by_key(self._spreadsheet_id)
        if self._worksheet_name:
            return spreadsheet.worksheet(self._worksheet_name)
        return spreadsheet.sheet1

    def load_rows(self) -> tuple[list[str], list[SheetRow]]:
        values = self._worksheet.get_all_values()
        if not values:
            return [], []

        header = values[0]
        rows = [SheetRow(index=i + 2, values=row) for i, row in enumerate(values[1:])]
        return header, rows

    def ensure_row_exists(self, row_number: int) -> None:
        current_rows = int(getattr(self._worksheet, "row_count", 0) or 0)
        if row_number > current_rows:
            self._worksheet.add_rows(row_number - current_rows)

    def update_row_values(self, row_number: int, start_column: str, values: list[str]) -> None:
        end_column = chr(ord(start_column.upper()) + len(values) - 1)
        self.ensure_row_exists(row_number)
        self._worksheet.update(
            range_name=f"{start_column.upper()}{row_number}:{end_column}{row_number}",
            values=[values],
            value_input_option="RAW",
        )


# =========================
# UTILITÁRIOS
# =========================


def _normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text).strip().casefold()


def _is_blank(value: object) -> bool:
    return str(value).strip() == ""


def _split_family_token(value: str) -> str:
    # Regra padrão: "Nome da Família - algo" -> "Nome da Família"
    return value.split(" - ", 1)[0].strip()


def relative_display_path(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir)).replace(os.sep, "/")
    except ValueError:
        return path.name


def discover_documents(base_dir: Path) -> list[Path]:
    docs: list[Path] = []
    for path in base_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            docs.append(path)
    return sorted(docs, key=lambda p: relative_display_path(p, base_dir).casefold())


def extract_family_name(path: Path, base_dir: Path) -> str:
    """
    Regra:
    - se estiver dentro de uma subpasta, usa o nome da primeira subpasta;
    - se estiver direto na pasta base, usa o nome do arquivo sem extensão;
    - corta tudo após " - ".
    """
    rel = path.relative_to(base_dir)
    source = rel.parts[0] if len(rel.parts) > 1 else path.stem
    return _split_family_token(source)


def extract_document_name(path: Path) -> str:
    """Nome gravado na coluna DOCUMENTO."""
    return path.stem.strip()


def resolve_registration_date(path: Path) -> date:
    """Data de registro usada na coluna DATA e no número da tradução."""
    return date.today()


def build_translation_number(folha_number: int, registration_date: date) -> str:
    """Regra atual: folha/ano + sufixo, ex.: 3929/2026C."""
    return f"{folha_number}/{registration_date.year}{TRANSLATION_SUFFIX}"


def parse_placeholders(raw_value: str | None) -> list[str]:
    if not raw_value:
        return []
    return [p.strip() for p in raw_value.split("|") if p.strip()]


# =========================
# ÍNDICE DA PLANILHA
# =========================


class SheetIndex:
    def __init__(self, header: list[str], rows: list[SheetRow]):
        self.header = header
        self.rows = rows
        self.family_idx = self._find_col_index(SHEET_FAMILY_HEADER)
        self.document_idx = self._find_col_index(SHEET_DOCUMENT_HEADER)
        self.date_idx = self._find_col_index(SHEET_DATE_HEADER)
        self.folha_idx = self._find_col_index(SHEET_FOLHA_HEADER)
        self.translation_idx = self._find_col_index(SHEET_TRANSLATION_HEADER)

    def _find_col_index(self, name: str) -> int:
        target = _normalize(name)
        for i, col in enumerate(self.header):
            if _normalize(col) == target:
                return i
        raise SheetError(
            f"Não encontrei a coluna '{name}' no cabeçalho da planilha. Cabeçalho detectado: {self.header}"
        )

    def _value_at(self, row: SheetRow, idx: int) -> str:
        if idx >= len(row.values):
            return ""
        return str(row.values[idx]).strip()

    def family_exists(self, family: str) -> list[SheetRow]:
        normalized_family = _normalize(family)
        matches: list[SheetRow] = []
        for row in self.rows:
            if self.family_idx >= len(row.values):
                continue
            if _normalize(row.values[self.family_idx]) == normalized_family:
                matches.append(row)
        return matches

    def document_exists(self, family: str, document: str, registration_date: date) -> Optional[SheetRow]:
        normalized_family = _normalize(family)
        normalized_document = _normalize(document)
        normalized_date = registration_date.strftime("%d/%m/%Y")
        for row in self.rows:
            if (
                self.family_idx >= len(row.values)
                or self.document_idx >= len(row.values)
                or self.date_idx >= len(row.values)
            ):
                continue

            if (
                _normalize(row.values[self.family_idx]) == normalized_family
                and _normalize(row.values[self.document_idx]) == normalized_document
                and _normalize(row.values[self.date_idx]) == _normalize(normalized_date)
            ):
                return row
        return None

    def next_folha_number(self) -> int:
        numbers: list[int] = []
        for row in self.rows:
            if self.folha_idx >= len(row.values):
                continue
            raw = str(row.values[self.folha_idx]).strip()
            if not raw:
                continue
            try:
                numbers.append(int(float(raw)))
            except ValueError:
                continue
        return (max(numbers) + 1) if numbers else 1

    def first_available_row(self) -> int:
        """
        Retorna a primeira linha onde B:E estejam vazias.
        Se não houver linha disponível no bloco atual, retorna a próxima linha após o último registro carregado.
        """
        for row in self.rows:
            b = self._value_at(row, self.family_idx)
            c = self._value_at(row, self.document_idx)
            d = self._value_at(row, self.date_idx)
            e = self._value_at(row, self.folha_idx)
            if all(_is_blank(v) for v in (b, c, d, e)):
                return row.index

        return (self.rows[-1].index + 1) if self.rows else 2

    def build_row(self, family: str, document: str, registration_date: date, folha_number: int) -> list[str]:
        width = max(
            len(self.header),
            self.family_idx + 1,
            self.document_idx + 1,
            self.date_idx + 1,
            self.folha_idx + 1,
            self.translation_idx + 1,
            6,
        )
        row = [""] * width
        row[self.family_idx] = family
        row[self.document_idx] = document
        row[self.date_idx] = registration_date.strftime("%d/%m/%Y")
        row[self.folha_idx] = str(folha_number)
        # A coluna F é mantida pela planilha (fórmulas/ordenação). O script não a sobrescreve.
        return row[:5]  # A:E

    def write_local_row(self, row_number: int, row_values: list[str]) -> None:
        """
        Mantém o cache local sincronizado depois de gravar na planilha.
        """
        for row in self.rows:
            if row.index == row_number:
                row.values = row_values
                return
        self.rows.append(SheetRow(index=row_number, values=row_values))


# =========================
# EDIÇÃO DO WORD
# =========================


def _normalize_footer_text(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = text.replace("\xa0", " ")
    text = text.replace("\r", " ").replace("\x07", " ")
    text = text.replace("n°", "nº").replace("N°", "Nº")
    text = re.sub(r"\s+", " ", text).strip().casefold()
    return text


def update_word_document_footer(path: Path, translation_number: str, patterns: list[str]) -> int:
    if win32com is None:
        raise DocumentError(f"pywin32 / win32com não está disponível: {_win32_import_error}")

    word = None
    doc = None

    try:
        word = win32com.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        doc = word.Documents.Open(str(path.resolve()))
        replacement_text = f"Tradução nº {translation_number}"

        section = doc.Sections(1)
        footer = section.Footers(1).Range

        if footer.Paragraphs.Count < 2:
            raise DocumentError(
                f"Rodapé inesperado em {path.name}: esperava pelo menos 2 parágrafos no footer."
            )

        target_paragraph = footer.Paragraphs(2)
        current_text = str(target_paragraph.Range.Text)

        # LOGGER.info("Texto atual do rodapé alvo: %r", current_text)

        target_paragraph.Range.Text = replacement_text

        # Confirma se a troca foi aplicada
        updated_text = str(target_paragraph.Range.Text)
        LOGGER.info("Rodapé alterado para: %r", updated_text)

        if replacement_text not in updated_text:
            raise DocumentError(
                f"A substituição no rodapé parece não ter sido aplicada em {path.name}. "
                f"Esperado: {replacement_text!r}. Encontrado: {updated_text!r}"
            )

        doc.Save()
        return 1

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def update_document(path: Path, translation_number: str, placeholders: list[str]) -> int:
    if path.suffix.lower() != ".docx":
        raise DocumentError(f"Extensão não suportada para este fluxo: {path.suffix}")
    return update_word_document_footer(path, translation_number, placeholders)


# =========================
# PLANILHA
# =========================


def load_sheet_client(spreadsheet_id: str, worksheet_name: str, credentials_json: Optional[str]) -> BaseSheetClient:
    spreadsheet_id = spreadsheet_id.strip()
    worksheet_name = worksheet_name.strip()

    if not spreadsheet_id:
        raise SheetError(
            "Configure o ID da planilha (ou use --sheet-id) para acessar a planilha privada."
        )

    return GSpreadSheetClient(
        spreadsheet_id=spreadsheet_id,
        worksheet_name=worksheet_name,
        credentials_json=credentials_json,
    )


# =========================
# PROCESSAMENTO
# =========================


def process_documents(
    base_dir: Path,
    client: BaseSheetClient,
    placeholders: list[str],
    dry_run: bool,
) -> int:
    header, rows = client.load_rows()
    if not header:
        raise SheetError("A planilha não possui cabeçalho ou está vazia.")

    index = SheetIndex(header, rows)
    docs = discover_documents(base_dir)
    processed = 0

    # LOGGER.info("%d arquivo(s) .docx encontrado(s) em %s", len(docs), base_dir)

    grouped: dict[str, list[Path]] = defaultdict(list)
    for doc_path in docs:
        family = extract_family_name(doc_path, base_dir)
        grouped[family].append(doc_path)

    for family in sorted(grouped.keys(), key=_normalize):
        family_docs = sorted(
            grouped[family],
            key=lambda p: relative_display_path(p, base_dir).casefold(),
        )

        LOGGER.info("família: %s", family)
        LOGGER.info("%s documento(s) encontrado(s)", len(family_docs))

        for doc_path in family_docs:
            rel_path = relative_display_path(doc_path, base_dir)
            document_name = extract_document_name(doc_path)
            registration_date = resolve_registration_date(doc_path)

            if index.document_exists(family, document_name, registration_date) is not None:
                LOGGER.info(
                    "Pulando já registrado: %s | data = %s",
                    document_name,
                    registration_date.strftime("%d/%m/%Y"),
                )
                continue

            folha_number = index.next_folha_number()
            translation_number = build_translation_number(folha_number, registration_date)
            row_values = index.build_row(
                family=family,
                document=document_name,
                registration_date=registration_date,
                folha_number=folha_number,
            )

            target_row = index.first_available_row()

            LOGGER.info("")
            LOGGER.info("==== %s ====", document_name)
            LOGGER.info("Registro concluído: linha = %s | tradução nº = %s", target_row, translation_number)

            if dry_run:
                LOGGER.info("[dry-run] Não gravei a planilha nem alterei o documento: %s", rel_path)
                continue

            sheet_altered = False
            try:
                client.update_row_values(target_row, "B", row_values[1:5])
                sheet_altered = True
                index.write_local_row(target_row, row_values)
            except Exception as exc:
                LOGGER.error(
                    "Falha ao gravar a planilha para %s. Planilha alterada: não. Motivo: %s",
                    rel_path,
                    exc,
                )
                raise

            try:
                changed = update_document(doc_path, translation_number, placeholders)
                if changed == 0:
                    raise DocumentError(
                        "Nenhum padrão de tradução foi encontrado no rodapé. Verifique o texto do modelo."
                    )
                # LOGGER.info("Atualizei %s com %d substituição(ões) no rodapé.", rel_path, changed)
            except Exception as exc:
                LOGGER.error(
                    "Falha ao editar o documento %s. Planilha alterada: %s. Motivo: %s",
                    rel_path,
                    "sim" if sheet_altered else "não",
                    exc,
                )
                raise

            processed += 1

    return processed


# =========================
# CLI
# =========================


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-dir", default=str(BASE_DIR))
    parser.add_argument("--sheet-id", default=GOOGLE_SHEETS_ID)
    parser.add_argument("--worksheet", default=GOOGLE_SHEETS_WORKSHEET)
    parser.add_argument("--creds-json", default=str(GOOGLE_SHEETS_CREDENTIALS_JSON))
    parser.add_argument(
        "--placeholders",
        default="|".join(DEFAULT_FOOTER_PLACEHOLDERS),
        help="Texto fixo do rodapé separado por |. Ex.: <<N_TRADUCAO>>|<<OUTRO>>",
    )
    parser.add_argument("--dry-run", action="store_true", help="Não grava na planilha nem altera arquivos")
    parser.add_argument("--log-level", default="INFO")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s"
    )

    base_dir = Path(args.base_dir).resolve()
    if not base_dir.exists():
        raise SystemExit(f"Pasta não encontrada: {base_dir}")

    placeholders = [p.strip() for p in str(args.placeholders).split("|") if p.strip()]
    if not placeholders:
        raise SystemExit("Nenhum placeholder informado para o rodapé.")

    try:
        client = load_sheet_client(
            spreadsheet_id=str(args.sheet_id),
            worksheet_name=str(args.worksheet),
            credentials_json=str(args.creds_json).strip() or None,
        )
        processed = process_documents(
            base_dir=base_dir,
            client=client,
            placeholders=placeholders,
            dry_run=args.dry_run,
        )
    except SheetError as exc:
        LOGGER.error(str(exc))
        return 2
    except DocumentError as exc:
        LOGGER.error(str(exc))
        return 3
    except Exception:
        LOGGER.exception("Execução interrompida por falha inesperada.")
        return 1

    LOGGER.info("")
    LOGGER.info("")
    LOGGER.info("Numerações concluídas")
    LOGGER.info("%d arquivo(s) processado(s).", processed)
    LOGGER.info("")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())