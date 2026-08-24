"""
Converte arquivos .docx contidos na pasta ./saida para PDF.

Fluxo:
- Varre a pasta ./saida (subpastas ou raiz);
- Para cada arquivo .docx, gera o arquivo .pdf correspondente em <pasta>_convertidos;
- Se todos os arquivos da pasta forem convertidos com sucesso (sem erros),
  a pasta original (.docx) é apagada e a pasta <pasta>_convertidos é renomeada para <pasta>.
"""

from __future__ import annotations

import gc
import logging
import os
import shutil
import sys
import time
from collections import defaultdict
from pathlib import Path

try:
    import win32com.client  # type: ignore
except Exception as exc:  # pragma: no cover - dependência em ambiente Windows/pywin32
    win32com = None
    _win32_import_error = exc
else:
    win32com = win32com.client
    _win32_import_error = None

# Configuração de Logs limpa (sem [INFO] e com formato direto)
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
)
LOGGER = logging.getLogger("converter_para_pdf")

PROJECT_DIR = Path(__file__).resolve().parent
BASE_DIR = PROJECT_DIR / "saida"
SUPPORTED_EXTENSIONS = {".docx"}
wdFormatPDF = 17  # Constante do MS Word para exportar como PDF


def relative_display_path(path: Path, base_dir: Path) -> str:
    try:
        return str(path.relative_to(base_dir)).replace(os.sep, "/")
    except ValueError:
        return path.name


def remover_pasta_com_retry(caminho: Path, retries: int = 5, delay: float = 0.5) -> None:
    for i in range(retries):
        try:
            if caminho.exists():
                shutil.rmtree(caminho)
            return
        except Exception as exc:
            if i == retries - 1:
                raise exc
            time.sleep(delay)


def renomear_pasta_com_retry(origem: Path, destino: Path, retries: int = 5, delay: float = 0.5) -> None:
    for i in range(retries):
        try:
            origem.rename(destino)
            return
        except Exception as exc:
            if i == retries - 1:
                raise exc
            time.sleep(delay)


def discover_documents(base_dir: Path) -> list[Path]:
    """
    Descobre todos os documentos .docx na pasta 'saida',
    ignorando arquivos temporários e pastas com sufixo `_convertidos`.
    """
    if not base_dir.exists():
        LOGGER.warning("A pasta '%s' não existe.", base_dir)
        return []

    docs: list[Path] = []
    for path in base_dir.rglob("*"):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if any(part.endswith("_convertidos") for part in path.parts):
            continue
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            docs.append(path)
    return sorted(docs, key=lambda p: relative_display_path(p, base_dir).casefold())


def converter_docx_para_pdf(docx_path: Path, base_dir: Path, word_app) -> Path:
    """
    Abre o arquivo .docx e salva a cópia em .pdf em uma estrutura de pastas com o sufixo `_convertidos`.
    """
    rel = docx_path.relative_to(base_dir)

    if len(rel.parts) > 1:
        pasta_orig = rel.parts[0]
        subpastas_internas = rel.parts[1:-1]
        nome_pasta_dest = f"{pasta_orig}_convertidos"
        destino_dir = base_dir / nome_pasta_dest / Path(*subpastas_internas)
    else:
        destino_dir = base_dir

    destino_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = destino_dir / f"{docx_path.stem}.pdf"

    abs_docx = str(docx_path.resolve())
    abs_pdf = str(pdf_path.resolve())

    doc = None
    try:
        doc = word_app.Documents.Open(
            abs_docx,
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
            Visible=False,
        )
        doc.SaveAs2(abs_pdf, FileFormat=wdFormatPDF)
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
            except Exception:
                pass

    return pdf_path


def main() -> None:
    LOGGER.info("")
    LOGGER.info("Iniciando conversão para PDF...")

    if _win32_import_error:
        LOGGER.error(
            "Erro ao importar pywin32 (win32com.client). Certifique-se de que está rodando no Windows com o pacote pywin32 instalado."
        )
        LOGGER.error("Detalhes do erro: %s", _win32_import_error)
        LOGGER.info("")
        sys.exit(1)

    documents = discover_documents(BASE_DIR)
    if not documents:
        LOGGER.info("Nenhum arquivo .docx encontrado na pasta '%s'.", BASE_DIR)
        LOGGER.info("(Os arquivos já foram convertidos para PDF ou a pasta está vazia).")
        LOGGER.info("")
        return

    LOGGER.info("Encontrados %d arquivo(s) .docx para conversão.", len(documents))
    LOGGER.info("")

    word_app = None
    converters_sucesso = 0
    converters_erro = 0

    pasta_stats: dict[str, dict] = defaultdict(
        lambda: {"total": 0, "sucessos": 0, "erros": 0, "caminho_orig": None, "caminho_dest": None}
    )

    for docx_path in documents:
        rel = docx_path.relative_to(BASE_DIR)
        if len(rel.parts) > 1:
            pasta_orig_nome = rel.parts[0]
            stats = pasta_stats[pasta_orig_nome]
            stats["total"] += 1
            stats["caminho_orig"] = BASE_DIR / pasta_orig_nome
            stats["caminho_dest"] = BASE_DIR / f"{pasta_orig_nome}_convertidos"

    try:
        word_app = win32com.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        for docx_path in documents:
            rel_name = relative_display_path(docx_path, BASE_DIR)
            rel = docx_path.relative_to(BASE_DIR)
            pasta_orig_nome = rel.parts[0] if len(rel.parts) > 1 else None

            try:
                LOGGER.info("Convertendo: %s", rel_name)
                pdf_path = converter_docx_para_pdf(docx_path, BASE_DIR, word_app)
                LOGGER.info("  -> Gerado: %s", relative_display_path(pdf_path, BASE_DIR))
                converters_sucesso += 1
                if pasta_orig_nome:
                    pasta_stats[pasta_orig_nome]["sucessos"] += 1
            except Exception as err:
                LOGGER.error("Falha ao converter %s: %s", rel_name, err)
                converters_erro += 1
                if pasta_orig_nome:
                    pasta_stats[pasta_orig_nome]["erros"] += 1

    finally:
        if word_app is not None:
            try:
                word_app.Quit()
            except Exception:
                pass
            del word_app
            gc.collect()

    # Finalização das pastas: se todas as certidões foram convertidas sem erros
    for pasta_nome, stats in pasta_stats.items():
        if stats["total"] > 0 and stats["erros"] == 0 and stats["sucessos"] == stats["total"]:
            orig = stats["caminho_orig"]
            dest = stats["caminho_dest"]
            if orig is not None and dest is not None and dest.exists():
                try:
                    remover_pasta_com_retry(orig)
                    renomear_pasta_com_retry(dest, orig)
                except Exception as e:
                    LOGGER.error("Erro ao substituir a pasta original '%s': %s", pasta_nome, e)
        elif stats["erros"] > 0:
            LOGGER.warning(
                "A pasta original '%s' foi mantida devido a erros na conversão (%d erro(s)).",
                pasta_nome,
                stats["erros"],
            )

    LOGGER.info("")
    LOGGER.info("--- RESUMO DA CONVERSÃO ---")
    LOGGER.info("Sucesso: %d", converters_sucesso)
    if converters_erro > 0:
        LOGGER.info("Erros: %d", converters_erro)
    LOGGER.info("")


if __name__ == "__main__":
    main()

