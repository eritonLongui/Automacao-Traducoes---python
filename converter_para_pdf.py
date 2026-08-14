"""
Converte arquivos .docx contidos na pasta ./saida para PDF.

Fluxo e estrutura de busca baseados em numerar_traducoes.py:
- Varre a pasta ./saida (arquivos em subpastas ou na raiz de ./saida);
- Para cada arquivo .docx (ignorando arquivos temporários iniciados com ~$):
    - Gera o arquivo .pdf correspondente com o mesmo nome na mesma pasta do .docx;
    - Mantém o arquivo .docx original intocado na pasta.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path

try:
    import win32com.client  # type: ignore
except Exception as exc:  # pragma: no cover - dependência em ambiente Windows/pywin32
    win32com = None
    _win32_import_error = exc
else:
    win32com = win32com.client
    _win32_import_error = None

# Configuração de Logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
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


def discover_documents(base_dir: Path) -> list[Path]:
    """
    Descobre todos os documentos .docx na pasta 'saida',
    seguindo o mesmo fluxo de navegação/ordenação de numerar_traducoes.py.
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
        if path.suffix.lower() in SUPPORTED_EXTENSIONS:
            docs.append(path)
    return sorted(docs, key=lambda p: relative_display_path(p, base_dir).casefold())


def converter_docx_para_pdf(docx_path: Path, base_dir: Path, word_app) -> Path:
    """
    Abre o arquivo .docx e salva a cópia em .pdf em uma estrutura de pastas com o sufixo `_convertidos`.

    Estrutura de saída:
    - Se o arquivo estava em `saida/FamiliaSilva/doc.docx`, salva em `saida/FamiliaSilva_convertidos/doc.pdf`.
    - Se o arquivo estava direto na raiz `saida/doc.docx`, salva em `saida/saida_convertidos/doc.pdf`.
    """
    rel = docx_path.relative_to(base_dir)

    if len(rel.parts) > 1:
        # Está dentro de uma subpasta (ex: FamiliaSilva/doc.docx) -> FamiliaSilva_convertidos/doc.pdf
        pasta_orig = rel.parts[0]
        subpastas_internas = rel.parts[1:-1]
        nome_pasta_dest = f"{pasta_orig}_convertidos"
        destino_dir = base_dir / nome_pasta_dest / Path(*subpastas_internas)
    else:
        # Está direto na raiz de 'saida' -> mantêm direto na raiz de 'saida'
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
    if _win32_import_error:
        LOGGER.error(
            "Erro ao importar pywin32 (win32com.client). Certifique-se de que está rodando no Windows com o pacote pywin32 instalado."
        )
        LOGGER.error("Detalhes do erro: %s", _win32_import_error)
        sys.exit(1)

    documents = discover_documents(BASE_DIR)
    if not documents:
        LOGGER.info("Nenhum arquivo .docx encontrado em '%s'.", BASE_DIR)
        return

    LOGGER.info("Encontrados %d arquivo(s) .docx para conversão em PDF.", len(documents))

    word_app = None
    converters_sucesso = 0
    converters_erro = 0

    try:
        LOGGER.info("Iniciando Microsoft Word...")
        word_app = win32com.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        for docx_path in documents:
            rel_name = relative_display_path(docx_path, BASE_DIR)
            try:
                LOGGER.info("Convertendo: %s", rel_name)
                pdf_path = converter_docx_para_pdf(docx_path, BASE_DIR, word_app)
                LOGGER.info("  -> Gerado: %s", relative_display_path(pdf_path, BASE_DIR))
                converters_sucesso += 1
            except Exception as err:
                LOGGER.error("Falha ao converter %s: %s", rel_name, err)
                converters_erro += 1

    finally:
        if word_app is not None:
            try:
                word_app.Quit()
            except Exception:
                pass

    LOGGER.info("--- RESUMO DA CONVERSÃO ---")
    LOGGER.info("Sucesso: %d", converters_sucesso)
    LOGGER.info("Erros: %d", converters_erro)


if __name__ == "__main__":
    main()
