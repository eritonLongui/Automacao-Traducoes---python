import re
from pathlib import Path
from typing import List, Dict, Any

import win32com.client

WORD_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
MATRICOLA_RE = re.compile(r"^\s*MATRICOLA\b", re.IGNORECASE)
FIRMA_RE = re.compile(r"^\s*\[Firma\]\s*$", re.IGNORECASE)


def sanitize_word_text(text: str) -> str:
    if text is None:
        return ""
    text = text.replace("\x07", "")  # end-of-cell marker / control residual
    text = text.replace("\r", "")    # paragraph mark
    text = WORD_CTRL_RE.sub("", text)
    return text


def texto_para_busca(text: str) -> str:
    """
    Normaliza o texto só para detecção de padrões estruturais.
    """
    return sanitize_word_text(text).replace("\u00A0", " ").strip()


def abrir_word():
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def abrir_documento(word, arquivo: Path, read_only: bool = False):
    return word.Documents.Open(
        str(arquivo),
        ReadOnly=read_only,
        AddToRecentFiles=False,
        ConfirmConversions=False,
    )


def iter_docx_files(base_dir: Path):
    if not base_dir.exists():
        return []
    arquivos = [p for p in base_dir.rglob("*.docx") if p.is_file()]
    return sorted(arquivos)


def extrair_paragrafos(doc) -> List[Dict[str, Any]]:
    paragrafos = []
    total = doc.Paragraphs.Count

    for i in range(1, total + 1):
        p = doc.Paragraphs(i)
        raw_text = p.Range.Text
        text = sanitize_word_text(raw_text)

        paragrafos.append(
            {
                "id": i,  # 1-based
                "text": text,
                "start": int(p.Range.Start),
                "end": int(p.Range.End),
            }
        )

    return paragrafos


def localizar_faixa_corpo(doc) -> Dict[str, Any]:
    """
    Localiza o corpo útil da certidão:
    - começa após o parágrafo que contém 'MATRICOLA ...'
    - termina antes do parágrafo '[Firma]'
    - ignora linhas em branco logo após a matrícula
    """
    total = doc.Paragraphs.Count
    matricola_idx = None
    firma_idx = None

    for i in range(1, total + 1):
        texto = texto_para_busca(doc.Paragraphs(i).Range.Text)

        if matricola_idx is None and MATRICOLA_RE.search(texto):
            matricola_idx = i
            continue

        if matricola_idx is not None and FIRMA_RE.fullmatch(texto):
            firma_idx = i
            break

    if matricola_idx is None:
        raise ValueError("Não foi encontrado o bloco que começa com 'MATRICOLA'.")

    if firma_idx is None:
        raise ValueError("Não foi encontrado o marcador final '[Firma]'.")

    start_idx = None
    for i in range(matricola_idx + 1, firma_idx):
        texto = texto_para_busca(doc.Paragraphs(i).Range.Text)
        if texto:
            start_idx = i
            break

    if start_idx is None:
        raise ValueError("Não foi encontrado texto útil após 'MATRICOLA'.")

    end_idx = None
    for i in range(firma_idx - 1, start_idx - 1, -1):
        texto = texto_para_busca(doc.Paragraphs(i).Range.Text)
        if texto:
            end_idx = i
            break

    if end_idx is None:
        raise ValueError("Não foi encontrado texto útil antes de '[Firma]'.")

    start_char = int(doc.Paragraphs(start_idx).Range.Start)
    end_char = int(doc.Paragraphs(end_idx).Range.End)

    return {
        "start_paragraph_id": start_idx,
        "end_paragraph_id": end_idx,
        "paragraph_ids": list(range(start_idx, end_idx + 1)),
        "start_char": start_char,
        "end_char": end_char,
    }


def obter_range_corpo(doc):
    """
    Retorna o Range do corpo útil da certidão.
    """
    faixa = localizar_faixa_corpo(doc)
    return doc.Range(faixa["start_char"], faixa["end_char"])


def extrair_paragrafos_corpo(doc) -> List[Dict[str, Any]]:
    """
    Extrai apenas os parágrafos úteis entre MATRICOLA e [Firma].
    Mantém os IDs originais do Word.
    """
    faixa = localizar_faixa_corpo(doc)
    paragrafos = []

    for i in faixa["paragraph_ids"]:
        p = doc.Paragraphs(i)
        texto = sanitize_word_text(p.Range.Text).strip()

        if not texto:
            continue

        paragrafos.append(
            {
                "id": i,
                "text": texto,
                "start": int(p.Range.Start),
                "end": int(p.Range.End),
            }
        )

    return paragrafos