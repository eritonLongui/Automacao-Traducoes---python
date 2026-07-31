import re
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from win32com.client import DispatchEx


BASE_DIR = Path(__file__).resolve().parent
ENTRADA_DIR = BASE_DIR / "entrada"
SAIDA_DIR = BASE_DIR / "divididos"

START_PATTERN = r"Che lo sappiano tutti quelli che vedranno questo Strumento Pubblico"

END_PATTERN = (
    r"Poiché non risulta nient[’']altro nel documento da me tradotto, "
    r"ho redatto il presente Strumento Pubblico di Traduzione nella città di Porto Alegre, "
    r"il\s+\d{2}/\d{2}/\d{4}\."
)

TYPE_PATTERN = r"CERTIFICATO INTEGRALE DI\s+(NASCITA|MATRIMONIO|MORTE)"

NEW_DATE = datetime.now().strftime("%d/%m/%Y")

DATE_PATTERN = re.compile(
    r"(Poiché non risulta nient[’']altro nel documento da me tradotto, ho redatto il presente Strumento Pubblico di Traduzione nella città di Porto Alegre, il\s+)"
    r"(\d{2}/\d{2}/\d{4})"
    r"(\.)",
    flags=re.IGNORECASE
)

WORD_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def clean_text(s: str) -> str:
    return WORD_CTRL_RE.sub("", s).strip()


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]


def find_input_files():
    if not ENTRADA_DIR.exists():
        raise FileNotFoundError(f"Pasta de entrada não encontrada: {ENTRADA_DIR}")

    arquivos = sorted(ENTRADA_DIR.glob("*.doc"))
    if not arquivos:
        raise FileNotFoundError(f"Nenhum arquivo .doc encontrado em: {ENTRADA_DIR}")

    return arquivos


def find_cert_start_indices(paragraphs):
    starts = []
    for i, p in enumerate(paragraphs):
        txt = clean_text(p.Range.Text)
        if re.search(START_PATTERN, txt, flags=re.IGNORECASE):
            starts.append(i)
    return starts


def find_cert_end_index(paragraphs, start_i, search_limit):
    for j in range(start_i, search_limit):
        txt = clean_text(paragraphs[j].Range.Text)
        if re.search(END_PATTERN, txt, flags=re.IGNORECASE):
            return j
    return search_limit - 1


def get_prefix_and_name(paragraph_texts):
    joined = "\n".join(paragraph_texts)

    tipo_match = re.search(TYPE_PATTERN, joined, flags=re.IGNORECASE)
    if not tipo_match:
        return None, None

    tipo = tipo_match.group(1).upper()
    
    if tipo == "NASCITA":
        prefix = "CN"
    elif tipo == "MATRIMONIO":
        prefix = "CC"
    elif tipo == "MORTE":
        prefix = "CO"
    else:
        prefix = "DOC"

    nome = None

    for i, txt in enumerate(paragraph_texts):
        if re.fullmatch(r"\s*NOME\s*", txt, flags=re.IGNORECASE):
            for nxt in paragraph_texts[i + 1:]:
                nxt = nxt.strip()
                if nxt:
                    nome = nxt
                    break
            break

    if not nome:
        for txt in paragraph_texts:
            m = re.match(r"^\s*NOME\b\s*[:\-]?\s*(.+)$", txt, flags=re.IGNORECASE)
            if m:
                nome = m.group(1).strip()
                break

    if nome:
        nome = re.sub(r"\s+", " ", nome).strip()

    return prefix, nome


def is_blank_paragraph(paragraph):
    txt = clean_text(paragraph.Range.Text)
    return txt == ""


def trim_edge_blank_paragraphs(doc):
    changed = True
    while changed and doc.Paragraphs.Count > 0:
        changed = False

        first = doc.Paragraphs(1)
        if is_blank_paragraph(first):
            first.Range.Delete()
            changed = True
            continue

        if doc.Paragraphs.Count > 0:
            last = doc.Paragraphs(doc.Paragraphs.Count)
            if is_blank_paragraph(last):
                last.Range.Delete()
                changed = True
                continue

def remove_leading_layout_noise(doc):
    while doc.Paragraphs.Count > 0:

        # Remove quebra de página manual
        if doc.Content.End > 1 and doc.Range(0, 1).Text == "\x0c":
            doc.Range(0, 1).Delete()
            continue

        p = doc.Paragraphs(1)
        txt = clean_text(p.Range.Text)

        if txt == "":
            p.Range.Delete()
            continue

        try:
            if p.Range.ParagraphFormat.PageBreakBefore:
                p.Range.ParagraphFormat.PageBreakBefore = False
        except Exception:
            pass

        break

def normalize_first_paragraph(doc):
    """
    Remove quebra de página automática aplicada ao primeiro parágrafo útil.
    Isso ajuda a evitar a página em branco no começo em certidões curtas.
    """
    for i in range(1, min(doc.Paragraphs.Count, 5) + 1):
        p = doc.Paragraphs(i)
        if not is_blank_paragraph(p):
            try:
                p.Range.ParagraphFormat.PageBreakBefore = False
            except Exception:
                pass
            break


def replace_final_date(doc):
    """
    Procura a frase final e troca apenas a data, preservando o restante da formatação.
    """
    for p in doc.Paragraphs:
        txt = p.Range.Text
        m = DATE_PATTERN.search(txt)
        if m:
            para_start = p.Range.Start
            date_start = para_start + m.start(2)
            date_end = para_start + m.end(2)

            date_range = doc.Range(date_start, date_end)
            date_range.Text = NEW_DATE
            return True
    return False


def split_doc(input_doc: Path, output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)

    word = DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    doc = word.Documents.Open(
        str(input_doc),
        ReadOnly=True,
        AddToRecentFiles=False,
        ConfirmConversions=False
    )

    try:
        paragraphs = list(doc.Paragraphs)
        starts = find_cert_start_indices(paragraphs)

        if not starts:
            raise RuntimeError(f"Não encontrei o início das certidões em: {input_doc.name}")

        usados = {}

        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_dir = Path(tmp_dir)
            
            total_certidoes = 0

            for idx, start_i in enumerate(starts):
                search_limit = starts[idx + 1] if idx + 1 < len(starts) else len(paragraphs)
                end_i = find_cert_end_index(paragraphs, start_i, search_limit)

                segment_paragraphs = paragraphs[start_i:end_i + 1]
                segment_texts = [clean_text(p.Range.Text) for p in segment_paragraphs]

                prefix, nome = get_prefix_and_name(segment_texts)
                if not prefix or not nome:
                    prefix = "DOC"
                    nome = "SEM_NOME"

                base_name = safe_filename(f"{prefix} {nome}")
                usados.setdefault(base_name, 0)
                usados[base_name] += 1

                if usados[base_name] > 1:
                    file_name = f"{base_name} ({usados[base_name]})"
                else:
                    file_name = base_name

                out_path = output_dir / f"{file_name}.docx"

                start_pos = paragraphs[start_i].Range.Start
                end_pos = paragraphs[end_i].Range.End

                temp_copy = tmp_dir / f"{input_doc.stem}_{idx}.doc"
                shutil.copy2(input_doc, temp_copy)

                copy_doc = word.Documents.Open(
                    str(temp_copy),
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    ConfirmConversions=False
                )

                try:
                    if end_pos < copy_doc.Content.End:
                        copy_doc.Range(end_pos, copy_doc.Content.End).Delete()

                    if start_pos > 0:
                        copy_doc.Range(0, start_pos).Delete()
                        
                    first_char = copy_doc.Range(0, 1).Text

                    if first_char == "\x0e":
                        copy_doc.Range(0, 1).Delete()

                    trim_edge_blank_paragraphs(copy_doc)
                    remove_leading_layout_noise(copy_doc)
                    normalize_first_paragraph(copy_doc)
                    replace_final_date(copy_doc)

                    copy_doc.SaveAs2(str(out_path), FileFormat=16)
                    total_certidoes += 1
                    print(f"Salvo: {out_path.name}")

                finally:
                    copy_doc.Close(False)

        return total_certidoes

    finally:
        doc.Close(False)
        word.Quit()


if __name__ == "__main__":
    arquivos = find_input_files()

    total_geral = 0

    for arquivo in arquivos:
        pasta_saida_arquivo = SAIDA_DIR / safe_filename(arquivo.stem)
        print(f"Processando: {arquivo.name}")
        print()
        total_geral += split_doc(arquivo, pasta_saida_arquivo)

    print()
    print(f"Total de certidões processadas: {total_geral}")

    # for arquivo in arquivos:
    #     pasta_saida_arquivo = SAIDA_DIR / safe_filename(arquivo.stem)
    #     print(f"Processando: {arquivo.name}")
    #     split_doc(arquivo, pasta_saida_arquivo)