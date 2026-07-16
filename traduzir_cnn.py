from pathlib import Path
import re
import shutil
from datetime import datetime

import win32com.client as win32

import fitz


BASE_DIR = Path(__file__).parent

INPUT_DIR = BASE_DIR / "entrada"
OUTPUT_DIR = BASE_DIR / "saida"
MODEL_PATH = BASE_DIR / "modelos" / "MODELO CNN DEISI.docx"

MONTHS = {
    "janeiro": "gennaio",
    "fevereiro": "febbraio",
    "março": "marzo",
    "abril": "aprile",
    "maio": "maggio",
    "junho": "giugno",
    "julho": "luglio",
    "agosto": "agosto",
    "setembro": "settembre",
    "outubro": "ottobre",
    "novembro": "novembre",
    "dezembro": "dicembre",
}

COUNTRIES = {
    "Itália": "Italia",
}

WD_FIND_CONTINUE = 1
WD_REPLACE_ALL = 2

NEW_DATE = datetime.now().strftime("%d/%m/%Y")

DATE_PATTERN = re.compile(
    r"(Poiché non risulta nient[’']altro nel documento da me tradotto, ho redatto il presente Strumento Pubblico di Traduzione nella città di Porto Alegre, il\s+)"
    r"(\d{2}/\d{2}/\d{4})"
    r"(\.)",
    flags=re.IGNORECASE
)


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def translate_month(date: str) -> str:
    for pt, it in MONTHS.items():
        date = date.replace(f" de {pt} de ", f" {it} ")
    return date


def normalize_country(country: str) -> str:
    return COUNTRIES.get(country.strip(), country.strip())


def between(text: str, start: str, end: str) -> str:
    i = text.find(start)

    if i == -1:
        return ""

    i += len(start)

    j = text.find(end, i)

    if j == -1:
        return ""

    return text[i:j].strip()


def extract_pdf_data(pdf_path: Path) -> dict[str, str]:

    with fitz.open(pdf_path) as pdf:
        text = " ".join(page.get_text() for page in pdf)

    text = normalize_text(text)

    data = {
        "nome": between(
            text,
            "naturalização brasileira de",
            ", filho(a) de",
        ),

        "pai": between(
            text,
            "filho(a) de",
            " e de ",
        ),

        "mae": between(
            text,
            " e de ",
            ", natural de",
        ),

        "local": normalize_country(
            between(
                text,
                "natural de(o)(a)",
                ", nascido(a)",
            )
        ),

        "data_nascimento": translate_month(
            between(
                text,
                "nascido(a) em",
                ".",
            )
        ),

        "hora_emissao": between(
            text,
            "emitida às",
            "h do dia",
        ),

        "data_emissao": between(
            text,
            "do dia",
            "(Hora e Data",
        ),

        "codigo": between(
            text,
            "código verificador",
            ",",
        ),
    }

    if not all(data.values()):
        missing = [k for k, v in data.items() if not v]
        raise ValueError(f"Campos não encontrados: {', '.join(missing)}")

    return data


def copy_model(pdf_path: Path) -> Path:
    OUTPUT_DIR.mkdir(exist_ok=True)

    output = OUTPUT_DIR / f"{pdf_path.stem}.docx"
    shutil.copy2(MODEL_PATH, output)

    return output


def open_word():
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    return word


def replace_text(doc, old, new):
    rng = doc.Content

    while rng.Find.Execute(FindText=old):
        rng.Text = new
        rng.Collapse(0)


def replace_final_date(doc):
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


def fill_document(doc_path: Path, data: dict[str, str]) -> None:

    word = open_word()
    doc = None

    try:
        doc = word.Documents.Open(str(doc_path))

        replacements = {
            "{{NOME}}": data["nome"],
            "{{PAI}}": data["pai"],
            "{{MAE}}": data["mae"],
            "{{LOCAL}}": data["local"],
            "{{DATA_NASCIMENTO}}": data["data_nascimento"],
            "{{HORA_EMISSAO}}": data["hora_emissao"],
            "{{DATA_EMISSAO}}": data["data_emissao"],
            "{{CODIGO}}": data["codigo"],
        }

        for old, new in replacements.items():
            replace_text(doc, old, new)
        
        replace_final_date(doc)
        
        doc.Save()

    finally:
        if doc is not None:
            doc.Close(False)

        word.Quit()


def process_pdf(pdf_path: Path) -> None:

    print(f"Processando: {pdf_path.name}")
    print()

    try:
        data = extract_pdf_data(pdf_path)

        doc_path = copy_model(pdf_path)

        fill_document(doc_path, data)

        print(f"Arquivo: {doc_path.name} salvo em: ./saida")

    except Exception as e:
        print(f"Erro em {pdf_path.name}: {e}")


def main():

    if not MODEL_PATH.exists():
        raise FileNotFoundError(f"Modelo não encontrado: {MODEL_PATH}")

    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {INPUT_DIR}")

    pdfs = sorted(INPUT_DIR.glob("*.pdf"))

    for pdf in pdfs:
        process_pdf(pdf)

if __name__ == "__main__":
    main()