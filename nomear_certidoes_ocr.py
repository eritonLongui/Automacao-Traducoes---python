import os
import sys
import re
import json
import time
import shutil
from pathlib import Path

import fitz  # PyMuPDF
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

# Encodagem UTF-8 no console Windows
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent
ENTRADA_DIR = BASE_DIR / "entrada"
SAIDA_DIR = BASE_DIR / "saida"

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else genai.Client()

# Modelo configurado para máxima velocidade
MODEL_NAME = "gemini-2.5-flash"


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]


def pdf_page_to_png_bytes(pdf_path: Path, page_num: int = 0) -> bytes:
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    pix = page.get_pixmap(dpi=150)
    img_bytes = pix.tobytes("png")
    doc.close()
    return img_bytes


def analisar_certidao_com_gemini(pdf_path: Path, max_retries: int = 5) -> dict:
    doc = fitz.open(pdf_path)
    texto_nativo = "\n".join([page.get_text("text") for page in doc]).strip()
    doc.close()

    prompt = """
    Analise a certidão fornecida e retorne estritamente um JSON:
    {
      "tipo": "NASCIMENTO" | "CASAMENTO" | "OBITO",
      "prefixo": "CN" | "CC" | "CO",
      "nome": "NOME COMPLETO"
    }

    Regras:
    1. NASCIMENTO: prefixo "CN", nome é NOME COMPLETO DO REGISTRADO.
    2. CASAMENTO: prefixo "CC", nome é "NOME DO CONJUGE 1 e NOME DO CONJUGE 2".
    3. OBITO: prefixo "CO", nome é NOME COMPLETO DO FALECIDO.
    4. Nomes em MAIÚSCULAS.
    5. Retorne APENAS o JSON válido.
    """

    if len(texto_nativo) > 50:
        contents = [f"Texto da Certidão:\n{texto_nativo}", prompt]
    else:
        img_bytes = pdf_page_to_png_bytes(pdf_path, page_num=0)
        contents = [types.Part.from_bytes(data=img_bytes, mime_type="image/png"), prompt]

    for tentativa in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json"
                )
            )
            return json.loads(response.text.strip())
        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg or "503" in err_msg or "UNAVAILABLE" in err_msg:
                match_wait = re.search(r"retry in (\d+(?:\.\d+)?)s", err_msg, flags=re.IGNORECASE)
                wait_seconds = float(match_wait.group(1)) + 1.5 if match_wait else 6.0
                print(f"   [!] Limite de cota atingido. Aguardando {wait_seconds:.1f}s...", flush=True)
                time.sleep(wait_seconds)
                continue
            raise e

    raise RuntimeError("Número máximo de tentativas atingido na API do Gemini.")


def find_pdf_files_in_entrada():
    if not ENTRADA_DIR.exists():
        return []

    pdf_entries = []
    for pdf_path in sorted(ENTRADA_DIR.glob("**/*.pdf")):
        rel_parent = pdf_path.relative_to(ENTRADA_DIR).parent
        pdf_entries.append((pdf_path, rel_parent))

    return pdf_entries


def main():
    if not ENTRADA_DIR.exists():
        print(f"[-] Pasta de entrada nao encontrada: {ENTRADA_DIR}", flush=True)
        return

    pdf_entries = find_pdf_files_in_entrada()
    if not pdf_entries:
        print(f"[-] Nenhum arquivo PDF encontrado em: {ENTRADA_DIR}", flush=True)
        return

    SAIDA_DIR.mkdir(parents=True, exist_ok=True)

    processados_com_sucesso = 0
    pastas_subdiretorios = set()

    print(f"[+] Modo Alta Velocidade (15 req/min) em {len(pdf_entries)} certidao(oes)...\n", flush=True)

    for pdf_file, rel_parent in pdf_entries:
        try:
            print(f" -> Lendo via Gemini API: {pdf_file.relative_to(ENTRADA_DIR)}", flush=True)
            
            dados = analisar_certidao_com_gemini(pdf_file)
            prefixo = dados.get("prefixo")
            nome = dados.get("nome")

            if not prefixo or not nome:
                print(f" [!] Falha ao identificar certidao em: {pdf_file.name}", flush=True)
                continue

            destino_dir = SAIDA_DIR / rel_parent
            destino_dir.mkdir(parents=True, exist_ok=True)

            nome_saida = safe_filename(f"{prefixo} {nome}.pdf")
            destino = destino_dir / nome_saida

            contador = 1
            while destino.exists():
                destino = destino_dir / safe_filename(f"{prefixo} {nome} ({contador}).pdf")
                contador += 1

            shutil.copy2(pdf_file, destino)
            print(f" [V] Salvo como: {destino.name}\n", flush=True)
            processados_com_sucesso += 1

            # Ajustado para 4.1s de intervalo (permitindo até 15 requisições por minuto com folga)
            time.sleep(4.1)

            if str(rel_parent) != ".":
                pastas_subdiretorios.add(ENTRADA_DIR / rel_parent)

        except Exception as err:
            print(f" [X] Erro ao processar arquivo {pdf_file.name}: {err}\n", flush=True)

    if processados_com_sucesso == len(pdf_entries):
        print("[+] Todos os arquivos foram renomeados e movidos com sucesso!", flush=True)
        
        for pasta in pastas_subdiretorios:
            if pasta.exists():
                print(f"[+] Removendo subpasta de entrada: {pasta}", flush=True)
                shutil.rmtree(pasta)

        for pdf_file, _ in pdf_entries:
            if pdf_file.exists():
                pdf_file.unlink()

        print(f"[+] Limpeza da pasta de entrada concluida!", flush=True)
    else:
        print(f"[!] Apenas {processados_com_sucesso}/{len(pdf_entries)} arquivos foram processados.", flush=True)
        print("[!] A pasta /entrada foi mantida por seguranca.", flush=True)


if __name__ == "__main__":
    main()
