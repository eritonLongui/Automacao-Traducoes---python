import os
import sys
import re
import json
import time
import shutil
from pathlib import Path

import fitz  # PyMuPDF para fatiamento ultrarrápido sem perda de qualidade
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

# ---------------------------------------------------------
# Configuração de Caminhos Globais
# ---------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
ENTRADA_DIR = BASE_DIR / "entrada"
SAIDA_DIR = BASE_DIR / "saida"

# Inicializa SDK GenAI da Google
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else genai.Client()

# Modelos específicos por fluxo
MODELO_FLUXO_2_ARQUIVO_UNICO = "gemini-3.5-flash"
MODELO_FLUXO_1_PASTA_CERTIDOES = "gemini-3.1-flash-lite"


def safe_filename(name: str) -> str:
    """Sanitiza nomes de arquivos para evitar caracteres proibidos no Windows."""
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180]


def pdf_page_to_png_bytes(pdf_doc: fitz.Document, page_num: int) -> bytes:
    """Converte uma página do PyMuPDF para bytes PNG em memória."""
    page = pdf_doc[page_num]
    pix = page.get_pixmap(dpi=150)
    return pix.tobytes("png")


# =========================================================
# FLUXO 2: ARQUIVO ÚNICO COM TODAS AS CERTIDÕES (gemini-3.5-flash)
# =========================================================

def analisar_pdf_inteiro_em_uma_chamada(pdf_path: Path, max_retries: int = 5) -> list[dict]:
    """
    Envia TODAS as páginas do PDF juntas em UMA ÚNICA CHAMADA à API usando gemini-3.5-flash.
    Retorna a lista estruturada das certidões presentes no documento com páginas de início/fim, prefixo e nome.
    """
    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    print(f"[+] Fluxo 2 (Arquivo Único): Enviando PDF completo ({total_pages} páginas) em 1 chamada via {MODELO_FLUXO_2_ARQUIVO_UNICO}...", flush=True)

    contents = []

    # Constrói o payload com cada página identificada (Texto ou Imagem)
    for i in range(total_pages):
        page_text = doc[i].get_text("text").strip()
        contents.append(f"--- PÁGINA {i + 1} ---")
        if len(page_text) > 50:
            contents.append(page_text)
        else:
            img_bytes = pdf_page_to_png_bytes(doc, i)
            contents.append(types.Part.from_bytes(data=img_bytes, mime_type="image/png"))

    prompt_analise_total = f"""
    Você é um especialista em certidões cartorárias brasileiras (Nascimento, Casamento e Óbito).
    Analise todas as {total_pages} páginas deste arquivo PDF fornecido acima.
    Identifique a localização de TODAS as certidões presentes no documento.

    Responda EXCLUSIVAMENTE um JSON com uma lista contendo cada certidão encontrada:

    {{
      "certidoes": [
        {{
          "pagina_inicio": 1,
          "pagina_fim": 1,
          "tipo": "NASCIMENTO" | "CASAMENTO" | "OBITO",
          "prefixo": "CN" | "CC" | "CO",
          "nome": "NOME DO REGISTRADO OU CONJUGES"
        }}
      ]
    }}

    Regras Importantes:
    1. "pagina_inicio" e "pagina_fim" devem ser os números reais das páginas (base 1 a {total_pages}).
    2. Garanta que TODAS as páginas (de 1 a {total_pages}) pertençam a alguma certidão (sem lacunas).
    3. Para NASCIMENTO: prefixo "CN", nome é NOME COMPLETO DO REGISTRADO.
    4. Para CASAMENTO: prefixo "CC", nome deve ser "NOME DO CONJUGE 1 e NOME DO CONJUGE 2".
    5. Para OBITO: prefixo "CO", nome é NOME COMPLETO DO FALECIDO.
    6. Nomes sempre em MAIÚSCULAS.
    7. Retorne APENAS o JSON estruturado.
    """

    contents.append(prompt_analise_total)

    for tentativa in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=MODELO_FLUXO_2_ARQUIVO_UNICO,
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json"
                )
            )
            dados = json.loads(response.text.strip())
            certidoes = dados.get("certidoes", [])
            doc.close()
            return certidoes

        except Exception as e:
            err_msg = str(e)
            if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg or "503" in err_msg or "UNAVAILABLE" in err_msg:
                match_wait = re.search(r"retry in (\d+(?:\.\d+)?)s", err_msg, flags=re.IGNORECASE)
                wait_seconds = float(match_wait.group(1)) + 2.0 if match_wait else 8.0
                print(f"   [!] Limite de cota ativado ({MODELO_FLUXO_2_ARQUIVO_UNICO}). Aguardando {wait_seconds:.1f}s...", flush=True)
                time.sleep(wait_seconds)
                continue
            doc.close()
            raise e

    doc.close()
    raise RuntimeError("Falha ao analisar PDF completo na API do Gemini.")


def fatiar_pdf_unico(pdf_path: Path):
    """
    Recebe o resultado da chamada única do Fluxo 2 e realiza o fatiamento físico via PyMuPDF.
    """
    certidoes_meta = analisar_pdf_inteiro_em_uma_chamada(pdf_path)

    if not certidoes_meta:
        raise ValueError("Nenhuma certidão estruturada foi identificada no PDF pela API do Gemini.")

    doc_original = fitz.open(pdf_path)
    total_paginas = len(doc_original)

    # Pasta de destino em /saida com o nome do PDF original sem extensão
    nome_pasta_saida = pdf_path.stem
    pasta_destino = SAIDA_DIR / nome_pasta_saida
    pasta_destino.mkdir(parents=True, exist_ok=True)

    print(f"[+] Fatiando fisicamente {len(certidoes_meta)} certidão(ões) via PyMuPDF...", flush=True)

    paginas_cobertas = set()

    for idx, item in enumerate(certidoes_meta, start=1):
        pag_inicio = int(item.get("pagina_inicio", 1)) - 1
        pag_fim = int(item.get("pagina_fim", 1)) - 1

        pag_inicio = max(0, min(pag_inicio, total_paginas - 1))
        pag_fim = max(pag_inicio, min(pag_fim, total_paginas - 1))

        for p in range(pag_inicio, pag_fim + 1):
            paginas_cobertas.add(p)

        prefixo = item.get("prefixo") or "CN"
        nome = item.get("nome") or f"CERTIDAO_{idx:03d}"

        novo_doc = fitz.open()
        novo_doc.insert_pdf(doc_original, from_page=pag_inicio, to_page=pag_fim)

        base_filename = safe_filename(f"{prefixo} {nome}.pdf")
        caminho_saida = pasta_destino / base_filename

        contador = 1
        while caminho_saida.exists():
            caminho_saida = pasta_destino / safe_filename(f"{prefixo} {nome} ({contador}).pdf")
            contador += 1

        novo_doc.save(caminho_saida)
        novo_doc.close()

        print(f"  [V] Salvo: {nome_pasta_saida}/{caminho_saida.name} (Páginas {pag_inicio + 1} a {pag_fim + 1})", flush=True)

    if len(paginas_cobertas) != total_paginas:
        print(f"  [!] Alerta: {total_paginas - len(paginas_cobertas)} página(s) não foram atribuídas. Verifique a divisão.", flush=True)

    doc_original.close()
    return pasta_destino


# =========================================================
# FLUXO 1: PASTA COM CERTIDÕES INDIVIDUAIS (gemini-3.1-flash-lite)
# =========================================================

def processar_pasta_certidoes():
    """
    Fluxo 1: Processa pasta já contendo vários arquivos individuais usando gemini-3.1-flash-lite.
    """
    pdf_entries = []
    for pdf_path in sorted(ENTRADA_DIR.glob("**/*.pdf")):
        rel_parent = pdf_path.relative_to(ENTRADA_DIR).parent
        pdf_entries.append((pdf_path, rel_parent))

    if not pdf_entries:
        print(f"[-] Nenhum arquivo PDF encontrado em: {ENTRADA_DIR}", flush=True)
        return

    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    processados_com_sucesso = 0
    pastas_subdiretorios = set()

    print(f"[+] Fluxo 1 (Pasta com Vários Arquivos): Processando {len(pdf_entries)} certidão(ões) via {MODELO_FLUXO_1_PASTA_CERTIDOES}...\n", flush=True)

    for pdf_file, rel_parent in pdf_entries:
        try:
            print(f" -> Lendo via {MODELO_FLUXO_1_PASTA_CERTIDOES}: {pdf_file.relative_to(ENTRADA_DIR)}", flush=True)
            
            doc = fitz.open(pdf_file)
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
                doc_img = fitz.open(pdf_file)
                img_bytes = pdf_page_to_png_bytes(doc_img, page_num=0)
                doc_img.close()
                contents = [types.Part.from_bytes(data=img_bytes, mime_type="image/png"), prompt]

            response = client.models.generate_content(
                model=MODELO_FLUXO_1_PASTA_CERTIDOES,
                contents=contents,
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json"
                )
            )

            dados = json.loads(response.text.strip())
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

            time.sleep(4.1)

            if str(rel_parent) != ".":
                pastas_subdiretorios.add(ENTRADA_DIR / rel_parent)

        except Exception as err:
            print(f" [X] Erro ao processar arquivo {pdf_file.name}: {err}\n", flush=True)

    if processados_com_sucesso == len(pdf_entries):
        print("[+] Todos os arquivos da pasta foram renomeados e movidos com sucesso!", flush=True)
        for pasta in pastas_subdiretorios:
            if pasta.exists():
                print(f"[+] Removendo subpasta de entrada: {pasta}", flush=True)
                shutil.rmtree(pasta)
        print(f"[+] Limpeza da pasta de entrada concluida!", flush=True)


# =========================================================
# ORQUESTRADOR PRINCIPAL DO FLUXO
# =========================================================

def main():
    if not ENTRADA_DIR.exists():
        print(f"[-] Pasta de entrada não encontrada: {ENTRADA_DIR}", flush=True)
        return

    itens_entrada = [p for p in ENTRADA_DIR.iterdir() if not p.name.startswith(".")]
    arquivos_pdf = [p for p in itens_entrada if p.is_file() and p.suffix.lower() == ".pdf"]
    pastas = [p for p in itens_entrada if p.is_dir()]

    # FLUXO 2: ARQUIVO PDF ÚNICO EM /entrada -> Modelo: gemini-3.5-flash
    if arquivos_pdf:
        pdf_target = arquivos_pdf[0]
        print(f"[+] [FLUXO 2] Detectado arquivo único em /entrada: {pdf_target.name}", flush=True)
        try:
            pasta_criada = fatiar_pdf_unico(pdf_target)
            pdf_target.unlink()
            print(f"\n[+] PDF original removido de /entrada. Todas as certidões divididas e nomeadas em: /saida/{pasta_criada.name}", flush=True)
        except Exception as e:
            print(f"\n[X] Erro ao fatiar/renomear o PDF {pdf_target.name}: {e}", flush=True)

    # FLUXO 1: PASTA COM VÁRIOS ARQUIVOS EM /entrada -> Modelo: gemini-3.1-flash-lite
    elif pastas:
        print(f"[+] [FLUXO 1] Detectada estrutura de pasta em /entrada: {pastas[0].name}", flush=True)
        processar_pasta_certidoes()

    else:
        print(f"[-] Nenhum PDF ou subpasta encontrado em: {ENTRADA_DIR}", flush=True)


if __name__ == "__main__":
    main()
