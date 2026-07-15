import json
import os
import re
import unicodedata
from pathlib import Path
from typing import List, Dict, Any, Tuple

from groq import Groq
from dotenv import load_dotenv

from formatacao_config import GROQ_MODEL, GROQ_TEMPERATURE, SYSTEM_PROMPT, ROLE_DEFINITIONS
from formatacao_leitor_word import normalizar_texto_para_analise

load_dotenv()

ALLOWED_DOCUMENT_TYPES = {"nascimento", "casamento", "desconhecido"}


def formatar_paragrafos_para_prompt(
    paragrafos: List[Dict[str, Any]]
) -> Tuple[str, List[Dict[str, Any]], Dict[int, int]]:
    """
    Constrói o JSON enviado para a IA com IDs sequenciais 1..N.
    O texto já vai normalizado apenas para análise.
    """
    paragrafos_prompt = []
    mapa_prompt_para_real = {}

    for idx, p in enumerate(paragrafos, start=1):
        paragrafos_prompt.append(
            {
                "id": idx,
                "text": normalizar_texto_para_analise(p["text"]),
            }
        )
        mapa_prompt_para_real[idx] = int(p["id"])

    bloco = {
        "paragraphs": paragrafos_prompt
    }

    return json.dumps(bloco, ensure_ascii=False, indent=2), paragrafos_prompt, mapa_prompt_para_real

def extrair_json_da_resposta(texto: str) -> Dict[str, Any]:
    texto = (texto or "").strip()

    if texto.startswith("```"):
        texto = re.sub(r"^```(?:json)?\s*", "", texto, flags=re.IGNORECASE)
        texto = re.sub(r"\s*```$", "", texto)

    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        inicio = texto.find("{")
        fim = texto.rfind("}")

        if inicio != -1 and fim != -1 and fim > inicio:
            trecho = texto[inicio:fim + 1]
            return json.loads(trecho)

        raise ValueError("A resposta da IA não é um JSON válido.")

def _encontrar_todas_ocorrencias(texto_base: str, trecho: str) -> List[Tuple[int, int]]:
    """
    Encontra todas as ocorrências do trecho no texto normalizado.
    Retorna lista de (start, end) no texto original.
    """
    texto_norm, mapa_texto = _normalizar_com_mapa(texto_base)
    trecho_norm, _ = _normalizar_com_mapa(trecho)

    if not texto_norm or not trecho_norm:
        return []

    ocorrencias = []
    pos = texto_norm.find(trecho_norm)

    while pos != -1:
        start_original = mapa_texto[pos]
        end_original = mapa_texto[pos + len(trecho_norm) - 1] + 1
        ocorrencias.append((start_original, end_original))
        pos = texto_norm.find(trecho_norm, pos + 1)

    return ocorrencias

def deduplicar_segmentos(resultado: Dict[str, Any]) -> Dict[str, Any]:
    vistos = set()
    segmentos_limpos = []

    for seg in resultado.get("segments", []):
        texto_norm, _ = _normalizar_com_mapa(seg["text"])
        chave = (int(seg["id"]), texto_norm)

        if chave in vistos:
            continue

        vistos.add(chave)
        segmentos_limpos.append(seg)

    return {"segments": segmentos_limpos}

def _normalizar_com_mapa(texto: str) -> Tuple[str, List[int]]:
    """
    Normaliza o texto para busca:
    - casefold
    - remove acentos
    - espaços múltiplos viram um único espaço
    - preserva mapa de posição para voltar ao índice original
    """
    texto = str(texto)

    chars = []
    mapa = []
    ultimo_foi_espaco = False

    for i, ch in enumerate(texto):
        if ch.isspace():
            if not ultimo_foi_espaco:
                chars.append(" ")
                mapa.append(i)
                ultimo_foi_espaco = True
            continue

        decomposto = unicodedata.normalize("NFKD", ch)
        sem_acentos = "".join(c for c in decomposto if not unicodedata.combining(c))

        if not sem_acentos:
            continue

        for c in sem_acentos.casefold():
            if c.isspace():
                if not ultimo_foi_espaco:
                    chars.append(" ")
                    mapa.append(i)
                    ultimo_foi_espaco = True
            else:
                chars.append(c)
                mapa.append(i)
                ultimo_foi_espaco = False

    inicio = 0
    fim = len(chars)

    while inicio < fim and chars[inicio] == " ":
        inicio += 1
    while fim > inicio and chars[fim - 1] == " ":
        fim -= 1

    return "".join(chars[inicio:fim]), mapa[inicio:fim]

def dividir_paragrafos_em_blocos(
    paragrafos: List[Dict[str, Any]],
    max_paragrafos: int = 2,
    max_caracteres: int = 1600,
) -> List[List[Dict[str, Any]]]:
    blocos = []
    bloco_atual = []
    tamanho_atual = 0

    for p in paragrafos:
        texto = str(p.get("text", ""))
        tamanho_texto = len(texto)

        if bloco_atual and (
            len(bloco_atual) >= max_paragrafos
            or tamanho_atual + tamanho_texto > max_caracteres
        ):
            blocos.append(bloco_atual)
            bloco_atual = []
            tamanho_atual = 0

        bloco_atual.append(p)
        tamanho_atual += tamanho_texto

    if bloco_atual:
        blocos.append(bloco_atual)

    return blocos

def traduzir_segmentos_para_word(
    resultado: Dict[str, Any],
    paragrafos_originais: List[Dict[str, Any]],
    mapa_prompt_para_real: Dict[int, int],
) -> Dict[str, Any]:
    """
    Converte a resposta da IA para o formato esperado pelo aplicador:
    paragraph + start + end + role

    Se um trecho aparecer mais de uma vez no mesmo parágrafo,
    aplica a formatação em todas as ocorrências.
    """
    mapa_real_para_texto = {
        int(p["id"]): str(p["text"])
        for p in paragrafos_originais
        if isinstance(p, dict) and "id" in p and "text" in p
    }

    segmentos_word = []

    for seg in resultado.get("segments", []):
        prompt_id = int(seg["id"])
        real_id = mapa_prompt_para_real[prompt_id]
        texto_para_busca = str(seg["text"]).strip()

        if real_id not in mapa_real_para_texto:
            raise ValueError(f"Parágrafo real não encontrado: {real_id}")

        texto_paragrafo = mapa_real_para_texto[real_id]
        ocorrencias = _encontrar_todas_ocorrencias(texto_paragrafo, texto_para_busca)

        if not ocorrencias:
            raise ValueError(f"Trecho não encontrado: {texto_para_busca!r}")

        for start, end in ocorrencias:
            segmentos_word.append(
                {
                    "paragraph": real_id,
                    "start": start,
                    "end": end,
                    "role": seg["role"],
                }
            )

    return {"segments": segmentos_word}

def validar_resposta(resultado: Dict[str, Any], paragrafos_prompt: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not isinstance(resultado, dict):
        raise ValueError("Resposta da IA precisa ser um objeto JSON.")

    if "segments" not in resultado:
        raise ValueError("Resposta da IA está sem o campo obrigatório: segments")

    if not isinstance(resultado["segments"], list):
        raise ValueError("'segments' precisa ser uma lista.")

    allowed_roles = set(ROLE_DEFINITIONS.keys())
    mapa_paragrafos = {
        p["id"]: " ".join(str(p["text"]).split())
        for p in paragrafos_prompt
        if isinstance(p, dict) and "id" in p and "text" in p
    }

    for idx, seg in enumerate(resultado["segments"], start=1):
        if not isinstance(seg, dict):
            raise ValueError(f"Segmento {idx} precisa ser um objeto JSON.")

        for campo in ("id", "text", "role"):
            if campo not in seg:
                raise ValueError(f"Segmento {idx} sem o campo obrigatório: {campo}")

        seg_id = seg["id"]
        text = seg["text"]
        role = seg["role"]

        if not isinstance(seg_id, int):
            raise ValueError(f"Segmento {idx}: id precisa ser inteiro.")

        if seg_id not in mapa_paragrafos:
            raise ValueError(f"Segmento {idx}: id inexistente no prompt: {seg_id}")

        if not isinstance(text, str) or not text.strip():
            raise ValueError(f"Segmento {idx}: text precisa ser uma string não vazia.")

        if not isinstance(role, str):
            raise ValueError(f"Segmento {idx}: role precisa ser uma string.")

        if role not in allowed_roles:
            raise ValueError(f"Segmento {idx}: role inválida: {role}")

        texto_paragrafo = mapa_paragrafos[seg_id]
        texto_item_norm, _ = _normalizar_com_mapa(text)
        texto_paragrafo_norm, _ = _normalizar_com_mapa(texto_paragrafo)

        if texto_item_norm not in texto_paragrafo_norm:
            raise ValueError(
                f"Segmento {idx}: o texto '{text}' não foi encontrado no parágrafo {seg_id}."
            )

    return resultado

def _analisar_bloco(
    paragrafos: List[Dict[str, Any]],
    client: Groq,
    debug_base: Path | None = None,
    nome_documento: str | None = None,
    sufixo: str = "",
) -> Dict[str, Any]:
    user_content, paragrafos_prompt, mapa_prompt_para_real = formatar_paragrafos_para_prompt(paragrafos)

    if debug_base and nome_documento:
        nome_arquivo_prompt = f"{nome_documento}_prompt{sufixo}.json"
        (debug_base / nome_arquivo_prompt).write_text(
            user_content,
            encoding="utf-8",
        )

    mensagens = [
        {
            "role": "system",
            "content": SYSTEM_PROMPT.format(
                roles="\n".join(
                    f"- {k}: {v}" for k, v in ROLE_DEFINITIONS.items()
                )
            ),
        },
        {
            "role": "user",
            "content": (
                "Retorne apenas um JSON válido com a chave segments.\n"
                "Não use markdown, não use comentários, não adicione texto fora do JSON.\n"
                "Cada item deve conter id, text e role.\n\n"
                f"{user_content}"
            ),
        },
    ]

    ultimo_content = None

    for tentativa in range(1, 4):
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=GROQ_TEMPERATURE,
            messages=mensagens,
        )

        content = response.choices[0].message.content
        ultimo_content = content

        if debug_base and nome_documento:
            nome_arquivo_resposta = f"{nome_documento}_resposta{sufixo}_tentativa{tentativa}.txt"
            (debug_base / nome_arquivo_resposta).write_text(
                content or "",
                encoding="utf-8",
            )

        if not content:
            if tentativa == 1:
                print("A IA retornou resposta vazia. Tentando novamente...")
                continue
            if tentativa == 2:
                print("A IA retornou resposta vazia. Tentando novamente...")
                continue
            raise ValueError("A IA retornou resposta vazia.")

        try:
            resultado = extrair_json_da_resposta(content)
            resultado = validar_resposta(resultado, paragrafos_prompt)
            resultado = deduplicar_segmentos(resultado)
            return traduzir_segmentos_para_word(resultado, paragrafos, mapa_prompt_para_real)
        except Exception as e:
            if debug_base and nome_documento:
                nome_erro = f"{nome_documento}_erro_parse{sufixo}_tentativa{tentativa}.txt"
                (debug_base / nome_erro).write_text(
                    str(e),
                    encoding="utf-8",
                )

            if tentativa == 1:
                print("A resposta da IA veio inválida. Repetindo a análise...")
                continue
            if tentativa == 2:
                print("A resposta da IA veio inválida. Repetindo a análise...")
                continue

            raise ValueError(f"Falha ao interpretar JSON da IA após {tentativa} tentativas: {e}") from e

    raise ValueError("Falha inesperada na análise do bloco.")

def analisar_documento(
    paragrafos: List[Dict[str, Any]],
    debug_base: Path | None = None,
    nome_documento: str | None = None,
) -> Dict[str, Any]:
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise EnvironmentError("Defina a variável de ambiente GROQ_API_KEY.")

    client = Groq(api_key=api_key)

    blocos = dividir_paragrafos_em_blocos(paragrafos, max_paragrafos=3, max_caracteres=2400)

    print(f"Blocos enviados à IA: {len(blocos)}")

    segmentos_finais = []

    for i, bloco in enumerate(blocos, start=1):
        print(f"Analisando bloco {i}/{len(blocos)}...")
        resultado_bloco = _analisar_bloco(
            bloco,
            client=client,
            debug_base=debug_base,
            nome_documento=nome_documento,
            sufixo=f"_bloco{i}",
        )
        segmentos_finais.extend(resultado_bloco["segments"])

    return {"segments": segmentos_finais}