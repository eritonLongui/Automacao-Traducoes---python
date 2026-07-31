import json
import os
import re
import time
import unicodedata
from pathlib import Path
from typing import List, Dict, Any, Tuple

from google import genai
from dotenv import load_dotenv

from formatacao_config import (
    GEMINI_FLASH_LITE_MODEL,
    GEMINI_FLASH_MODEL,
    GEMINI_3_FLASH_MODEL,
    GEMINI_TEMPERATURE,
    SYSTEM_PROMPT,
    ROLE_DEFINITIONS_POR_TIPO,
    ROLE_ALTERACAO_PRIMEIRA_OCORRENCIA,
)
from formatacao_leitor_word import normalizar_texto_para_analise

load_dotenv()


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
        
        if seg["role"] in ROLE_ALTERACAO_PRIMEIRA_OCORRENCIA:
            ocorrencias = ocorrencias[:1]

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


def validar_resposta(
    resultado: Dict[str, Any],
    paragrafos_prompt: List[Dict[str, Any]],
    role_definitions: dict[str, Any] | None = None,
) -> Dict[str, Any]:
    if not isinstance(resultado, dict):
        raise ValueError("Resposta da IA precisa ser um objeto JSON.")

    if "segments" not in resultado:
        raise ValueError("Resposta da IA está sem o campo obrigatório: segments")

    if not isinstance(resultado["segments"], list):
        raise ValueError("'segments' precisa ser uma lista.")

    allowed_roles = set(role_definitions.keys())
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


def interseccionar_resultados(
    resultado_1: Dict[str, Any],
    resultado_2: Dict[str, Any],
) -> Dict[str, Any]:
    """
    Mantém apenas os segmentos que apareceram nas DUAS análises.
    Comparação por id + role + text normalizado.
    """
    def chave(seg: Dict[str, Any]) -> Tuple[int, str, str]:
        texto_norm, _ = _normalizar_com_mapa(str(seg.get("text", "")))
        return (
            int(seg["id"]),
            str(seg["role"]),
            texto_norm,
        )

    segmentos_2 = {chave(seg): seg for seg in resultado_2.get("segments", [])}

    segmentos_finais = []
    vistos = set()

    for seg in resultado_1.get("segments", []):
        k = chave(seg)
        if k in segmentos_2 and k not in vistos:
            segmentos_finais.append(seg)
            vistos.add(k)

    return {"segments": segmentos_finais}


def _analisar_rodada(
    paragrafos: List[Dict[str, Any]],
    gemini_client: genai.Client,
    debug_base: Path | None = None,
    nome_documento: str | None = None,
    rodada: int = 1,
    role_definitions: dict[str, Any] | None = None,
) -> Dict[str, Any]:
    user_content, paragrafos_prompt, _ = formatar_paragrafos_para_prompt(paragrafos)

    if debug_base and nome_documento:
        nome_arquivo_prompt = f"{nome_documento}_prompt_rodada{rodada}.json"
        (debug_base / nome_arquivo_prompt).write_text(
            user_content,
            encoding="utf-8",
        )

    system_instruction = SYSTEM_PROMPT.format(
        roles="\n".join(f"- {k}: {v}" for k, v in role_definitions.items())
    )

    prompt_gemini = (
        "Retorne apenas um JSON válido com a chave segments.\n"
        "Não use markdown, não use comentários, não adicione texto fora do JSON.\n"
        "Cada item deve conter id, text e role.\n"
        f"{user_content}"
    )

    for tentativa in range(1, 4):
        content = ""

        try:
            response = gemini_client.models.generate_content(
                model=(
                    GEMINI_FLASH_LITE_MODEL
                    if tentativa == 1
                    else GEMINI_FLASH_MODEL
                    if tentativa == 2
                    else GEMINI_3_FLASH_MODEL
                ),
                contents=prompt_gemini,
                config=genai.types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    temperature=GEMINI_TEMPERATURE,
                )
            )
            content = response.text or ""
        except Exception:
            content = ""

        if debug_base and nome_documento:
            nome_arquivo_resposta = (
                f"{nome_documento}_resposta_rodada{rodada}_tentativa{tentativa}.txt"
            )
            (debug_base / nome_arquivo_resposta).write_text(
                content,
                encoding="utf-8",
            )

        if not content:
            if tentativa < 3:
                print(f"      [Rodada {rodada}, tentativa {tentativa} falhou, tentando novamente...]")
                time.sleep(1)
                continue
            raise ValueError(f"A IA retornou resposta vazia na rodada {rodada}.")

        try:
            resultado = extrair_json_da_resposta(content)
            resultado = validar_resposta(resultado, paragrafos_prompt, role_definitions)
            resultado = deduplicar_segmentos(resultado)
            return resultado
        except Exception as e:
            if debug_base and nome_documento:
                nome_erro = f"{nome_documento}_erro_parse_rodada{rodada}_tentativa{tentativa}.txt"
                (debug_base / nome_erro).write_text(
                    str(e),
                    encoding="utf-8",
                )

            if tentativa < 3:
                print(f"      [Rodada {rodada}, tentativa {tentativa} inválida, tentando novamente...")
                time.sleep(1)
                continue

            raise ValueError(
                f"Falha ao interpretar resposta da IA na rodada {rodada}: {e}"
            )

    raise ValueError("Falha inesperada na análise da rodada.")


def analisar_documento(
    paragrafos: List[Dict[str, Any]],
    debug_base: Path | None = None,
    nome_documento: str | None = None,
    tipo_certidao: str | None = None,
) -> Dict[str, Any]:
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    if not gemini_api_key:
        raise EnvironmentError("Defina a variável de ambiente GEMINI_API_KEY.")

    gemini_client = genai.Client(api_key=gemini_api_key)

    role_definitions = ROLE_DEFINITIONS_POR_TIPO[tipo_certidao]
    if role_definitions is None:
        raise ValueError(f"Tipo de certidão inválido: {tipo_certidao}")

    print("Analisando documento inteiro - rodada 1...")
    resultado_1 = _analisar_rodada(
        paragrafos,
        gemini_client=gemini_client,
        debug_base=debug_base,
        nome_documento=nome_documento,
        rodada=1,
        role_definitions=role_definitions,
    )

    print("Analisando documento inteiro - rodada 2...")
    resultado_2 = _analisar_rodada(
        paragrafos,
        gemini_client=gemini_client,
        debug_base=debug_base,
        nome_documento=nome_documento,
        rodada=2,
        role_definitions=role_definitions,
    )

    resultado_final = interseccionar_resultados(resultado_1, resultado_2)

    if not resultado_final["segments"]:
        raise ValueError("Nenhum segmento coincidiu nas duas análises.")

    _, _, mapa_prompt_para_real = formatar_paragrafos_para_prompt(paragrafos)
    return traduzir_segmentos_para_word(resultado_final, paragrafos, mapa_prompt_para_real)