import json
import os
import re
from pathlib import Path
from typing import List, Dict, Any, Tuple

from groq import Groq
from dotenv import load_dotenv

from formatacao_config import GROQ_MODEL, GROQ_TEMPERATURE, SYSTEM_PROMPT, ROLE_DEFINITIONS

load_dotenv()

ALLOWED_DOCUMENT_TYPES = {"nascimento", "casamento", "desconhecido"}


def formatar_paragrafos_para_prompt(paragrafos: List[Dict[str, Any]]) -> Tuple[str, List[Dict[str, Any]], Dict[int, int]]:
    """
    Constrói o JSON enviado para a IA usando IDs sequenciais 1..N,
    independentemente dos IDs reais do Word.

    Retorna:
    - string JSON para o prompt
    - lista de parágrafos normalizados para validação
    - mapa de id do prompt -> id real do Word
    """
    paragrafos_prompt = []
    mapa_prompt_para_real = {}

    for idx, p in enumerate(paragrafos, start=1):
        paragrafos_prompt.append(
            {
                "id": idx,
                "text": p["text"],
            }
        )
        mapa_prompt_para_real[idx] = int(p["id"])

    bloco = {
        "instructions": {
            "paragraph_id_is_sequential_1_based": True,
            "definitions": ROLE_DEFINITIONS,
        },
        "paragraphs": paragrafos_prompt,
    }

    return json.dumps(bloco, ensure_ascii=False, indent=2), paragrafos_prompt, mapa_prompt_para_real


def extrair_json_da_resposta(texto: str) -> Dict[str, Any]:
    texto = texto.strip()
    try:
        return json.loads(texto)
    except json.JSONDecodeError as exc:
        raise ValueError("A resposta da IA não é um JSON válido.") from exc


def _normalizar_com_mapa(texto: str) -> Tuple[str, List[int]]:
    """
    Normaliza o texto para busca:
    - casefold
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
        else:
            chars.append(ch.casefold())
            mapa.append(i)
            ultimo_foi_espaco = False

    inicio = 0
    fim = len(chars)

    while inicio < fim and chars[inicio] == " ":
        inicio += 1
    while fim > inicio and chars[fim - 1] == " ":
        fim -= 1

    return "".join(chars[inicio:fim]), mapa[inicio:fim]


def _encontrar_trecho_no_texto(texto_base: str, trecho: str) -> Tuple[int, int]:
    """
    Localiza um trecho dentro do texto, tolerando diferenças de espaço.
    Retorna (start, end) no texto original.
    """
    texto_norm, mapa_texto = _normalizar_com_mapa(texto_base)
    trecho_norm, _ = _normalizar_com_mapa(trecho)

    if not texto_norm or not trecho_norm:
        raise ValueError("Texto base ou trecho vazio na busca.")

    pos = texto_norm.find(trecho_norm)
    if pos < 0:
        raise ValueError(f"Trecho não encontrado: {trecho!r}")

    start_original = mapa_texto[pos]
    end_original = mapa_texto[pos + len(trecho_norm) - 1] + 1
    return start_original, end_original


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
        texto_item = " ".join(text.split()).casefold()

        if texto_item not in texto_paragrafo.casefold():
            raise ValueError(
                f"Segmento {idx}: o texto '{text}' não foi encontrado no parágrafo {seg_id}."
            )

    return resultado


def traduzir_segmentos_para_word(
    resultado: Dict[str, Any],
    paragrafos_originais: List[Dict[str, Any]],
    mapa_prompt_para_real: Dict[int, int],
) -> Dict[str, Any]:
    """
    Converte a resposta da IA para o formato esperado pelo aplicador:
    paragraph + start + end + role
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
        start, end = _encontrar_trecho_no_texto(texto_paragrafo, texto_para_busca)

        segmentos_word.append(
            {
                "paragraph": real_id,
                "start": start,
                "end": end,
                "role": seg["role"],
            }
        )

    return {
        "segments": segmentos_word
    }


def analisar_documento(
    paragrafos: List[Dict[str, Any]],
    debug_base: Path | None = None,
    nome_documento: str | None = None,
) -> Dict[str, Any]:
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise EnvironmentError("Defina a variável de ambiente GROQ_API_KEY.")

    client = Groq(api_key=api_key)

    user_content, paragrafos_prompt, mapa_prompt_para_real = formatar_paragrafos_para_prompt(paragrafos)

    # debug
    print(f"Parágrafos enviados: {len(paragrafos)}")
    print(f"Tamanho do prompt: {len(user_content)} caracteres")
    if debug_base and nome_documento:
        (debug_base / f"{nome_documento}_prompt.json").write_text(
            user_content,
            encoding="utf-8",
        )

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            temperature=GROQ_TEMPERATURE,
            messages=[
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
                        "Analise o JSON abaixo.\n\n"
                        "Retorne APENAS um objeto JSON válido, sem markdown, sem comentários e sem texto adicional.\n\n"
                        "O objeto deve possuir exatamente esta estrutura:\n"
                        "{\n"
                        '  "segments": [\n'
                        "    {\n"
                        '      "id": 1,\n'
                        '      "text": "Giuseppe Rossi",\n'
                        '      "role": "registered_name"\n'
                        "    }\n"
                        "  ]\n"
                        "}\n\n"
                        "Regras:\n"
                        "- cada item em 'segments' representa um trecho a ser formatado;\n"
                        "- 'id' é o id do parágrafo na lista enviada ao modelo;\n"
                        "- 'text' deve aparecer exatamente no parágrafo indicado;\n"
                        "- 'role' deve ser uma das roles permitidas.\n\n"
                        f"{user_content}"
                    ),
                },
            ],
            response_format={"type": "json_object"},
        )
    except Exception as e:
        print(f"Erro Groq: {e}")
        raise

    content = response.choices[0].message.content

    # debug
    if debug_base and nome_documento:
        (debug_base / f"{nome_documento}_resposta.txt").write_text(
            content or "",
            encoding="utf-8",
        )

    if not content:
        raise ValueError("A IA retornou resposta vazia.")

    resultado = extrair_json_da_resposta(content)
    resultado = validar_resposta(resultado, paragrafos_prompt)

    # Converte para o formato esperado pelo aplicador
    resultado_word = traduzir_segmentos_para_word(resultado, paragrafos, mapa_prompt_para_real)
    return resultado_word