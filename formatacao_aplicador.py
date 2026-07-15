from typing import Dict, Any, List

def aplicar_formatacoes_gerais(doc) -> None:
    """
    Aplica formatações estáticas fixas no documento inteiro, independentes da IA.
    """
    try:
        # 1. "ufficio" -> minúsculo e sem negrito
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        
        find_obj.Text = "ufficio"
        find_obj.Replacement.Text = "ufficio"
        find_obj.Replacement.Font.Bold = False
        find_obj.Format = True
        find_obj.MatchCase = False
        find_obj.Execute(Replace=2, Wrap=1)

        # 2. "dell'Ufficio Dello Stato Civile" -> "dell'Ufficio dello Stato Civile"
        find_obj2 = doc.Content.Find
        find_obj2.ClearFormatting()
        find_obj2.Replacement.ClearFormatting()
        
        find_obj2.Text = "dell'ufficio dello stato civile"
        find_obj2.Replacement.Text = "dell'Ufficio dello Stato Civile"
        find_obj2.Format = False
        find_obj2.MatchCase = False
        find_obj2.Execute(Replace=2, Wrap=1)
    except Exception as e:
        print(f"Erro ao aplicar formatações gerais: {e}")

from formatacao_config import ROLE_STYLES
from formatacao_leitor_word import sanitize_word_text


def clamped(value: int, minimum: int, maximum: int) -> int:
    return max(minimum, min(value, maximum))


def resetar_formatacao_range(word_range) -> None:
    """
    Zera apenas atributos de fonte do trecho inteiro.
    Evita mexer em espaçamento/alinhamento de parágrafo,
    para preservar a estrutura do documento.
    """
    try:
        font = word_range.Font
        font.Bold = False
        font.Italic = False
        font.Underline = 0
        font.StrikeThrough = False
        font.DoubleStrikeThrough = False
        font.AllCaps = False
        font.SmallCaps = False
        font.Hidden = False
        font.Shadow = False
        font.Outline = False
        font.Emboss = False
        font.Engrave = False
        font.Subscript = False
        font.Superscript = False
        font.ColorIndex = 0
    except Exception:
        pass

    try:
        word_range.HighlightColorIndex = 0
    except Exception:
        pass


def colocar_minusculo_texto(texto: str) -> str:
    if texto is None:
        return ""
    return texto.lower()


def capitalizar_frases_texto(texto: str) -> str:
    """
    Capitaliza a primeira letra do texto e a primeira letra após
    pontuação final de frase.
    """
    resultado = []
    precisa_maiuscula = True

    for ch in texto:
        if ch.isalpha():
            if precisa_maiuscula:
                resultado.append(ch.upper())
                precisa_maiuscula = False
            else:
                resultado.append(ch.lower())
        else:
            resultado.append(ch)
            if ch in ".!?":
                precisa_maiuscula = True
            elif ch == "\n":
                precisa_maiuscula = True

    return "".join(resultado)


def colocar_minusculo_paragrafos(doc, paragraph_ids: List[int]) -> None:
    """
    Aplica minúsculas parágrafo por parágrafo.
    """
    for paragraph_id in paragraph_ids:
        try:
            para = doc.Paragraphs(paragraph_id)
            texto = sanitize_word_text(para.Range.Text)
            if not texto:
                continue

            para.Range.Text = colocar_minusculo_texto(texto) + "\r"
        except Exception:
            continue


def capitalizar_paragrafos_e_frases(doc, paragraph_ids: List[int]) -> None:
    """
    Aplica capitalização de início de parágrafo e após pontos finais.
    """
    for paragraph_id in paragraph_ids:
        try:
            para = doc.Paragraphs(paragraph_id)
            texto = sanitize_word_text(para.Range.Text)
            if not texto:
                continue

            texto = colocar_minusculo_texto(texto)
            texto = capitalizar_frases_texto(texto)

            para.Range.Text = texto + "\r"
        except Exception:
            continue


def normalizar_caixa_paragrafos(doc, paragraph_ids: List[int]) -> None:
    """
    Minúsculas + capitalização de início de frase/parágrafo.
    """
    colocar_minusculo_paragrafos(doc, paragraph_ids)
    capitalizar_paragrafos_e_frases(doc, paragraph_ids)


def uppercase_range(word_range) -> None:
    """
    Converte o conteúdo do range para maiúsculas preservando a estrutura do trecho.
    Fazemos char a char para minimizar risco de alterar formatação ao redor.
    """
    chars = word_range.Characters
    count = chars.Count

    for i in range(count, 0, -1):
        ch = chars(i)
        txt = ch.Text

        if txt in ("\r", "\x07"):
            continue

        upper = txt.upper()
        if upper != txt:
            ch.Text = upper


def aplicar_segmento(doc, segmento: Dict[str, Any]) -> None:
    role = segmento["role"]
    style = ROLE_STYLES.get(role)

    if not style:
        return

    paragraph_id = int(segmento["paragraph"])
    start = int(segmento["start"])
    end = int(segmento["end"])

    if end <= start:
        return

    para = doc.Paragraphs(paragraph_id)
    texto_paragrafo = sanitize_word_text(para.Range.Text)
    tamanho = len(texto_paragrafo)

    start = clamped(start, 0, tamanho)
    end = clamped(end, 0, tamanho)

    if end <= start:
        return

    base_start = para.Range.Start
    trecho = doc.Range(base_start + start, base_start + end)

    if style.get("uppercase"):
        uppercase_range(trecho)

    if style.get("bold"):
        trecho.Font.Bold = True


def aplicar_segmentos(doc, segmentos: List[Dict[str, Any]]) -> None:
    """
    Aplica os segmentos em ordem estável.
    """
    segmentos_ordenados = sorted(
        segmentos,
        key=lambda s: (int(s["paragraph"]), int(s["start"]), int(s["end"])),
    )

    for seg in segmentos_ordenados:
        try:
            aplicar_segmento(doc, seg)
        except Exception:
            continue