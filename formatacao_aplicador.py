from typing import Dict, Any, List

from formatacao_config import ROLE_STYLES
from formatacao_leitor_word import sanitize_word_text


WD_FIND_STOP = 0
WD_COLLAPSE_END = 0

def capitalizar_primeira(texto: str) -> str:
    """Ex.: 'certifica' -> 'Certifica'"""
    texto = texto or ""
    return texto[:1].upper() + texto[1:].lower() if texto else texto


def capitalizar_cada_palavra(texto: str) -> str:
    """Ex.: 'rogerio da silva' -> 'Rogerio Da Silva'"""
    texto = (texto or "").strip()
    if not texto:
        return texto
    return " ".join(p[:1].upper() + p[1:].lower() if p else p for p in texto.split())


def _formatar_ocorrencias(
    doc,
    busca: str,
    texto_saida: str,
    *,
    bold=None,
    match_case=False,
    whole_word=False,
    ) -> int:
    """
    Localiza ocorrências de `busca` aceitando qualquer capitalização
    e força o texto de saída exatamente como informado.

    Retorna a quantidade de ocorrências alteradas.
    """
    if not busca:
        return 0

    total = 0
    rng = doc.Content.Duplicate
    find = rng.Find

    find.ClearFormatting()
    find.Text = busca
    find.MatchCase = match_case
    find.MatchWholeWord = whole_word
    find.Forward = True
    find.Wrap = WD_FIND_STOP

    while find.Execute():
        rng.Text = texto_saida

        if bold is not None:
            rng.Font.Bold = bool(bold)

        total += 1

        # Continua a busca depois do texto recém-substituído
        rng.Collapse(WD_COLLAPSE_END)
        rng.End = doc.Content.End

    return total


def _coletar_ocorrencias(doc, busca: str, match_case=False, whole_word=False):
    """
    Coleta posições (Start, End) de todas as ocorrências encontradas.
    Usado apenas para a regra especial de Anottazione/Anottazioni.
    """
    ocorrencias = []

    if not busca:
        return ocorrencias

    rng = doc.Content.Duplicate
    find = rng.Find

    find.ClearFormatting()
    find.Text = busca
    find.MatchCase = match_case
    find.MatchWholeWord = whole_word
    find.Forward = True
    find.Wrap = WD_FIND_STOP

    while find.Execute():
        ocorrencias.append((rng.Start, rng.End))
        rng.Start = rng.End
        rng.End = doc.Content.End

    return ocorrencias


def _aplicar_anottazione(doc):
    """
    - 'Anottazione' e 'Anottazioni' sempre em minúsculo e negrito
    - Se ocorrerem muito próximas (<= 30 caracteres), apenas a última fica em negrito
    """
    ocorrencias = []

    for termo in ("Anottazione", "Anottazioni"):
        ocorrencias.extend(
            _coletar_ocorrencias(doc, termo, match_case=False, whole_word=True)
        )

    if not ocorrencias:
        return

    ocorrencias.sort(key=lambda r: (r[0], r[1]))

    grupos = []
    grupo_atual = [ocorrencias[0]]

    for atual in ocorrencias[1:]:
        anterior = grupo_atual[-1]
        distancia = atual[0] - anterior[1]

        if distancia <= 30:
            grupo_atual.append(atual)
        else:
            grupos.append(grupo_atual)
            grupo_atual = [atual]

    grupos.append(grupo_atual)

    # Processa de trás para frente para não bagunçar as posições originais
    for grupo in reversed(grupos):
        for idx, (inicio, fim) in enumerate(reversed(grupo)):
            rng = doc.Range(inicio, fim)
            rng.Text = rng.Text.lower()
            rng.Font.Bold = (idx == 0)  # somente a última do grupo fica em negrito


def _normalizar_lista_nomes(nomes_registrado):
    if not nomes_registrado:
        return []   

    if isinstance(nomes_registrado, str):
        lista = [nomes_registrado]
    else:
        lista = nomes_registrado

    return [str(n).strip() for n in lista if n and str(n).strip()]

def _aplicar_nome_registrado(doc, nomes_registrado):
    """
    Recebe:
    - string com 1 nome
    - lista/tupla com 1 ou 2 nomes

    Converte o(s) nome(s) encontrado(s) para MAIÚSCULO e negrito.
    """
    for nome in _normalizar_lista_nomes(nomes_registrado):
        _formatar_ocorrencias(
            doc,
            nome,
            nome.upper(),
            bold=True,
            match_case=False,
            whole_word=False,
        )


def aplicar_formatacoes_gerais(doc, nomes_registrado=None) -> None:
    """
    Aplica formatações estáticas fixas no documento inteiro,
    independentes da IA.
    """
    try:
        # Casos específicos: ajuste de caixa dentro da frase
        _formatar_ocorrencias(
        doc,
        "dell'ufficio dello stato civile",
        "dell'Ufficio dello Stato Civile",
        bold=False,
        match_case=False,
        whole_word=False,
        )
        _formatar_ocorrencias(
        doc,
        "all'ufficio di stato civile",
        "all'Ufficio di Stato Civile",
        bold=False,
        match_case=False,
        whole_word=False,
        )
        _formatar_ocorrencias(
        doc,
        "annotazione del cpf",
        "Annotazione del CPF",
        bold=False,
        match_case=False,
        whole_word=False,
        )
        _formatar_ocorrencias(
        doc,
        "[A tergo riprende]",
        "[A tergo riprende]",
        bold=True,
        match_case=False,
        whole_word=False,
        )

        # Frases que deve ficar sempre minúscula
        for termo in (
            "ufficiale dello stato civile",
            "dell'atto di stato civile",
            "e ne do fede",
            "ufficiale",
            "nulla più",
            "nulla di più",
            "nulla più da certificare",
            "certificato a trascrizione integrale",
            "comunione dei beni",
            "aerogramma",
            "null'altro",
            "nato a",
            "nata a",
            "dello",
            "funzionario incaricato"
        ):
            _formatar_ocorrencias(
                doc,
                termo,
                termo,
                bold=False,
                match_case=False,
                whole_word=False,
            )

        # Palavras que devem ficar com inicial maiúscula e sem negrito
        for termo in (
            "certifica",
            "certifico",
            "ufficio",
            "nascita",
            "matrimonio",
            "osservazioni",
            "lo sposo",
            "la sposa",
            "comunicazione",
            "stato civile",
            "nihil",
            "da certificare"
            "atto di"
        ):
            _formatar_ocorrencias(
                doc,
                termo,
                capitalizar_primeira(termo),
                bold=False,
                match_case=False,
                whole_word=True,
            )

        # Palavras que devem ficar com inicial maiúscula e em negrito
        for termo in (
            "annotazioni",
            "annotazione"
        ):
            _formatar_ocorrencias(
                doc,
                termo,
                capitalizar_primeira(termo),
                bold=False,
                match_case=False,
                whole_word=True,
            )

        # Palavras que devem ficar minúsculas e em negrito
        for termo in ("rettifica",):
            _formatar_ocorrencias(
                doc,
                termo,
                termo,
                bold=True,
                match_case=False,
                whole_word=True,
            )

        _aplicar_anottazione(doc)
        _aplicar_nome_registrado(doc, nomes_registrado)

    except Exception as e:
        print(f"Erro ao aplicar formatações gerais: {e}")
        raise


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

    if not style: # regras fora do ROLE_STYLES
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

    if style.get("capitalize_each_word"):
        texto_atual = trecho.Text
        texto_novo = capitalizar_cada_palavra(texto_atual)
        if texto_novo != texto_atual:
            trecho.Text = texto_novo
            trecho = doc.Range(base_start + start, base_start + start + len(texto_novo))

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