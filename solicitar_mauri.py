"""
Fluxo de solicitação de traduções para o Mauri.
- Lê as pastas de famílias em ./entrada
- Identifica apenas arquivos PDF (.pdf) dentro da pasta da família
  (arquivos .doc e .docx são ignorados e preservados)
- Pergunta se o envio é urgente
- Prepara o e-mail anexando apenas os PDFs encontrados
- Envia via SMTP (com credenciais do .env)
- Registra a solicitação na planilha do Google (aba MAURIZIO)
- Se enviado e registrado com sucesso: exclui apenas os PDFs enviados
  (e a pasta da família caso não sobre nenhum outro arquivo ou subpasta nela).
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import shutil
import smtplib
import unicodedata
from collections import Counter
from datetime import date
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import List

from dotenv import load_dotenv

# Reaproveitamento das funções do projeto
from numerar_traducoes import (
    GOOGLE_SHEETS_CREDENTIALS_JSON,
    GOOGLE_SHEETS_ID,
    GSpreadSheetClient,
    _split_family_token,
    extract_document_name,
)

# Carrega variáveis de ambiente
load_dotenv()

# =========================
# CONFIGURAÇÕES EDITÁVEIS
# =========================

PROJECT_DIR = Path(__file__).resolve().parent
ENTRADA_DIR = PROJECT_DIR / "entrada"
DIVIDIDOS_DIR = PROJECT_DIR / "divididos"

WORKSHEET_NAME = "MAURIZIO"

# Mapeamento de prefixos para tipo no singular e plural
DOC_TYPE_MAP = {
    "CNN": ("CNN", "CNNs"),
    "CN": ("Nascimento", "Nascimentos"),
    "CC": ("Casamento", "Casamentos"),
    "CO": ("Óbito", "Óbitos"),
    "DIV": ("Divórcio", "Divórcios"),
    "ESC": ("Escritura", "Escrituras"),
    "PR": ("Procuração", "Procurações"),
}

# Prefixo por extenso (caso venha escrito por extenso no nome do arquivo)
DOC_FULLNAME_MAP = {
    "NASCIMENTO": ("Nascimento", "Nascimentos"),
    "CASAMENTO": ("Casamento", "Casamentos"),
    "OBITO": ("Óbito", "Óbitos"),
    "DIVORCIO": ("Divórcio", "Divórcios"),
    "ESCRITURA": ("Escritura", "Escrituras"),
    "PROCURACAO": ("Procuração", "Procurações"),
}

# Sufixos de idioma
SUFFIX_MAP = {
    "ITALIANO": "ITALIANO",
    "INGLES": "INGLÊS",
    "ESPANHOL": "ESPANHOL",
}

# Destinatários e Cópia
EMAIL_TO = ["traduzione@nostrali.com.br"]
EMAIL_CC = ["daniel.pozzer@nostrali.com.br", "eriton.santos@nostrali.com.br"]

# Configurações SMTP (.env ou fallback)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)
LOGGER = logging.getLogger("solicitar_mauri")


# =========================
# FUNÇÕES DE APOIO
# =========================

def discover_mauri_tasks(base_dir: Path) -> list[tuple[Path, str, list[Path]]]:
    """
    Varre ./entrada e localiza pastas de família que possuem arquivos PDF (.pdf).
    Arquivos .doc e .docx são ignorados por este fluxo e preservados.
    Também suporta caso os PDFs estejam numa subpasta 'mauri' por compatibilidade.
    Retorna uma lista de tuplas:
    (family_dir, family_name, pdf_files)
    """
    tasks = []
    if not base_dir.exists():
        LOGGER.warning("Diretório de entrada não encontrado: %s", base_dir)
        return tasks

    for item in sorted(base_dir.iterdir()):
        if not item.is_dir() or item.name.startswith("."):
            continue

        family_name = _split_family_token(item.name)

        # Procura PDFs diretamente na pasta da família ou em subpasta 'mauri'
        pdf_files = []
        for f in item.rglob("*.pdf"):
            if f.is_file() and not f.name.startswith("~$") and not f.name.startswith("."):
                pdf_files.append(f)

        if pdf_files:
            tasks.append((item, family_name, sorted(pdf_files, key=lambda x: x.name.lower())))

    return tasks


def _strip_accents(text: str) -> str:
    """Remove acentos e caracteres especiais para facilitar comparações seguras."""
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(c for c in normalized if not unicodedata.combining(c))


def parse_mauri_document_name(doc_path: Path) -> tuple[str, str, str | None]:
    """
    Analisa a nomenclatura do arquivo na pasta /mauri.
    Retorna uma tupla:
      (tipo_singular, nome_pessoa, idioma_sufixo)

    Regras:
    1. Sufixos aceitos no final (com ou sem hífen/espaço): ITALIANO, INGLÊS, ESPANHOL.
    2. Prefixos aceitos no início: CNN, CN, CC, CO, DIV, ESC, PR (ou palavras inteiras: DIVORCIO, etc.).
    3. Se for inglês ou espanhol, o tipo exibido é "Certidão estrangeira (inglês ou espanhol)".
    4. Se não corresponder a nenhuma regra, mantém o nome completo do arquivo.
    """
    stem = doc_path.stem.strip()
    raw_clean = _strip_accents(stem).upper()

    # Detecta sufixo de idioma no final
    detected_suffix = None
    for suf_key, suf_label in SUFFIX_MAP.items():
        # Ex: "DIVORCIO ADRIANE - ITALIANO" ou "DIVORCIO ADRIANE ITALIANO"
        pattern = rf"(?:^|[\s\-_]+){suf_key}$"
        match = re.search(pattern, raw_clean)
        if match:
            detected_suffix = suf_label
            # Remove o sufixo do final da string original
            cut_idx = match.start()
            stem = stem[:cut_idx].rstrip(" -_")
            raw_clean = raw_clean[:cut_idx].rstrip(" -_")
            break

    # Se tiver sufixo INGLÊS ou ESPANHOL, a regra define o tipo como estrangeiro
    if detected_suffix in ("INGLÊS", "ESPANHOL"):
        lang_name = detected_suffix.lower()
        foreign_type = f"Certidão estrangeira ({lang_name})"
        # Tenta extrair a pessoa removendo o prefixo se existir
        matched_prefix = False
        # Ordena prefixos por tamanho desc (CNN antes de CN)
        for pref in sorted(list(DOC_TYPE_MAP.keys()) + list(DOC_FULLNAME_MAP.keys()), key=len, reverse=True):
            if raw_clean.startswith(pref) and (len(raw_clean) == len(pref) or raw_clean[len(pref)] in " -_"):
                person = stem[len(pref):].strip(" -_")
                return foreign_type, person, detected_suffix

        return foreign_type, stem, detected_suffix

    # Detecta prefixos conhecidos (CNN, CN, CC, CO, DIV, ESC, PR)
    for pref, (sing, _) in sorted(DOC_TYPE_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        if raw_clean.startswith(pref) and (len(raw_clean) == len(pref) or raw_clean[len(pref)] in " -_"):
            person = stem[len(pref):].strip(" -_")
            return sing, person, detected_suffix

    # Detecta por extenso (DIVORCIO, ESCRITURA, etc.)
    for full_pref, (sing, _) in sorted(DOC_FULLNAME_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        if raw_clean.startswith(full_pref) and (len(raw_clean) == len(full_pref) or raw_clean[len(full_pref)] in " -_"):
            person = stem[len(full_pref):].strip(" -_")
            return sing, person, detected_suffix

    # Regra de fallback: mantém o nome do arquivo
    return stem, "", detected_suffix


def format_grouped_types(type_counts: Counter[str]) -> str:
    """
    Formata as contagens de tipos para resumo e planilha.
    Exemplos:
      - Divórcio (1) -> "Divórcio"
      - Divórcio (2) -> "2 Divórcios"
      - Divórcio (1) e Escritura (1) -> "Divórcio e Escritura"
      - Divórcio (2) e Escritura (1) -> "2 Divórcios e Escritura"
      - 3 ou mais: "2 Divórcios, Casamento e Escritura"
    """
    # Mapeia plural se conhecido
    plural_map = {}
    for sing, plur in DOC_TYPE_MAP.values():
        plural_map[sing] = plur
    for sing, plur in DOC_FULLNAME_MAP.values():
        plural_map[sing] = plur

    parts = []
    for t_name, count in type_counts.items():
        if count == 1:
            parts.append(t_name)
        else:
            plural_name = plural_map.get(t_name, f"{t_name}s")
            parts.append(f"{count} {plural_name}")

    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} e {parts[1]}"
    return f"{', '.join(parts[:-1])} e {parts[-1]}"


def get_effective_cc_list(user_email: str, cc_list: list[str]) -> list[str]:
    """Remove o próprio remetente da lista de cópia (Cc) para evitar auto-cópia."""
    user_norm = user_email.strip().lower()
    return [email for email in cc_list if email.strip().lower() != user_norm]


def send_email_smtp(
    subject: str,
    body_text: str,
    body_html: str,
    attachments: list[Path],
    to_list: list[str],
    cc_list: list[str],
) -> None:
    """Envia o e-mail (HTML + texto puro alternativo) com múltiplos anexos utilizando SMTP."""
    if not SMTP_USER or not SMTP_PASSWORD:
        raise ValueError(
            "Configurações SMTP_USER e SMTP_PASSWORD não foram preenchidas no arquivo .env."
        )

    # Filtra Cc caso o remetente esteja na lista
    effective_cc = get_effective_cc_list(SMTP_USER, cc_list)

    msg = MIMEMultipart("mixed")
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(to_list)
    if effective_cc:
        msg["Cc"] = ", ".join(effective_cc)
    msg["Subject"] = subject

    # Container alternativo: clientes que suportam HTML veem a versão estilizada;
    # clientes de texto puro recebem a versão alternativa limpa.
    msg_alt = MIMEMultipart("alternative")
    msg_alt.attach(MIMEText(body_text, "plain", "utf-8"))
    msg_alt.attach(MIMEText(body_html, "html", "utf-8"))
    msg.attach(msg_alt)

    # Anexos
    for file_path in attachments:
        with open(file_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{file_path.name}"',
        )
        msg.attach(part)

    recipients = to_list + effective_cc

    LOGGER.info("Conectando ao servidor SMTP %s:%s...", SMTP_HOST, SMTP_PORT)
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg, from_addr=SMTP_USER, to_addrs=recipients)

    LOGGER.info("E-mail enviado com sucesso para %s (Cc: %s)", to_list, effective_cc)


def register_in_maurizio_sheet(
    client: GSpreadSheetClient,
    family_name: str,
    document_sequence_str: str,
    quantity: int,
    reg_date: date,
) -> int:
    """
    Grava os dados na aba MAURIZIO:
    Coluna C: CLIENTE (família)
    Coluna D: TIPO DE DOCUMENTO (nomes das certidões em sequência)
    Coluna E: DATA (dd/mm/aaaa)
    Coluna F: QUANTIDADE
    """
    worksheet = client._worksheet
    all_values = worksheet.get_all_values()

    # Encontra a primeira linha livre onde as colunas C até F estejam vazias
    target_row = len(all_values) + 1
    for idx, row in enumerate(all_values[2:], start=3):  # Cabeçalho está na linha 2
        col_c = row[2].strip() if len(row) > 2 else ""
        col_d = row[3].strip() if len(row) > 3 else ""
        col_e = row[4].strip() if len(row) > 4 else ""
        col_f = row[5].strip() if len(row) > 5 else ""

        if not any([col_c, col_d, col_e, col_f]):
            target_row = idx
            break

    client.ensure_row_exists(target_row)

    row_data = [
        family_name,
        document_sequence_str,
        reg_date.strftime("%d/%m/%Y"),
        str(quantity),
    ]

    # Atualiza C{target_row}:F{target_row}
    worksheet.update(
        range_name=f"C{target_row}:F{target_row}",
        values=[row_data],
        value_input_option="RAW",
    )
    LOGGER.info("Registrado na aba MAURIZIO (Linha %d): %s", target_row, row_data)
    return target_row


def clean_processed_family_folder(family_dir: Path, sent_files: list[Path]) -> None:
    """
    Exclui apenas os arquivos PDF que foram enviados por e-mail.
    Se a pasta da família não contiver outros arquivos ou subpastas além dos excluídos,
    exclui também a pasta da família. Caso contenha outros documentos (.doc, .docx, etc.),
    a pasta da família é preservada intacta.
    """
    for file_path in sent_files:
        try:
            if file_path.exists():
                file_path.unlink()
                LOGGER.info("Arquivo PDF removido: %s", file_path.name)
        except Exception as exc:
            LOGGER.warning("Não foi possível excluir o arquivo '%s': %s", file_path.name, exc)

    # Remove subpastas vazias (ex: se existia uma subpasta 'mauri' vazia)
    for sub in list(family_dir.iterdir()):
        if sub.is_dir():
            sub_items = [i for i in sub.iterdir() if not i.name.startswith(".")]
            if not sub_items:
                try:
                    shutil.rmtree(sub)
                    LOGGER.info("Subpasta vazia removida: %s", sub.name)
                except Exception:
                    pass

    # Verifica se restou algum arquivo ou pasta na família
    remaining_items = [
        item for item in family_dir.iterdir()
        if not item.name.startswith(".") and not item.name.startswith("~$")
    ]

    if not remaining_items:
        try:
            shutil.rmtree(family_dir)
            LOGGER.info("Pasta da família vazia removida: %s", family_dir.name)
        except Exception as exc:
            LOGGER.warning("Não foi possível remover a pasta da família '%s': %s", family_dir.name, exc)
    else:
        LOGGER.info(
            "Pasta da família '%s' mantida pois contém outros documentos (%d restante(s), ex: .docx/.doc).",
            family_dir.name,
            len(remaining_items),
        )


def mover_pastas_restantes_para_divididos(entrada_dir: Path, divididos_dir: Path) -> None:
    """
    Ao final do fluxo (após envio e remoção dos PDFs), move todas as pastas de famílias
    remanescentes em ./entrada diretamente para ./divididos, sem alterar absolutamente nada
    no conteúdo, similar à lógica de pausar_revisao.py.
    """
    if not entrada_dir.exists():
        return

    pastas_familias = [p for p in entrada_dir.iterdir() if p.is_dir() and not p.name.startswith(".")]
    if not pastas_familias:
        return

    divididos_dir.mkdir(parents=True, exist_ok=True)
    print("\n" + "=" * 50)
    print("Movendo pastas remanescentes de 'entrada' para 'divididos'...")
    print("=" * 50)

    for pasta in pastas_familias:
        destino = divididos_dir / pasta.name
        if destino.exists():
            LOGGER.warning("A pasta '%s' já existe em 'divididos'. Mesclando/atualizando arquivos...", pasta.name)
            # Move arquivo por arquivo se a pasta já existir em divididos
            for item in pasta.iterdir():
                dest_item = destino / item.name
                if dest_item.exists():
                    if dest_item.is_dir():
                        shutil.rmtree(dest_item)
                    else:
                        dest_item.unlink()
                shutil.move(str(item), str(dest_item))
            # Remove a pasta de origem se ficou vazia
            try:
                shutil.rmtree(pasta)
            except Exception:
                pass
        else:
            shutil.move(str(pasta), str(destino))

        LOGGER.info("-> Movida para divididos: %s", pasta.name)


# =========================
# EXECUÇÃO PRINCIPAL
# =========================

def main() -> None:
    parser = argparse.ArgumentParser(description="Solicitação de traduções - Mauri")
    parser.add_argument(
        "--dry-run", "-t",
        action="store_true",
        help="Modo simulação: valida pastas, monta e-mail e verifica planilha sem enviar e sem apagar nada.",
    )
    args = parser.parse_args()
    dry_run = args.dry_run

    print("=" * 60)
    if dry_run:
        print("   FLUXO DE SOLICITAÇÃO - MAURI [MODO SIMULAÇÃO / TESTE]")
        print("   (Nenhum e-mail será enviado, nenhum arquivo será apagado)")
    else:
        print("      FLUXO DE SOLICITAÇÃO DE TRADUÇÕES - MAURI")
    print("=" * 60)

    tasks = discover_mauri_tasks(ENTRADA_DIR)

    if not tasks:
        print("\nNenhum arquivo PDF para envio ao Mauri encontrado em ./entrada.")
        print("Os arquivos .doc e .docx são preservados e não ativam este envio.")
        # Se não há PDFs para enviar mas há pastas em entrada, move para divididos
        if not dry_run:
            mover_pastas_restantes_para_divididos(ENTRADA_DIR, DIVIDIDOS_DIR)
        return

    print(f"\nDetectada(s) {len(tasks)} família(s) com PDF(s) para o Mauri:\n")
    for idx, (_, family_name, files) in enumerate(tasks, start=1):
        print(f"  {idx}. Família: {family_name} ({len(files)} PDF(s))")
        for f in files:
            print(f"     - {f.name}")

    print("\n" + "-" * 60)

    # Pergunta sobre urgência
    resposta_urgente = input("O envio é URGENTE? Digite 's' ou 'sim' (ou apenas Enter para Não): ").strip().lower()
    is_urgent = resposta_urgente in ["s", "sim"]

    if is_urgent:
        print(">> Marcado como: [URGENTE]")
    else:
        print(">> Marcado como: [NORMAL]")

    # Inicializa cliente da planilha
    print("\nConectando à planilha...")
    try:
        sheet_client = GSpreadSheetClient(
            spreadsheet_id=GOOGLE_SHEETS_ID,
            worksheet_name=WORKSHEET_NAME,
            credentials_json=str(GOOGLE_SHEETS_CREDENTIALS_JSON),
        )
    except Exception as exc:
        LOGGER.error("Erro ao conectar à planilha Google: %s", exc)
        return

    hoje = date.today()
    sucessos = 0
    falhas = 0

    for family_dir, family_name, files in tasks:
        print("\n" + "=" * 50)
        LOGGER.info("Processando: %s", family_name)

        # Monta assunto
        if is_urgent:
            subject = f"URGENTE - Traduções Nostrali - {family_name}"
        else:
            subject = f"Traduções Nostrali - {family_name}"

        # Analisa cada arquivo da pasta /mauri
        parsed_docs = []
        type_counter: Counter[str] = Counter()
        has_italian = False

        for f in files:
            doc_type, person, suffix = parse_mauri_document_name(f)
            parsed_docs.append((f, doc_type, person, suffix))
            type_counter[doc_type] += 1
            if suffix == "ITALIANO":
                has_italian = True

        # String formatada dos tipos para a Planilha (Coluna D)
        tipo_resumo_planilha = format_grouped_types(type_counter)

        # Monta itens para o corpo do e-mail:
        # Padrão: <b>nome_da_familia</b> - tipos_de_documentos - nome_da_pessoa_do_documento
        # Se for fallback sem pessoa: <b>nome_da_familia</b> - nome_do_documento
        body_lines_text = []
        body_lines_html = []
        for f, doc_type, person, suffix in parsed_docs:
            if person:
                line_plain = f"- {family_name} - {doc_type} - {person}"
                line_html = f"<li><b>{family_name}</b> - {doc_type} - {person}</li>"
            else:
                line_plain = f"- {family_name} - {doc_type}"
                line_html = f"<li><b>{family_name}</b> - {doc_type}</li>"

            body_lines_text.append(line_plain)
            body_lines_html.append(line_html)

        items_text = "\n".join(body_lines_text)
        items_html = "\n".join(body_lines_html)

        # Introdução do e-mail com regra de sufixo Italiano
        if has_italian:
            intro_text = "Olá,\nSeguem traduções à serem feitas para o Português:\n"
            intro_html = "<p>Olá,</p><p>Seguem traduções à serem feitas <b>para o Português</b>:</p>"
        else:
            intro_text = "Olá,\nSeguem traduções à serem feitas:\n"
            intro_html = "<p>Olá,</p><p>Seguem traduções à serem feitas:</p>"

        # Versão texto puro (com quebras de linha e assinatura em texto)
        body_text = (
            f"{intro_text}"
            f"{items_text}\n\n\n"
            f"Atenciosamente,"
        )

        # Versão HTML (com quebras de linha e 'Atenciosamente,' em itálico)
        body_html = (
            f"{intro_html}"
            f"<ul>"
            f"{items_html}"
            f"</ul>"
            f"<br><br>"
            f"<p><i>Atenciosamente,</i></p>"
        )

        try:
            if dry_run:
                simulated_cc = get_effective_cc_list(SMTP_USER, EMAIL_CC)
                print("\n[SIMULAÇÃO] --- PRÉVIA DO E-MAIL ---")
                print(f"De:   {SMTP_USER or '(Definido no .env)'}")
                print(f"Para: {', '.join(EMAIL_TO)}")
                print(f"Cc:   {', '.join(simulated_cc) if simulated_cc else '(Nenhum - remetente removido)'}")
                print(f"Assunto: {subject}")
                print(f"Anexos ({len(files)}): {[f.name for f in files]}")
                print("\n--- CORPO DO E-MAIL (Texto Puro) ---")
                print(body_text)
                print("--- CORPO DO E-MAIL (Visualização HTML) ---")
                print(body_html)
                print("-------------------------------------------\n")

                # Simula a busca de linha na planilha
                all_values = sheet_client._worksheet.get_all_values()
                target_row = len(all_values) + 1
                for idx, row in enumerate(all_values[2:], start=3):
                    col_c = row[2].strip() if len(row) > 2 else ""
                    col_d = row[3].strip() if len(row) > 3 else ""
                    col_e = row[4].strip() if len(row) > 4 else ""
                    col_f = row[5].strip() if len(row) > 5 else ""
                    if not any([col_c, col_d, col_e, col_f]):
                        target_row = idx
                        break

                print(f"[SIMULAÇÃO] Planilha: Seria gravado na linha {target_row} da aba MAURIZIO:")
                print(f"            Coluna C (Cliente): {family_name}")
                print(f"            Coluna D (Tipo):    {tipo_resumo_planilha}")
                print(f"            Coluna E (Data):    {hoje.strftime('%d/%m/%Y')}")
                print(f"            Coluna F (Qtd):     {len(files)}")
                print("[SIMULAÇÃO] Nenhuma exclusão de arquivo/pasta foi realizada.")
                sucessos += 1
                continue

            # 1. Enviar e-mail
            LOGGER.info("Enviando e-mail...")
            send_email_smtp(
                subject=subject,
                body_text=body_text,
                body_html=body_html,
                attachments=files,
                to_list=EMAIL_TO,
                cc_list=EMAIL_CC,
            )

            # 2. Registrar na planilha
            LOGGER.info("Registrando na planilha...")
            register_in_maurizio_sheet(
                client=sheet_client,
                family_name=family_name,
                document_sequence_str=tipo_resumo_planilha,
                quantity=len(files),
                reg_date=hoje,
            )

            # 3. Limpar arquivos PDF processados (e pasta da família se vazia)
            LOGGER.info("Limpando arquivos PDF processados...")
            clean_processed_family_folder(family_dir, files)

            sucessos += 1
            LOGGER.info("Processamento da família '%s' concluído com sucesso!", family_name)

        except Exception as exc:
            falhas += 1
            LOGGER.error("Falha ao processar família '%s': %s", family_name, exc)
            LOGGER.warning("Os arquivos de '%s' NÃO foram apagados para garantir a segurança.", family_name)

    print("\n" + "=" * 60)
    print(f"Resumo da Execução: {sucessos} com sucesso, {falhas} falha(s).")
    print("=" * 60)

    # Move pastas remanescentes para divididos (apenas se não estiver em modo dry-run e não houver falhas críticas)
    if not dry_run:
        mover_pastas_restantes_para_divididos(ENTRADA_DIR, DIVIDIDOS_DIR)
    else:
        print("\n[SIMULAÇÃO] Se não estivesse em modo teste, as pastas remanescentes seriam movidas para 'divididos'.")


if __name__ == "__main__":
    main()
