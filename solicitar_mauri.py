"""
Fluxo de solicitação de traduções para o Mauri.
- Lê as pastas em ./entrada
- Identifica subpastas '/mauri' (case-insensitive)
- Pergunta se o envio é urgente
- Prepara o e-mail com os anexos encontrados em '/mauri'
- Envia via SMTP (com credenciais do .env)
- Registra a solicitação na planilha do Google (aba MAURIZIO)
- Se enviado e registrado com sucesso: exclui a pasta '/mauri'
  (e a pasta da família caso não sobre nenhum outro arquivo/pasta nela).
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import smtplib
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

WORKSHEET_NAME = "MAURIZIO"

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

def discover_mauri_tasks(base_dir: Path) -> list[tuple[Path, Path, str, list[Path]]]:
    """
    Varre ./entrada e localiza pastas de família que possuem subpasta 'mauri'.
    Retorna uma lista de tuplas:
    (family_dir, mauri_dir, family_name, files_in_mauri)
    """
    tasks = []
    if not base_dir.exists():
        LOGGER.warning("Diretório de entrada não encontrado: %s", base_dir)
        return tasks

    for item in sorted(base_dir.iterdir()):
        if not item.is_dir() or item.name.startswith("."):
            continue

        family_name = _split_family_token(item.name)

        # Procura subpasta chamada 'mauri' (case-insensitive)
        mauri_subdirs = [
            sub for sub in item.iterdir()
            if sub.is_dir() and sub.name.strip().lower() == "mauri"
        ]

        if not mauri_subdirs:
            continue

        for mauri_dir in mauri_subdirs:
            # Lista arquivos válidos (ignora arquivos temporários do Word/Office ~$*)
            files = [
                f for f in mauri_dir.iterdir()
                if f.is_file() and not f.name.startswith("~$") and not f.name.startswith(".")
            ]

            if files:
                tasks.append((item, mauri_dir, family_name, sorted(files, key=lambda x: x.name.lower())))
            else:
                LOGGER.warning("Subpasta 'mauri' vazia em: %s", item.name)

    return tasks


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


def clean_processed_family_folder(family_dir: Path, mauri_dir: Path) -> None:
    """
    Exclui a pasta /mauri processada.
    Se a pasta da família não contiver outros arquivos ou subpastas além dela,
    exclui também a pasta da família.
    """
    if mauri_dir.exists():
        shutil.rmtree(mauri_dir)
        LOGGER.info("Subpasta removida: %s", mauri_dir.name)

    # Verifica se restou algum arquivo ou pasta na família
    remaining_items = [
        item for item in family_dir.iterdir()
        if not item.name.startswith(".") and not item.name.startswith("~$")
    ]

    if not remaining_items:
        shutil.rmtree(family_dir)
        LOGGER.info("Pasta da família vazia removida: %s", family_dir.name)
    else:
        LOGGER.info(
            "Pasta da família mantida pois contém outros documentos (%d restante(s)).",
            len(remaining_items),
        )


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
        print("\nNenhuma subpasta '/mauri' com arquivos encontrada em ./entrada.")
        print("Certifique-se de que a estrutura é ./entrada/<Nome da Familia>/mauri/<arquivos>")
        return

    print(f"\nDetectada(s) {len(tasks)} solicitação(ões) pendente(s):\n")
    for idx, (_, mauri_dir, family_name, files) in enumerate(tasks, start=1):
        print(f"  {idx}. Família: {family_name} ({len(files)} arquivo(s))")
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

    for family_dir, mauri_dir, family_name, files in tasks:
        print("\n" + "=" * 50)
        LOGGER.info("Processando: %s", family_name)

        # Monta assunto
        if is_urgent:
            subject = f"URGENTE - Traduções Nostrali - {family_name}"
        else:
            subject = f"Traduções Nostrali - {family_name}"

        # Extrai nomes das certidões
        doc_names = [extract_document_name(f) for f in files]
        doc_sequence_str = ", ".join(doc_names)

        # Monta itens para o corpo do e-mail:
        # Padrão: <b>nome_da_familia</b> - tipos_de_documentos - nome_da_pessoa_do_documento
        # (Por enquanto os nomes dos arquivos são usados como placeholder até definirmos a regra de extração)
        body_lines_text = []
        body_lines_html = []
        for name in doc_names:
            body_lines_text.append(f"- {family_name} - {name}")
            body_lines_html.append(f"<li><b>{family_name}</b> - {name}</li>")

        items_text = "\n".join(body_lines_text)
        items_html = "\n".join(body_lines_html)

        # Versão texto puro (com quebras de linha e assinatura em texto)
        body_text = (
            f"Olá,\n"
            f"Seguem traduções à serem feitas:\n"
            f"{items_text}\n\n\n"
            f"Atenciosamente,"
        )

        # Versão HTML (com quebras de linha e 'Atenciosamente,' em itálico)
        body_html = (
            f"<p>Olá,</p>"
            f"<p>Seguem traduções à serem feitas:</p>"
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
                print(f"            Coluna D (Tipo):    {doc_sequence_str}")
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
                document_sequence_str=doc_sequence_str,
                quantity=len(files),
                reg_date=hoje,
            )

            # 3. Limpar pasta /mauri (e pasta da família se vazia)
            LOGGER.info("Limpando arquivos processados...")
            clean_processed_family_folder(family_dir, mauri_dir)

            sucessos += 1
            LOGGER.info("Processamento da família '%s' concluído com sucesso!", family_name)

        except Exception as exc:
            falhas += 1
            LOGGER.error("Falha ao processar família '%s': %s", family_name, exc)
            LOGGER.warning("Os arquivos de '%s' NÃO foram apagados para garantir a segurança.", family_name)

    print("\n" + "=" * 60)
    print(f"Resumo da Execução: {sucessos} com sucesso, {falhas} falha(s).")
    print("=" * 60)


if __name__ == "__main__":
    main()
