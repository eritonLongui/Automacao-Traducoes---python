import os
import shutil
import sys
from pathlib import Path

try:
    import win32com.client
except ImportError:
    win32com = None

BASE_DIR = Path(__file__).resolve().parent
SAIDA_DIR = BASE_DIR / "saida"
PAUSADO_DIR = BASE_DIR / "pausado"


def listar_documentos_abertos_word() -> list[str]:
    """
    Verifica se há alguma instância do Word rodando e se existem documentos (.doc/.docx) abertos.
    Retorna uma lista com os nomes dos documentos abertos.
    """
    if win32com is None:
        return []

    try:
        word = win32com.client.GetActiveObject("Word.Application")
        documentos = []
        for i in range(1, word.Documents.Count + 1):
            try:
                doc = word.Documents.Item(i)
                nome = doc.Name
                ext = Path(nome).suffix.lower()
                if ext in {".doc", ".docx"} or not ext:
                    documentos.append(nome)
            except Exception:
                continue
        return documentos
    except Exception:
        # GetActiveObject falha com com_error se o Word não estiver em execução
        return []


def verificar_e_aguardar_word() -> bool:
    """
    Verifica se há documentos abertos no Word.
    Se houver, aguarda o usuário fechar e dar ENTER para continuar, ou digitar 'c' / 'cancelar' para abortar.
    Retorna True se puder prosseguir, False se o usuário cancelou.
    """
    while True:
        docs_abertos = listar_documentos_abertos_word()
        if not docs_abertos:
            return True

        print("\n" + "!" * 60)
        print("[ATENÇÃO] Existem documentos abertos no Word:")
        for doc_nome in docs_abertos:
            print(f"  - {doc_nome}")
        print("!" * 60)
        print("Por favor, feche os arquivos no Word antes de prosseguir.")
        print("Pressione ENTER após fechar para tentar novamente,")
        print("ou digite 'c' / 'cancelar' para interromper o fluxo.")

        resposta = input("\nOpção [ENTER para continuar / 'c' para cancelar]: ").strip().lower()
        if resposta in {"c", "cancelar"}:
            print("\nOperação cancelada pelo usuário.")
            return False


def obter_pastas_filhas(diretorio: Path) -> list[Path]:
    """Retorna todas as subpastas (pastas de família) presentes no diretório."""
    if not diretorio.exists():
        return []
    return [p for p in diretorio.iterdir() if p.is_dir()]


def pausar_revisao():
    print("=" * 50)
    print("           FLUXO: PAUSAR / RETOMAR REVISÃO")
    print("=" * 50)

    # Verificação de arquivos abertos no Word antes de mover qualquer pasta
    if not verificar_e_aguardar_word():
        return

    pastas_saida = obter_pastas_filhas(SAIDA_DIR)
    pastas_pausado = obter_pastas_filhas(PAUSADO_DIR)

    # CENÁRIO 1: Há pastas em saída -> Mover para pausado
    if pastas_saida:
        PAUSADO_DIR.mkdir(parents=True, exist_ok=True)
        print(f"\n[PAUSAR] {len(pastas_saida)} pasta(s) encontrada(s) em 'saida'.")
        print(f"Movendo para '{PAUSADO_DIR.name}'...\n")

        for pasta in pastas_saida:
            destino = PAUSADO_DIR / pasta.name
            if destino.exists():
                print(f"  [AVISO] A pasta '{pasta.name}' já existia em 'pausado'. Atualizando...")
                shutil.rmtree(destino)
            shutil.move(str(pasta), str(destino))
            print(f"  -> Movido: {pasta.name}")

        print("\nRevisão pausada com sucesso!")
        print("As pastas foram movidas para 'pausado'.")
        return

    # CENÁRIO 2: Saída está vazia de pastas -> Restaurar de pausado para saída
    if pastas_pausado:
        SAIDA_DIR.mkdir(parents=True, exist_ok=True)
        print(f"\n[RETOMAR] Pasta 'saida' está livre e há {len(pastas_pausado)} pasta(s) em 'pausado'.")
        print(f"Movendo de volta para '{SAIDA_DIR.name}'...\n")

        for pasta in pastas_pausado:
            destino = SAIDA_DIR / pasta.name
            if destino.exists():
                print(f"  [AVISO] A pasta '{pasta.name}' já existia em 'saida'. Atualizando...")
                shutil.rmtree(destino)
            shutil.move(str(pasta), str(destino))
            print(f"  <- Restaurado: {pasta.name}")

        # Remove a pasta /pausado após mover tudo de volta
        try:
            shutil.rmtree(PAUSADO_DIR)
            print(f"\nPasta '{PAUSADO_DIR.name}' excluída com sucesso.")
        except Exception as e:
            print(f"\n[AVISO] Não foi possível remover a pasta '{PAUSADO_DIR.name}': {e}")

        print("\nRevisão retomada com sucesso!")
        return

    # CENÁRIO 3: Nem saída nem pausado contêm pastas
    print("\n[INFO] Nenhuma pasta de família encontrada nem em 'saida' nem em 'pausado'.")
    print("Nada para mover.")


if __name__ == "__main__":
    pausar_revisao()
