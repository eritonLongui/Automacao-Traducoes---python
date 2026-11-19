import os
import shutil
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SAIDA_DIR = BASE_DIR / "saida"
PAUSADO_DIR = BASE_DIR / "pausado"


def obter_pastas_filhas(diretorio: Path) -> list[Path]:
    """Retorna todas as subpastas (pastas de família) presentes no diretório."""
    if not diretorio.exists():
        return []
    return [p for p in diretorio.iterdir() if p.is_dir()]


def pausar_revisao():
    print("=" * 50)
    print("           FLUXO: PAUSAR / RETOMAR REVISÃO")
    print("=" * 50)

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
