"""
Atualiza a data final de encerramento para a data atual em todas as certidões .docx
presentes na pasta ./saida (incluindo subpastas de famílias).
"""

from pathlib import Path
import pythoncom

from formatacao_aplicador import replace_final_date
from formatacao_leitor_word import abrir_word, abrir_documento, iter_docx_files
from formatacao_config import SAIDA_DIR


def atualizar_datas_saida(saida_dir: Path = SAIDA_DIR) -> int:
    saida_dir = Path(saida_dir).resolve()
    if not saida_dir.exists():
        print(f"Diretório de saída não encontrado: {saida_dir}")
        return 0

    arquivos = list(iter_docx_files(saida_dir))
    if not arquivos:
        print(f"Nenhum arquivo .docx encontrado em: {saida_dir}")
        return 0

    print(f"Atualizando data da tradução para a data atual em {len(arquivos)} arquivo(s)...")

    pythoncom.CoInitialize()
    word = None
    atualizados = 0

    try:
        word = abrir_word()

        for arquivo in arquivos:
            doc = None
            try:
                doc = abrir_documento(word, arquivo, read_only=False)
                if replace_final_date(doc):
                    doc.Save()
                    atualizados += 1
                    print(f"  [DATA ATUALIZADA] {arquivo.name}")
                else:
                    print(f"  [SEM FECHO/IGNORADO] {arquivo.name}")
            except Exception as e:
                print(f"  [ERRO] Falha ao atualizar data em {arquivo.name}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    print(f"\nConcluído: {atualizados} documento(s) com data atualizada.")
    return atualizados


if __name__ == "__main__":
    atualizar_datas_saida()
