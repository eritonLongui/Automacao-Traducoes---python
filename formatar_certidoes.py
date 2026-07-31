import argparse
import json
import shutil
import re
import pythoncom
from pathlib import Path

from formatacao_config import ENTRADA_DIR, SAIDA_DIR, DEBUG_DIR
from formatacao_ia import analisar_documento
from formatacao_leitor_word import (
    abrir_word,
    abrir_documento,
    extrair_paragrafos_corpo,
    localizar_faixa_corpo,
    iter_docx_files,
)
from formatacao_aplicador import (
    resetar_formatacao_range,
    aplicar_segmentos,
    aplicar_formatacoes_gerais,
)


def identificar_tipo_pelo_nome(nome_arquivo: str) -> str:
    stem = Path(nome_arquivo).stem.upper().strip()

    if re.match(r"^CN\b", stem):
        return "nascimento"

    if re.match(r"^CC\b", stem):
        return "casamento"
    
    if re.match(r"^CO\b", stem):
        return "obito"

    return "desconhecido"

def extrair_nomes_arquivo(nome_arquivo: str):
    tipo = identificar_tipo_pelo_nome(nome_arquivo)
    stem = Path(nome_arquivo).stem.strip()

    if tipo == "nascimento":
        return [
            re.sub(r"^CN\s+", "", stem, flags=re.IGNORECASE).strip()
        ]

    if tipo == "casamento":
        nomes = re.sub(r"^CC\s+", "", stem, flags=re.IGNORECASE)

        return [
            nome.strip()
            for nome in re.split(r"\s+e\s+", nomes, flags=re.IGNORECASE)
        ]
    
    if tipo == "obito":
        return [re.sub(r"^CO\s+", "", stem, flags=re.IGNORECASE).strip()]

    return []

def criar_pastas(base_saida: Path, base_debug: Path) -> None:
    base_saida.mkdir(parents=True, exist_ok=True)
    base_debug.mkdir(parents=True, exist_ok=True)

def limpar_arquivos_processados(entrada_base: Path) -> None:
    """
    Remove todo o conteúdo de ENTRADA_DIR, preservando apenas .gitkeep.
    """
    if not entrada_base.exists():
        return

    for item in entrada_base.iterdir():
        if item.name == ".gitkeep":
            continue

        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()

def caminho_saida(
    arquivo_entrada: Path,
    entrada_base: Path,
    saida_base: Path,
) -> Path:
    rel = arquivo_entrada.relative_to(entrada_base)
    destino_dir = saida_base / rel.parent
    destino_dir.mkdir(parents=True, exist_ok=True)
    return destino_dir / f"{arquivo_entrada.stem}.docx"


def caminho_saida_sem_formatacao(
    arquivo_entrada: Path,
    entrada_base: Path,
    saida_base: Path,
) -> Path:
    rel = arquivo_entrada.relative_to(entrada_base)
    destino_dir = saida_base / rel.parent
    destino_dir.mkdir(parents=True, exist_ok=True)
    return destino_dir / f"{arquivo_entrada.stem}_sem_formatacao.docx"


def caminho_debug(
    arquivo_entrada: Path,
    entrada_base: Path,
    debug_base: Path,
) -> Path:
    rel = arquivo_entrada.relative_to(entrada_base)
    destino_dir = debug_base / rel.parent
    destino_dir.mkdir(parents=True, exist_ok=True)
    return destino_dir / f"{arquivo_entrada.stem}.json"


def salvar_saida_nao_formatada(
    arquivo_origem: Path,
    out_fallback: Path,
) -> None:
    if out_fallback.exists():
        out_fallback.unlink()
    shutil.copy2(arquivo_origem, out_fallback)


def processar_arquivo(
    word,
    arquivo: Path,
    entrada_base: Path,
    saida_base: Path,
    debug_base: Path,
) -> None:
    out_path = caminho_saida(arquivo, entrada_base, saida_base)
    out_fallback = caminho_saida_sem_formatacao(arquivo, entrada_base, saida_base)
    debug_path = caminho_debug(arquivo, entrada_base, debug_base)

    if out_path.exists():
        out_path.unlink()
    if out_fallback.exists():
        out_fallback.unlink()

    doc = None

    try:
        doc = abrir_documento(word, arquivo, read_only=False)
        tipo_certidao = identificar_tipo_pelo_nome(arquivo.name)

        faixa = localizar_faixa_corpo(doc)
        corpo_range = doc.Range(faixa["start_char"], faixa["end_char"])

        resetar_formatacao_range(corpo_range)

        paragrafos = extrair_paragrafos_corpo(doc)

        texto_extraido = "\n\n".join(
            f"[{p['id']}]\n{p['text']}"
            for p in paragrafos
        )
        (debug_base / f"{arquivo.stem}_extraido.txt").write_text(
            texto_extraido,
            encoding="utf-8",
        )

        print(f"\n========== {arquivo.name} ==========")
        
        analise = analisar_documento(
            paragrafos,
            debug_base=debug_base,
            nome_documento=arquivo.stem,
            tipo_certidao=tipo_certidao,
        )

        # debug
        debug_path.write_text(
            json.dumps(analise, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        nomes_registrado = extrair_nomes_arquivo(arquivo.name)

        aplicar_formatacoes_gerais(
            doc,
            nomes_registrado=nomes_registrado,
        )

        aplicar_segmentos(doc, analise["segments"]) 

        doc.SaveAs2(str(out_path), FileFormat=16)

        if out_fallback.exists():
            out_fallback.unlink()

        print(f"OK: {out_path.name}")

    except Exception:
        salvar_saida_nao_formatada(arquivo, out_fallback)
        print(f"FALHA: {arquivo.name} -> salvo sem formatação em {out_fallback.name}")
        raise

    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def resolver_pastas(entrada_raiz: Path, lote: str | None) -> tuple[Path, Path]:
    """
    entrada_raiz: ./divididos
    lote: nome da subpasta dentro de ./divididos, por exemplo "lote_001"
    """
    entrada_raiz = entrada_raiz.resolve()

    if lote:
        entrada_trabalho = (entrada_raiz / lote).resolve()
        try:
            entrada_trabalho.relative_to(entrada_raiz)
        except ValueError as exc:
            raise ValueError(
                f"O lote '{lote}' precisa estar dentro de: {entrada_raiz}"
            ) from exc

        if not entrada_trabalho.exists():
            raise FileNotFoundError(f"Pasta do lote não encontrada: {entrada_trabalho}")

        return entrada_raiz, entrada_trabalho

    if not entrada_raiz.exists():
        raise FileNotFoundError(f"Pasta de entrada não encontrada: {entrada_raiz}")

    return entrada_raiz, entrada_raiz


def nome_pasta_processando(entrada_trabalho: Path, arquivos: list[Path]) -> str:
    if not arquivos:
        return entrada_trabalho.name

    pastas_relativas = {
        arquivo.parent.relative_to(entrada_trabalho)
        for arquivo in arquivos
        if arquivo.parent != entrada_trabalho
    }

    if not pastas_relativas:
        return entrada_trabalho.name

    if len(pastas_relativas) == 1:
        rel = next(iter(pastas_relativas))
        return str(rel)

    nomes_topo = sorted({
        rel.parts[0] if rel.parts else entrada_trabalho.name
        for rel in pastas_relativas
    })

    return ", ".join(nomes_topo)


def main():
    parser = argparse.ArgumentParser(
        description="Formata certidões .docx com apoio de IA e salva em outra pasta."
    )

    parser.add_argument(
        "--entrada",
        type=str,
        default=str(ENTRADA_DIR),
        help="Pasta raiz com os lotes em .docx.",
    )

    parser.add_argument(
        "--lote",
        type=str,
        default=None,
        help="Nome da subpasta dentro da pasta de entrada.",
    )

    parser.add_argument(
        "--saida",
        type=str,
        default=str(SAIDA_DIR),
        help="Pasta de saída dos documentos formatados.",
    )

    parser.add_argument(
        "--debug",
        type=str,
        default=str(DEBUG_DIR),
        help="Pasta para salvar o JSON bruto da IA.",
    )

    parser.add_argument(
        "--arquivo",
        type=str,
        default=None,
        help="Processa apenas um arquivo específico.",
    )

    parser.add_argument(
        "--limpar-lote",
        action="store_true",
        help="Remove a pasta do lote ao final, se tudo terminar sem erro.",
    )

    args = parser.parse_args()

    entrada_raiz, entrada_trabalho = resolver_pastas(Path(args.entrada), args.lote)
    saida_base = Path(args.saida).resolve()
    debug_base = Path(args.debug).resolve()

    criar_pastas(saida_base, debug_base)

    pythoncom.CoInitialize()
    word = None
    erros = 0

    total_processadas = 0
    total_arquivos = 0

    try:
        word = abrir_word()

        if args.arquivo:
            arquivo = Path(args.arquivo).resolve()

            if not arquivo.exists():
                raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")

            try:
                arquivo.relative_to(entrada_trabalho)
            except ValueError as exc:
                raise ValueError(
                    f"O arquivo '{arquivo}' precisa estar dentro de: {entrada_trabalho}"
                ) from exc

            print(f"Processando: {arquivo.parent.name}")
            print()

            processar_arquivo(
                word,
                arquivo,
                entrada_raiz,
                saida_base,
                debug_base,
            )

            total_processadas = 1
            total_arquivos = 1

        else:
            arquivos = list(iter_docx_files(entrada_trabalho))
            total_arquivos = len(arquivos)

            nome_processamento = nome_pasta_processando(entrada_trabalho, arquivos)
            print(f"Processando: {nome_processamento}")
            print()

            if not arquivos:
                print(f"Nenhum .docx encontrado em: {entrada_trabalho}")
            else:
                for arquivo in arquivos:
                    try:
                        processar_arquivo(
                            word,
                            arquivo.resolve(),
                            entrada_raiz,
                            saida_base,
                            debug_base,
                        )

                        total_processadas += 1
                        print()
                    except Exception as e:
                        erros += 1
                        print(f"ERRO em {arquivo.name}: {e}")
                        print()

        if args.limpar_lote:
            if not args.lote:
                raise ValueError(
                    "--limpar-lote só pode ser usado quando você informar --lote."
                )

            if erros == 0:
                shutil.rmtree(entrada_trabalho)
                print(f"Pasta removida: {entrada_trabalho}")
            else:
                print(
                    "O lote não foi removido porque houve erro em pelo menos um arquivo."
                )

        print(f"Total de certidões encontradas: {total_arquivos}")
        print(f"Total de certidões processadas com sucesso: {total_processadas}")
        print(f"Total de erros: {erros}")

    finally:
        if word is not None:
            word.Quit()

        pythoncom.CoUninitialize()

        limpar_arquivos_processados(entrada_raiz)

    raise SystemExit(1 if erros else 0)


if __name__ == "__main__":
    main()