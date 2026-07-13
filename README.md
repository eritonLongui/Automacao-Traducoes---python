# Separador de Certidões

Script em Python para separar automaticamente múltiplas certidões contidas em um único arquivo `.doc` do Microsoft Word, gerando um arquivo `.docx` para cada certidão encontrada.

## Requisitos

* Windows
* Microsoft Word instalado
* Python 3.11 ou superior

## Estrutura do projeto

```text
.
├── entrada/
├── saida/
├── executar.bat
├── setup.bat
├── separar_arquivos.py
├── requirements.txt
└── README.md
```

## Primeira configuração

Após baixar ou clonar o projeto, execute:

```
setup.bat
```

Esse script irá:

* instalar as dependências do projeto;
* criar as pastas `entrada` e `saida`, caso não existam;
* criar e ocultar os arquivos necessários para manter essas pastas no Git.

Essa configuração precisa ser realizada apenas uma vez.

## Como utilizar

1. Coloque o(s) arquivo(s) `.doc` na pasta `entrada`.
2. Execute `executar.bat`.
3. Aguarde o processamento.
4. Os arquivos separados serão gerados na pasta `saida`.

Ao término da execução, o terminal exibirá a quantidade total de certidões processadas.

## Dependências

As dependências do projeto estão listadas em `requirements.txt`.

Caso seja necessário instalá-las manualmente:

```bash
pip install -r requirements.txt
```

## Observações

* O projeto utiliza automação do Microsoft Word através da biblioteca `pywin32`.
* O Microsoft Word deve estar instalado na máquina.
* Apenas arquivos `.doc` presentes na pasta `entrada` serão processados.
* Os arquivos gerados serão salvos em subpastas dentro de `saida`, organizados pelo nome do arquivo de origem.
