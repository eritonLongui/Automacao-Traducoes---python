# Automação de Traduções

Scripts em Python para facilitar e automatizar fluxos de trabalho com certidões e traduções, abrangendo separação de arquivos, formatação, numeração das traduções e tradução de cnn.

## Requisitos

* Windows
* Microsoft Word instalado
* **Python 3.14.6**
* Arquivo de credenciais `numero-traducao-....json` (ex: `numero-traducao-82a46e53b009.json`) na raiz do diretório.
* Chaves de API configuradas no arquivo `.env`.

## Estrutura do projeto

```text
.
├── entrada/
├── divididos/
├── saida/
├── executar fluxo completo.bat
├── executar traduzir_cnn.bat
├── executar separar_arquivos.bat
├── executar formatar_certidoes.bat
├── executar numerar_certidoes.bat
├── setup.bat
├── formatacao_*.py
├── formatar_certidoes.py
├── numerar_traducoes.py
├── separar_arquivos.py
├── traduzir_cnn.py
├── .env.example
├── requirements.txt
└── README.md
```

## Configuração Inicial

1. Após baixar ou clonar o projeto, execute o script de setup:
   ```
   setup.bat
   ```
   Esse script irá instalar as dependências do projeto, e criar as pastas necessárias.
2. Certifique-se de ter instalado o **Python na versão 3.14.6**.
3. Crie ou renomeie o arquivo `.env.example` para `.env` na raiz do projeto e preencha com as suas chaves de API:
   ```
   GEMINI_API_KEY=sua_chave_aqui
   ```
4. Adicione o arquivo de credenciais de APIs (ex: `numero-traducao-82a46e53b009.json`) na mesma pasta dos scripts para que o código funcione corretamente.

## Fluxos Disponíveis

### 1. Separar Arquivos (`executar separar_arquivos.bat`)
Separa automaticamente múltiplas certidões contidas em um único arquivo `.doc` do Microsoft Word, gerando um arquivo `.docx` para cada certidão encontrada.
* **Como utilizar:** Coloque o(s) arquivo(s) `.doc` na pasta `entrada` e execute o `.bat`. Os arquivos gerados serão salvos em subpastas dentro de `divididos`.

### 2. Formatar Certidões (`formatar_certidoes.py`)
Aplica formatações automatizadas através de revisão com IA e regras de formatação em certidões processadas, interagindo com o Microsoft Word via `pywin32`.

### 3. Numerar Certidões (`executar numerar_certidoes.bat`)
Fluxo acionado pelo `numerar_traducoes.py` que organiza e aplica uma numeração padrão para as certidões e traduções trabalhadas presentes na pasta `saida`.

### 4. Traduzir CNN (`executar traduzir_cnn.bat`)
Executa o fluxo do `traduzir_cnn.py` para traduzir PDFs inclusos na pasta `entrada` e retornando com um arquivo `.docx` na pasta `saida`.

### 5. Fluxo Completo (`executar fluxo completo.bat`)
Executa de forma automatizada e sequencial a separação de arquivos, a formatação das certidões e a numeração de certidões.
Caso a CNN traduzida já esteja presente na pasta `saida`, o fluxo também irá alterar o numero de tradução desse arquivo.

## Observações

* Apenas arquivos `.doc` ou cnn `.pdf` presentes na pasta `entrada` serão processados.
* Para o fluxo funcionar corretamente, caso as certidões sejam inseridas diretamente na pasta `divididos`, elas devem seguir o padrão de nomenclatura: `pasta:` "nome da família" / `arquivos:` "CN + NOME DO REGISTRADO" (para nascimentos); e "CN + NOME DO REGISTRADO + e + NOME DO CONJUGE".
* O fluxo `numerar_traducoes.py` registra as certidões na planilha correspondente, e quanto executado, verifica se a planilha possui ocorrências dos mesmos `nome da família` e `nome do arquivo` na data atual, para verificar duplicidades, registros em outras datas não contam como duplicidade.