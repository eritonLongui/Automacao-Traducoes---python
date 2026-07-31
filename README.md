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
├── saida/
├── modelos/
├── executar fluxo completo.bat
├── executar numerar_certidoes.bat
├── executar separar_arquivos.bat
├── executar traduzir_cnn.bat
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
   GROQ_API_KEY=sua_chave_aqui
   GEMINI_API_KEY=sua_chave_aqui
   ```
4. Adicione o arquivo de credenciais de APIs (ex: `numero-traducao-82a46e53b009.json`) na mesma pasta dos scripts para que o código funcione corretamente.

## Fluxos Disponíveis

### 1. Separar Arquivos (`executar separar_arquivos.bat`)
Separa automaticamente múltiplas certidões contidas em um único arquivo `.doc` do Microsoft Word, gerando um arquivo `.docx` para cada certidão encontrada.
* **Como utilizar:** Coloque o(s) arquivo(s) `.doc` na pasta `entrada` e execute o `.bat`. Os arquivos gerados serão salvos em subpastas dentro de `saida`.

### 2. Formatar Certidões (`formatar_certidoes.py`)
Aplica formatações automatizadas através das bibliotecas e modelos de formatação em certidões processadas, interagindo com o Microsoft Word via `pywin32`.

### 3. Numerar Certidões (`executar numerar_certidoes.bat`)
Fluxo acionado pelo `numerar_traducoes.py` que organiza e aplica uma numeração padrão para as certidões e traduções trabalhadas.

### 4. Traduzir CNN (`executar traduzir_cnn.bat`)
Executa o fluxo do `traduzir_cnn.py` para auxiliar com as traduções usando APIs externas.

### 5. Fluxo Completo (`executar fluxo completo.bat`)
Executa de forma automatizada e sequencial a separação de arquivos e a formatação das certidões.

## Observações

* O projeto utiliza automação do Microsoft Word através da biblioteca `pywin32`.
* O Microsoft Word deve estar devidamente instalado e licenciado na máquina.
* Apenas arquivos `.doc` ou cnn `.pdf` presentes na pasta `entrada` serão processados.
