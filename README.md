# Automação de Traduções

Scripts em Python para facilitar e automatizar fluxos de trabalho com certidões e traduções juramentadas/cartoriais (PT-BR ➔ IT), abrangendo envio de documentos ao Mauri via e-mail e planilha, separação de arquivos múltiplos, formatação inteligente com IA, atualização de datas e numeração oficial de registros.

## Requisitos

* Windows (10 ou 11)
* Microsoft Word instalado
* **Python 3.10+** (compatível com 3.14+)
* Arquivo de credenciais Service Account do Google Sheets (ex: `numero-traducao-82a46e53b009.json`) na raiz do diretório.
* Configurações de API e SMTP no arquivo `.env`.

## Estrutura do projeto

```text
.
├── entrada/                         # Entrada inicial de pastas e arquivos recebidos
├── divididos/                       # Certidões individuais (.docx) divididas por Família
├── saida/                           # Documentos finais formatados e numerados (.docx)
├── modelos/ [oculto]                # Modelos Word base
├── executar fluxo completo.bat      # Orquestrador de todo o pipeline
├── executar solicitar_mauri.bat     # Envio de pedidos de tradução por e-mail e registro
├── executar testar_solicitar_mauri.bat # Simulação do envio (sem enviar e-mail nem apagar)
├── executar separar_arquivos.bat    # Separação de arquivos .doc múltiplos
├── executar formatar_certidoes.bat  # Formatação com Gemini IA e Word COM
├── executar numerar_certidoes.bat   # Numeração e registro na aba DEISI
├── executar traduzir_cnn.bat        # Preenchimento de modelo CNN
├── executar pausar_revisao.bat      # Alterna pastas entre saida/ e pausado/
├── setup.bat                        # Instalação e flags de ocultação
├── solicitar_mauri.py [oculto]
├── separar_arquivos.py [oculto]
├── formatar_certidoes.py [oculto]
├── numerar_traducoes.py [oculto]
├── traduzir_cnn.py [oculto]
├── .env.example
├── requirements.txt
└── README.md
```

## Configuração Inicial

1. Após baixar ou clonar o projeto, dê dois cliques no arquivo:
   ```cmd
   setup.bat
   ```
   Esse script instalará as dependências do `requirements.txt`, criará as pastas necessárias e organizará a visibilidade dos arquivos executáveis `.bat`.
2. Crie ou configure o arquivo `.env` na raiz do projeto com as chaves e credenciais:
   ```env
   GEMINI_API_KEY=sua_chave_gemini_aqui

   # Configurações de Envio de E-mail (solicitar_mauri.py)
   SMTP_HOST=smtp.gmail.com
   SMTP_PORT=587
   SMTP_USER=seu_email@nostrali.com.br
   SMTP_PASSWORD=sua_senha_de_app_aqui
   ```
3. Mantenha o arquivo de credenciais do Google Sheets (ex: `numero-traducao-82a46e53b009.json`) na pasta do projeto.

---

## Fluxos Disponíveis

### 1. Solicitação de Traduções - Mauri (`executar solicitar_mauri.bat` / `executar testar_solicitar_mauri.bat`)
Identifica arquivos `.pdf` contidos dentro das pastas de famílias em `./entrada`:
* **Triagem Inteligente:** Processa e anexa exclusivamente os arquivos `.pdf`. Documentos `.docx` ou `.doc` (procurações, certidões) são preservados.
* **Perguntas Interativas:** Questiona se o envio é urgente (alterando o assunto do e-mail com prefixo `URGENTE`).
* **Envio de E-mail (SMTP):** Dispara o e-mail em formato HTML estilizado para `traduzione@nostrali.com.br` com cópia para os responsáveis (removendo automaticamente o próprio remetente do `Cc`).
* **Classificação e Agrupamento:** Analisa os prefixos (`CN`, `CC`, `CO`, `PR`, `CNN`, `DIV`, `ESC`) e sufixos (`ITALIANO`, `INGLÊS`, `ESPANHOL`), gerando o resumo no formato `"2 Divórcios e Escritura"` ou `"Certidão estrangeira"`.
* **Registro na Planilha:** Registra na aba **`MAURIZIO`** as colunas Cliente (C), Tipo de Documento (D), Data (E) e Quantidade (F).
* **Limpeza e Encaminhamento:** Após o envio bem-sucedido, exclui apenas os `.pdf` enviados e move toda a pasta da família com os documentos restantes para `./divididos` para seguir nos próximos passos.

### 2. Separar Arquivos (`executar separar_arquivos.bat`)
Varre a pasta `./entrada` em busca de arquivos unificados em `.doc` e quebra em certidões individuais `.docx`, salvando na respectiva subpasta da família dentro de `./divididos`.

### 3. Formatar Certidões (`executar formatar_certidoes.bat`)
Lê as pastas e certidões em `./divididos`, analisa as entidades do texto com **Gemini Flash** e aplica as regras tipográficas rígidas (negrito, caixa alta, termos cartoriais) no Microsoft Word, salvando as certidões prontas em `./saida`.

### 4. Numerar Certidões (`executar numerar_certidoes.bat`)
Varre `./saida`, consulta a aba **`DEISI`** na planilha Google para obter a numeração sequencial oficial, grava o registro e injeta o número oficial no rodapé de cada `.docx`.

### 5. Traduzir CNN (`executar traduzir_cnn.bat`)
Extrai dados de Certidões Negativas de Naturalização em PDF via `PyMuPDF` (`fitz`), preenche o modelo padrão `modelos/MODELO CNN DEISI.docx` e salva em `./saida`.

### 6. Pausar / Retomar Revisão (`executar pausar_revisao.bat`)
Alterna rapidamente as pastas de famílias entre `./saida` e `./pausado` para permitir trabalhar em novos lotes prioritários sem misturar revisões pendentes.

### 7. Fluxo Completo (`executar fluxo completo.bat`)
Orquestra o ciclo de ponta a ponta:
1. **Solicitar Mauri:** Verifica PDFs em `./entrada`, dispara e-mails, registra na aba `MAURIZIO`, apaga os PDFs e move as pastas de famílias com os arquivos restantes para `./divididos`.
2. **Separar Arquivos:** Caso haja arquivos `.doc` unificados em `./entrada`, realiza a separação para `./divididos`.
3. **Formatar Certidões:** Formata com IA e Word os arquivos de `./divididos` e exporta para `./saida`.
4. **Limpeza da Entrada:** Remove resíduos temporários da pasta de entrada.
5. **Atualizar Data:** Sincroniza a data de encerramento da tradução.
6. **Numerar Certidões:** Registra na planilha `DEISI` e aplica a numeração nos rodapés.
