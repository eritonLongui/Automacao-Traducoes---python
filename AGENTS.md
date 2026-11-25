# 📘 DIRETRIZES DE DESENVOLVIMENTO, PADRÕES E ARQUITETURA (AGENT INSTRUCTIONS)

> **ATENÇÃO PARA AGENTES DE IA:**  
> Leia atentamente este documento antes de criar, alterar ou refatorar qualquer arquivo neste repositório. Ele resume o ecossistema, os padrões arquiteturais, regras de negócio para traduções juramentadas/cartoriais (PT-BR ➔ IT), convenções de código Python/PowerShell/Batch e o fluxo operacional de entrada e saída.

---

## 1. Visão Geral do Sistema e Propósito

Este ecossistema automatiza a esteira de processamento e tradução de certidões cartoriais e documentos públicos brasileiros para fins de cidadania italiana (Registro Civil Italiano).

### O Pipeline Central de Execução:
1. **Entrada (`./entrada`):** Recepção de arquivos `.doc` (unificados/múltiplas certidões agrupadas) ou CNNs em `.pdf`.
2. **Separação (`separar_arquivos.py` / `executar separar_arquivos.bat`):** 
   - Quebra documentos consolidados em certidões individuais `.docx`.
   - Salva em `./divididos/<Nome Da Familia>/`.
   - Padroniza a nomenclatura dos arquivos.
3. **Formatação Inteligente (`formatar_certidoes.py` / `executar formatar_certidoes.bat`):**
   - Extrai texto e analisa certidões via **Gemini 3.x Flash** (`google.genai`).
   - Identifica entidades (registrados, genitores, avós, cônjuges, datas, locais, títulos de retificação, averbações).
   - Comunica-se diretamente com o Microsoft Word via automação COM (`win32com.client`) para aplicar estilos tipográficos (negrito, caixa alta, correções de termos técnicos).
4. **Numeração e Registro em Planilha (`numerar_traducoes.py` / `executar numerar_certidoes.bat`):**
   - Varre `./saida/<Nome Da Familia>/*.docx`.
   - Consulta e registra as traduções no Google Sheets (planilha `DEISI` da Nostrali).
   - Injeta o número sequencial oficial da tradução no rodapé do documento Word.
   - Detecta duplicidades/revisões recentes: se já houver registro do documento e família nos últimos 12 dias, não cria nova linha na planilha, indica a data do registro anterior e atualiza o rodapé do documento com a numeração/ano daquele registro anterior.
5. **Tradução CNN (`traduzir_cnn.py` / `executar traduzir_cnn.bat`):**
   - Extrai dados de Certidões Negativas de Naturalização (CNN) em PDF via `PyMuPDF` (`fitz`).
   - Preenche o modelo padrão `modelos/MODELO CNN DEISI.docx` e salva em `./saida`.
6. **Pausar / Retomar Revisão (`pausar_revisao.py` / `executar pausar_revisao.bat`):**
   - Move pastas de famílias de `./saida` para `./pausado` (se houver pastas em saída).
   - Se `./saida` estiver vazia e houver pastas em `./pausado`, restaura-as para `./saida` e remove o diretório `./pausado`.
7. **Fluxo Completo (`executar fluxo completo.bat`):**
   - Orquestra os passos de separação, formatação, limpeza temporária de `entrada` e numeração final.

---

## 2. Estrutura de Diretórios e Ocultação

O projeto utiliza um padrão onde scripts utilitários executáveis para o usuário final são mantidos visíveis na raiz como arquivos `.bat`, enquanto arquivos de código, configurações e diretórios internos são ocultos via atributo de arquivo do Windows (`attrib +h`):

```text
.
├── entrada/                     # Arquivos brutos recebidos (.doc consolidados, .pdf)
├── divididos/ [oculto]          # Certidões separadas (.docx) em subpastas por Família (etapa intermediária)
├── saida/                       # Documentos finais formatados e numerados (.docx)
├── modelos/ [oculto]            # Modelos Word base (ex: MODELO CNN DEISI.docx)
├── executar fluxo completo.bat  # Runner principal para o usuário final
├── executar *.bat               # Runners de passos individuais
├── setup.bat                    # Script de setup de ambiente, pip e flags de visibilidade
├── formatacao_*.py [oculto]     # Módulos de formatação via IA e Word COM
├── formatar_certidoes.py [oculto]
├── numerar_traducoes.py [oculto]
├── separar_arquivos.py [oculto]
├── traduzir_cnn.py [oculto]
├── converter_para_pdf.py [oculto]
├── .env [oculto]                # Chaves de API (GEMINI_API_KEY, etc.)
└── numero-traducao-*.json       # Credenciais Service Account do Google Sheets
```

> **Aviso Importante sobre Arquivos Ocultos:**  
> Ao listar arquivos via PowerShell ou comandos de sistema, **sempre use `-Force` ou `attrib`**, caso contrário os scripts `.py`, `.env` e subpastas não aparecerão. Quando criar novos scripts `.py` ou de suporte, certifique-se de compatibilizá-los com o `setup.bat`.

---

## 3. Padrões de Nomenclatura e Regras de Negócio

### 3.1. Nomenclatura de Arquivos
- **Nascimento:** `CN NOME DO REGISTRADO.docx` (Ex: `CN MARIANA POLI DE SOUZA.docx`)
- **Casamento:** `CC NOME CONJUGE 1 e NOME CONJUGE 2.docx` (Ex: `CC PEDRO POLI e LOURDES CONCEICAO TARTALIA POLI.docx`)
- **Óbito:** `CO NOME DO FALECIDO.docx`
- **Certidão Negativa de Naturalização:** `CNN NOME DO REQUERENTE.docx` (Ignorada pelo script de formatação padrão de certidões comuns)
- **Procurações:** Prefixos `PR ...` (também ignoradas pela formatação de certidões)
- **Pastas de Família:** Nome completo do requerente/família em Title Case ou padrão claro (Ex: `Mariana Poli De Souza Campos`).

### 3.2. Regras de Formatação Tipográfica (`ROLE_STYLES`)
Definidas rigorosamente em `formatacao_config.py` e validadas em `formatacao_aplicador.py`:
- **Negrito e Caixa Alta (`BOLD + UPPERCASE`):**
  - Nome do Registrado (`registered_name`)
  - Nome do Cônjuge (`spouse_name`)
  - Nome de solteiro(a) (`maiden_name`)
  - Nome de casado(a) (`married_name`)
  - Novo nome em averbações/retificações (`rectification_new_name`)
- **Apenas Caixa Alta (`UPPERCASE`):**
  - Genitores (`father_name`, `mother_name`)
  - Avós maternos e paternos (`grandparent_name`)
  - Sogros (`father_in_law`, `mother_in_law`)
  - Declarante parente (`parent_declarant`)
  - Nome antigo em retificações (`rectification_old_name`)
- **Apenas Negrito (`BOLD`):**
  - Datas do fato: nascimento, casamento, óbito (`birth_date`, `marriage_date`, `death_date`)
  - Locais do fato: cidade/comuna (`birth_place`, `marriage_place`, `death_place`)
  - Data e local de registro de casamento (`marriage_registration_date`, `marriage_registration_place`)
  - Títulos de averbação/anotação (`annotation_title`, `rectification_title`)
  - Objeto de anotação quando isolada (`annotation_subject`, ex: *"Si è unito in matrimonio"*)
- **Capitalização Normal / Title Case:**
  - Testemunhas (`witness`)
  - Oficiais e escreventes (`officials`, `employees`)

---

## 4. Padrões de Código e Tecnologias

### 4.1. Ambiente e Dependências
- **Python:** Versão 3.10+ (compatível com 3.14+).
- **pywin32 (`win32com.client`):** Obrigatório para automação nativa do Microsoft Word (`Word.Application`). Todos os scripts manipulam o documento via objeto COM para preservar formatações complexas, cabeçalhos e rodapés que bibliotecas puras (como `python-docx`) não tratam com fidelidade.
- **google-genai:** SDK oficial moderno da Google para interações com modelos Gemini (uso de `client.models.generate_content`). Modelos preferenciais: `gemini-3.8-flash`, `gemini-3.5-flash-lite`.
- **gspread / google-auth:** Para persistência em nuvem na planilha de controle de traduções.
- **PyMuPDF (`fitz`):** Para extração precisa de texto vetorial em PDFs de entrada (CNNs).

### 4.2. Convenções para Novos Arquivos e Scripts
Ao introduzir um novo script no fluxo:
1. **Runner `.bat` associado:** Crie sempre um arquivo `.bat` correspondente na raiz (ex: `executar meu_script.bat`) contendo:
   ```bat
   @echo off
   title Nome do Fluxo
   cd /d "%~dp0"
   python meu_script.py
   pause
   ```
2. **Tratamento de Encerramento do Word:** Em scripts que utilizam `win32com`, garanta blocos `try...finally` garantindo `doc.Close(SaveChanges=...)` e `word.Quit()` para evitar processos órfãos (`WINWORD.EXE`) travando arquivos.
3. **Respeito ao Flag de Ocultação:** No final de modificações ou adições de arquivos, execute `attrib +h <novo_arquivo.py>` se ele não for um `.bat` executável pelo usuário.
4. **Idempotência e Não-Destrutividade:** Nunca exclua arquivos em `saida/` ou pastas de clientes sem confirmação prévia; limpezas automáticas devem se restringir a arquivos temporários de `entrada/` após processamento bem-sucedido.
5. **Configurações Centralizadas:** Novos seletores, termos de busca e listas de exclusão devem ser adicionados preferencialmente em `formatacao_config.py` ou em arquivos de configuração dedicados, nunca hardcoded dispersos.
