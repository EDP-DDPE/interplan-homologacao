# 🔍 Ferramenta de Homologação do Interplan

> Compara automaticamente as tabelas exportadas do **Interplan** entre a **versão atual** e a **versão a homologar** e gera um relatório consolidado em Excel com todas as diferenças encontradas.

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Plataforma-Windows-0078D6?logo=windows&logoColor=white)
![Excel](https://img.shields.io/badge/Saída-Excel%20(.xlsx)-217346?logo=microsoftexcel&logoColor=white)

---

## 📑 Sumário

1. [Pré-requisitos](#-1-pré-requisitos)
2. [Download da ferramenta](#-2-download-da-ferramenta)
3. [Preparação dos arquivos de entrada](#-3-preparação-dos-arquivos-de-entrada)
4. [Execução](#️-4-execução)
5. [Relatório gerado](#-5-relatório-gerado)
6. [Verificações realizadas](#-6-verificações-realizadas)
7. [Solução de problemas](#-7-solução-de-problemas)

---

## ✅ 1. Pré-requisitos

> [!IMPORTANT]
> É **obrigatório** ter o **Python 3.9 ou superior** instalado na máquina antes de executar a ferramenta.

1. Baixe o Python em 👉 [python.org/downloads](https://www.python.org/downloads/).
2. Durante a instalação, **marque a opção `Add Python to PATH`**.
3. Para confirmar a instalação, abra o **Prompt de Comando** e digite:

   ```bat
   python --version
   ```

   Deve aparecer algo como `Python 3.12.x`.

> 💡 Não é necessário instalar bibliotecas manualmente — o `executar.bat` cria um ambiente virtual e instala tudo sozinho na primeira execução (requer acesso à internet).

---

## 📥 2. Download da ferramenta

**Opção A — ZIP (recomendado)**

1. Nesta página do repositório, clique em **`<> Code` ➜ `Download ZIP`**.
2. Extraia o conteúdo do ZIP em uma pasta de sua preferência (ex.: `C:\Homologacao_Interplan`).

**Opção B — Download manual**

Baixe individualmente os arquivos abaixo e salve todos **na mesma pasta**:

| 📄 Arquivo | Função |
|---|---|
| `main.py` | Script principal de comparação |
| `executar.bat` | Inicializador (cria o ambiente e executa o script) |
| `requirements.txt` | Lista de dependências Python |

Estrutura esperada da pasta:

```text
📁 Homologacao_Interplan/
 ├── 📄 main.py
 ├── 📄 executar.bat
 └── 📄 requirements.txt
```

---

## 📂 3. Preparação dos arquivos de entrada

Exporte as tabelas do Interplan nas duas versões (**atual** e **a homologar**) e **cole os arquivos CSV na mesma pasta do `main.py`**.

> [!WARNING]
> Os arquivos devem ser salvos **exatamente com os nomes abaixo**. Arquivos com nomes diferentes não serão reconhecidos.

### 📊 Tabelas suportadas

| Tabela | 📄 Versão atual | 📄 Versão a homologar | 🔑 Chave de comparação |
|---|---|---|---|
| **Cabos** | `Cabos_Atual.CSV` | `Cabos_a_Homologar.CSV` | Cabo + EXTERN ID |
| **Curvas de carga e geração** | `Curvas_Atual.CSV` | `Curvas_a_Homologar.CSV` | Nome + Classificacao_Dia |
| **Chaves** | `Chaves_Atual.CSV` | `Chaves_a_Homologar.CSV` | Chave |
| **Relatório ANEEL** | `Relatorio_ANEEL_Atual.CSV` | `Relatorio_ANEEL_a_Homologar.CSV` | Matrícula |
| **Relatório ANEEL – Fluxo** | `RelatorioANEEL_Fluxo_Atual.CSV` | `RelatorioANEEL_Fluxo_a_Homologar.CSV` | Matrícula |
| **Demanda sem correção** | `DemandaSemCorrecao_Atual.CSV` | `DemandaSemCorrecao_a_Homologar.CSV` | Matrícula |
| **Demanda com correção** | `DemandaComCorrecao_Atual.CSV` | `DemandaComCorrecao_a_Homologar.CSV` | Matrícula |
| **Curto-circuito nas barras** | `CurtoBarras_Atual.CSV` | `CurtoBarras_a_Homologar.CSV` | Alimentador + Código de chave |
| **Relatório Diagnóstico de Circuitos** | `RelDiagCirc_Atual.csv` | `RelDiagCirc_a_Homologar.csv` | Circuito / Código Socorro |
| **Relatório Pós-Simulação** | `RelatorioPosSimulacao_Atual.csv` | `RelatorioPosSimulacao_a_Homologar.csv` | Circ1 + Circ2 |

> 💡 **Não é preciso ter todas as tabelas.** A ferramenta processa apenas os pares de arquivos presentes na pasta; os ausentes são ignorados (aparece um aviso `[ERRO] ... Arquivo não encontrado` na tela, o que é esperado).
>
> ⚠️ Cada tabela precisa dos **dois** arquivos (Atual **e** a Homologar) para ser comparada.

### Exemplo: validação de Curvas e Cabos

```text
📁 Homologacao_Interplan/
 ├── 📄 main.py
 ├── 📄 executar.bat
 ├── 📄 requirements.txt
 ├── 📊 Cabos_Atual.CSV
 ├── 📊 Cabos_a_Homologar.CSV
 ├── 📊 Curvas_Atual.CSV
 └── 📊 Curvas_a_Homologar.CSV
```

### 📌 Formato esperado dos CSV

- Separador: **ponto e vírgula (`;`)**
- Codificação: **ANSI / Latin-1** (padrão da exportação do Interplan)
- Os arquivos devem ser usados **como exportados** — não abra e salve pelo Excel antes, pois isso pode alterar o formato.

---

## ▶️ 4. Execução

1. Feche o arquivo `Relatorio_Homologacao.xlsx`, caso esteja aberto de uma execução anterior.
2. Dê **duplo clique** em **`executar.bat`**.
3. Aguarde o processamento. Na tela será exibido o andamento de cada tabela:

   ```text
   ============================================================
   Processando: Cabos
   ============================================================
     Cabos        Novos:   2  Removidos:   0  Alterações:    15  -> ALTERAÇÕES IDENTIFICADAS
   ...
   ============================================================
   Relatório gerado: Relatorio_Homologacao.xlsx
   ============================================================
   ```

4. Pressione qualquer tecla para fechar a janela.

> ⏱️ A **primeira execução** é mais demorada, pois cria o ambiente virtual (`.venv`) e instala as dependências. As próximas são bem mais rápidas.

---

## 📊 5. Relatório gerado

O arquivo **`Relatorio_Homologacao.xlsx`** é criado na **mesma pasta do `main.py`**, com as seguintes abas:

| 🗂️ Aba | Conteúdo |
|---|---|
| **Resumo** | Visão geral de cada tabela: nº de registros (atual × homologar), novos, removidos, alterações e status final (`SEM ALTERAÇÃO` ou `ALTERAÇÕES IDENTIFICADAS`). |
| **Validacao_Colunas** | Verifica se as duas versões têm **as mesmas colunas e na mesma ordem**, listando colunas exclusivas de cada versão e posições divergentes. |
| **Diferencas** | Registros **incluídos** (`NOVO`) ou **removidos** (`REMOVIDO`) na versão a homologar. |
| **Alteracoes** | Todas as alterações de valor, indicando registro, coluna, **valor anterior**, **valor novo** e **erro percentual** (para campos numéricos). |
| **Alteracoes_por_Coluna** | Quantidade de alterações por coluna de cada tabela — útil para identificar onde as mudanças se concentram. |
| **`<Tabela>_Atual`** / **`<Tabela>_Homolog`** | Cópia das tabelas analisadas, na versão atual e na versão a homologar. |

> [!NOTE]
> **Tabela de Curvas:** o arquivo exportado pelo Interplan é dividido em seções (`CATEGORIA`, `CURVA_TIPICA` e `CURVA_PONTOS`), sem cabeçalho. A ferramenta **trata e une essas seções** em uma única tabela antes da comparação. Por isso, as abas `Curvas_Atual` e `Curvas_Homolog` mostram os dados **já tratados** (como foram analisados), e não no formato bruto do Interplan.
>
> Classificação do dia: `1` = Dia Útil · `2` = Sábado · `3` = Domingo · `4` = Agregada.

---

## 🧪 6. Verificações realizadas

- ✅ **Validação de estrutura** — nomes e ordem das colunas
- ➕ **Registros incluídos** — presentes apenas na versão a homologar
- ➖ **Registros removidos** — presentes apenas na versão atual
- ✏️ **Registros alterados** — comparação campo a campo, com tolerância numérica (diferenças de formatação como `1,50` × `1.5` não são consideradas alteração)
- 📈 **Erro percentual** — calculado como `(Homologar − Atual) / Atual × 100`
- 📋 **Resumo consolidado** das diferenças por tabela

---

## 🛠️ 7. Solução de problemas

| ❗ Problema | 💡 Solução |
|---|---|
| `'py' não é reconhecido...` / `'python' não é reconhecido...` | O Python não está instalado ou não foi adicionado ao PATH. Reinstale marcando **`Add Python to PATH`**. |
| `ERRO: requirements.txt nao encontrado.` | O `requirements.txt` não está na mesma pasta do `executar.bat`. |
| `ERRO ao instalar dependencias.` | Verifique a conexão com a internet (ou proxy corporativo) e execute novamente. |
| `[ERRO] 'Tabela': Arquivo não encontrado` | Normal se você não pretende comparar essa tabela. Caso pretenda, confira se o nome do arquivo está **exatamente** como na [tabela da seção 3](#-tabelas-suportadas). |
| `Colunas-chave ausentes` | A exportação não contém as colunas usadas como chave. Verifique se o arquivo correto foi exportado. |
| `PermissionError` ao gerar o relatório | O `Relatorio_Homologacao.xlsx` está aberto no Excel. Feche-o e execute novamente. |
| `Nenhuma tabela pôde ser processada` | Nenhum par de arquivos válido foi encontrado na pasta. Revise os nomes e o local dos CSV. |

---

<p align="center">📌 <i>Ferramenta desenvolvida para apoio à homologação de versões do Interplan.</i></p>

