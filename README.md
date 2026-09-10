# Ferramenta de Homologação do Interplan

Ferramenta para comparação automática das tabelas exportadas do Interplan entre a versão atual e a versão a homologar.

## Arquivos necessários

Copiar para a mesma pasta:

- Cabos_Atual.CSV
- Cabos_a_Homologar.CSV
- Curvas_Atual.CSV
- Curvas_a_Homologar.CSV

## Execução

Executar:

executar.bat

## Resultado

Será gerado o arquivo:

Relatorio_Homologacao.xlsx

## Verificações realizadas

- Validação de estrutura das tabelas
- Registros incluídos
- Registros removidos
- Registros alterados
- Resumo consolidado das diferenças



Rascunho do passo a passo para o usuário olhar o README e saber exatamente o que fazer.

Primeiro é muito importante destacar que é necessário ter o python previamente instalado na máquina.
Depois o usuário precisa fazer o download da pasta zip desse repositório (ou baixar cada arquivo de forma manual).

Tópico "Validação das curvas de carga e geração e da tabela de cabos"
É necessário colar os arquivos gerados "Curvas_Atual.CSV" e "Curvas_a_Homologar.CSV" na pasta aonde o código main.py está e rodar executar.bat .
O código irá gerar um arquivo "Relatorio_Homologacao.xlsx" na pasta em que o main.py está com as seguintes abas: Resumo (contém as principais informações sobre as tabelas analisadas); Validacao_Coluna que contém uma análise se as tabelas contém as mesmas colunas e na mesma ordem; Diferencas indica se alguma informação foi removida ou incluída; Alteracoes indica todas as alterações que aconteceram, indicando o que estava antes e o valor atual; Alteracoes por coluna mostra a quantidade de alterações por coluna. O arquivo traz também as tabelas analisadas. Vale ressaltar que a tabela de curvas é tratada então em Resultado ela já vem da forma que foi analisada (ou seja tratada, não do jeito que veio do interplan...).

