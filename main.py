import re
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

# ==========================================================================
# CONFIGURAÇÃO GERAL
# Cada item de TABELAS é uma comparação. O campo "tipo" define como o arquivo
# é lido:
#   - "csv_plano"  : CSV comum, com cabeçalho na 1ª linha (ex.: Cabos).
#   - "secionado"  : export dividido em seções (ex.: Curvas Típicas de Carga),
#                    onde cada seção vira uma tabela comparada separadamente.
# ==========================================================================
TABELAS = [
    {
        "nome": "Cabos",
        "tipo": "csv_plano",
        "arquivo_atual": "Cabos_Atual.CSV",
        "arquivo_homologar": "Cabos_a_Homologar.CSV",
        "chave": ["Cabo", "EXTERN ID"],
    },
    {
        "nome": "Curvas",
        "tipo": "secionado",
        "arquivo_atual": "Curvas_Atual.CSV",
        "arquivo_homologar": "Curvas_a_Homologar.CSV",
        # chave de comparação por seção
        "chaves": {
            "CATEGORIA": ["ID"],
            "CURVA_TIPICA": ["ID"],
            "CURVA_PONTOS": ["Curva_ID", "Hora"],
        },
        # dump das seções de pontos é grande; deixe False se não quiser as
        # abas com os dados brutos das duas versões dessa tabela.
        "exportar_dados": True,
    },
]

ARQUIVO_SAIDA = "Relatorio_Homologacao.xlsx"
CSV_SEP = ";"
CSV_ENCODING = "latin1"

# Nomes de coluna para cada seção do arquivo seccionado (ajuste conforme a
# semântica do seu sistema; o que importa para a comparação são as chaves).
COLUNAS_SECOES = {
    "CATEGORIA": ["ID", "Nome", "Campo3", "Faixa_Min", "Faixa_Max",
                  "Num_Pontos", "Campo7", "Campo8", "Participacao", "Campo10"],
    "CURVA_TIPICA": ["ID", "Ref_Categoria", "Valor", "Campo4", "Fp"],
    "CURVA_PONTOS": ["Curva_ID", "Hora", "P_pu", "Desvio_P", "Campo5", "Campo6"],
}


# ==========================================================================
# LEITURA
# ==========================================================================
def ler_arquivo(caminho):
    """Lê um CSV comum (com cabeçalho), tratando arquivo inexistente/vazio."""
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    try:
        return pd.read_csv(caminho, sep=CSV_SEP, encoding=CSV_ENCODING)
    except pd.errors.EmptyDataError:
        raise ValueError(f"Arquivo vazio: {caminho}")


def parse_arquivo_secionado(caminho):
    """
    Lê um export dividido em seções, onde cada seção começa por uma linha com
    apenas o nome dela (ex.: 'CATEGORIA;') e as linhas seguintes são os dados,
    sem cabeçalho de colunas. Retorna {nome_secao: DataFrame(str)}.

    Todos os valores são lidos como texto, preservando o valor exato (inclusive
    o separador decimal, que varia entre seções) para uma comparação fiel.
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    linhas_por_secao = {}
    secao_atual = None
    with open(caminho, encoding=CSV_ENCODING) as f:
        for linha in f:
            linha = linha.rstrip("\r\n")
            if not linha:
                continue
            campos = linha.split(CSV_SEP)
            if campos and campos[-1] == "":       # remove o ';' final
                campos = campos[:-1]
            # cabeçalho de seção: um único campo alfabético
            if len(campos) == 1 and re.match(r"^[A-Za-zÀ-ÿ_]+$", campos[0]):
                secao_atual = campos[0]
                linhas_por_secao[secao_atual] = []
            elif secao_atual is not None:
                linhas_por_secao[secao_atual].append(campos)

    if not linhas_por_secao:
        raise ValueError(f"Nenhuma seção reconhecida em: {caminho}")

    secoes = {}
    for secao, linhas in linhas_por_secao.items():
        df = pd.DataFrame(linhas, dtype=str)
        nomes = COLUNAS_SECOES.get(secao)
        ncols = df.shape[1]
        if nomes and ncols >= len(nomes):
            df = df.iloc[:, :len(nomes)]
            df.columns = nomes
        elif nomes:
            df.columns = nomes[:ncols]
        else:
            df.columns = [f"Col_{i + 1}" for i in range(ncols)]
        secoes[secao] = df
    return secoes


def montar_chave(df, colunas_chave):
    """Concatena as colunas-chave em uma única string 'valor1|valor2|...'."""
    return (
        df[colunas_chave]
        .astype(str)
        .apply(lambda col: col.str.strip())
        .agg("|".join, axis=1)
    )


def _nome_aba(nome, usados):
    """Sanitiza o nome da aba (máx. 31 chars) e garante unicidade."""
    nome = re.sub(r"[\[\]:*?/\\]", "_", nome)[:31]
    base = nome
    i = 2
    while nome in usados:
        sufixo = f"_{i}"
        nome = base[:31 - len(sufixo)] + sufixo
        i += 1
    usados.add(nome)
    return nome


# ==========================================================================
# VALIDAÇÕES E COMPARAÇÕES
# ==========================================================================
def validar_colunas(nome, df_atual, df_homologar):
    """Compara a estrutura (nomes e ordem) das colunas das duas versões."""
    cols_atual = list(df_atual.columns)
    cols_homologar = list(df_homologar.columns)

    apenas_atual = sorted(set(cols_atual) - set(cols_homologar))
    apenas_homologar = sorted(set(cols_homologar) - set(cols_atual))

    mesmas_colunas = not apenas_atual and not apenas_homologar
    mesma_ordem = cols_atual == cols_homologar

    fora_ordem = []
    if mesmas_colunas and not mesma_ordem:
        for i, (c1, c2) in enumerate(zip(cols_atual, cols_homologar), start=1):
            if c1 != c2:
                fora_ordem.append(f"Posição {i}: Atual='{c1}' | Homologar='{c2}'")

    return pd.DataFrame([{
        "Tabela": nome,
        "Mesmas_Colunas": "SIM" if mesmas_colunas else "NÃO",
        "Mesma_Ordem": "SIM" if mesma_ordem else "NÃO",
        "Colunas_Apenas_Atual": ", ".join(apenas_atual),
        "Colunas_Apenas_Homologar": ", ".join(apenas_homologar),
        "Colunas_Fora_De_Ordem": " | ".join(fora_ordem),
    }])


def verificar_novos_removidos(nome, df_atual, df_homologar):
    """Identifica registros novos e removidos com base na coluna CHAVE."""
    chaves_atual = set(df_atual["CHAVE"])
    chaves_homologar = set(df_homologar["CHAVE"])

    novos = df_homologar[~df_homologar["CHAVE"].isin(chaves_atual)].copy()
    novos.insert(0, "Status", "NOVO")

    removidos = df_atual[~df_atual["CHAVE"].isin(chaves_homologar)].copy()
    removidos.insert(0, "Status", "REMOVIDO")

    diferencas = pd.concat([novos, removidos], ignore_index=True)
    diferencas.insert(0, "Tabela", nome)
    diferencas = diferencas.drop(columns=["CHAVE"], errors="ignore")

    resumo = {
        "Tabela": nome,
        "Registros_Atual": len(df_atual),
        "Registros_Homologar": len(df_homologar),
        "Novos": len(novos),
        "Removidos": len(removidos),
    }
    return resumo, diferencas


def detectar_alteracoes(nome, df_atual, df_homologar, chave_cols):
    """
    Compara, registro a registro (chaves em comum), todas as colunas em comum
    e retorna as diferenças em formato longo. Vetorizado por coluna.
    """
    cols_saida = ["Tabela", "CHAVE"] + chave_cols + ["Coluna", "Valor_Atual", "Valor_Homologar"]

    at = df_atual.drop_duplicates("CHAVE").set_index("CHAVE")
    ho = df_homologar.drop_duplicates("CHAVE").set_index("CHAVE")

    chaves_comuns = at.index.intersection(ho.index)
    if len(chaves_comuns) == 0:
        return pd.DataFrame(columns=cols_saida)

    at = at.loc[chaves_comuns]
    ho = ho.loc[chaves_comuns]

    colunas_comparacao = [
        c for c in at.columns
        if c in ho.columns and c not in chave_cols
    ]

    registros = []
    for col in colunas_comparacao:
        a = at[col].astype(str)
        h = ho[col].astype(str)
        mask = a != h
        if mask.any():
            registros.append(pd.DataFrame({
                "Tabela": nome,
                "CHAVE": at.index[mask],
                "Coluna": col,
                "Valor_Atual": a[mask].values,
                "Valor_Homologar": h[mask].values,
            }))

    if not registros:
        return pd.DataFrame(columns=cols_saida)

    df_alt = pd.concat(registros, ignore_index=True)
    info = at[chave_cols].reset_index()
    df_alt = df_alt.merge(info, on="CHAVE", how="left")
    return df_alt[cols_saida]


def comparar_unidade(nome, df_atual, df_homologar, chave_cols):
    """Executa todo o fluxo de comparação para uma unidade (tabela/seção)."""
    faltando_a = [c for c in chave_cols if c not in df_atual.columns]
    faltando_h = [c for c in chave_cols if c not in df_homologar.columns]
    if faltando_a or faltando_h:
        raise KeyError(
            f"Colunas-chave ausentes -> Atual: {faltando_a} | Homologar: {faltando_h}"
        )

    validacao = validar_colunas(nome, df_atual, df_homologar)

    export_atual = df_atual.copy()
    export_homologar = df_homologar.copy()

    df_atual = df_atual.copy()
    df_homologar = df_homologar.copy()
    df_atual["CHAVE"] = montar_chave(df_atual, chave_cols)
    df_homologar["CHAVE"] = montar_chave(df_homologar, chave_cols)

    resumo, diferencas = verificar_novos_removidos(nome, df_atual, df_homologar)
    alteracoes = detectar_alteracoes(nome, df_atual, df_homologar, chave_cols)

    resumo["Alteracoes"] = len(alteracoes)
    resumo["Status"] = (
        "SEM ALTERAÇÃO"
        if resumo["Novos"] == 0 and resumo["Removidos"] == 0 and len(alteracoes) == 0
        else "ALTERAÇÕES IDENTIFICADAS"
    )

    return {
        "nome": nome,
        "validacao": validacao,
        "resumo": resumo,
        "diferencas": diferencas,
        "alteracoes": alteracoes,
        "export_atual": export_atual,
        "export_homologar": export_homologar,
    }


def expandir_config(config):
    """
    Carrega os arquivos de uma config e devolve a lista de unidades a comparar:
    (nome_completo, df_atual, df_homologar, chave). Uma tabela 'csv_plano' vira
    uma única unidade; uma 'secionado' vira uma unidade por seção.
    """
    tipo = config.get("tipo", "csv_plano")
    nome_base = config["nome"]

    if tipo == "csv_plano":
        df_a = ler_arquivo(config["arquivo_atual"])
        df_h = ler_arquivo(config["arquivo_homologar"])
        return [(nome_base, df_a, df_h, config["chave"])]

    if tipo == "secionado":
        secoes_a = parse_arquivo_secionado(config["arquivo_atual"])
        secoes_h = parse_arquivo_secionado(config["arquivo_homologar"])
        unidades = []
        for secao, chave in config["chaves"].items():
            if secao not in secoes_a or secao not in secoes_h:
                print(f"[AVISO] Seção '{secao}' ausente em um dos arquivos; ignorada.")
                continue
            unidades.append((f"{nome_base}_{secao}", secoes_a[secao], secoes_h[secao], chave))
        return unidades

    raise ValueError(f"Tipo de config desconhecido: {tipo}")


# ==========================================================================
# EXPORTAÇÃO PARA EXCEL
# ==========================================================================
def _ajustar_larguras(worksheet):
    """Ajusta a largura das colunas ao conteúdo (limitado a 60)."""
    for col_cells in worksheet.columns:
        letra = get_column_letter(col_cells[0].column)
        largura = max((len(str(c.value)) for c in col_cells if c.value is not None), default=0)
        worksheet.column_dimensions[letra].width = min(largura + 2, 60)


def escrever_excel(consolidado, tabelas, caminho_saida):
    """Gera o Excel com abas consolidadas + abas de dados por tabela/seção."""
    usados = set()
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        for chave, aba in [
            ("resumo", "Resumo"),
            ("validacao", "Validacao_Colunas"),
            ("diferencas", "Diferencas"),
            ("alteracoes", "Alteracoes"),
            ("alteracoes_por_coluna", "Alteracoes_por_Coluna"),
        ]:
            consolidado[chave].to_excel(writer, sheet_name=_nome_aba(aba, usados), index=False)

        for nome, dados in tabelas.items():
            dados["atual"].to_excel(writer, sheet_name=_nome_aba(f"{nome}_Atual", usados), index=False)
            dados["homologar"].to_excel(writer, sheet_name=_nome_aba(f"{nome}_Homolog", usados), index=False)

        for ws in writer.book.worksheets:
            _ajustar_larguras(ws)


# ==========================================================================
# ORQUESTRAÇÃO
# ==========================================================================
def main():
    resumos, validacoes, diferencas, alteracoes = [], [], [], []
    tabelas_export = {}

    for config in TABELAS:
        nome_base = config["nome"]
        print(f"\n{'=' * 60}\nProcessando: {nome_base}\n{'=' * 60}")
        try:
            unidades = expandir_config(config)
        except (FileNotFoundError, ValueError, KeyError) as e:
            print(f"[ERRO] '{nome_base}': {e}")
            continue

        exportar = config.get("exportar_dados", True)

        for nome, df_a, df_h, chave in unidades:
            try:
                res = comparar_unidade(nome, df_a, df_h, chave)
            except (KeyError, ValueError) as e:
                print(f"[ERRO] Unidade '{nome}': {e}")
                continue

            resumos.append(res["resumo"])
            validacoes.append(res["validacao"])
            diferencas.append(res["diferencas"])
            alteracoes.append(res["alteracoes"])
            if exportar:
                tabelas_export[nome] = {
                    "atual": res["export_atual"],
                    "homologar": res["export_homologar"],
                }

            r = res["resumo"]
            print(f"  {nome:<28} Novos:{r['Novos']:>4}  Removidos:{r['Removidos']:>4}  "
                  f"Alterações:{r['Alteracoes']:>6}  -> {r['Status']}")

    if not resumos:
        print("\nNenhuma tabela pôde ser processada. Excel não gerado.")
        return

    df_resumo = pd.DataFrame(resumos)
    df_validacao = pd.concat(validacoes, ignore_index=True)
    df_diferencas = pd.concat(diferencas, ignore_index=True)
    df_alteracoes = pd.concat(alteracoes, ignore_index=True)

    if not df_alteracoes.empty:
        df_alt_por_coluna = (
            df_alteracoes.groupby(["Tabela", "Coluna"]).size()
            .reset_index(name="Qtd_Alteracoes")
            .sort_values(["Tabela", "Qtd_Alteracoes"], ascending=[True, False])
        )
    else:
        df_alt_por_coluna = pd.DataFrame(columns=["Tabela", "Coluna", "Qtd_Alteracoes"])

    consolidado = {
        "resumo": df_resumo,
        "validacao": df_validacao,
        "diferencas": df_diferencas,
        "alteracoes": df_alteracoes,
        "alteracoes_por_coluna": df_alt_por_coluna,
    }

    escrever_excel(consolidado, tabelas_export, ARQUIVO_SAIDA)
    print(f"\n{'=' * 60}\nRelatório gerado: {ARQUIVO_SAIDA}\n{'=' * 60}")


if __name__ == "__main__":
    main()
