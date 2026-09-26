"""
Total de Peças por Rota — versão atualizada

Compatível com os arquivos atuais:
  - no_chao_remessas_tp_todas.xlsx (aba Remessas)
  - programacao_entregas_tppf_w38.xlsx (aba Planilha1)
  - tabela_lojas_rotas.xlsx (tabela persistente Loja -> Rota)

A rota oficial usada no resumo vem da tabela de suporte. A programação é usada
para cadastrar automaticamente lojas que ainda não existem nela. A coluna Rota
que agora também existe em Remessas é preservada apenas como referência e não
entra em conflito no merge.
"""

from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
ARQ_REMESSAS = BASE_DIR / "no_chao_remessas_tp_todas.xlsx"
ARQ_PROGRAMACAO = BASE_DIR / "programacao_entregas_tppf_w38.xlsx"
ARQ_TABELA_SUPORTE = BASE_DIR / "tabela_lojas_rotas.xlsx"
ARQ_SAIDA = BASE_DIR / "total_pecas_por_rota.xlsx"

STATUS_PENDENTE = "PENDENTE - preencher rota manualmente"
STATUS_AUTO = "Cadastrada automaticamente pela programação"


def normaliza_loja(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    texto = str(valor).strip().upper()
    return texto[:-2] if texto.endswith(".0") and texto[:-2].isdigit() else texto


def validar_colunas(df, obrigatorias, origem):
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Coluna(s) ausente(s) em {origem}: {', '.join(faltantes)}")


def carregar_remessas(caminho):
    df = pd.read_excel(caminho, sheet_name="Remessas")
    validar_colunas(df, ["Loja", "Peças"], caminho.name)
    df = df[df["Loja"].notna()].copy()
    df["Loja_norm"] = df["Loja"].apply(normaliza_loja)
    df["Peças"] = pd.to_numeric(df["Peças"], errors="coerce").fillna(0)

    # A planilha nova já possui Rota. Renomeamos para evitar Rota_x/Rota_y no merge.
    if "Rota" in df.columns:
        df = df.rename(columns={"Rota": "Rota_remessa"})
    return df


def carregar_mapa_da_programacao(caminho):
    df = pd.read_excel(caminho, sheet_name="Planilha1")
    validar_colunas(df, ["Loja", "Rota"], caminho.name)
    df["Loja_norm"] = df["Loja"].apply(normaliza_loja)
    df["Rota"] = df["Rota"].astype("string").str.strip()
    df = df[df["Loja_norm"].notna() & df["Rota"].notna() & (df["Rota"] != "")]
    return df[["Loja_norm", "Rota"]].drop_duplicates(subset="Loja_norm", keep="last")


def carregar_tabela_suporte(caminho):
    if not caminho.exists():
        return pd.DataFrame(columns=["Loja", "Rota", "Status"])

    tabela = pd.read_excel(caminho, dtype={"Loja": str})
    validar_colunas(tabela, ["Loja", "Rota"], caminho.name)
    tabela["Loja"] = tabela["Loja"].apply(normaliza_loja)
    tabela["Rota"] = tabela["Rota"].astype("string").str.strip()
    return tabela


def atualizar_tabela_suporte(tabela, mapa_semana):
    conhecidas = set(tabela["Loja"].dropna())
    novas = mapa_semana[~mapa_semana["Loja_norm"].isin(conhecidas)].copy()

    if novas.empty:
        return tabela

    novas = novas.rename(columns={"Loja_norm": "Loja"})
    # Mantém todas as colunas novas da tabela atual (Max Peças, Qtd. Ruas etc.).
    for coluna in tabela.columns:
        if coluna not in novas.columns:
            novas[coluna] = pd.NA
    if "Status" in tabela.columns:
        novas["Status"] = STATUS_AUTO

    tabela = pd.concat([tabela, novas[tabela.columns]], ignore_index=True)
    print(f"+ {len(novas)} loja(s) nova(s) cadastrada(s) pela programação.")
    return tabela


def garantir_lojas_das_remessas(tabela, remessas):
    conhecidas = set(tabela["Loja"].dropna())
    faltantes = sorted(set(remessas["Loja_norm"].dropna()) - conhecidas)
    if not faltantes:
        return tabela

    linhas = []
    for loja in faltantes:
        linha = {col: pd.NA for col in tabela.columns}
        linha["Loja"] = loja
        linha["Rota"] = pd.NA
        if "Status" in tabela.columns:
            linha["Status"] = STATUS_PENDENTE
        linhas.append(linha)

    tabela = pd.concat([tabela, pd.DataFrame(linhas)], ignore_index=True)
    print(f"ATENÇÃO: {len(faltantes)} loja(s) sem rota: {faltantes}")
    print(f"Preencha a coluna Rota em '{ARQ_TABELA_SUPORTE.name}' e rode novamente.")
    return tabela


def calcular_total_por_rota():
    for arquivo in (ARQ_REMESSAS, ARQ_PROGRAMACAO):
        if not arquivo.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")

    remessas = carregar_remessas(ARQ_REMESSAS)
    mapa_semana = carregar_mapa_da_programacao(ARQ_PROGRAMACAO)
    tabela = carregar_tabela_suporte(ARQ_TABELA_SUPORTE)
    tabela = atualizar_tabela_suporte(tabela, mapa_semana)
    tabela = garantir_lojas_das_remessas(tabela, remessas)

    tabela = tabela.sort_values("Loja", key=lambda s: s.astype(str)).reset_index(drop=True)
    tabela.to_excel(ARQ_TABELA_SUPORTE, index=False)

    completo = remessas.merge(
        tabela[["Loja", "Rota"]].drop_duplicates(subset="Loja", keep="last"),
        left_on="Loja_norm", right_on="Loja", how="left"
    )

    rota_vazia = completo["Rota"].isna() | (completo["Rota"].astype("string").str.strip() == "")
    if rota_vazia.any():
        lojas_pendentes = sorted(completo.loc[rota_vazia, "Loja_norm"].dropna().unique())
        print(f"ATENÇÃO: rota ainda não preenchida para: {lojas_pendentes}")

    completo["Rota"] = completo["Rota"].fillna("").astype(str).str.strip().replace("", "SEM ROTA")

    resumo = (
        completo.groupby("Rota", as_index=False)["Peças"]
        .sum()
        .sort_values("Peças", ascending=False)
        .reset_index(drop=True)
    )
    resumo["Peças"] = resumo["Peças"].round().astype(int)

    total_geral = int(resumo["Peças"].sum())
    resumo_com_total = pd.concat(
        [resumo, pd.DataFrame([{"Rota": "TOTAL GERAL", "Peças": total_geral}])],
        ignore_index=True,
    )
    resumo_com_total.to_excel(ARQ_SAIDA, index=False)

    print("\nTotal de peças por rota:")
    print(resumo.to_string(index=False))
    print(f"\nTotal geral: {total_geral} peças")
    print(f"Arquivos gerados/atualizados: {ARQ_SAIDA.name} | {ARQ_TABELA_SUPORTE.name}")
    return resumo


if __name__ == "__main__":
    calcular_total_por_rota()
