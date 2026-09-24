"""
Total de Peças por Rota
========================

Cruza a planilha de remessas (peças por loja) com uma TABELA DE SUPORTE
(Loja -> Rota) para calcular o total de peças transportado por rota.

A tabela de suporte é persistente (fica salva em disco) e é alimentada
automaticamente pela planilha de programação de entregas a cada execução,
sem nunca apagar o que já foi cadastrado antes. Isso resolve o problema
de a programação de transporte só trazer a loja na semana em que ela
tem envio: uma vez cadastrada, a loja permanece na tabela de suporte
mesmo em semanas em que não aparece na programação.

Entradas esperadas:
  - no-chao-remessas-tp-todas.xlsx   (aba "Remessas")
      colunas usadas: Remessa, Loja, Peças
  - Programação_-_Entregas_TP_PF_-_W38.xlsx  (aba "Planilha1")
      colunas usadas: Loja, Rota
      (troque pelo arquivo da semana atual quando for rodar de novo)

Saída:
  - tabela_lojas_rotas.xlsx     (tabela de suporte Loja -> Rota, atualizada)
  - total_pecas_por_rota.xlsx   (resumo de peças por rota)
  - impressão no console do total geral e de lojas pendentes de cadastro
"""

import pandas as pd
from pathlib import Path

ARQ_REMESSAS = "no-chao-remessas-tp-todas.xlsx"
ARQ_PROGRAMACAO = "Programação_-_Entregas_TP_PF_-_W38.xlsx"
ARQ_TABELA_SUPORTE = "tabela_lojas_rotas.xlsx"
ARQ_SAIDA = "total_pecas_por_rota.xlsx"

STATUS_PENDENTE = "PENDENTE - preencher rota manualmente"
STATUS_AUTO = "Cadastrada automaticamente pela programação"


def normaliza_loja(valor):
    """Padroniza o código da loja para permitir o cruzamento entre as
    duas planilhas (ex.: 120.0 -> '120', 'YC105 ' -> 'YC105')."""
    if pd.isna(valor):
        return None
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip().upper()


def carregar_remessas(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(caminho, sheet_name="Remessas")

    # remove a linha de total ("Total (122)") e linhas sem loja
    df = df[df["Loja"].notna()].copy()

    df["Loja_norm"] = df["Loja"].apply(normaliza_loja)
    df["Peças"] = pd.to_numeric(df["Peças"], errors="coerce").fillna(0)
    return df


def carregar_mapa_da_programacao(caminho: str) -> pd.DataFrame:
    """Lê Loja -> Rota da planilha de programação da semana atual."""
    df = pd.read_excel(caminho, sheet_name="Planilha1")
    df["Loja_norm"] = df["Loja"].apply(normaliza_loja)

    # cada loja aparece 1x por data de entrega na semana, mas a rota
    # é sempre a mesma para a mesma loja -> ficamos só com o par único
    mapa = df[["Loja_norm", "Rota"]].drop_duplicates(subset="Loja_norm")
    return mapa


def carregar_tabela_suporte(caminho: str) -> pd.DataFrame:
    """Carrega a tabela de suporte Loja -> Rota já existente, ou cria
    uma vazia se ainda não existir (primeira execução)."""
    if Path(caminho).exists():
        tabela = pd.read_excel(caminho, dtype={"Loja": str})
        tabela["Loja"] = tabela["Loja"].apply(normaliza_loja)
        return tabela
    return pd.DataFrame(columns=["Loja", "Rota", "Status"])


def atualizar_tabela_suporte(tabela: pd.DataFrame, mapa_semana: pd.DataFrame) -> pd.DataFrame:
    """Acrescenta à tabela de suporte as lojas novas vindas da programação
    da semana atual, sem sobrescrever o que já está cadastrado."""
    conhecidas = set(tabela["Loja"])
    novas = mapa_semana[~mapa_semana["Loja_norm"].isin(conhecidas)].copy()

    if not novas.empty:
        novas = novas.rename(columns={"Loja_norm": "Loja"})
        novas["Status"] = STATUS_AUTO
        tabela = pd.concat([tabela, novas[["Loja", "Rota", "Status"]]], ignore_index=True)
        print(f"➕ {len(novas)} loja(s) nova(s) cadastrada(s) na tabela de suporte a partir da programação.\n")

    return tabela


def garantir_lojas_das_remessas(tabela: pd.DataFrame, remessas: pd.DataFrame) -> pd.DataFrame:
    """Garante que toda loja que aparece nas remessas exista na tabela de
    suporte. Se não existir em lugar nenhum, cria a linha com rota em
    branco e status PENDENTE, para preenchimento manual."""
    conhecidas = set(tabela["Loja"])
    lojas_remessas = set(remessas["Loja_norm"].dropna())
    faltantes = sorted(lojas_remessas - conhecidas)

    if faltantes:
        pendentes = pd.DataFrame({
            "Loja": faltantes,
            "Rota": [None] * len(faltantes),
            "Status": [STATUS_PENDENTE] * len(faltantes),
        })
        tabela = pd.concat([tabela, pendentes], ignore_index=True)
        print(f"⚠️  {len(faltantes)} loja(s) sem rota cadastrada em lugar nenhum: {faltantes}")
        print(f"    Foram adicionadas em '{ARQ_TABELA_SUPORTE}' com status PENDENTE.")
        print("    Preencha a coluna 'Rota' manualmente e rode o script de novo.\n")

    return tabela


def calcular_total_por_rota():
    remessas = carregar_remessas(ARQ_REMESSAS)
    mapa_semana = carregar_mapa_da_programacao(ARQ_PROGRAMACAO)

    tabela_suporte = carregar_tabela_suporte(ARQ_TABELA_SUPORTE)
    tabela_suporte = atualizar_tabela_suporte(tabela_suporte, mapa_semana)
    tabela_suporte = garantir_lojas_das_remessas(tabela_suporte, remessas)

    # salva a tabela de suporte já atualizada (cresce a cada execução)
    tabela_suporte.sort_values("Loja").to_excel(ARQ_TABELA_SUPORTE, index=False)

    completo = remessas.merge(
        tabela_suporte[["Loja", "Rota"]],
        left_on="Loja_norm", right_on="Loja", how="left",
    )

    pendentes = completo[completo["Rota"].isna() | (completo["Rota"] == "")]
    if not pendentes.empty:
        lojas_pendentes = sorted(pendentes["Loja_norm"].unique())
        print(f"⚠️  Rota ainda não preenchida para: {lojas_pendentes}")
        print("    (essas peças ficaram em 'SEM ROTA' até você preencher a tabela de suporte)\n")

    completo["Rota"] = completo["Rota"].fillna("").replace("", "SEM ROTA")

    resumo = (
        completo.groupby("Rota", as_index=False)["Peças"]
        .sum()
        .sort_values("Peças", ascending=False)
        .reset_index(drop=True)
    )
    resumo["Peças"] = resumo["Peças"].astype(int)

    total_geral = resumo["Peças"].sum()
    linha_total = pd.DataFrame([{"Rota": "TOTAL GERAL", "Peças": total_geral}])
    resumo_com_total = pd.concat([resumo, linha_total], ignore_index=True)

    resumo_com_total.to_excel(ARQ_SAIDA, index=False)

    print("Total de peças por rota:")
    print(resumo.to_string(index=False))
    print(f"\nTotal geral: {total_geral} peças")
    print(f"\nArquivos gerados: {ARQ_SAIDA} | {ARQ_TABELA_SUPORTE}")

    return resumo


if __name__ == "__main__":
    calcular_total_por_rota()
