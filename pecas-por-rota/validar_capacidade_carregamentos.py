"""Validação diária de disponibilidade de peças por carregamento/rota.

Ao executar, o programa pergunta qual DATA deve ser analisada.

Arquivos utilizados:
- no_chao_remessas_tp_todas.xlsx
    Disponibilidade ATUAL de peças no chão.

- programacao_entregas_tppf_w38.xlsx
    Programação dos carregamentos.

- tabela_lojas_rotas.xlsx
    Referência de Max Peças por loja.

REGRA:
1. O usuário informa a data que deseja analisar.
2. A programação é filtrada SOMENTE para essa data.
3. O estoque de no_chao_remessas_tp_todas.xlsx é considerado como
   a disponibilidade atual.
4. Os carregamentos daquele dia são processados em ordem de horário.
5. Se uma rota tiver mais de um carregamento no mesmo dia,
   o primeiro carregamento consome o estoque e o saldo restante
   é utilizado no próximo.
6. Dias anteriores e posteriores NÃO entram no cálculo.
"""

from pathlib import Path
import pandas as pd


# ============================================================
# ARQUIVOS
# ============================================================

PASTA = Path(__file__).resolve().parent

ARQ_DISPONIBILIDADE = PASTA / "no_chao_remessas_tp_todas.xlsx"
ARQ_PROGRAMACAO = PASTA / "programacao_entregas_tppf_w38.xlsx"
ARQ_LOJAS_ROTAS = PASTA / "tabela_lojas_rotas.xlsx"
ARQ_SAIDA = PASTA / "validacao_capacidade_carregamentos.xlsx"


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normaliza_texto(v):
    if pd.isna(v):
        return ""

    return str(v).strip().upper()


def normaliza_loja(v):
    if pd.isna(v):
        return None

    if isinstance(v, float) and v.is_integer():
        v = int(v)

    return str(v).strip().upper()


def exigir_colunas(df, cols, arquivo):
    faltam = [c for c in cols if c not in df.columns]

    if faltam:
        raise ValueError(
            f"\nERRO no arquivo {arquivo.name}.\n"
            f"Colunas ausentes: {faltam}"
        )


def verificar_arquivo(caminho):
    if not caminho.exists():
        raise FileNotFoundError(
            f"\nArquivo não encontrado:\n{caminho}\n"
        )


# ============================================================
# DISPONIBILIDADE ATUAL DE PEÇAS
# ============================================================

def carregar_disponibilidade():

    verificar_arquivo(ARQ_DISPONIBILIDADE)

    # Tenta primeiro a aba Remessas.
    # Se o nome da aba mudar, lê automaticamente a primeira aba.
    try:
        df = pd.read_excel(
            ARQ_DISPONIBILIDADE,
            sheet_name="Remessas"
        )
    except ValueError:
        df = pd.read_excel(ARQ_DISPONIBILIDADE)

    exigir_colunas(
        df,
        ["Rota", "Peças"],
        ARQ_DISPONIBILIDADE
    )

    # Remove linhas sem rota
    df = df[df["Rota"].notna()].copy()

    df["Rota_norm"] = df["Rota"].apply(normaliza_texto)

    df["Peças"] = pd.to_numeric(
        df["Peças"],
        errors="coerce"
    ).fillna(0)

    # Soma todas as peças disponíveis de cada rota
    disponibilidade = (
        df.groupby(
            "Rota_norm",
            as_index=False
        )["Peças"]
        .sum()
        .rename(
            columns={
                "Peças": "Disponível inicial"
            }
        )
    )

    return disponibilidade


# ============================================================
# PROGRAMAÇÃO DOS CARREGAMENTOS
# ============================================================

def carregar_programacao():

    verificar_arquivo(ARQ_PROGRAMACAO)

    # Tenta Planilha1.
    # Caso o nome tenha mudado, utiliza a primeira aba.
    try:
        df = pd.read_excel(
            ARQ_PROGRAMACAO,
            sheet_name="Planilha1"
        )
    except ValueError:
        df = pd.read_excel(ARQ_PROGRAMACAO)

    exigir_colunas(
        df,
        [
            "Loja",
            "Rota",
            "Carregamento",
            "Capacidade do veículo"
        ],
        ARQ_PROGRAMACAO
    )

    df["Rota_norm"] = df["Rota"].apply(
        normaliza_texto
    )

    df["Loja_norm"] = df["Loja"].apply(
        normaliza_loja
    )

    # Converte a coluna de carregamento para data
    df["Carregamento"] = pd.to_datetime(
        df["Carregamento"],
        errors="coerce",
        dayfirst=True
    )

    df["Capacidade do veículo"] = pd.to_numeric(
        df["Capacidade do veículo"],
        errors="coerce"
    ).fillna(0)

    # Tratamento da hora
    if "Hora" in df.columns:

        def converter_hora(v):

            if pd.isna(v):
                return ""

            # Caso seja horário do Excel
            if hasattr(v, "strftime"):
                try:
                    return v.strftime("%H:%M")
                except Exception:
                    pass

            return str(v).strip()

        df["Hora_ordem"] = df["Hora"].apply(
            converter_hora
        )

    else:
        df["Hora_ordem"] = ""

    return df


# ============================================================
# MAX PEÇAS POR LOJA
# ============================================================

def carregar_max_pecas():

    if not ARQ_LOJAS_ROTAS.exists():

        print(
            f"\nAVISO: arquivo não encontrado:\n"
            f"{ARQ_LOJAS_ROTAS}"
        )

        return pd.DataFrame(
            columns=[
                "Loja_norm",
                "Max Peças"
            ]
        )

    # Lê a primeira aba, independentemente do nome
    df = pd.read_excel(
        ARQ_LOJAS_ROTAS
    )

    if not {"Loja", "Max Peças"}.issubset(df.columns):

        print(
            f"\nAVISO: {ARQ_LOJAS_ROTAS.name} "
            f"não possui as colunas "
            f"'Loja' e 'Max Peças'."
        )

        return pd.DataFrame(
            columns=[
                "Loja_norm",
                "Max Peças"
            ]
        )

    df["Loja_norm"] = df["Loja"].apply(
        normaliza_loja
    )

    df["Max Peças"] = pd.to_numeric(
        df["Max Peças"],
        errors="coerce"
    )

    return (
        df[
            [
                "Loja_norm",
                "Max Peças"
            ]
        ]
        .drop_duplicates("Loja_norm")
    )


# ============================================================
# MONTA OS CARREGAMENTOS
# ============================================================

def montar_carregamentos(
    programacao,
    max_pecas
):

    p = programacao.merge(
        max_pecas,
        on="Loja_norm",
        how="left"
    )

    # Cada carga é identificada por:
    # data + rota + carreta + horário

    chaves = [
        "Carregamento",
        "Rota_norm"
    ]

    coluna_carreta = (
        "Ocupação Plano / Planejamento de Carretas TP"
    )

    if coluna_carreta in p.columns:
        chaves.append(coluna_carreta)

    chaves.append("Hora_ordem")

    agg = {

        "Loja_norm":
            lambda s: ", ".join(
                sorted(
                    set(
                        x
                        for x in s.dropna()
                    )
                )
            ),

        "Capacidade do veículo":
            "max",

        "Max Peças":
            lambda s: s.sum(
                min_count=1
            ),
    }

    if "RESTRIÇÃO" in p.columns:

        p["RESTRIÇÃO"] = pd.to_numeric(
            p["RESTRIÇÃO"],
            errors="coerce"
        )

        agg["RESTRIÇÃO"] = (
            lambda s: s.sum(
                min_count=1
            )
        )

    cargas = (
        p.groupby(
            chaves,
            dropna=False,
            as_index=False
        )
        .agg(agg)
    )

    cargas = cargas.rename(
        columns={

            "Loja_norm":
                "Lojas",

            "Capacidade do veículo":
                "Necessário (capacidade veículo)",

            "Max Peças":
                "Soma Max Peças lojas",

            "RESTRIÇÃO":
                "Soma restrições lojas",
        }
    )

    cargas = cargas.sort_values(
        [
            "Hora_ordem",
            "Rota_norm"
        ]
    ).reset_index(drop=True)

    return cargas


# ============================================================
# VALIDAÇÃO
# ============================================================

def validar():

    print()
    print("=" * 70)
    print("VALIDAÇÃO DIÁRIA DE PEÇAS PARA CARREGAMENTO")
    print("=" * 70)
    print()

    # --------------------------------------------------------
    # PERGUNTA PRIMEIRO QUAL DIA DEVE SER ANALISADO
    # --------------------------------------------------------

    data_digitada = input(
        "Informe a data que deseja analisar "
        "(DD/MM/AAAA): "
    ).strip()

    try:

        data_analise = pd.to_datetime(
            data_digitada,
            format="%d/%m/%Y"
        ).normalize()

    except ValueError:

        print()
        print(
            "ERRO: Data inválida."
        )

        print(
            "Digite no formato DD/MM/AAAA."
        )

        print(
            "Exemplo: 25/09/2026"
        )

        return

    print()
    print(
        f"Data selecionada: "
        f"{data_analise.strftime('%d/%m/%Y')}"
    )

    print()
    print("Carregando arquivos...")

    # --------------------------------------------------------
    # CARREGA OS ARQUIVOS
    # --------------------------------------------------------

    try:

        disp = carregar_disponibilidade()

        prog = carregar_programacao()

        maxp = carregar_max_pecas()

    except Exception as erro:

        print()
        print("=" * 70)
        print("ERRO AO CARREGAR OS ARQUIVOS")
        print("=" * 70)
        print()
        print(erro)
        print()

        return

    # --------------------------------------------------------
    # FILTRA SOMENTE A DATA INFORMADA
    # --------------------------------------------------------

    prog_dia = prog[
        prog["Carregamento"]
        .dt.normalize()
        .eq(data_analise)
    ].copy()

    # --------------------------------------------------------
    # VERIFICA SE EXISTE PROGRAMAÇÃO
    # --------------------------------------------------------

    if prog_dia.empty:

        print()
        print("=" * 70)

        print(
            f"NÃO EXISTEM CARREGAMENTOS "
            f"PROGRAMADOS PARA "
            f"{data_analise.strftime('%d/%m/%Y')}."
        )

        print("=" * 70)

        # Mostra quais datas existem no arquivo
        datas_disponiveis = (
            prog["Carregamento"]
            .dropna()
            .dt.normalize()
            .drop_duplicates()
            .sort_values()
        )

        if not datas_disponiveis.empty:

            print()
            print(
                "Datas encontradas na programação:"
            )

            for data in datas_disponiveis:

                print(
                    " - "
                    + data.strftime(
                        "%d/%m/%Y"
                    )
                )

        return

    # --------------------------------------------------------
    # MOSTRA QUANTIDADE DE REGISTROS
    # --------------------------------------------------------

    print(
        f"Encontrados "
        f"{len(prog_dia)} registros "
        f"na programação desse dia."
    )

    # --------------------------------------------------------
    # MONTA AS CARGAS SOMENTE DESSE DIA
    # --------------------------------------------------------

    cargas = montar_carregamentos(
        prog_dia,
        maxp
    )

    print(
        f"Total de carregamentos identificados: "
        f"{len(cargas)}"
    )

    # --------------------------------------------------------
    # ESTOQUE INICIAL
    # --------------------------------------------------------

    saldo = dict(
        zip(
            disp["Rota_norm"],
            disp["Disponível inicial"]
        )
    )

    linhas = []

    # --------------------------------------------------------
    # PROCESSA SOMENTE AS CARGAS DO DIA
    # --------------------------------------------------------

    for _, r in cargas.iterrows():

        rota = r["Rota_norm"]

        # Estoque disponível antes da carga
        antes = int(
            round(
                saldo.get(
                    rota,
                    0
                )
            )
        )

        # Quantidade necessária
        necessario = int(
            round(
                r[
                    "Necessário "
                    "(capacidade veículo)"
                ]
            )
        )

        # Quanto conseguimos atender
        atendido = min(
            antes,
            necessario
        )

        # Quanto está faltando
        falta = max(
            necessario - antes,
            0
        )

        # Saldo depois dessa carga
        depois = max(
            antes - necessario,
            0
        )

        # Atualiza saldo da rota
        saldo[rota] = depois

        if falta == 0:
            status = "OK - SUFICIENTE"
        else:
            status = "FALTA PEÇAS"

        linha = {

            "Data carregamento":
                data_analise,

            "Hora":
                r["Hora_ordem"],

            "Rota":
                rota,

            "Carreta":
                r.get(
                    "Ocupação Plano / "
                    "Planejamento de Carretas TP",
                    ""
                ),

            "Lojas":
                r["Lojas"],

            "Necessário":
                necessario,

            "Disponível antes":
                antes,

            "Atendido":
                atendido,

            "Faltante":
                falta,

            "Saldo após carregamento":
                depois,

            "Status":
                status,

            "Soma Max Peças lojas":
                r.get(
                    "Soma Max Peças lojas"
                ),

            "Soma restrições lojas":
                r.get(
                    "Soma restrições lojas"
                ),
        }

        linhas.append(linha)

    # --------------------------------------------------------
    # RESULTADO
    # --------------------------------------------------------

    resultado = pd.DataFrame(
        linhas
    )

    # --------------------------------------------------------
    # RESUMO POR ROTA
    # --------------------------------------------------------

    resumo = (
        resultado.groupby(
            "Rota",
            as_index=False
        )
        .agg(

            Necessário=(
                "Necessário",
                "sum"
            ),

            **{
                "Disponível inicial":
                    (
                        "Disponível antes",
                        "first"
                    ),

                "Atendido":
                    (
                        "Atendido",
                        "sum"
                    ),

                "Faltante total":
                    (
                        "Faltante",
                        "sum"
                    ),

                "Saldo final":
                    (
                        "Saldo após carregamento",
                        "last"
                    )
            }
        )
    )

    resumo["Status"] = (
        resumo[
            "Faltante total"
        ]
        .apply(
            lambda x:
            "OK - SUFICIENTE"
            if x == 0
            else "FALTA PEÇAS"
        )
    )

    # --------------------------------------------------------
    # SALVA EXCEL
    # --------------------------------------------------------

    try:

        with pd.ExcelWriter(
            ARQ_SAIDA,
            engine="openpyxl"
        ) as w:

            resultado.to_excel(
                w,
                sheet_name="Validação por carga",
                index=False
            )

            resumo.to_excel(
                w,
                sheet_name="Resumo por rota",
                index=False
            )

            disp.to_excel(
                w,
                sheet_name="Disponibilidade atual",
                index=False
            )

    except PermissionError:

        print()
        print(
            "ERRO: não foi possível salvar o Excel."
        )

        print(
            f"Feche o arquivo "
            f"{ARQ_SAIDA.name} "
            f"caso ele esteja aberto e "
            f"execute novamente."
        )

        return

    # --------------------------------------------------------
    # EXIBE RESULTADO
    # --------------------------------------------------------

    print()
    print("=" * 100)

    print(
        "RESULTADO - "
        + data_analise.strftime(
            "%d/%m/%Y"
        )
    )

    print("=" * 100)

    cols = [

        "Hora",

        "Rota",

        "Necessário",

        "Disponível antes",

        "Atendido",

        "Faltante",

        "Saldo após carregamento",

        "Status"
    ]

    print(
        resultado[
            cols
        ].to_string(
            index=False
        )
    )

    # --------------------------------------------------------
    # TOTAIS DO DIA
    # --------------------------------------------------------

    total_necessario = int(
        resultado[
            "Necessário"
        ].sum()
    )

    total_atendido = int(
        resultado[
            "Atendido"
        ].sum()
    )

    total_faltante = int(
        resultado[
            "Faltante"
        ].sum()
    )

    print()
    print("=" * 70)
    print("RESUMO DO DIA")
    print("=" * 70)

    print(
        f"Data: "
        f"{data_analise.strftime('%d/%m/%Y')}"
    )

    print(
        f"Quantidade de carregamentos: "
        f"{len(resultado)}"
    )

    print(
        f"Total necessário: "
        f"{total_necessario:,}"
        .replace(",", ".")
        + " peças"
    )

    print(
        f"Total atendido: "
        f"{total_atendido:,}"
        .replace(",", ".")
        + " peças"
    )

    print(
        f"Total faltante: "
        f"{total_faltante:,}"
        .replace(",", ".")
        + " peças"
    )

    if total_faltante == 0:

        print()
        print(
            "RESULTADO: HÁ PEÇAS "
            "SUFICIENTES PARA TODOS "
            "OS CARREGAMENTOS DO DIA."
        )

    else:

        print()
        print(
            "ATENÇÃO: NÃO HÁ PEÇAS "
            "SUFICIENTES PARA TODOS "
            "OS CARREGAMENTOS DO DIA."
        )

    print()
    print(
        f"Arquivo gerado:\n"
        f"{ARQ_SAIDA}"
    )

    print()

    return resultado


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    validar()
