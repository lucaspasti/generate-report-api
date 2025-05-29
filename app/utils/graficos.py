import matplotlib.pyplot as plt # type: ignore
import numpy as np # type: ignore
import pandas as pd # type: ignore


def gera_distribuicao_granulometrica_qsd(dados):

    categorias = [
        "Areia muito grossa (%)",
        "Areia grossa (%)",
        "Areia média (%)",
        "Areia fina (%)",
        "Areia muito fina (%)",
        "Silte (%)",
        "Argila (%)",
    ]
    cores = [
        "#00b0f0",
        "#0070c0",
        "#00b050",
        "#7030a0",
        "#002060",
        "#76933c",
        "#f4b084",
    ]

    x = np.arange(dados.shape[0])

    fig, ax = plt.subplots(figsize=(8, 6), dpi=120)
    bottom = np.zeros(dados.shape[0])

    for i, cat in enumerate(categorias):
        bars = ax.bar(
            x,
            dados[cat],
            bottom=bottom,
            label=cat,
            color=cores[i],
            edgecolor="white",
            linewidth=1,
        )
        bottom += dados[cat]

    ax.set_xticks(x)
    ax.set_xticklabels(dados.index.tolist(), fontsize=11)
    ax.set_ylabel("Granulometria (%)", fontsize=12)
    ax.set_xlabel("Amostras", fontsize=12)
    ax.set_ylim(0, 100)
    ax.set_title(
        "Distribuição Granulométrica por Amostra", fontsize=14, weight="bold", pad=15
    )
    ax.legend(loc="upper left", bbox_to_anchor=(1, 1), fontsize=10, frameon=True)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return fig


import matplotlib.pyplot as plt # type: ignore


def graficos_linha_com_vmp_por_classe_qsd(
    dados, colunas, vmp_dict, eixo_x="Ponto", classe_col="Classe"
):
    if eixo_x not in dados.columns or classe_col not in dados.columns:
        raise ValueError("Colunas obrigatórias não encontradas no DataFrame.")

    figuras = []

    for col in colunas:
        if col == eixo_x or col == classe_col:
            continue

        fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
        ax.plot(dados[eixo_x].values, dados[col].values, marker="o", label="Amostras")

        # Linha da média dos pontos
        media = dados[col].mean()
        ax.axhline(
            y=media,
            color="blue",
            linestyle="-.",
            linewidth=1.5,
            label=f"Média ({media:.2f})",
        )

        # Linhas VMP por classe
        for i, row in dados.iterrows():
            classe = row[classe_col]
            vmp = vmp_dict.get(classe, {}).get(col, None)
            if vmp is not None:
                ax.axhline(y=vmp, color="red", linestyle="--", linewidth=1.0, alpha=0.5)

        # Legenda do VMP
        classes_usadas = dados[classe_col].unique()
        legenda = (
            "VMP (por classe)"
            if len(classes_usadas) > 1
            else f"VMP ({classes_usadas[0]})"
        )
        valor_legenda = vmp_dict.get(classes_usadas[0], {}).get(col, None)
        if valor_legenda is not None:
            ax.axhline(
                y=valor_legenda,
                color="red",
                linestyle="--",
                linewidth=1.5,
                label=legenda,
            )

        ax.set_xlabel(eixo_x)
        ax.set_ylabel(col)
        ax.set_title(f"{col} vs {eixo_x}")
        ax.legend()
        fig.tight_layout()

        figuras.append(fig)

    return figuras


def grafico_qualidade_agua(df, parametro, classe, vmp_qag):
    figs = []
    for parametro in parametro:

        # Garantir que a coluna existe e é numérica
        if parametro not in df.columns:
            raise ValueError(f"Parâmetro '{parametro}' não encontrado no DataFrame.")
        df[parametro] = pd.to_numeric(df[parametro], errors="coerce")

        # Agrupar e pivotar
        df_grouped = df.groupby(["Ponto", "Profundidade"])[parametro].mean().unstack()
        pontos = df_grouped.index
        x = np.arange(len(pontos))
        width = 0.2

        # Obter valores para profundidades
        superficie = df_grouped.get(
            "Superfície", pd.Series([np.nan] * len(pontos), index=pontos)
        ).values
        meio = df_grouped.get(
            "Meio", pd.Series([np.nan] * len(pontos), index=pontos)
        ).values
        fundo = df_grouped.get(
            "Fundo", pd.Series([np.nan] * len(pontos), index=pontos)
        ).values

        # Média total e média por ponto
        media_total = np.full(
            len(pontos), np.nanmean(np.concatenate([superficie, meio, fundo]))
        )
        media_ponto = np.nanmean([superficie, meio, fundo], axis=0)

        # Valor do VMP
        vmp_valor = vmp_qag.get(classe, {}).get(parametro, None)
        if isinstance(vmp_valor, str) or vmp_valor is None:
            conama_limite = np.full(len(pontos), np.nan)
            limite_str = f"(sem limite numérico)"
        else:
            # cria linha constante no valor do limite
            conama_limite = np.full(len(pontos), float(vmp_valor))
            limite_str = f"(Limite CONAMA: {vmp_valor})"

        # Plot
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.bar(x - width, superficie, width, label="Superfície", color="#0072B2")
        ax.bar(x, meio, width, label="Meio", color="#E69F00")
        ax.bar(x + width, fundo, width, label="Fundo", color="#009E73")

        ax.plot(
            x,
            media_total,
            label="Média total",
            color="#56B4E9",
            linewidth=2,
        )
        ax.plot(
            x,
            media_ponto,
            label="Média ponto",
            color="olive",
            linestyle="--",
            linewidth=2,
        )
        ax.plot(
            x,
            conama_limite,
            label="CONAMA",
            color="red",
            linestyle="--",
            linewidth=3,
        )

        ax.set_xticks(x)
        ax.set_xticklabels(pontos, rotation=45)
        ax.set_ylabel("Concentração (mg/L)")
        ax.set_title(f"{parametro} - {classe} {limite_str}")
        ax.legend()
        ax.grid(True, axis="y", linestyle="--", alpha=0.7)

        figs.append(fig)

    return figs
