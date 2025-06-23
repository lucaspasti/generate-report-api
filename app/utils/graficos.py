import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


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
        "Distribuição Granulométrica por Amostra",
        fontsize=14,
        weight="bold",
        pad=15
    )
    ax.legend(loc="upper left", bbox_to_anchor=(
        1, 1), fontsize=10, frameon=True)
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    plt.close(fig)

    return fig


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
        ax.plot(dados[eixo_x].values, dados[col].values,
                marker="o", label="Amostras")

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
        for _, row in dados.iterrows():
            classe = row[classe_col]
            vmp = vmp_dict.get(classe, {}).get(col, None)
            if vmp is not None:
                ax.axhline(y=vmp, color="red", linestyle="--",
                           linewidth=1.0, alpha=0.5)

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
        plt.close(fig)

    return figuras


def grafico_qualidade_agua(df, parametros, classe, vmp_qag):
    """
    Gera gráficos de qualidade de água para cada parâmetro em `parametros`, na classe `classe`.
    """
    # 1) Garante que não trabalhamos sobre uma view
    df = df.copy()

    # 2) Converte todas as colunas de parâmetros de uma só vez
    df.loc[:, parametros] = df.loc[:, parametros].apply(
        pd.to_numeric, errors="coerce")

    figs = []
    try:
        for parametro in parametros:
            if parametro not in df.columns:
                raise ValueError(
                    f"Parâmetro '{parametro}' não encontrado no DataFrame.")

            # Agrupa e pivot
            df_grouped = (
                df.groupby(["Ponto", "Profundidade"])[parametro]
                .mean()
                .unstack(fill_value=np.nan)
            )
            pontos = df_grouped.index
            x = np.arange(len(pontos))
            width = 0.2

            superficie = df_grouped.get(
                "Superfície", pd.Series(np.nan, index=pontos)).values
            meio = df_grouped.get("Meio",      pd.Series(
                np.nan, index=pontos)).values
            fundo = df_grouped.get(
                "Fundo",     pd.Series(np.nan, index=pontos)).values
            media_total = np.full(len(pontos), np.nanmean(
                np.concatenate([superficie, meio, fundo])))

            media_ponto = np.nanmean([superficie, meio, fundo], axis=0)

            vmp_valor = vmp_qag.get(classe, {}).get(parametro, None)
            if vmp_valor is None or isinstance(vmp_valor, str):
                conama_limite = np.full(len(pontos), np.nan)
                limite_str = "(sem limite numérico)"
            else:
                conama_limite = np.full(len(pontos), float(vmp_valor))
                limite_str = f"(Limite CONAMA: {vmp_valor})"

            # Plot
            fig, ax = plt.subplots(figsize=(12, 6), dpi=120)
            ax.bar(x - width, superficie, width, label="Superfície")
            ax.bar(x, meio,      width, label="Meio")
            ax.bar(x + width, fundo, width, label="Fundo")

            ax.plot(x, media_total, label="Média total",   linewidth=2)
            ax.plot(x, media_ponto, label="Média ponto",
                    linestyle="--", linewidth=2)
            ax.plot(x, conama_limite, label="CONAMA",
                    linestyle="--", linewidth=3)

            ax.set_xticks(x)
            ax.set_xticklabels(pontos, rotation=45)
            ax.set_ylabel("Concentração (mg/L)")
            ax.set_title(f"{parametro} — Classe {classe} {limite_str}")
            ax.legend()
            ax.grid(axis="y", linestyle="--", alpha=0.7)
            plt.tight_layout()

            figs.append(fig)
            plt.close(fig)
    except Exception as e:
        print(f"Erro ao gerar gráfico para {classe} - {parametro}: {e}")
        return []

    return figs
