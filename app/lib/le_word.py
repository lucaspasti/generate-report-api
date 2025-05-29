import pandas as pd
from docxtpl import DocxTemplate

labels = {
    "qualidade_agua_superficial_u": "GAIA_INPUTS_QAG",
    "qualidade_agua_superficial_m": "GAIA_INPUTS_QAG",
    "qualidade_agua_subterranea_u": "GAIA_INPUTS_QAGS",
    "qualidade_agua_subterranea_m": "GAIA_INPUTS_QAGS",
    "qualidade_sedimentos_m": "GAIA_INPUTS_QSD",
    "qualidade_sedimentos_u": "GAIA_INPUTS_QSD",
    "qualidade_ar_m": "GAIA_INPUTS_QAR",
    "qualidade_ar_u": "GAIA_INPUTS_QAR",
    "monitoramento_efluentes_u": "GAIA_INPUTS_QEF",
    "monitoramento_efluentes_m": "GAIA_INPUTS_QEF",
    "residuos_solidos": "GAIA_INPUTS_PRS",
    "monitoramento_ruidos_terrestres": "GAIA_INPUTS_NPS",
    "monitoramento_ruidos_subaquaticos": "GAIA_INPUTS_RSA",
    "acompanhamento_dragagem_u": "GAIA_ANVISA_RAD",
    "pragas_vetores": "GAIA_INPUTS_ANVISA_PEV",
    "potabilidade_agua_u": "GAIA_INPUTS_ANVISA_POA",
    "potabilidade_agua_m": "GAIA_INPUTS_ANVISA_POA",
    "monitoramento_manguezais": "GAIA_INPUTS_MAN",
    "morfodinamica_praial": "GAIA_INPUTS_MDP",
    "fauna_organismos_sentinelas": "GAIA_INPUTS_MOS",
    "encalhes_linha_costa": "GAIA_INPUTS_ELC",
    "metaoceanografico": "GAIA_INPUTS_PMM",
    "pesca_artesanal": "GAIA_INPUTS_MPA",
    "comunicacao_social": "GAIA_INPUTS_PCS",
    "educacao_ambiental": "GAIA_INPUTS_PEA",
    "fauna_bioacumulacao": "GAIA_INPUTS_BIO",
    "fauna_bioaquatica": "GAIA_INPUTS_BAQ",
    "fauna_cetaceos": "GAIA_INPUTS_CET",
    "fauna_bentos": "GAIA_INPUTS_BEN",
    "fauna_aves_aquaticas": "GAIA_INPUTS_AVA",
    "fauna_quelonios": "GAIA_INPUTS_QLN",
    "fauna_peixes_teleostosos": "GAIA_INPUTS_PTA",
}

# Defina a chave que deseja usar
for chave in labels.keys():
    print(f"Processando chave: {chave}")
    # Caminhos dos arquivos
    docx_template_path = (
        f"C:/Users/ecpro_dhl3wmn/relatorio-ec-infra/relatorios/{chave}.docx"
    )

    # Nome da aba no Excel
    sheet_name = labels[chave]

    if sheet_name == "GAIA_INPUTS_QEF":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=2)

    elif sheet_name == "GAIA_INPUTS_RSA":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_ANVISA_RAD":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_INPUTS_ANVISA_POA":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_INPUTS_MAN":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_INPUTS_MDP":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_INPUTS_MOS":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name == "GAIA_INPUTS_ELC":
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    elif sheet_name in [
        "GAIA_INPUTS_PMM",
        "GAIA_INPUTS_MPA",
        "GAIA_INPUTS_PCS",
        "GAIA_INPUTS_PEA",
        "GAIA_INPUTS_BIO",
        "GAIA_INPUTS_BAQ",
        "GAIA_INPUTS_CET",
        "GAIA_INPUTS_BEN",
        "GAIA_INPUTS_AVA",
        "GAIA_INPUTS_QLN",
        "GAIA_INPUTS_PTA",
    ]:
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name, header=1)

    else:
        df = pd.read_excel("inputs.xlsx", sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    print(f"Colunas encontradas na aba {sheet_name}: {df.columns.tolist()}")

    # Constrói o dicionário para o template
    df_dict = {row["Código"]: row["Descrição"] for _, row in df.iterrows()}

    # Carrega e renderiza o template
    document = DocxTemplate(docx_template_path)
    document.render(df_dict)
    document.save(
        f"C:/Users/ecpro_dhl3wmn/relatorio-ec-infra/relatorios/{chave}_preenchido.docx"
    )
