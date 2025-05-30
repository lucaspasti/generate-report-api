import os
import io
from uuid import UUID
from collections import OrderedDict
import datetime
from datetime import datetime
import numpy as np
import pandas as pd
import requests
from pytz import timezone
from supabase import Client
from fastapi import HTTPException
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
import tempfile
from app.schemas.qsd import QSDRequest, QSDResponse
from app.utils.date_utils import mes_por_extenso
from app.utils.file_utils import create_signed_url
from app.utils.graficos import gera_distribuicao_granulometrica_qsd, graficos_linha_com_vmp_por_classe_qsd
from app.services.indicadores.indicadores_qsd import indicadores_qsd
from app.services.vmps.vmp_qsd import vmp_qsd

# Caminho para o template DOCX
TEMPLATE_PATH = os.path.join(
    os.getcwd(), "app/services/relatorios/qualidade_sedimentos_u.docx"
)
BUCKET = "relatorios-qsd"

# Parâmetros pré-definidos (podem ser importados de outro módulo se preferir)


def gerar_relatorio_qsd(supabase: Client, payload: QSDRequest) -> QSDResponse:
    """
    Gera, faz upload e registra no banco um relatório QSD.
    """
    # 1) Datas e chaves
    periodicidade = payload.periodicidade
    data_str = payload.data_campanha.isoformat()
    tz_br = timezone("America/Sao_Paulo")
    now_br = datetime.now(tz_br)
    object_key = f"{payload.ativo_id}/{data_str}/{now_br:%Y-%m-%d_%H-%M-%S}.docx"
    # use o tempdir correto do sistema
    tmp_dir = tempfile.gettempdir()
    local_path = os.path.join(tmp_dir, os.path.basename(object_key))

    # 2) Carrega template
    document = DocxTemplate(TEMPLATE_PATH)

    # 3) Consultas no Supabase
    ativo = (
        supabase.table("ativos").select(
            "*").eq("id", str(payload.ativo_id)).execute()
    )
    if not ativo.data:
        raise HTTPException(404, detail="Ativo não encontrado")

    configuracoes = (
        supabase.table("configuracao_formulario_ativos")
        .select("*")
        .eq("ativo_id", str(payload.ativo_id))
        .eq("tipo_formulario", "form_qualidade_de_sedimentos")
        .execute()
    )
    if not configuracoes.data:
        raise HTTPException(
            404, detail="Configuração do formulário não encontrada")

    form_qsd = (
        supabase.table("form_qualidade_de_sedimentos")
        .select("*")
        .eq("ativo_id", str(payload.ativo_id))
        .eq("campanha_de_coleta", data_str)
        .execute()
    )
    if not form_qsd.data:
        raise HTTPException(404, detail="Campanha não encontrada")

    # 4) DataFrame de resultados
    resultados = form_qsd.data[0]["resultados"]

    df_resultados = pd.DataFrame(resultados).fillna("Indisponível")

    data_dt = datetime.strptime(data_str, "%Y-%m-%d").date()

    lab = form_qsd.data[0]

    q_14 = df_resultados["Ponto"].nunique()

    url = form_qsd.data[0]["registro_fotografico_fundeio_amostra_de_sedimentos"][0]

    # 2) Baixe o conteúdo
    resp = requests.get(url)
    resp.raise_for_status()  # levanta exceção se status != 200
    # reusar o mesmo buffer
    buf = io.BytesIO(resp.content)
    # ajuste a largura que quiser
    q_21 = InlineImage(document, buf, width=Cm(5))

    q_20_1 = form_qsd.data[0]["nome_laboratorio"]
    q_20_2 = form_qsd.data[0]["razao_social_laboratorio"]
    q_20_3 = form_qsd.data[0]["cnpj_laboratorio"]
    q_20_4 = form_qsd.data[0]["endereco_laboratorio"]
    q_20_5 = form_qsd.data[0]["responsavel_tecnico"]
    q_20_6 = form_qsd.data[0]["email"]
    q_20_7 = form_qsd.data[0]["contato"]

    exclude = ["Ponto", "Classe", "Tipo de Análise",
               "Toxicidade", "Grupo de Análise"]

    for col in df_resultados.columns:
        if col in exclude:
            continue

        df_resultados[col] = pd.to_numeric(df_resultados[col], errors="coerce")
        non_na = df_resultados[col].dropna()
        if not non_na.empty and (non_na % 1 == 0).all():
            df_resultados[col] = df_resultados[col].astype("Int64")

    url = form_qsd.data[0]["registro_fotografico_equipamento_de_transporte"][0]

    # 2) Baixe o conteúdo
    resp = requests.get(url)
    resp.raise_for_status()  # levanta exceção se status != 200

    # reusar o mesmo buffer
    buf = io.BytesIO(resp.content)
    # ajuste a largura que quiser
    q_25 = InlineImage(document, buf, width=Cm(5))

    tabela_qsd28 = indicadores_qsd

    grafico_qsd29 = gera_distribuicao_granulometrica_qsd(
        df_resultados[
            [
                "Areia muito grossa (%)",
                "Areia grossa (%)",
                "Areia média (%)",
                "Areia fina (%)",
                "Areia muito fina (%)",
                "Silte (%)",
                "Argila (%)",
            ]
        ]
    )
    buf = io.BytesIO()
    grafico_qsd29.savefig(buf, format="PNG", dpi=120, bbox_inches="tight")
    buf.seek(0)

    q_30 = sum(df_resultados["Areia muito grossa (%)"]) / len(resultados)
    q_31 = sum(df_resultados["Areia grossa (%)"]) / len(resultados)
    q_32 = sum(df_resultados["Areia média (%)"]) / len(resultados)
    q_33 = sum(df_resultados["Areia fina (%)"]) / len(resultados)
    q_34 = sum(df_resultados["Areia muito fina (%)"]) / len(resultados)
    q_35 = sum(df_resultados["Silte (%)"]) / len(resultados)
    q_36 = sum(df_resultados["Argila (%)"]) / len(resultados)

    registros_37 = []

    for _, row in df_resultados.iterrows():
        classe = row["Classe"]
        # pega o sub-dicionário de parâmetros→limites para esta classe (ou {} se não existir)
        limites = vmp_qsd.get(classe, {})

        for parametro, vmp in limites.items():
            valor = row.get(parametro, None)
            registros_37.append(
                {
                    "Ponto": row["Ponto"],
                    "Classe": classe,
                    "Parametro": parametro,
                    "Valor": valor,
                    "VMP": vmp,
                    "Conforme": None if valor is None else (valor <= vmp),
                }
            )

    # monta o DataFrame final
    df_comparacao = pd.DataFrame(registros_37)

    qsd_37 = df_comparacao["Conforme"].mean() * 100

    metais_pesados = [
        "Arsênio (mg/kg)",
        "Cadmio (mg/kg)",
        "Chumbo (mg/kg)",
        "Cobre (mg/kg)",
        "Cromo (mg/kg)",
        "Mercúrio (mg/kg)",
        "Níquel (mg/kg)",
        "Zinco (mg/kg)",
    ]

    # 2) Filtra, remove colunas estáticas e marca “Inconforme!”
    tabela_qsd38 = df_comparacao[df_comparacao["Parametro"].isin(metais_pesados)].drop(
        columns=["Classe", "VMP"]
    )
    mask_nc_38 = df_comparacao["Conforme"] == False
    tabela_qsd38.loc[mask_nc_38, "Valor"] = tabela_qsd38.loc[
        mask_nc_38, "Valor"
    ].astype(str)

    # 3) Pivot para formato wide
    tabela_qsd38 = tabela_qsd38.pivot(
        index="Ponto", columns="Parametro", values="Valor"
    ).reset_index()

    # 4) Renomeia colunas para nomes Python-friendly
    mapping38 = {
        "Arsênio (mg/kg)": "Arsenio",
        "Cadmio (mg/kg)": "Cadmio",
        "Chumbo (mg/kg)": "Chumbo",
        "Cobre (mg/kg)": "Cobre",
        "Cromo (mg/kg)": "Cromo",
        "Mercúrio (mg/kg)": "Mercurio",
        "Níquel (mg/kg)": "Niquel",
        "Zinco (mg/kg)": "Zinco",
    }
    tabela_qsd38.rename(columns=mapping38, inplace=True)

    tabela_qsd38 = tabela_qsd38.to_dict(orient="records")

    pesticidas_organoclorados = [
        "Tributilestanho (μg/kg)",
        "HCH (Alfa HCH) (μg/kg)",
        "HCH (Beta HCH) (μg/kg)",
        "HCH (Delta HCH) (μg/kg)",
        "HCH (Gama HCH/lindano) (μg/kg)",
        "Clordano (Alfa) (μg/kg)",
        "Clordano (Gama) (μg/kg)",
        "DDD (μg/kg)",
        "DDE (μg/kg)",
        "DDT (μg/kg)",
        "Dieldrin (μg/kg)",
        "Endrin (μg/kg)",
        "Bifenilas Policloradas (μg/kg)",
    ]
    mapping = {
        "Tributilestanho (μg/kg)": "Tributilestanho",
        "HCH (Alfa HCH) (μg/kg)": "HCH_Alfa_HCH",
        "HCH (Beta HCH) (μg/kg)": "HCH_Beta_HCH",
        "HCH (Delta HCH) (μg/kg)": "HCH_Delta_HCH",
        "HCH (Gama HCH/lindano) (μg/kg)": "HCH_Gama_HCH_lindano",
        "Clordano (Alfa) (μg/kg)": "Clordano_Alfa",
        "Clordano (Gama) (μg/kg)": "Clordano_Gama",
        "DDD (μg/kg)": "DDD",
        "DDE (μg/kg)": "DDE",
        "DDT (μg/kg)": "DDT",
        "Dieldrin (μg/kg)": "Dieldrin",
        "Endrin (μg/kg)": "Endrin",
        "Bifenilas Policloradas (μg/kg)": "Bifenilas_Policloradas",
    }

    tabela_qsd39 = df_comparacao[
        df_comparacao["Parametro"].isin(pesticidas_organoclorados)
    ].drop(columns=["Classe", "VMP"])
    mask_nc = df_comparacao["Conforme"] == False
    tabela_qsd39.loc[mask_nc, "Valor"] = (
        tabela_qsd39.loc[mask_nc, "Valor"].astype(str) + " Inconforme!"
    )
    for idx, row in tabela_qsd39.iterrows():
        raw = row["Valor"]
        text = str(raw)
        if "Inconforme" in text:
            # separa valor + sufixo
            rt = RichText()
            valor_sem_inconforme = raw.replace(" Inconforme!", "")
            rt.add(valor_sem_inconforme, color="FF0000")  # run vermelho
            tabela_qsd39.at[idx, "Valor"] = rt
        else:
            # mantém como string simples
            tabela_qsd39.at[idx, "Valor"] = text

    tabela_qsd39 = tabela_qsd39.pivot(
        index="Ponto", columns="Parametro", values="Valor"
    ).reset_index()

    # 2) renomeia as colunas geradas pelo pivot
    tabela_qsd39.rename(columns=mapping, inplace=True)

    tabela_qsd39 = tabela_qsd39.to_dict(orient="records")

    hpas = [
        "Benzo(a)antraceno (μg/kg)",
        "Benzo(a)pireno (μg/kg)",
        "Criseno (μg/kg)",
        "Dibenzo(a,h)antraceno (μg/kg)",
        "Acenafteno (μg/kg)",
        "Acenaftileno (μg/kg)",
        "Antraceno (μg/kg)",
        "Fenantreno (μg/kg)",
        "Fluoranteno (μg/kg)",
        "Fluoreno (μg/kg)",
        "2-Metilnaftaleno (μg/kg)",
        "Naftaleno (μg/kg)",
        "Pireno (μg/kg)",
        "Somátoria de HPAs (μg/kg)",
    ]

    tabela_qsd40 = df_comparacao[df_comparacao["Parametro"].isin(hpas)].drop(
        columns=["Classe", "VMP"]
    )
    mask_nc_40 = df_comparacao["Conforme"] == False
    tabela_qsd40.loc[mask_nc_40, "Valor"] = (
        tabela_qsd40.loc[mask_nc_40, "Valor"].astype(str) + " Inconforme!"
    )
    for idx, row in tabela_qsd40.iterrows():
        raw = row["Valor"]
        text = str(raw)
        if "Inconforme" in text:
            # separa valor + sufixo
            rt = RichText()
            valor_sem_inconforme = raw.replace(" Inconforme!", "")
            rt.add(valor_sem_inconforme, color="FF0000")  # run vermelho
            tabela_qsd40.at[idx, "Valor"] = rt
        else:
            # mantém como string simples
            tabela_qsd40.at[idx, "Valor"] = text

    # 3) Pivot para formato wide
    tabela_qsd40 = tabela_qsd40.pivot(
        index="Ponto", columns="Parametro", values="Valor"
    ).reset_index()

    # 4) Renomeia colunas para nomes Python-friendly
    mapping40 = {
        "Benzo(a)antraceno (μg/kg)": "Benzo_a_antraceno",
        "Benzo(a)pireno (μg/kg)": "Benzo_a_pireno",
        "Criseno (μg/kg)": "Criseno",
        "Dibenzo(a,h)antraceno (μg/kg)": "Dibenzo_a_h_antraceno",
        "Acenafteno (μg/kg)": "Acenafteno",
        "Acenaftileno (μg/kg)": "Acenaftileno",
        "Antraceno (μg/kg)": "Antraceno",
        "Fenantreno (μg/kg)": "Fenantreno",
        "Fluoranteno (μg/kg)": "Fluoranteno",
        "Fluoreno (μg/kg)": "Fluoreno",
        "2-Metilnaftaleno (μg/kg)": "Metilnaftaleno_2",
        "Naftaleno (μg/kg)": "Naftaleno",
        "Pireno (μg/kg)": "Pireno",
        "Somátoria de HPAs (μg/kg)": "Somatoria_HPAs",
    }
    tabela_qsd40.rename(columns=mapping40, inplace=True)

    tabela_qsd40 = tabela_qsd40.to_dict(orient="records")

    nutrientes = [
        "Carbono Orgânico Total (%)",
        "Nitrogênio Kjeldahl Total (mg/kg)",
        "Fósforo Total (mg/kg)",
    ]

    mapping_41 = {
        "Carbono Orgânico Total (%)": "COT",
        "Nitrogênio Kjeldahl Total (mg/kg)": "Nitrogenio",
        "Fósforo Total (mg/kg)": "Fosforo",
    }

    tabela_qsd41 = df_comparacao[df_comparacao["Parametro"].isin(nutrientes)].drop(
        columns=["Classe", "VMP"]
    )

    mask_nc = tabela_qsd41["Conforme"] == False
    tabela_qsd41.loc[mask_nc, "Valor"] = (
        tabela_qsd41.loc[mask_nc, "Valor"].astype(str) + " Inconforme!"
    )
    for idx, row in tabela_qsd41.iterrows():
        raw = row["Valor"]
        text = str(raw)
        if "Inconforme" in text:
            # separa valor + sufixo
            rt = RichText()
            valor_sem_inconforme = raw.replace(" Inconforme!", "")
            rt.add(valor_sem_inconforme, color="FF0000")  # run vermelho
            tabela_qsd41.at[idx, "Valor"] = rt
        else:
            # mantém como string simples
            tabela_qsd41.at[idx, "Valor"] = text

    tabela_qsd41 = tabela_qsd41.pivot(
        index="Ponto", columns="Parametro", values="Valor"
    ).reset_index()

    tabela_qsd41.rename(columns=mapping_41, inplace=True)

    tabela_qsd41 = tabela_qsd41.to_dict(orient="records")

    mort_toxi = ["Mortalidade (%)", "Amônia não ionizada (mg/L)", "Toxicidade"]

    mapping_42 = {
        "Mortalidade (%)": "Mortalidade",
        "Amônia não ionizada (mg/L)": "Amonia_nao_ionizada",
        "Toxicidade": "Toxicidade",
    }

    tabela_qsd42 = df_comparacao.copy()

    tabela_qsd42 = tabela_qsd42[tabela_qsd42["Parametro"].isin(mort_toxi)].drop(
        columns=["Classe", "VMP"]
    )

    mask_nc = tabela_qsd42["Conforme"] == False
    tabela_qsd42.loc[mask_nc, "Valor"] = (
        tabela_qsd42.loc[mask_nc, "Valor"].astype(str) + " Inconforme!"
    )
    toxi_map = df_resultados.set_index("Ponto")["Toxicidade"]

    tabela_qsd42 = tabela_qsd42.pivot(
        index="Ponto", columns=["Parametro"], values="Valor"
    ).reset_index()

    tabela_qsd42["Toxicidade"] = tabela_qsd42["Ponto"].map(toxi_map)

    tabela_qsd42.rename(columns=mapping_42, inplace=True)

    tabela_qsd421 = tabela_qsd42.to_dict(orient="records")

    qsd_43 = tabela_qsd42.copy()

    toxi_map = df_resultados.set_index("Ponto")["Toxicidade"]

    qsd_43["Toxicidade"] = qsd_43["Ponto"].map(toxi_map)

    qsd_43["Toxicidade"] = qsd_43["Toxicidade"] == "Tóxico"

    # 2) Calcula a porcentagem de True (i.e. tóxicos) e multiplica por 100
    qsd_43 = qsd_43["Toxicidade"].mean() * 100

    tabela_qsd47 = indicadores_qsd

    tabela_qsd47 = pd.DataFrame(tabela_qsd47)

    tabela_qsd47_aux = df_comparacao.copy()

    tabela_qsd47 = tabela_qsd47.merge(
        tabela_qsd47_aux[["Parametro", "Valor", "Conforme"]],
        how="left",
        left_on=["Parametro"],
        right_on=["Parametro"],
    )

    tabela_qsd47 = tabela_qsd47.dropna()

    for _, row in tabela_qsd47.iterrows():
        if row["Conforme"] == True:
            tabela_qsd47.at[_, "Conforme"] = "Alcançado"
        elif row["Conforme"] == False:
            tabela_qsd47.at[_, "Conforme"] = "Não Alcançado"

    tabela_qsd47 = tabela_qsd47.drop(
        columns=["Tipo", "Programa", "Valor", "Unidade", "Parametro"]
    )
    tabela_qsd47 = tabela_qsd47.rename(columns={"Conforme": "Resultado"})

    tabela_qsd47 = tabela_qsd47.to_dict(orient="records")

    laudo = form_qsd.data[0]['laudo'][0]
    q_50 = RichText()
    q_50.add(laudo, underline=True, color="#1F74C8",
             url_id=document.build_url_id(laudo))

    dados_grafico = [
        'Carbono Orgânico Total (%)', 'Nitrogênio Kjeldahl Total (mg/kg)',  'Fósforo Total (mg/kg)']
    graficos_qsd51 = graficos_linha_com_vmp_por_classe_qsd(
        df_resultados, dados_grafico, vmp_qsd)

    # Converte os gráficos em InlineImage para o template
    imagens_qsd51 = []
    for fig in graficos_qsd51:
        buf = io.BytesIO()
        fig.savefig(buf, format="PNG", dpi=120, bbox_inches="tight")
        buf.seek(0)
        imagem = InlineImage(document, buf, width=Cm(12), height=Cm(6))
        imagens_qsd51.append(imagem)

    # 7) Contexto e renderização
    contexto = {
        "QSD_01": ativo.data[0]["nome"],
        "QSD_02": form_qsd.data[0]["campanha_de_coleta"],
        "QSD_03": data_dt.strftime("%m"),
        "QSD_04": data_dt.strftime("%Y"),
        "QSD_05": "Florianópolis",
        "QSD_06": datetime.now().day,
        "QSD_07": mes_por_extenso(data_dt.strftime("%m")),
        "QSD_08": datetime.now().year,
        "QSD_09": ativo.data[0]["nome"],
        "QSD_10": ativo.data[0]["cnpj"],
        "QSD_11": ativo.data[0]["endereco"],
        "QSD_12": ativo.data[0]["nome"],
        "QSD_13": ativo.data[0]["numero_licenca"],
        "QSD_14": q_14,
        "QSD_15": ativo.data[0]["orgao_regulador"],
        "QSD_16": ativo.data[0]["endereco"],
        # 'QSD_17': form_qsd.data[0]['localizacao_dos_pontos_de_monitoramento'],#MAPA
        "QSD_18": configuracoes.data[0][
            "localizacao_dos_pontos_de_monitoramento"
        ],  # pontos
        "QSD_19": configuracoes.data[0]["parametro_periodicidade"],  # pontos
        "QSD_20_1": q_20_1,
        "QSD_20_2": q_20_2,
        "QSD_20_3": q_20_3,
        "QSD_20_4": q_20_4,
        "QSD_20_5": q_20_5,
        "QSD_20_6": q_20_6,
        "QSD_20_7": q_20_7,
        "QSD_21": q_21,
        "QSD_22": configuracoes.data[0]["dados_laboratoriais"][0].get(
            "amostrador_de_coleta"
        ),
        "QSD_23": configuracoes.data[0]["dados_laboratoriais"][0].get(
            "equipamento_de_armazenamento"
        ),
        "QSD_24": configuracoes.data[0]["dados_laboratoriais"][0].get(
            "tipo_de_ampostragem"
        ),
        "QSD_25": q_25,
        "QSD_26": configuracoes.data[0]["dados_laboratoriais"][0].get(
            "metodologia_adotada"
        ),
        # "QSD_27": configuracoes_form[0]['tipo_de_frasco_de_armazenamento'][0], ####################################
        "QSD_28": tabela_qsd28,
        "QSD_29": grafico_qsd29,
        "QSD_30": q_30,
        "QSD_31": q_31,
        "QSD_32": q_32,
        "QSD_33": q_33,
        "QSD_34": q_34,
        "QSD_35": q_35,
        "QSD_36": q_36,
        "QSD_37": qsd_37,
        "QSD_38": tabela_qsd38,
        "QSD_39": tabela_qsd39,
        "QSD_40": tabela_qsd40,
        "QSD_41": tabela_qsd41,
        "QSD_42": tabela_qsd421,
        "QSD_43": qsd_43,
        "QSD_44": "Responsável Técnico",
        "QSD_45": "CREA",
        "QSD_46": "CTF IBAMA",
        # "QSD_44": responsavel.data['nome'],
        # "QSD_45": responsavel.data['crea'],
        # "QSD_46": responsavel.data['ctf_ibama'],
        "QSD_47": tabela_qsd47,
        "QSD_48": "",  # Variável aberta, para inclusão de texto pelo responsável técnico do relatório
        "QSD_49": "",  # Variável aberta, para inclusão de texto pelo responsável técnico do relatório
        "QSD_50": q_50,
        "QSD_51": imagens_qsd51,
        "QSD_54": periodicidade,
    }  # PERIODICADADE SELECIONADA / AO GERAR O RELATÓRIO

    document.render(contexto)
    document.save(local_path)

    # 8) Upload no Storage
    with open(local_path, "rb") as f:
        data = f.read()
    upload = supabase.storage.from_(BUCKET).upload(
        object_key,
        data,
        {
            "contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        },
    )

    # 9) pegar url publica
    public_url = supabase.storage.from_(BUCKET).get_public_url(object_key)

    # 10) Insere registro na tabela `relatorios`
    try:
        supabase.table("relatorios").insert(
            {
                "nome_relatorio": payload.nome_relatorio,
                "descricao_relatorio": payload.descricao_relatorio,
                "ativo_id": str(payload.ativo_id),
                "user_id": str(payload.user_id),
                "tipo_relatorio": "qsd",
                "url_relatorio": public_url,
            }
        ).execute()

    except Exception as e:
        return QSDResponse(
            sucesso=False, mensagem=f"Erro ao registrar o relatório: {str(e)}"
        )

    finally:
        # nenhum erro: devolvemos sucesso
        return QSDResponse(
            sucesso=True, mensagem="Relatório gerado e registrado com sucesso."
        )
