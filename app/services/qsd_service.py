import os
import io
from datetime import datetime
import numpy as np
import pandas as pd
import requests
from docxtpl import InlineImage, RichText
from docx.shared import Cm
from app.schemas.qsd import QSDRequest, QSDResponse
from app.utils.date_utils import mes_por_extenso
from app.utils.graficos import gera_distribuicao_granulometrica_qsd, graficos_linha_com_vmp_por_classe_qsd
from app.services.indicadores.indicadores_qsd import indicadores_qsd
from app.services.vmps.vmp_qsd import vmp_qsd
from app.services.abs_service import AbstractService

TEMPLATE_PATH = os.path.join(
    os.getcwd(),
    "app/services/relatorios/qualidade_sedimentos_u.docx"
)
BUCKET = "relatorios-qsd"


class QSDService(AbstractService):
    def __init__(self, supabase, payload):
        super().__init__(
            supabase,
            payload,
            TEMPLATE_PATH,
            BUCKET,
            "qsd",
            "form_qualidade_de_sedimentos",
            QSDResponse,
            QSDRequest,
        )

    def gerar_relatorio(self) -> QSDResponse:
        # Imagens
        url_fundeio = self.lab["registro_fotografico_fundeio_amostra_de_sedimentos"][0]
        resp = requests.get(url_fundeio)
        resp.raise_for_status()
        q_21 = InlineImage(self.document, io.BytesIO(
            resp.content), width=Cm(5))

        url_equip = self.lab["registro_fotografico_equipamento_de_transporte"][0]
        resp = requests.get(url_equip)
        resp.raise_for_status()
        q_25 = InlineImage(self.document, io.BytesIO(
            resp.content), width=Cm(5))

        data_dt = self.data_str
        q_14 = self.df_resultados["Ponto"].nunique()
        q_20_1 = self.lab["nome_laboratorio"]
        q_20_2 = self.lab["razao_social_laboratorio"]
        q_20_3 = self.lab["cnpj_laboratorio"]
        q_20_4 = self.lab["endereco_laboratorio"]
        q_20_5 = self.lab["responsavel_tecnico"]
        q_20_6 = self.lab["email"]
        q_20_7 = self.lab["contato"]

        # Conversão de tipos para as colunas de resultado, exceto as exclusões
        exclude = ["Ponto", "Classe", "Tipo de Análise",
                   "Toxicidade", "Grupo de Análise"]
        for col in self.df_resultados.columns:
            if col not in exclude:
                self.df_resultados[col] = pd.to_numeric(
                    self.df_resultados[col], errors="coerce")

        # Gráfico granulometria
        grafico_qsd29 = gera_distribuicao_granulometrica_qsd(
            self.df_resultados[
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
        img_qsd29 = InlineImage(self.document, buf, width=Cm(12), height=Cm(6))

        # Médias de granulometria
        q_30 = self.df_resultados["Areia muito grossa (%)"].mean()
        q_31 = self.df_resultados["Areia grossa (%)"].mean()
        q_32 = self.df_resultados["Areia média (%)"].mean()
        q_33 = self.df_resultados["Areia fina (%)"].mean()
        q_34 = self.df_resultados["Areia muito fina (%)"].mean()
        q_35 = self.df_resultados["Silte (%)"].mean()
        q_36 = self.df_resultados["Argila (%)"].mean()

        # Conformidade geral
        registros_37 = []
        for _, row in self.df_resultados.iterrows():
            classe = row["Classe"]
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
        df_comparacao = pd.DataFrame(registros_37)
        qsd_37 = df_comparacao["Conforme"].mean() * 100

        # Tabelas de parâmetros
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
        tabela_qsd38 = df_comparacao[df_comparacao["Parametro"].isin(metais_pesados)].drop(
            columns=["Classe", "VMP"]
        )
        mask_nc_38 = tabela_qsd38["Conforme"] == False
        tabela_qsd38.loc[mask_nc_38,
                         "Valor"] = tabela_qsd38.loc[mask_nc_38, "Valor"].astype(str)
        tabela_qsd38 = tabela_qsd38.pivot(
            index="Ponto", columns="Parametro", values="Valor").reset_index()
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

        # Pesticidas
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
        mask_nc = tabela_qsd39["Conforme"] == False
        tabela_qsd39.loc[mask_nc, "Valor"] = (
            tabela_qsd39.loc[mask_nc, "Valor"].astype(str) + " Inconforme!"
        )
        for idx, row in tabela_qsd39.iterrows():
            raw = row["Valor"]
            text = str(raw)
            if "Inconforme" in text:
                rt = RichText()
                valor_sem_inconforme = raw.replace(" Inconforme!", "")
                rt.add(valor_sem_inconforme, color="FF0000")
                tabela_qsd39.at[idx, "Valor"] = rt
            else:
                tabela_qsd39.at[idx, "Valor"] = text
        tabela_qsd39 = tabela_qsd39.pivot(
            index="Ponto", columns="Parametro", values="Valor").reset_index()
        tabela_qsd39.rename(columns=mapping, inplace=True)
        tabela_qsd39 = tabela_qsd39.to_dict(orient="records")

        # HPAs
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
        tabela_qsd40 = df_comparacao[df_comparacao["Parametro"].isin(hpas)].drop(
            columns=["Classe", "VMP"]
        )
        mask_nc_40 = tabela_qsd40["Conforme"] == False
        tabela_qsd40.loc[mask_nc_40, "Valor"] = (
            tabela_qsd40.loc[mask_nc_40, "Valor"].astype(str) + " Inconforme!"
        )
        for idx, row in tabela_qsd40.iterrows():
            raw = row["Valor"]
            text = str(raw)
            if "Inconforme" in text:
                rt = RichText()
                valor_sem_inconforme = raw.replace(" Inconforme!", "")
                rt.add(valor_sem_inconforme, color="FF0000")
                tabela_qsd40.at[idx, "Valor"] = rt
            else:
                tabela_qsd40.at[idx, "Valor"] = text
        tabela_qsd40 = tabela_qsd40.pivot(
            index="Ponto", columns="Parametro", values="Valor").reset_index()
        tabela_qsd40.rename(columns=mapping40, inplace=True)
        tabela_qsd40 = tabela_qsd40.to_dict(orient="records")

        # Nutrientes
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
                rt = RichText()
                valor_sem_inconforme = raw.replace(" Inconforme!", "")
                rt.add(valor_sem_inconforme, color="FF0000")
                tabela_qsd41.at[idx, "Valor"] = rt
            else:
                tabela_qsd41.at[idx, "Valor"] = text
        tabela_qsd41 = tabela_qsd41.pivot(
            index="Ponto", columns="Parametro", values="Valor"
        ).reset_index()
        tabela_qsd41.rename(columns=mapping_41, inplace=True)
        tabela_qsd41 = tabela_qsd41.to_dict(orient="records")

        # Mortalidade/Toxicidade
        mort_toxi = [
            "Mortalidade (%)", "Amônia não ionizada (mg/L)", "Toxicidade"]
        mapping_42 = {
            "Mortalidade (%)": "Mortalidade",
            "Amônia não ionizada (mg/L)": "Amonia_nao_ionizada",
            "Toxicidade": "Toxicidade",
        }
        tabela_qsd42 = df_comparacao[df_comparacao["Parametro"].isin(mort_toxi)].drop(
            columns=["Classe", "VMP"]
        )
        mask_nc = tabela_qsd42["Conforme"] == False
        tabela_qsd42.loc[mask_nc, "Valor"] = (
            tabela_qsd42.loc[mask_nc, "Valor"].astype(str) + " Inconforme!"
        )
        toxi_map = self.df_resultados.set_index("Ponto")["Toxicidade"]
        tabela_qsd42 = tabela_qsd42.pivot(
            index="Ponto", columns=["Parametro"], values="Valor"
        ).reset_index()
        tabela_qsd42["Toxicidade"] = tabela_qsd42["Ponto"].map(toxi_map)
        tabela_qsd42.rename(columns=mapping_42, inplace=True)
        tabela_qsd421 = tabela_qsd42.to_dict(orient="records")

        # % Pontos tóxicos
        qsd_43 = tabela_qsd42.copy()
        qsd_43["Toxicidade"] = qsd_43["Ponto"].map(toxi_map)
        qsd_43["Toxicidade"] = qsd_43["Toxicidade"] == "Tóxico"
        qsd_43 = qsd_43["Toxicidade"].mean() * 100

        # Indicadores 47
        tabela_qsd47 = pd.DataFrame(indicadores_qsd)
        aux = df_comparacao[["Parametro", "Valor", "Conforme"]]
        tabela_qsd47 = tabela_qsd47.merge(
            aux, on="Parametro", how="left").dropna()
        tabela_qsd47["Resultado"] = tabela_qsd47["Conforme"].map(
            {True: "Alcançado", False: "Não Alcançado"})
        tabela_qsd47 = tabela_qsd47.drop(
            columns=["Tipo", "Programa", "Valor", "Unidade", "Conforme", "Parametro"])
        lista_qsd47 = tabela_qsd47.to_dict(orient="records")

        # Laudo (RichText com hyperlink)
        laudo = self.form.data[0]["laudo"][0]
        q_50 = RichText()
        q_50.add(
            laudo, underline=True, color="#1F74C8",
            url_id=self.document.build_url_id(laudo)
        )

        # Gráficos de nutrientes
        dados_grafico = [
            "Carbono Orgânico Total (%)", "Nitrogênio Kjeldahl Total (mg/kg)", "Fósforo Total (mg/kg)"
        ]
        try:
            graficos_qsd51 = graficos_linha_com_vmp_por_classe_qsd(
                self.df_resultados, dados_grafico, vmp_qsd
            )
        except:
            pass
        imagens_qsd51 = self._figs_to_inline_images(graficos_qsd51)

        # Preencher contexto
        self.contexto.update({
            "QSD_01": self.ativo.data[0]["nome"],
            "QSD_02": self.form.data[0]["campanha_de_coleta"],
            "QSD_03": data_dt.strftime("%m"),
            "QSD_04": data_dt.strftime("%Y"),
            "QSD_05": "Florianópolis",
            "QSD_06": datetime.now().day,
            "QSD_07": mes_por_extenso(data_dt.strftime("%m")),
            "QSD_08": datetime.now().year,
            "QSD_09": self.ativo.data[0]["nome"],
            "QSD_10": self.ativo.data[0]["cnpj"],
            "QSD_11": self.ativo.data[0]["endereco"],
            "QSD_12": self.ativo.data[0]["nome"],
            "QSD_13": self.ativo.data[0]["numero_licenca"],
            "QSD_14": q_14,
            "QSD_15": self.ativo.data[0]["orgao_regulador"],
            "QSD_16": self.ativo.data[0]["endereco"],
            "QSD_18": self.configuracoes.data[0]["localizacao_dos_pontos_de_monitoramento"],
            "QSD_19": self.configuracoes.data[0]["parametro_periodicidade"],
            "QSD_20_1": q_20_1,
            "QSD_20_2": q_20_2,
            "QSD_20_3": q_20_3,
            "QSD_20_4": q_20_4,
            "QSD_20_5": q_20_5,
            "QSD_20_6": q_20_6,
            "QSD_20_7": q_20_7,
            "QSD_21": q_21,
            "QSD_25": q_25,
            "QSD_28": indicadores_qsd,
            "QSD_29": img_qsd29,
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
            "QSD_47": lista_qsd47,
            "QSD_48": "",
            "QSD_49": "",
            "QSD_50": q_50,
            "QSD_51": imagens_qsd51,
            "QSD_54": self.periodicidade,
        })

        self.render_document()
        return self.upload_and_register_report()
