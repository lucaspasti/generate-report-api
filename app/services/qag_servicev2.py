# qag_service.py

import os
import io
from datetime import datetime
import numpy as np
import pandas as pd
import requests
from pytz import timezone
from supabase import Client
from docxtpl import InlineImage, RichText
from docx.shared import Cm
from app.schemas.qag import QAGRequest, QAGResponse
from app.utils.date_utils import mes_por_extenso
from app.utils.graficos import grafico_qualidade_agua
from app.services.indicadores.indicadores_qag import indicadores_qag
from app.services.vmps.vmp_qag import vmp_qag
from app.services.abs_service import AbstractService


TEMPLATE_PATH = os.path.join(
    os.getcwd(),
    "app/services/relatorios/qualidade_agua_superficial_u.docx",
)
BUCKET = "relatorios-qag"


class QAGService(AbstractService):
    def __init__(self, supabase: Client, payload: QAGRequest):
        super().__init__(
            supabase,
            payload,
            TEMPLATE_PATH,
            BUCKET,
            "qag",
            "superficial_water_quality_form",
            QAGResponse,
            QAGRequest,
        )

    def gerar_relatorio(self) -> QAGResponse:
        # fotos = []
        # for key in (
        #     "registros_fotograficos_sondas",
        #     "registros_fotograficos_amostradores",
        #     "registros_fotograficos_caixas_termicas",
        # ):
        #     url = self.lab[key][0]
        #     resp = requests.get(url)
        #     resp.raise_for_status()
        #     buf = io.BytesIO(resp.content)
        #     fotos.append(InlineImage(self.document, buf, width=Cm(5)))
        # q_22, q_23, q_24 = fotos

        data_dt = datetime.strptime(self.data_str, "%Y-%m-%d").date()
        q_14 = self.df_resultados["Ponto"].nunique()

        lab0 = self.lab
        q_20_1 = lab0["nome_laboratorio"]
        q_20_2 = lab0["razao_social_laboratorio"]
        q_20_3 = lab0["cnpj_laboratorio"]
        q_20_4 = lab0["endereco_laboratorio"]
        q_20_5 = lab0["responsavel_tecnico"]
        q_20_6 = lab0["email"]
        q_20_7 = lab0["contato"]
        q_25 = self.configuracoes.data[0]['metodology']
        q_26 = indicadores_qag

        sel27 = self.df_resultados[[
            "Ponto", "Profundidade"] + self.parametros].to_dict(orient="records")
        self.inserir_tabela_agrupada(
            data=sel27, parametros=self.parametros, ctx_key="tabela_qag_27")

        figs29 = []
        for cls in self.df_resultados["Classe"].unique():
            figs29.extend(
                grafico_qualidade_agua(
                    self.df_resultados[self.df_resultados["Classe"] == cls],
                    self.parametros,
                    cls,
                    vmp_qag,
                )
            )
        imagens_qag29 = self._figs_to_inline_images(figs29)

        q_30 = self.df_resultados["Profundidade"].unique().tolist()

        registros = []
        for _, row in self.df_resultados.iterrows():
            cls = row["Classe"]
            limites = vmp_qag.get(cls, {})
            for p in self.parametros:
                if p not in limites:
                    continue
                try:
                    vmp = float(limites[p])
                    val = float(row.get(p, np.nan))
                    conform = val <= vmp
                except:
                    val, conform = np.nan, False
                registros.append(
                    {"Parametro": p, "Valor": val, "Conforme": conform})
        df_comparacao = pd.DataFrame(registros)
        q_31 = df_comparacao["Conforme"].mean() * 100

        self.inserir_tabela_media(
            df=self.df_resultados, parametros=self.parametros, ctx_key="QAG_32")

        pm = self.parametros_metais_pesados
        regs34 = self.df_resultados[[
            "Ponto", "Profundidade"] + pm].to_dict(orient="records")
        self.inserir_tabela_agrupada(
            data=regs34, parametros=pm, ctx_key="QAG_34")
        q_35 = self._percentual_conformidade(
            self.df_resultados, pm, vmp_qag)
        figs36 = []
        for cls in self.df_resultados["Classe"].unique():
            figs36.extend(
                grafico_qualidade_agua(
                    self.df_resultados[self.df_resultados["Classe"] == cls],
                    pm,
                    cls,
                    vmp_qag,
                )
            )
        imagens_qag_36 = self._figs_to_inline_images(figs36)
        self.inserir_tabela_media(
            df=self.df_resultados, parametros=pm, ctx_key="QAG_37")

        sol = self.solventes
        q_38 = self._percentual_conformidade(
            self.df_resultados, sol, vmp_qag)
        self.inserir_tabela_media(
            df=self.df_resultados, parametros=sol, ctx_key="QAG_39")
        regs40 = self.df_resultados[[
            "Ponto", "Profundidade"] + sol].to_dict(orient="records")
        self.inserir_tabela_agrupada(
            data=regs40, parametros=sol, ctx_key="QAG_40")

        laudo = self.form.data[0]["laudos"][0]
        q_43 = RichText()
        q_43.add(
            laudo,
            underline=True,
            color="1281F0",
            url_id=self.document.build_url_id(laudo),
        )

        tabela47 = pd.DataFrame(indicadores_qag)
        aux = df_comparacao[["Parametro", "Valor", "Conforme"]]
        tabela47 = tabela47.merge(aux, on="Parametro", how="left").dropna()
        tabela47["Resultado"] = tabela47["Conforme"].map(
            {True: "Alcançado", False: "Não Alcançado"})
        tabela47 = tabela47.drop(
            columns=["Tipo", "Programa", "Valor", "Unidade", "Conforme"])
        lista_qag47 = tabela47.to_dict(orient="records")

        figs49 = []
        for cls in self.df_resultados["Classe"].unique():
            figs49.extend(
                grafico_qualidade_agua(
                    self.df_resultados[self.df_resultados["Classe"] == cls],
                    sol,
                    cls,
                    vmp_qag,
                )
            )
        imagens_qag_49 = self._figs_to_inline_images(figs49)

        self.contexto.update(
            {
                "parametros_escolhidos": self.parametros,
                "QAG_01": self.ativo.data[0]["terminal_name"],
                "QAG_02": self.form.data[0]["campanha_de_coleta"],
                "QAG_03": data_dt.strftime("%m"),
                "QAG_04": data_dt.strftime("%Y"),
                "QAG_05": "Florianópolis",
                "QAG_06": datetime.now().day,
                "QAG_07": mes_por_extenso(data_dt.strftime("%m")),
                "QAG_08": datetime.now().year,
                "QAG_09": self.ativo.data[0]["terminal_name"],
                "QAG_10": self.ativo.data[0]["cnpj"],
                "QAG_11": self.ativo.data[0]["port_location"],
                "QAG_12": self.ativo.data[0]["terminal_name"],
                "QAG_13": self.license.data[0]["number"],
                "QAG_14": q_14,
                "QAG_15": self.license.data[0]["organ"],
                "QAG_16": self.ativo.data[0]["port_location"],
                "QAG_18": self.configuracoes.data[0]["points_location"],
                "QAG_19": self.configuracoes.data[0]["periodicity_parameters"],
                "QAG_20_1": q_20_1,
                "QAG_20_2": q_20_2,
                "QAG_20_3": q_20_3,
                "QAG_20_4": q_20_4,
                "QAG_20_5": q_20_5,
                "QAG_20_6": q_20_6,
                "QAG_20_7": q_20_7,
                # "QAG_22": q_22,
                # "QAG_23": q_23,
                # "QAG_24": q_24,
                "QAG_25": q_25,
                "QAG_26": q_26,
                "tabela_qag_27": self.contexto["tabela_qag_27"],
                "QAG_28": self.parametros,
                "QAG_29": imagens_qag29,
                "QAG_30": q_30,
                "QAG_31": q_31,
                "QAG_32": self.contexto["QAG_32"],
                "QAG_33": pm,
                "QAG_34": self.contexto["QAG_34"],
                "QAG_35": q_35,
                "QAG_36": imagens_qag_36,
                "QAG_37": self.contexto["QAG_37"],
                "QAG_38": q_38,
                "QAG_39": self.contexto["QAG_39"],
                "QAG_40": self.contexto["QAG_40"],
                "QAG_43": q_43,
                "QAG_47": lista_qag47,
                "QAG_48": sol,
                "QAG_49": imagens_qag_49,
                "QAG_54": self.periodicidade,
            }
        )

        self.render_document()
        return self.upload_and_register_report()
