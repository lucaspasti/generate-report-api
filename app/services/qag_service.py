import os
import io
from datetime import datetime
from uuid import UUID
from collections import OrderedDict

import numpy as np
import pandas as pd
import requests
from pytz import timezone
from supabase import Client
from fastapi import HTTPException
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Cm
from docx.enum.table import WD_TABLE_ALIGNMENT

from app.schemas.qag import QAGRequest, QAGResponse
from app.utils.date_utils import mes_por_extenso
from app.utils.file_utils import create_signed_url
from app.utils.graficos import grafico_qualidade_agua
from app.services.indicadores import indicadores_qag
from app.services.vmps.vmp_qag import vmp_qag

# Caminho para o template DOCX
TEMPLATE_PATH = os.path.join(os.getcwd(), "app/services/relatorios/qualidade_agua_superficial_u.docx")
BUCKET = "relatorios-qag"

# Parâmetros pré-definidos (podem ser importados de outro módulo se preferir)
parametros_fisico_quimicos = [
    "Materiais flutuantes", "Óleos e graxas", "Substâncias que comuniquem gosto ou odor",
    "Corantes provenientes de fontes antrópicas", "Resíduos sólidos objetáveis", "Turbidez (UNT)",
    "Cor verdadeira (mg Pt/L)", "Sólidos dissolvidos totais (mg/L)", "pH (N/A)",
]
# --- Parâmetros pré-definidos ---
parametros_metais_pesados = [
    "Antimônio (mg/L Sb)",
    "Arsênio total (mg/L As)",
    "Bário total (mg/L Ba)",
    "Berílio total (mg/L Be)",
    "Boro total (mg/L B)",
    "Cádmio total (mg/L Cd)",
    "Chumbo total (mg/L Pb)",
    "Cobalto total (mg/L Co)",
    "Cobre dissolvido (mg/L Cu)",
    "Cromo total (mg/L Cr)",
    "Ferro dissolvido (mg/L Fe)",
    "Manganês total (mg/L Mn)",
    "Mercúrio total (mg/L Hg)",
    "Níquel total (mg/L Ni)",
    "Prata total (mg/L Ag)",
    "Selênio total (mg/L Se)",
    "Tálio total (mg/L Tl)",
    "Urânio total (mg/L U)",
    "Vanádio total (mg/L V)",
    "Zinco total (mg/L Zn)",
]

parametros_fisico_quimicos = [
    "Materiais flutuantes",
    "Óleos e graxas",
    "Substâncias que comuniquem gosto ou odor",
    "Corantes provenientes de fontes antrópicas",
    "Resíduos sólidos objetáveis",
    "Turbidez (UNT)",
    "Cor verdadeira (mg Pt/L)",
    "Sólidos dissolvidos totais (mg/L)",
    "pH (N/A)",
]

# Oxigênio
parametros_oxigenio = ["DBO 5 dias a 20°C (mg/L O2)", "OD (mg/L O2)"]

# Microbiológicos
parametros_microbiologicos = ["Coliformes termotolerantes (NMP/100mL)"]

# Nutrientes
nutrientes = [
    "Fósforo total (ambiente lêntico) (mg/L P)",
    "Fósforo total (ambiente intermediário) (mg/L P)",
    "Fósforo total (ambiente lótico) (mg/L P)",
    "Polifosfatos (mg/L P)",
    "Nitrato (mg/L N)",
    "Nitrito (mg/L N)",
    "Nitrogênio amoniacal total (pH ≤ 7,5) (mg/L N)",
    "Nitrogênio amoniacal total (7,5 < pH ≤ 8,0) (mg/L N)",
    "Nitrogênio amoniacal total (8,0 < pH ≤ 8,5) (mg/L N)",
    "Nitrogênio amoniacal total (pH > 8,5) (mg/L N)",
]

# Outros elementos dissolvidos
elementos_dissolvidos = [
    "Alumínio dissolvido (mg/L Al)",
    "Antimônio (mg/L Sb)",
    "Arsênio total (mg/L As)",
    "Bário total (mg/L Ba)",
    "Berílio total (mg/L Be)",
    "Boro total (mg/L B)",
    "Chumbo total (mg/L Pb)",
    "Cobalto total (mg/L Co)",
    "Cobre dissolvido (mg/L Cu)",
    "Cromo total (mg/L Cr)",
    "Ferro dissolvido (mg/L Fe)",
    "Fluoreto total (mg/L F)",
    "Lítio total (mg/L Li)",
    "Manganês total (mg/L Mn)",
    "Mercúrio total (mg/L Hg)",
    "Níquel total (mg/L Ni)",
    "Prata total (mg/L Ag)",
    "Selênio total (mg/L Se)",
    "Tálio total (mg/L Tl)",
    "Urânio total (mg/L U)",
    "Vanádio total (mg/L V)",
    "Zinco total (mg/L Zn)",
]

# PAHs (Hidrocarbonetos aromáticos policíclicos)
pahs = [
    "Benzo (a) antraceno (μg/L)",
    "Benzo (a) pireno (μg/L)",
    "Benzo(b) fluoranteno (μg/L)",
    "Benzo(k) fluoranteno (μg/L)",
    "Criseno (μg/L)",
    "Dibenzo (a,h) antraceno (μg/L)",
]

# Pesticidas e PCBs
pesticidas_pcbs = [
    "Acrilamida (μg/L)",
    "Alacloro (μg/L)",
    "Aldrin + Dieldrin (μg/L)",
    "Atrazina (μg/L)",
    "Carbaril (μg/L)",
    "Clordano (cis + trans) (μg/L)",
    "DDT (p,p'-DDT + p,p'-DDE + p,p'-DDD) (μg/L)",
    "Dodecacloro pentaciclodecano (μg/L)",
    "Endossulfan (a + b + sulfato) (μg/L)",
    "Endrin (μg/L)",
    "Lindano (g-HCH) (μg/L)",
    "Malation (μg/L)",
    "Metoxicloro (μg/L)",
    "Pentaclorofenol (μg/L)",
    "Toxafeno (μg/L)",
    "Trifluralina (μg/L)",
    "PCBs - Bifenilas Policloradas (μg/L)",
]

# Solventes halogenados e compostos voláteis
solventes = [
    "1,2-Dicloroetano (mg/L)",
    "1,1-Dicloroeteno (mg/L)",
    "2,4-D (μg/L)",
    "2,4-Diclorofenol (μg/L)",
    "2,4,5-T (μg/L)",
    "2,4,6 - Triclorofenol (mg/L)",
    "Tricloroeteno (mg/L)",
    "Tetracloreto de carbono (mg/L)",
    "Tetracloroeteno (mg/L)",
    "Diclorometano (mg/L)",
    "Estireno (mg/L)",
    "Etilbenzeno (μg/L)",
    "Tolueno (μg/L)",
    "Monoclorobenzeno (μg/L)",
    "2-Clorofenol (μg/L)",
]

# Outros orgânicos e surfactantes
outros_organicos = [
    "Fenóis totais (mg/L)",
    "Substâncias tensoativas que reagem com o azul de metileno (mg/L LAS)",
    "Substâncias tensoativas que reagem com o vermelho de metila (mg/L MBAS)",
]

# Íons comuns
outros_ions = [
    "Cloreto total (mg/L Cl)",
    "Cloro residual total (mg/L Cl)",
    "Sulfato total (mg/L SO4)",
    "Sulfeto (H2S não dissociado) (mg/L S)",
    "Carbono orgânico total (mg/L C)",
]


def gerar_relatorio_qag(supabase: Client, payload: QAGRequest) -> QAGResponse:
    """
    Gera, faz upload e registra no banco um relatório QAG.
    """
    # 1) Datas e chaves
    data_str = payload.data_campanha.isoformat()
    tz_br = timezone("America/Sao_Paulo")
    now_br = datetime.now(tz_br)
    object_key = f"{payload.ativo_id}/{data_str}/{now_br:%Y-%m-%d_%H-%M-%S}.docx"
    local_path = f"/tmp/{os.path.basename(object_key)}"

    # 2) Carrega template
    document = DocxTemplate(TEMPLATE_PATH)

    # 3) Consultas no Supabase
    ativo = supabase.table("ativos").select("*").eq("id", str(payload.ativo_id)).execute()
    if not ativo.data:
        raise HTTPException(404, detail="Ativo não encontrado")

    configuracoes = (
        supabase.table("configuracao_formulario_ativos")
        .select("*")
        .eq("ativo_id", str(payload.ativo_id))
        .eq("tipo_formulario", "form_qualidade_da_agua_superficial")
        .execute()
    )
    if not configuracoes.data:
        raise HTTPException(404, detail="Configuração do formulário não encontrada")

    form_qag = (
        supabase.table("form_qualidade_da_agua_superficial")
        .select("*")
        .eq("ativo_id", str(payload.ativo_id))
        .eq("campanha_de_coleta", data_str)
        .execute()
    )
    if not form_qag.data:
        raise HTTPException(404, detail="Campanha não encontrada")

    # 4) DataFrame de resultados
    resultados = form_qag.data[0]["resultados"]
    df = (
        pd.DataFrame(resultados)
        .fillna("Indisponível")
        .replace({"": pd.NA, "Indisponível": pd.NA})
    )

    # Parâmetros escolhidos
    parametros = configuracoes.data[0].get("parametros_escolhidos", parametros_fisico_quimicos)

    # 5) Montagem de imagens InlineImage
    lab = form_qag.data[0]
    fotos = []
    for key in [
        "registros_fotograficos_sondas",
        "registros_fotograficos_amostradores",
        "registros_fotograficos_caixas_termicas",
    ]:
        url = lab[key][0]
        resp = requests.get(url)
        resp.raise_for_status()
        buf = io.BytesIO(resp.content)
        fotos.append(InlineImage(document, buf, width=Cm(5)))
    q_22, q_23, q_24 = fotos

    # 6) Gráficos
    selecionados = df[["Ponto","Classe","Profundidade","Tipo de análise"] + parametros].copy()
    all_figs = []
    for classe in selecionados["Classe"].unique():
        figs = grafico_qualidade_agua(
            selecionados[selecionados["Classe"]==classe],
            parametros, classe, vmp_qag
        )
        all_figs.extend(figs)
    imagens_qag29 = []
    for fig in all_figs:
        buf = io.BytesIO()
        fig.savefig(buf, format="PNG", dpi=120, bbox_inches="tight")
        buf.seek(0)
        imagens_qag29.append(InlineImage(document, buf, width=Cm(12), height=Cm(6)))

    # 7) Contexto e renderização
    contexto = {
        "QAG_01": ativo.data[0]["nome"],
        "QAG_02": data_str,
        # ... preencha todos os campos do contexto conforme o template
        "QAG_22": q_22,
        "QAG_23": q_23,
        "QAG_24": q_24,
        "QAG_29": imagens_qag29,
        # etc.
    }
    document.render(contexto)
    document.save(local_path)

    # 8) Upload no Storage
    with open(local_path, "rb") as f:
        data = f.read()
    upload = supabase.storage.from_(BUCKET).upload(object_key, data, {"contentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"})

    # 9) pegar url publica
    public_url = supabase.storage.from_(BUCKET).get_public_url(object_key)
    
    # 10) Insere registro na tabela `relatorios`
    try: 
        supabase.table("relatorios").insert({
            "nome_relatorio": payload.nome_relatorio,
            "descricao_relatorio": payload.descricao_relatorio,
            "ativo_id": str(payload.ativo_id),
            "user_id": str(payload.user_id),
            "tipo_relatorio": "qag",
            "url_relatorio": public_url,
        }).execute()

    except Exception as e:
        return QAGResponse(
            sucesso=False,
            mensagem=f"Erro ao registrar o relatório: {str(e)}"
        )

    finally:
        # nenhum erro: devolvemos sucesso
        return QAGResponse(
            sucesso=True,
            mensagem="Relatório gerado e registrado com sucesso."
        )

