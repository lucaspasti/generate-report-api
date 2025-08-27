# app/schemas/qag.py
from uuid import UUID
from datetime import date
from pydantic import BaseModel, HttpUrl, Field


class QAGRequest(BaseModel):
    ativo_id: UUID
    data_campanha: list[date]
    user_id: UUID
    nome_relatorio: str = Field(
        f"Relatório de Qualidade da Água Superficial")
    periodicidade: str
    parametros: list[str] = Field(
        default=[
            "Materiais flutuantes",
            "Óleos e graxas",
            "Substâncias que comuniquem gosto ou odor",
            "Corantes provenientes de fontes antrópicas",
            "Resíduos sólidos objetáveis",
            "Turbidez (UNT)",
            "Cor verdadeira (mg Pt/L)",
            "Sólidos dissolvidos totais (mg/L)",
            "pH (N/A)",
        ],
        description="Lista de parâmetros a serem incluídos no relatório.",
    )


class QAGResponse(BaseModel):
    mensagem: str
    sucesso: bool
    url_relatorio: HttpUrl
