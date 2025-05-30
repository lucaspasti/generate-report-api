# app/schemas/qsd.py
from uuid import UUID
from datetime import date
from pydantic import BaseModel, Field


class QSDRequest(BaseModel):
    ativo_id: UUID
    data_campanha: date
    user_id: UUID
    nome_relatorio: str = Field("Relatório de Qualidade de Sedimentos")
    descricao_relatorio: str


class QSDResponse(BaseModel):
    mensagem: str
    sucesso: bool
