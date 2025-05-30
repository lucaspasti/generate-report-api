# app/schemas/qsd.py
from uuid import UUID
from datetime import date
from pydantic import BaseModel, Field


class QAGSRequest(BaseModel):
    ativo_id: UUID
    data_campanha: date
    user_id: UUID
    nome_relatorio: str = Field("Relatório de Qualidade da Água Subterrânea")
    descricao_relatorio: str
    periodicidade: str


class QAGSResponse(BaseModel):
    mensagem: str
    sucesso: bool
