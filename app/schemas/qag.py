# app/schemas/qag.py
from uuid import UUID
from datetime import date
from pydantic import BaseModel, Field

class QAGRequest(BaseModel):
    ativo_id: UUID
    data_campanha: date
    user_id: UUID
    nome_relatorio: str = Field("Relatório de Qualidade da Água Superficial")
    descricao_relatorio: str
    periodicidade: str

class QAGResponse(BaseModel):
    mensagem: str
    sucesso: bool
