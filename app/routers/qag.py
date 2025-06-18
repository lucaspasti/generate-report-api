# app/api/routes/reports.py

from fastapi import APIRouter, Depends, HTTPException
from app.schemas.qag import QAGRequest, QAGResponse
from app.services.qag_servicev2 import QAGService 
from app.dependencies import get_supabase

router = APIRouter()


@router.post("/", response_model=QAGResponse)
def criar_qag(
    payload: QAGRequest,
    supabase=Depends(get_supabase),
):
    service = QAGService(supabase=supabase, payload=payload)
    try:
        return service.gerar_relatorio()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
