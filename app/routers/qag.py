from fastapi import APIRouter, Depends, HTTPException
from app.schemas.qag import QAGRequest, QAGResponse
from app.services.qag_service import gerar_relatorio_qag
from app.dependencies import get_supabase

router = APIRouter()

@router.post("/", response_model=QAGResponse)
def criar_qag(
    payload: QAGRequest,
    supabase = Depends(get_supabase),
):
    try:
        return gerar_relatorio_qag(supabase, payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
