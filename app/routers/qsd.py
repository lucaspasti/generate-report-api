from fastapi import APIRouter, Depends, HTTPException
from app.schemas.qsd import QSDRequest, QSDResponse
from app.services.qsd_service import gerar_relatorio_qsd
from app.dependencies import get_supabase

router = APIRouter()

@router.post("/", response_model=QSDResponse)
def criar_qsd(
    payload: QSDRequest,
    supabase = Depends(get_supabase),
):
    try:
        return gerar_relatorio_qsd(supabase, payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
