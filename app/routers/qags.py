from fastapi import APIRouter, Depends, HTTPException
from app.schemas.qags import QAGSRequest, QAGSResponse
from app.services.qags_service import gerar_relatorio_qags
from app.dependencies import get_supabase

router = APIRouter()

@router.post("/", response_model=QAGSResponse)
def criar_qag(
    payload: QAGSRequest,
    supabase = Depends(get_supabase),
):
    try:
        return gerar_relatorio_qags(supabase, payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
