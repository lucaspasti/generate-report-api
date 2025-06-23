from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware  # type: ignore

from app.config import settings
from app.routers import qag, qsd, qags


app = FastAPI(
    title="API EC-Infra",
    version="1.0.0",
    description="Backend para geração de relatórios QAG, QSD, QAR e afins",
)

# 1) CORSMiddleware deve vir imediatamente depois de criar o app…
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "https://ec-infra.vercel.app"],
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    allow_credentials=True,
)

# 2) Só depois disso você monta arquivos estáticos ou routers
app.include_router(qag.router, prefix="/reports/qag", tags=["QAG"])
app.include_router(qsd.router, prefix="/reports/qsd", tags=["QSD"])
app.include_router(qags.router, prefix="/reports/qags", tags=["QAGS"])


@app.get("/", tags=["Health"])
def health_check():
    return {"status": "ok", "message": "API EC-Infra is up and running"}


@app.get("/_env")
def env_check():
    return {
        "status": "ok",
        "message": "API EC-Infra is up and running",
        "env": {
            "SUPABASE_URL": settings.SUPABASE_URL,
            "SUPABASE_KEY": 'OK',
            "ALLOWED_ORIGINS": settings.ALLOWED_ORIGINS,
        },
    }
