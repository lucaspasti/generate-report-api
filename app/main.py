from fastapi import FastAPI  # type: ignore
from fastapi.middleware.cors import CORSMiddleware  # type: ignore

from app.config import settings
from app.routers import qag
from app.routers import qsd
from app.routers import qags

app = FastAPI(
    title="API EC-Infra",
    version="1.0.0",
    description="Backend para geração de relatórios QAG, QSD, QAR e afins",
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    allow_credentials=True,
)

# Routers
app.include_router(qag.router, prefix="/reports/qag", tags=["QAG"])
app.include_router(qsd.router, prefix="/reports/qsd", tags=["QSD"])
# app.include_router(qar.router, prefix="/reports/qar", tags=["QAR"])
app.include_router(qags.router, prefix="/reports/qags", tags=["QAGS"])


@app.get("/", tags=["Health"])
def health_check():
    return {"status": "ok", "message": "API EC-Infra is up and running"}
