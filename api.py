"""
api.py — Servidor FastAPI de GuíaBot.

Endpoints:
  GET  /health             → estado del servidor + Drive
  GET  /api/fichas         → lista fichas disponibles (Excel en assets/Listados)
  GET  /api/guias          → guías de una ficha  (?ficha_id=3414936)
  POST /api/auditar        → lanza la auditoría y devuelve resumen

Despliegue Railway:
  uvicorn api:app --host 0.0.0.0 --port $PORT
"""

import os
import re
import glob
import logging
from contextlib import asynccontextmanager
from typing import Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# SSL antes de cualquier conexión HTTPS
from ssl_config import configure_environment
configure_environment()

from Core.drive_manager import conectar_drive
import main as _m

logging.basicConfig(level=logging.INFO, format="%(levelname)s  %(name)s — %(message)s")
logger = logging.getLogger("guiabot.api")


# ── Drive singleton: se crea una vez al arrancar el proceso ──────────────────

_drive_service = None

@asynccontextmanager
async def lifespan(app: FastAPI):
    global _drive_service
    logger.info("Conectando a Google Drive...")
    try:
        _drive_service = conectar_drive()
        logger.info("Drive: %s", "OK" if _drive_service else "FALLO (auditoría no disponible)")
    except Exception as exc:
        logger.warning("Error al conectar Drive: %s", exc)
    yield
    _drive_service = None


app = FastAPI(title="GuíaBot API", version="1.0", lifespan=lifespan)

# CORS: orígenes permitidos (Vercel + localhost dev)
_ORIGINS = [
    o.strip()
    for o in os.environ.get(
        "ALLOWED_ORIGINS",
        "https://guiabot-co.vercel.app,http://localhost:5500,http://localhost:3000,http://127.0.0.1:5500",
    ).split(",")
    if o.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=_ORIGINS,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization"],
    allow_credentials=False,
)


# ── Modelos de entrada ────────────────────────────────────────────────────────

class AuditRequest(BaseModel):
    ficha_id: str
    guia_id: Optional[str] = None


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    """Verifica que el servidor esté vivo y si Drive está disponible."""
    return {"status": "ok", "drive": bool(_drive_service)}


@app.get("/api/fichas")
def listar_fichas():
    """Lista las fichas disponibles (Excel en assets/Listados) con su programa."""
    excels = sorted(glob.glob(os.path.join(_m.CARPETA_LISTADOS, "*.xlsx")))
    fichas = []
    for ruta in excels:
        nombre = os.path.basename(ruta)
        if nombre.startswith("~$"):
            continue
        nums     = re.findall(r"\d{7}", nombre)
        ficha_id = nums[0] if nums else None
        if ficha_id:
            fichas.append({
                "ficha_id"       : ficha_id,
                "nombre_archivo" : nombre,
                "programa"       : _m.FICHA_A_PROGRAMA.get(ficha_id, "desconocido"),
            })
    return fichas


@app.get("/api/guias")
def listar_guias_ficha(ficha_id: str = Query(..., description="ID numérico de la ficha")):
    """Lista las guías disponibles para una ficha, consultando Supabase."""
    programa, _ = _m.resolver_programa_ficha(ficha_id)
    if not programa:
        raise HTTPException(404, f"Ficha {ficha_id} no encontrada o sin programa asignado")
    guias = _m.obtener_guias_por_programa(programa)
    return {"ficha_id": ficha_id, "programa": programa, "guias": guias}


@app.post("/api/auditar")
def api_auditar(req: AuditRequest):
    """
    Lanza la auditoría completa (o de una guía) para la ficha indicada.
    Bloquea hasta que termina y devuelve el resumen.
    """
    if not _drive_service:
        raise HTTPException(503, "Google Drive no disponible — reinicia el servidor")

    logger.info("Auditoría iniciada: ficha=%s guia=%s", req.ficha_id, req.guia_id)

    try:
        resultado = _m.ejecutar_auditoria(
            ficha_id = req.ficha_id,
            guia_id  = req.guia_id,
            service  = _drive_service,
        )
    except Exception as exc:
        logger.exception("Error inesperado en auditoría ficha=%s", req.ficha_id)
        raise HTTPException(500, f"Error interno: {exc}")

    if resultado.get("status") == "error":
        raise HTTPException(422, resultado.get("mensaje", "Error desconocido"))

    logger.info(
        "Auditoría OK: ficha=%s %d/%d entregadas en %.1fs",
        resultado["ficha"], resultado["entregadas"],
        resultado["total"], resultado["duracion_segundos"],
    )
    return resultado
