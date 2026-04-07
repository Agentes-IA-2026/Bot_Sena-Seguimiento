"""
core.py
Bot madre híbrido de GuíaBot.
Lee la configuración de cada docente desde Supabase y decide
automáticamente si usar Drive o Classroom para verificar entregas.

Un solo código — comportamiento diferente por docente.
"""

import os
from datetime import datetime
from supabase import create_client, Client
from bot import drive_adapter, classroom_adapter
from bot.document_analyzer import extraer_desde_bytes, extraer_desde_lista_manual


# ── CONEXIÓN SUPABASE ─────────────────────────────────────────────────────────

SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")


def _supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)


# ── CARGAR CONFIGURACIÓN DEL DOCENTE ─────────────────────────────────────────

def cargar_config_docente(docente_id: str) -> dict | None:
    """
    Carga toda la configuración del docente desde Supabase.
    Retorna None si el docente no existe o está deshabilitado.

    Estructura retornada:
    {
        "id": "uuid",
        "nombre": "Carlos Pérez",
        "email": "docente@gmail.com",
        "fuente": "drive" | "classroom",
        "plan": "estandar" | "avanzado" | "institucion",
        "activo": True,
        "token_google": { access_token, refresh_token, client_id, client_secret },
        "materias": [...],
        "estudiantes": [...],
        "guias": [...],
    }
    """
    db = _supabase()

    docente = db.table("docentes")\
        .select("*")\
        .eq("id", docente_id)\
        .eq("activo", True)\
        .single()\
        .execute()

    if not docente.data:
        return None

    config = docente.data

    # Cargar materias del docente
    materias = db.table("materias")\
        .select("*")\
        .eq("docente_id", docente_id)\
        .execute()
    config["materias"] = materias.data or []

    # Cargar estudiantes
    estudiantes = db.table("estudiantes")\
        .select("*")\
        .eq("docente_id", docente_id)\
        .execute()
    config["estudiantes"] = estudiantes.data or []

    # Cargar guías activas
    guias = db.table("guias")\
        .select("*")\
        .eq("docente_id", docente_id)\
        .eq("activa", True)\
        .execute()
    config["guias"] = guias.data or []

    return config


# ── OBTENER EVIDENCIAS DE UNA GUÍA ───────────────────────────────────────────

def obtener_evidencias(guia: dict) -> list:
    """
    Obtiene la lista de evidencias requeridas para una guía.
    Primero intenta desde el PDF guardado en Supabase Storage.
    Si no puede, usa la lista manual guardada en la tabla guias.
    """
    # Intento 1: PDF en Supabase Storage
    if guia.get("pdf_url"):
        try:
            import httpx
            respuesta = httpx.get(guia["pdf_url"], timeout=15)
            if respuesta.status_code == 200:
                evidencias = extraer_desde_bytes(respuesta.content)
                if evidencias:
                    return evidencias
        except Exception as e:
            print(f"Error descargando PDF: {e}")

    # Intento 2: Lista manual guardada en Supabase
    if guia.get("evidencias_manuales"):
        return guia["evidencias_manuales"]

    # Intento 3: Lista manual hardcodeada en código
    return extraer_desde_lista_manual(guia.get("nombre", ""))


# ── VERIFICAR UN ESTUDIANTE ───────────────────────────────────────────────────

def verificar_estudiante(config: dict, estudiante: dict, guia: dict,
                         evidencias: list) -> dict:
    """
    Verifica las entregas de un estudiante en una guía.
    Decide automáticamente si usar Drive o Classroom
    según la configuración del docente.

    Retorna:
        {
            "estudiante": "María Rodríguez",
            "guia": "Guía 01 - Diagnóstico",
            "evidencias": {
                "Evidencia_01.pdf": True,
                "Evidencia_02.xlsx": False,
            },
            "entregadas": 1,
            "total": 2,
            "porcentaje": 50,
            "fuente": "drive" | "classroom",
        }
    """
    fuente    = config.get("fuente", "drive")
    token     = config.get("token_google", {})
    resultado = {ev: False for ev in evidencias}

    try:
        if fuente == "drive":
            # Verificar en Google Drive
            folder_id = drive_adapter.extraer_id_carpeta(
                estudiante.get("link_drive", "")
            )
            if folder_id:
                resultado = drive_adapter.verificar(token, folder_id, evidencias)

        elif fuente == "classroom":
            # Verificar en Google Classroom
            course_id      = guia.get("classroom_course_id", "")
            coursework_ids = guia.get("classroom_coursework_ids", [])
            student_id     = estudiante.get("classroom_id", "")

            if course_id and coursework_ids and student_id:
                estados = classroom_adapter.verificar(
                    token, course_id, coursework_ids, [student_id]
                )
                # Convertir estados de Classroom a True/False
                estados_estudiante = estados.get(student_id, {})
                for i, ev in enumerate(evidencias):
                    if i < len(coursework_ids):
                        estado = estados_estudiante.get(coursework_ids[i], "CREATED")
                        resultado[ev] = classroom_adapter.estado_a_bool(estado)

    except Exception as e:
        print(f"Error verificando {estudiante.get('nombre', '')}: {e}")

    entregadas = sum(resultado.values())
    total      = len(evidencias)
    porcentaje = round((entregadas / total * 100) if total > 0 else 0)

    return {
        "estudiante" : estudiante.get("nombre", ""),
        "guia"       : guia.get("nombre", ""),
        "evidencias" : resultado,
        "entregadas" : entregadas,
        "total"      : total,
        "porcentaje" : porcentaje,
        "fuente"     : fuente,
    }


# ── AUDITORÍA COMPLETA ────────────────────────────────────────────────────────

def auditar(docente_id: str, guia_id: str = None,
            estudiante_id: str = None) -> dict:
    """
    Punto de entrada principal del bot madre.
    Audita un docente completo o filtra por guía y/o estudiante.

    Parámetros opcionales:
        guia_id       → auditar solo esa guía
        estudiante_id → auditar solo ese estudiante

    Retorna:
        {
            "docente": "Carlos Pérez",
            "fuente": "drive" | "classroom",
            "fecha": "2026-04-06 10:30",
            "resultados": [...],
            "resumen": {
                "total_estudiantes": 30,
                "entregaron": 22,
                "pendientes": 8,
                "porcentaje_general": 73,
            }
        }
    """
    config = cargar_config_docente(docente_id)
    if not config:
        return {"error": "Docente no encontrado o cuenta suspendida"}

    # Filtrar guías si se especifica una
    guias = config["guias"]
    if guia_id:
        guias = [g for g in guias if g["id"] == guia_id]

    # Filtrar estudiantes si se especifica uno
    estudiantes = config["estudiantes"]
    if estudiante_id:
        estudiantes = [e for e in estudiantes if e["id"] == estudiante_id]

    resultados   = []
    total_e      = 0
    total_ok     = 0

    for guia in guias:
        evidencias = obtener_evidencias(guia)
        if not evidencias:
            continue

        for estudiante in estudiantes:
            resultado = verificar_estudiante(
                config, estudiante, guia, evidencias
            )
            resultados.append(resultado)

            total_e  += resultado["total"]
            total_ok += resultado["entregadas"]

    # Guardar en Supabase para historial
    _guardar_historial(docente_id, resultados)

    pct_general = round((total_ok / total_e * 100) if total_e > 0 else 0)
    entregaron  = sum(1 for r in resultados if r["porcentaje"] == 100)
    pendientes  = len(estudiantes) - entregaron

    return {
        "docente"    : config.get("nombre", ""),
        "fuente"     : config.get("fuente", "drive"),
        "fecha"      : datetime.now().strftime("%Y-%m-%d %H:%M"),
        "resultados" : resultados,
        "resumen"    : {
            "total_estudiantes" : len(estudiantes),
            "entregaron"        : entregaron,
            "pendientes"        : pendientes,
            "porcentaje_general": pct_general,
        }
    }


# ── GUARDAR HISTORIAL ─────────────────────────────────────────────────────────

def _guardar_historial(docente_id: str, resultados: list):
    """
    Guarda los resultados de la auditoría en Supabase.
    Solo guarda texto — no archivos — para no consumir espacio.
    """
    try:
        db    = _supabase()
        fecha = datetime.now().isoformat()

        registros = []
        for r in resultados:
            for evidencia, estado in r["evidencias"].items():
                registros.append({
                    "docente_id"  : docente_id,
                    "estudiante"  : r["estudiante"],
                    "guia"        : r["guia"],
                    "evidencia"   : evidencia,
                    "entregado"   : estado,
                    "fecha"       : fecha,
                })

        if registros:
            db.table("verificaciones").insert(registros).execute()

    except Exception as e:
        print(f"Error guardando historial: {e}")
