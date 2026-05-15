"""
core.py
Bot madre híbrido de GuíaBot.
Lee la configuración de cada docente desde Supabase y decide
automáticamente si usar Drive o Classroom para verificar entregas.

Un solo código — comportamiento diferente por docente.
"""

import os
from datetime import datetime
from bot import drive_adapter, classroom_adapter
from bot.document_analyzer import extraer_desde_bytes, extraer_desde_lista_manual

# Fase 2: motor de auditoría enriquecida
from bot.auditor import MotorAuditoria, EstadoAuditoria, resumen as resumen_auditoria
from bot.drive_adapter import listar_archivos_con_info
from sincronizador_tabla import cargar_actividades
from supabase_client import get_supabase


# ── CONEXIÓN SUPABASE ─────────────────────────────────────────────────────────

def _supabase():
    return get_supabase()


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
            from ssl_config import get_httpx_client
            with get_httpx_client(timeout=15) as http:
                respuesta = http.get(guia["pdf_url"])
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


# ── OBTENER ACTIVIDADES (FASE 2) ──────────────────────────────────────────────

def obtener_actividades(guia: dict, programa: str = "") -> list[dict]:
    """
    Retorna actividades parametrizadas desde Supabase para una guía.
    Fallback: convierte la lista plana de obtener_evidencias() al formato dict.

    El campo guia["nombre"] se usa como filtro parcial (tolerante a mayúsculas).
    Si el programa es conocido, filtra por él también para evitar ambigüedades.
    """
    guia_nombre = guia.get("nombre", "")
    actividades = cargar_actividades(programa=programa or None, guia=guia_nombre)
    if actividades:
        return actividades

    # Fallback: convertir evidencias antiguas a formato dict mínimo
    evidencias = obtener_evidencias(guia)
    return [
        {
            "actividad_id"        : str(i),
            "guia"                : guia_nombre,
            "programa"            : programa,
            "actividad_resumen"   : ev,
            "nombre_esperado"     : ev,
            "variantes_permitidas": [],
            "tipo_archivo"        : "CUALQUIER_FORMATO",
            "obligatoria"         : True,
            "carpeta_drive"       : None,
            "regla_validacion"    : None,
        }
        for i, ev in enumerate(evidencias)
    ]


# ── VERIFICAR UN ESTUDIANTE (FASE 2) ──────────────────────────────────────────

def verificar_estudiante_v2(
    config: dict,
    estudiante: dict,
    guia: dict,
    actividades: list[dict],
    todas_actividades: list[dict] | None = None,
) -> dict:
    """
    Motor enriquecido: usa MotorAuditoria y retorna estados detallados.
    Solo soporta fuente 'drive' (Classroom no tiene metadatos de carpeta/nombre).

    todas_actividades — universo completo para detectar EVIDENCIA_CRUZADA.
    Si se omite, las actividades sin match pasan directo a FALTA.

    Retorna el mismo shape que verificar_estudiante() más el campo resultados_v2.
    """
    token     = config.get("token_google", {})
    folder_id = drive_adapter.extraer_id_carpeta(estudiante.get("link_drive", ""))

    if not folder_id:
        total = len(actividades)
        return {
            "estudiante"   : estudiante.get("nombre", ""),
            "guia"         : guia.get("nombre", ""),
            "evidencias"   : {a.get("nombre_esperado", ""): False for a in actividades},
            "resultados_v2": [],
            "entregadas"   : 0,
            "total"        : total,
            "porcentaje"   : 0,
            "fuente"       : "drive",
            "error"        : "sin_link",
        }

    try:
        archivos    = listar_archivos_con_info(token, folder_id)
        resultados  = MotorAuditoria(actividades, archivos, todas_actividades).auditar()
        stats       = resumen_auditoria(resultados)

        # Mapa bool para compatibilidad con _generar_reporte_excel y _guardar_en_supabase
        # es_entregada() devuelve True también para FORMATO_INCORRECTO, NOMBRE_DIFERENTE, etc.
        evidencias_bool = {
            r["nombre_esperado"]: es_entregada(r)
            for r in resultados
        }

        return {
            "estudiante"   : estudiante.get("nombre", ""),
            "guia"         : guia.get("nombre", ""),
            "evidencias"   : evidencias_bool,
            "resultados_v2": resultados,
            "entregadas"   : stats["ok"],
            "total"        : stats["total"],
            "porcentaje"   : stats["porcentaje"],
            "fuente"       : "drive",
        }

    except Exception as e:
        print(f"Error verificar_v2 {estudiante.get('nombre', '')}: {e}")
        total = len(actividades)
        return {
            "estudiante"   : estudiante.get("nombre", ""),
            "guia"         : guia.get("nombre", ""),
            "evidencias"   : {a.get("nombre_esperado", ""): False for a in actividades},
            "resultados_v2": [],
            "entregadas"   : 0,
            "total"        : total,
            "porcentaje"   : 0,
            "fuente"       : "drive",
            "error"        : str(e),
        }


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

    fuente   = config.get("fuente", "drive")
    programa = config.get("programa", "")

    # Cargar universo completo una sola vez (necesario para EVIDENCIA_CRUZADA)
    todas_actividades = cargar_actividades() if fuente == "drive" else []

    for guia in guias:
        # Fase 2: intentar actividades parametrizadas primero
        if fuente == "drive":
            actividades = obtener_actividades(guia, programa)
            usar_v2     = bool(actividades)
        else:
            actividades = []
            usar_v2     = False

        if not usar_v2:
            evidencias = obtener_evidencias(guia)
            if not evidencias:
                continue

        for estudiante in estudiantes:
            if usar_v2:
                resultado = verificar_estudiante_v2(
                    config, estudiante, guia, actividades, todas_actividades
                )
            else:
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
