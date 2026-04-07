"""
classroom_adapter.py
Adaptador Google Classroom para GuíaBot.
Más simple que Drive porque Classroom ya organiza las entregas
por tarea — no hay que buscar archivos por nombre.
"""

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials


# ── CONEXIÓN ──────────────────────────────────────────────────────────────────

def conectar(token_dict: dict):
    """
    Conecta a Classroom usando el token OAuth2 del docente.
    Mismo token que Drive — Google unifica la autorización.
    """
    creds = Credentials(
        token=token_dict["access_token"],
        refresh_token=token_dict.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=token_dict["client_id"],
        client_secret=token_dict["client_secret"],
    )
    return build('classroom', 'v1', credentials=creds)


# ── LISTAR CURSOS ─────────────────────────────────────────────────────────────

def listar_cursos(token_dict: dict) -> list:
    """
    Retorna todos los cursos activos del docente en Classroom.
    Útil para la configuración inicial cuando el docente se registra.

    Retorna lista de dicts:
        [{ "id": "...", "nombre": "Matemáticas 10A", "seccion": "..." }]
    """
    try:
        service = conectar(token_dict)
        resultado = service.courses().list(
            courseStates=['ACTIVE']
        ).execute()

        cursos = []
        for c in resultado.get('courses', []):
            cursos.append({
                "id"      : c['id'],
                "nombre"  : c['name'],
                "seccion" : c.get('section', ''),
                "sala"    : c.get('room', ''),
            })
        return cursos

    except Exception as e:
        print(f"Error listar cursos Classroom: {e}")
        return []


# ── LISTAR TAREAS DEL CURSO ───────────────────────────────────────────────────

def listar_tareas(token_dict: dict, course_id: str) -> list:
    """
    Retorna todas las tareas (coursework) de un curso.
    Útil para que el docente vincule sus guías con tareas de Classroom.

    Retorna lista de dicts:
        [{ "id": "...", "titulo": "Guía 01 - Diagnóstico", "estado": "PUBLISHED" }]
    """
    try:
        service = conectar(token_dict)
        resultado = service.courses().courseWork().list(
            courseId=course_id
        ).execute()

        tareas = []
        for t in resultado.get('courseWork', []):
            tareas.append({
                "id"     : t['id'],
                "titulo" : t['title'],
                "estado" : t.get('state', ''),
                "fecha"  : t.get('dueDate', {}),
            })
        return tareas

    except Exception as e:
        print(f"Error listar tareas Classroom: {e}")
        return []


# ── VERIFICACIÓN PRINCIPAL ────────────────────────────────────────────────────

def verificar(token_dict: dict, course_id: str, coursework_ids: list,
              lista_estudiantes: list) -> dict:
    """
    Verifica el estado de entrega de cada estudiante en cada tarea.

    A diferencia de Drive, Classroom ya sabe quién entregó — no hay
    que buscar archivos por nombre. Solo se consulta el estado.

    Parámetros:
        token_dict        → credenciales OAuth2 del docente (de Supabase)
        course_id         → ID del curso en Classroom
        coursework_ids    → lista de IDs de tareas (una por guía/evidencia)
        lista_estudiantes → lista de emails o IDs de estudiantes

    Retorna:
        dict {
            "estudiante@email.com": {
                "tarea_id_1": "TURNED_IN",   ← entregó
                "tarea_id_2": "CREATED",     ← asignada pero no entregó
                "tarea_id_3": "RETURNED",    ← devuelta por el docente
            }
        }

    Estados posibles de Classroom:
        TURNED_IN  → El estudiante entregó ✅
        CREATED    → Asignada, no entregó ❌
        RETURNED   → El docente la devolvió para corregir ⚠️
        RECLAIMED  → El estudiante reclamó su entrega
    """
    resultados = {est: {} for est in lista_estudiantes}

    try:
        service = conectar(token_dict)

        for coursework_id in coursework_ids:
            page_token = None
            while True:
                entregas = service.courses().courseWork().studentSubmissions().list(
                    courseId=course_id,
                    courseWorkId=coursework_id,
                    pageToken=page_token
                ).execute()

                for entrega in entregas.get('studentSubmissions', []):
                    student_id = entrega.get('userId', '')
                    estado     = entrega.get('state', 'CREATED')

                    # Buscar el estudiante en nuestra lista
                    for est in lista_estudiantes:
                        if est == student_id or est in str(entrega):
                            if est not in resultados:
                                resultados[est] = {}
                            resultados[est][coursework_id] = estado
                            break

                page_token = entregas.get('nextPageToken')
                if not page_token:
                    break

    except Exception as e:
        print(f"Error Classroom verificar: {e}")

    return resultados


def estado_a_bool(estado: str) -> bool:
    """
    Convierte el estado de Classroom a True/False para el reporte.
    TURNED_IN y RETURNED se cuentan como entregado.
    """
    return estado in ['TURNED_IN', 'RETURNED']


# ── LISTAR ESTUDIANTES DEL CURSO ──────────────────────────────────────────────

def listar_estudiantes(token_dict: dict, course_id: str) -> list:
    """
    Retorna todos los estudiantes inscritos en un curso.
    Útil para sincronizar con Supabase en la configuración inicial.

    Retorna lista de dicts:
        [{ "id": "...", "nombre": "María Rodríguez", "email": "..." }]
    """
    try:
        service = conectar(token_dict)
        resultado = service.courses().students().list(
            courseId=course_id
        ).execute()

        estudiantes = []
        for s in resultado.get('students', []):
            perfil = s.get('profile', {})
            nombre = perfil.get('name', {})
            estudiantes.append({
                "id"     : s['userId'],
                "nombre" : nombre.get('fullName', ''),
                "email"  : perfil.get('emailAddress', ''),
            })
        return estudiantes

    except Exception as e:
        print(f"Error listar estudiantes Classroom: {e}")
        return []
