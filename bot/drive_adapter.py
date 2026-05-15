"""
drive_adapter.py
Adaptador Google Drive para GuíaBot.
Funciona con OAuth2 por docente (cada docente autoriza su propia cuenta).
Mantiene toda la lógica de comparación flexible del código original.
"""

import os
import re
import unicodedata
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials


# ── NORMALIZACIÓN ─────────────────────────────────────────────────────────────

def _normalizar(texto: str) -> str:
    """
    Normalización para matching de nombres de archivo en Drive.
    Usa NFD + strip combining marks — más robusto que una tabla manual de
    tildes, y alineado con auditor.normalizar() y _norm_guia_py/_norm_evidencia_py.
    """
    t = texto.lower()
    t = unicodedata.normalize('NFD', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'\.(pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)(\.|$)', ' ', t)
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    if 'programa' in t or 'formacion' in t:
        print(
            f"DEBUG_DRIVE_ADAPTER_MI_PROGRAMA"
            f"  raw={repr(texto)}"
            f"  norm={repr(t)}"
        )
    return t


def _palabras_clave(texto_normalizado: str) -> set:
    IGNORAR = {
        'pdf','png','jpg','xlsx','pptx','docx','doc',
        'con','del','los','las','una','uno','por',
        'que','para','como','the','and',
    }
    palabras = texto_normalizado.split()
    return {p for p in palabras if len(p) > 3 and p not in IGNORAR}


def _coincide(nombre_evidencia: str, nombre_drive: str) -> bool:
    ev_norm    = _normalizar(nombre_evidencia)
    drive_norm = _normalizar(nombre_drive)

    if ev_norm in drive_norm or drive_norm in ev_norm:
        return True

    palabras_ev    = _palabras_clave(ev_norm)
    palabras_drive = _palabras_clave(drive_norm)

    if palabras_ev and palabras_drive:
        comunes    = palabras_ev & palabras_drive
        porcentaje = len(comunes) / len(palabras_ev)
        if porcentaje >= 0.5:
            return True

    def bigramas(palabras):
        lista = sorted(palabras)
        return {f"{lista[i]} {lista[i+1]}" for i in range(len(lista)-1)}

    if len(palabras_ev) >= 2 and len(palabras_drive) >= 2:
        if bigramas(palabras_ev) & bigramas(palabras_drive):
            return True

    return False


# ── CONEXIÓN ──────────────────────────────────────────────────────────────────

def conectar(token_dict: dict):
    """
    Conecta a Drive usando el token OAuth2 del docente.
    token_dict viene de Supabase — guardado cuando el docente autorizó su cuenta.
    """
    creds = Credentials(
        token=token_dict["access_token"],
        refresh_token=token_dict.get("refresh_token"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=token_dict["client_id"],
        client_secret=token_dict["client_secret"],
    )
    return build('drive', 'v3', credentials=creds)


# ── LISTAR ARCHIVOS ───────────────────────────────────────────────────────────

def _listar_con_info(
    service,
    folder_id: str,
    carpeta_padre: str = "",
    folder_path: str = "",
) -> list[dict]:
    """
    Lista recursivamente todos los archivos de una carpeta Drive.
    Cada elemento: nombre, extension, mime_type, carpeta_padre, folder_path, drive_file_id

    Notas:
    - supportsAllDrives + includeItemsFromAllDrives para Shared Drives.
    - Google Docs/Sheets/Slides no tienen extensión: extension="" y mime_type indica su tipo.
      El campo mime_type permite que _tipo_ok los trate correctamente en auditor.py.
    """
    resultado = []
    try:
        page_token = None
        while True:
            resp = service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType)",
                pageSize=200,
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()

            for item in resp.get("files", []):
                nombre    = item["name"]
                mime_type = item["mimeType"]
                if mime_type == "application/vnd.google-apps.folder":
                    sub_path = f"{folder_path}/{nombre}".lstrip("/")
                    resultado.extend(
                        _listar_con_info(service, item["id"], nombre, sub_path)
                    )
                else:
                    _, ext = os.path.splitext(nombre)
                    resultado.append({
                        "nombre"        : nombre,
                        "extension"     : ext.lower(),
                        "mime_type"     : mime_type,
                        "carpeta_padre" : carpeta_padre,
                        "folder_path"   : folder_path,
                        "drive_file_id" : item["id"],
                    })

            page_token = resp.get("nextPageToken")
            if not page_token:
                break

    except Exception as e:
        print(f"Error al listar carpeta {folder_id}: {e}")

    return resultado


def _listar_archivos_recursivo(service, folder_id: str) -> list[str]:
    """Wrapper de compatibilidad con el código existente — retorna solo nombres."""
    return [a["nombre"] for a in _listar_con_info(service, folder_id)]


def listar_archivos_con_info(token_dict: dict, folder_id: str) -> list[dict]:
    """
    Versión pública OAuth2 de _listar_con_info.
    Usar desde bot/core.py o cualquier módulo con token OAuth2.
    """
    service = conectar(token_dict)
    return _listar_con_info(service, folder_id)


def listar_archivos_con_info_service(service, folder_id: str) -> list[dict]:
    """
    Versión pública para service account.
    Usar desde main.py donde el service ya está conectado con cuenta de servicio.
    """
    return _listar_con_info(service, folder_id)


# ── VERIFICACIÓN PRINCIPAL ────────────────────────────────────────────────────

def verificar(token_dict: dict, folder_id: str, lista_evidencias: list) -> dict:
    """
    Verifica qué evidencias están presentes en la carpeta Drive del estudiante.

    Parámetros:
        token_dict       → credenciales OAuth2 del docente (de Supabase)
        folder_id        → ID de la carpeta Drive del estudiante
        lista_evidencias → lista de nombres de evidencias requeridas

    Retorna:
        dict { nombre_evidencia: True/False }
    """
    resultados = {ev: False for ev in lista_evidencias}

    try:
        service = conectar(token_dict)
        archivos = _listar_archivos_recursivo(service, folder_id)

        if not archivos:
            return resultados

        for ev in lista_evidencias:
            for nombre_real in archivos:
                if _coincide(ev, nombre_real):
                    resultados[ev] = True
                    break

    except Exception as e:
        print(f"Error Drive verificar: {e}")

    return resultados


def extraer_id_carpeta(link: str) -> str | None:
    """Extrae el ID de Drive de cualquier formato de URL."""
    if not link or not isinstance(link, str):
        return None
    link = link.strip()
    for patron in [
        r'/folders/([a-zA-Z0-9_-]+)',
        r'/file/d/([a-zA-Z0-9_-]+)',
        r'[?&]id=([a-zA-Z0-9_-]+)',
        r'open\?id=([a-zA-Z0-9_-]+)',
    ]:
        m = re.search(patron, link)
        if m:
            return m.group(1)
    if "http" not in link and len(link) > 10:
        return link
    return None
