"""
drive_adapter.py
Adaptador Google Drive para GuíaBot.
Funciona con OAuth2 por docente (cada docente autoriza su propia cuenta).
Mantiene toda la lógica de comparación flexible del código original.
"""

import os
import re
import unicodedata
from urllib.parse import unquote, urlparse, parse_qs
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
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

FOLDER_MIME = "application/vnd.google-apps.folder"
SHORTCUT_MIME = "application/vnd.google-apps.shortcut"


class DriveAccessError(RuntimeError):
    pass


def _diagnosticar_error_drive(exc: Exception) -> str:
    if isinstance(exc, HttpError):
        status = getattr(exc.resp, "status", None)
        reason = ""
        try:
            reason = getattr(exc.resp, "reason", "") or ""
        except Exception:
            reason = ""
        if status == 404:
            return (
                "folder inexistente o sin permiso para la cuenta autenticada "
                "(Google Drive devuelve 404 en ambos casos)"
            )
        if status == 403:
            return "acceso denegado para la cuenta autenticada"
        return f"error HTTP Drive {status or '?'} {reason}".strip()
    return type(exc).__name__


def _drive_get_metadata(service, file_id: str) -> dict:
    return service.files().get(
        fileId=file_id,
        fields="id, name, mimeType, shortcutDetails",
        supportsAllDrives=True,
    ).execute()


def _resolver_folder_raiz(service, folder_id: str, debug: bool = False) -> tuple[str, dict]:
    meta = _drive_get_metadata(service, folder_id)
    mime = meta.get("mimeType")
    if debug:
        print(f"      📁 carpeta raíz   : {meta.get('name')} [{mime}]")
    if mime == SHORTCUT_MIME:
        details = meta.get("shortcutDetails") or {}
        target_id = details.get("targetId")
        target_mime = details.get("targetMimeType")
        if debug:
            print(f"      🔀 raíz es shortcut: target_id={target_id} [{target_mime}]")
        if target_mime != FOLDER_MIME or not target_id:
            raise DriveAccessError(
                f"El folder_id {folder_id} apunta a un shortcut que no resuelve "
                f"a carpeta. targetMimeType={target_mime or '(vacío)'}"
            )
        target_meta = _drive_get_metadata(service, target_id)
        if debug:
            print(
                f"      🔀 target carpeta : {target_meta.get('name')} "
                f"[{target_meta.get('mimeType')}]"
            )
            print(f"      🔀 folder_id efectivo: {target_id}")
        return target_id, target_meta
    if mime != FOLDER_MIME:
        raise DriveAccessError(
            f"El folder_id {folder_id} no apunta a una carpeta Drive. mimeType={mime}"
        )
    return folder_id, meta


def _listar_hijos(service, folder_id: str) -> list[dict]:
    hijos = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields=(
                "nextPageToken, incompleteSearch, "
                "files(id, name, mimeType, shortcutDetails)"
            ),
            pageSize=1000,
            pageToken=page_token,
            corpora="allDrives",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        if resp.get("incompleteSearch"):
            print(f"   ⚠️ Drive devolvió búsqueda incompleta para carpeta {folder_id}")
        hijos.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return hijos


def inspeccionar_carpeta_service(service, folder_id: str) -> dict:
    """
    Diagnóstico liviano de una carpeta Drive: metadata raíz, resolución de
    shortcut si aplica, e hijos directos del folder efectivo.
    """
    try:
        folder_id_efectivo, meta = _resolver_folder_raiz(service, folder_id, debug=False)
        hijos = _listar_hijos(service, folder_id_efectivo)
    except Exception as e:
        diagnostico = _diagnosticar_error_drive(e)
        raise DriveAccessError(
            f"No se puede inspeccionar la carpeta raíz Drive {folder_id}: {diagnostico}. "
            f"Detalle: {e}"
        ) from e
    return {
        "folder_id_original": folder_id,
        "folder_id_efectivo": folder_id_efectivo,
        "metadata": meta,
        "hijos": hijos,
    }


def _es_carpeta_o_shortcut_carpeta(item: dict) -> bool:
    if item.get("mimeType") == FOLDER_MIME:
        return True
    if item.get("mimeType") != SHORTCUT_MIME:
        return False
    return item.get("shortcutDetails", {}).get("targetMimeType") == FOLDER_MIME


def _id_carpeta_real(item: dict) -> str:
    if item.get("mimeType") == SHORTCUT_MIME:
        return item.get("shortcutDetails", {}).get("targetId") or item["id"]
    return item["id"]


def _listar_con_info(
    service,
    folder_id: str,
    carpeta_padre: str = "",
    folder_path: str = "",
    debug: bool = False,
    _visitadas: set[str] | None = None,
    strict: bool = True,
) -> list[dict]:
    """
    Lista recursivamente todos los archivos de una carpeta Drive.
    Cada elemento: nombre, extension, mime_type, carpeta_padre, folder_path, drive_file_id

    Notas:
    - supportsAllDrives + includeItemsFromAllDrives para Shared Drives.
    - corpora=allDrives evita búsquedas parciales cuando el portafolio está en
      una unidad compartida.
    - Entra también a shortcuts que apuntan a carpetas. En Drive el usuario ve
      el shortcut como carpeta, pero la API lo devuelve con mimeType shortcut.
    - Google Docs/Sheets/Slides no tienen extensión: extension="" y mime_type indica su tipo.
      El campo mime_type permite que _tipo_ok los trate correctamente en auditor.py.
    """
    resultado = []
    _visitadas = _visitadas or set()
    if folder_id in _visitadas:
        if debug:
            print(f"   ⚠️ Drive: carpeta ya visitada, se omite ciclo: {folder_id}")
        return resultado
    _visitadas.add(folder_id)

    try:
        hijos = _listar_hijos(service, folder_id)
        carpetas = [h for h in hijos if _es_carpeta_o_shortcut_carpeta(h)]
        if debug:
            ruta = folder_path or "(raíz)"
            print(
                f"   📁 Drive lee /{ruta}: {len(hijos)} hijo(s), "
                f"{len(carpetas)} subcarpeta(s)"
            )
            if carpetas:
                print("   📁 Subcarpetas: " + ", ".join(c["name"] for c in carpetas[:20]))

        for item in hijos:
            nombre    = item["name"]
            mime_type = item["mimeType"]
            if _es_carpeta_o_shortcut_carpeta(item):
                sub_path = f"{folder_path}/{nombre}".lstrip("/")
                resultado.extend(
                    _listar_con_info(
                        service,
                        _id_carpeta_real(item),
                        nombre,
                        sub_path,
                        debug=debug,
                        _visitadas=_visitadas,
                        strict=strict,
                    )
                )
            else:
                target = item.get("shortcutDetails", {}) if mime_type == SHORTCUT_MIME else {}
                _, ext = os.path.splitext(nombre)
                resultado.append({
                    "nombre"        : nombre,
                    "extension"     : ext.lower(),
                    "mime_type"     : target.get("targetMimeType") or mime_type,
                    "carpeta_padre" : carpeta_padre,
                    "folder_path"   : folder_path,
                    "drive_file_id" : target.get("targetId") or item["id"],
                })

    except Exception as e:
        ruta = folder_path or "(raíz)"
        detalle = _diagnosticar_error_drive(e)
        mensaje = f"Error al listar carpeta Drive {folder_id} /{ruta}: {detalle}. Detalle: {e}"
        if strict:
            raise DriveAccessError(mensaje) from e
        print(f"   ❌ {mensaje}")

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


def listar_archivos_con_info_service(
    service,
    folder_id: str,
    debug: bool = False,
    link_original: str | None = None,
) -> list[dict]:
    """
    Versión pública para service account.
    Usar desde main.py donde el service ya está conectado con cuenta de servicio.
    """
    if debug:
        if link_original:
            print(f"      🔗 URL portafolio : {link_original}")
        print(f"      🔑 folder_id      : {folder_id}")

    meta = None
    try:
        folder_id_efectivo, meta = _resolver_folder_raiz(service, folder_id, debug=debug)
    except Exception as e:
        diagnostico = _diagnosticar_error_drive(e)
        raise DriveAccessError(
            f"No se puede leer la carpeta raíz Drive {folder_id}: {diagnostico}. "
            "Verifica que el link sea la carpeta exacta y que esté compartida "
            f"con la cuenta autenticada por el bot. Detalle: {e}"
        ) from e

    if debug:
        if folder_id_efectivo != folder_id:
            print(f"      🔑 folder efectivo: {folder_id_efectivo}")
        print(f"      ✅ files.get raíz : OK")
    archivos = _listar_con_info(service, folder_id_efectivo, debug=debug, strict=True)
    if debug:
        print(f"      📄 archivos antes del matching: {len(archivos)}")
        for archivo in archivos[:30]:
            ruta = archivo.get("folder_path") or "(raíz)"
            print(f"      │  {archivo['nombre']}  /{ruta}")
        if len(archivos) > 30:
            print(f"      │  ... y {len(archivos) - 30} archivo(s) más")
    return archivos


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
    link = unquote(link.strip())
    try:
        parsed = urlparse(link)
        params = parse_qs(parsed.query)
        if params.get("id"):
            return params["id"][0]
    except Exception:
        pass
    for patron in [
        r'/folders/([a-zA-Z0-9_-]+)',
        r'/drive/(?:u/\d+/)?folders/([a-zA-Z0-9_-]+)',
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
