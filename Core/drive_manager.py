import os
import re
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build


DEFAULT_CREDENTIALS_PATH = 'assets/credenciales.json'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
EXPECTED_CLIENT_EMAIL_ENV = "BOT_SENA_EXPECTED_DRIVE_CLIENT_EMAIL"
FOLDER_MIME = 'application/vnd.google-apps.folder'
SHORTCUT_MIME = 'application/vnd.google-apps.shortcut'


def _resolver_ruta_credenciales_drive() -> tuple[str, str]:
    """
    Fuente única y explícita para credenciales Drive.
    Prioridad:
      1. BOT_SENA_DRIVE_CREDENTIALS
      2. GOOGLE_APPLICATION_CREDENTIALS
      3. assets/credenciales.json
    """
    candidatos = [
        ("BOT_SENA_DRIVE_CREDENTIALS", os.environ.get("BOT_SENA_DRIVE_CREDENTIALS")),
        ("GOOGLE_APPLICATION_CREDENTIALS", os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")),
        ("default", DEFAULT_CREDENTIALS_PATH),
    ]
    for fuente, ruta in candidatos:
        if ruta:
            return fuente, os.path.abspath(os.path.expanduser(ruta))
    return "default", os.path.abspath(DEFAULT_CREDENTIALS_PATH)


def obtener_info_credenciales_drive() -> dict:
    fuente, ruta = _resolver_ruta_credenciales_drive()
    info = {
        "fuente": fuente,
        "ruta": ruta,
        "existe": os.path.exists(ruta),
        "client_email": None,
        "project_id": None,
        "type": None,
        "expected_client_email": os.environ.get(EXPECTED_CLIENT_EMAIL_ENV),
    }
    if not info["existe"]:
        return info
    try:
        with open(ruta, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        info["client_email"] = data.get("client_email")
        info["project_id"] = data.get("project_id")
        info["type"] = data.get("type")
    except Exception as exc:
        info["error"] = str(exc)
    return info


def validar_info_credenciales_drive(info: dict) -> None:
    """Valida el JSON activo sin exponer secretos."""
    ruta = info.get("ruta")
    if not info.get("existe"):
        raise RuntimeError(f"No se encontró el archivo de credenciales Drive: {ruta}")
    if info.get("error"):
        raise RuntimeError(f"No se pudo leer el JSON de credenciales Drive {ruta}: {info['error']}")
    if info.get("type") != "service_account":
        raise RuntimeError(
            f"El JSON Drive debe ser service_account. "
            f"Ruta={ruta}, type={info.get('type') or '(vacío)'}"
        )
    if not info.get("client_email"):
        raise RuntimeError(f"El JSON Drive no tiene client_email. Ruta={ruta}")
    if not info.get("project_id"):
        raise RuntimeError(f"El JSON Drive no tiene project_id. Ruta={ruta}")

    esperado = info.get("expected_client_email")
    if esperado and info["client_email"].lower() != esperado.lower():
        raise RuntimeError(
            "La credencial Drive activa no coincide con la cuenta esperada. "
            f"Esperada={esperado}; activa={info['client_email']}; "
            f"fuente={info['fuente']}; ruta={ruta}. "
            f"Corrige {EXPECTED_CLIENT_EMAIL_ENV} o la ruta de credenciales."
        )


def imprimir_info_credenciales_drive(prefijo: str = "   ") -> dict:
    info = obtener_info_credenciales_drive()
    print(f"{prefijo}🔐 Credenciales Drive")
    print(f"{prefijo}   fuente       : {info['fuente']}")
    print(f"{prefijo}   ruta JSON    : {info['ruta']}")
    print(f"{prefijo}   existe       : {info['existe']}")
    print(f"{prefijo}   client_email : {info.get('client_email') or '(no disponible)'}")
    print(f"{prefijo}   project_id   : {info.get('project_id') or '(no disponible)'}")
    esperado = info.get("expected_client_email")
    if esperado:
        coincide = (info.get("client_email") or "").lower() == esperado.lower()
        estado = "OK" if coincide else "NO COINCIDE"
        print(f"{prefijo}   esperado     : {esperado} [{estado}] ({EXPECTED_CLIENT_EMAIL_ENV})")
    if info.get("error"):
        print(f"{prefijo}   error JSON   : {info['error']}")
    return info


def conectar_drive(debug: bool = True):
    """Conecta con la API usando la fuente única de credenciales Drive."""
    info = imprimir_info_credenciales_drive() if debug else obtener_info_credenciales_drive()
    ruta_json = info["ruta"]
    validar_info_credenciales_drive(info)
 
    creds = service_account.Credentials.from_service_account_file(ruta_json, scopes=DRIVE_SCOPES)
    return build('drive', 'v3', credentials=creds)


def _resolver_folder_real(service, folder_id: str) -> str:
    meta = service.files().get(
        fileId=folder_id,
        fields="id, name, mimeType, shortcutDetails",
        supportsAllDrives=True,
    ).execute()
    mime = meta.get('mimeType')
    if mime == SHORTCUT_MIME:
        details = meta.get('shortcutDetails') or {}
        target_id = details.get('targetId')
        target_mime = details.get('targetMimeType')
        if target_id and target_mime == FOLDER_MIME:
            print(f"   🔀 Folder raíz es shortcut → {target_id}")
            return target_id
    if mime != FOLDER_MIME:
        raise RuntimeError(f"El id Drive no es carpeta. mimeType={mime}")
    return folder_id
 
 
def _normalizar(texto: str) -> str:
    """
    Normaliza un texto para comparación flexible:
    - Minúsculas
    - Sin tildes
    - Sin extensiones de archivo
    - Sin caracteres especiales, solo letras y números
    """
    texto = texto.lower()
 
    # Quitar tildes
    reemplazos = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'ä': 'a', 'ë': 'e', 'ï': 'i', 'ö': 'o', 'ü': 'u',
        'à': 'a', 'è': 'e', 'ì': 'i', 'ò': 'o', 'ù': 'u',
        'ñ': 'n',
    }
    for origen, destino in reemplazos.items():
        texto = texto.replace(origen, destino)
 
    # Quitar extensiones comunes
    texto = re.sub(r'\.(pdf|png|jpg|jpeg|xlsx|xls|pptx|ppt|docx|doc|mp4)(\.|$)', ' ', texto)
 
    # Quitar caracteres especiales, dejar solo letras, números y espacios
    texto = re.sub(r'[^a-z0-9\s]', ' ', texto)
 
    # Colapsar espacios múltiples
    texto = re.sub(r'\s+', ' ', texto).strip()
 
    return texto
 
 
def _palabras_clave(texto_normalizado: str) -> set[str]:
    """
    Extrae palabras significativas (más de 3 letras) de un texto normalizado.
    Ignora palabras muy comunes que no aportan al match.
    """
    PALABRAS_IGNORAR = {
        'pdf', 'png', 'jpg', 'xlsx', 'pptx', 'docx', 'doc',
        'con', 'del', 'los', 'las', 'una', 'uno', 'por',
        'que', 'para', 'como', 'the', 'and', 'sena',  # "sena" aparece en casi todo
    }
    palabras = texto_normalizado.split()
    return {p for p in palabras if len(p) > 3 and p not in PALABRAS_IGNORAR}
 
 
def _singularizar(palabra: str) -> str:
    """
    Reduce una palabra española a su forma aproximada en singular.
    Sirve para comparar 'cotizaciones' con 'cotizacion', 'fotos' con 'foto', etc.
    """
    if len(palabra) <= 4:
        return palabra
    if palabra.endswith('es') and len(palabra) > 4:
        return palabra[:-2]
    if palabra.endswith(('as', 'os', 's')) and len(palabra) > 3:
        return palabra[:-1]
    return palabra


def _coincide(nombre_evidencia: str, nombre_drive: str, debug: bool = False) -> bool:
    """
    Determina si un archivo del Drive corresponde a una evidencia requerida.
    Usa múltiples estrategias de comparación flexible.
    """
    ev_norm = _normalizar(nombre_evidencia)
    drive_norm = _normalizar(nombre_drive)
 
    # Estrategia 1: Contención directa (uno contiene al otro)
    if ev_norm in drive_norm or drive_norm in ev_norm:
        return True
 
    # Estrategia 2: Palabras clave en común
    # Calculamos qué porcentaje de palabras clave de la evidencia aparecen en el archivo
    palabras_ev = _palabras_clave(ev_norm)
    palabras_drive = _palabras_clave(drive_norm)
 
    if palabras_ev and palabras_drive:
        comunes = palabras_ev & palabras_drive
        porcentaje = len(comunes) / len(palabras_ev)
 
        if debug and comunes:
            print(f"      palabras_ev={palabras_ev} | drive={palabras_drive} | "
                  f"comunes={comunes} | {porcentaje:.0%}")
 
        # Si el 50% o más de las palabras clave coinciden → match
        if porcentaje >= 0.5:
            return True
 
    # Estrategia 3: Coincidencia por bigramas (pares de palabras consecutivas)
    # Útil para "quien soy" vs "3.1 quien soy.docx"
    def bigramas(palabras):
        lista = sorted(palabras)
        return {f"{lista[i]} {lista[i+1]}" for i in range(len(lista)-1)}
 
    if len(palabras_ev) >= 2 and len(palabras_drive) >= 2:
        bi_ev = bigramas(palabras_ev)
        bi_drive = bigramas(palabras_drive)
        if bi_ev & bi_drive:
            return True

    # Estrategia 4: Comparación en singular
    # Detecta que "cotizaciones" y "cotizacion" son la misma evidencia
    raices_ev    = {_singularizar(p) for p in palabras_ev}
    raices_drive = {_singularizar(p) for p in palabras_drive}

    if raices_ev and raices_drive:
        comunes_sing = raices_ev & raices_drive
        porcentaje_sing = len(comunes_sing) / len(raices_ev)

        if debug and comunes_sing:
            print(f"      [singular] raices_ev={raices_ev} | drive={raices_drive} | "
                  f"comunes={comunes_sing} | {porcentaje_sing:.0%}")

        if porcentaje_sing >= 0.5:
            return True

    return False
 
 
def _listar_archivos_recursivo(service, folder_id: str, strict: bool = True) -> list[str]:
    """
    Lista TODOS los archivos dentro de una carpeta y sus subcarpetas.
    Retorna una lista de nombres de archivo originales.
    """
    folder_mime = 'application/vnd.google-apps.folder'
    shortcut_mime = 'application/vnd.google-apps.shortcut'
    nombres = []
 
    try:
        folder_id = _resolver_folder_real(service, folder_id)
        page_token = None
        while True:
            results = service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType, shortcutDetails)",
                pageSize=1000,
                pageToken=page_token,
                corpora="allDrives",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
 
            items = results.get('files', [])
 
            for item in items:
                is_folder = item['mimeType'] == folder_mime
                is_folder_shortcut = (
                    item['mimeType'] == shortcut_mime
                    and item.get('shortcutDetails', {}).get('targetMimeType') == folder_mime
                )
                if is_folder or is_folder_shortcut:
                    target_id = item.get('shortcutDetails', {}).get('targetId') or item['id']
                    sub = _listar_archivos_recursivo(service, target_id, strict=strict)
                    nombres.extend(sub)
                else:
                    nombres.append(item['name'])
 
            page_token = results.get('nextPageToken')
            if not page_token:
                break
 
    except Exception as e:
        mensaje = f"Error al listar carpeta Drive {folder_id}: {e}"
        if strict:
            raise RuntimeError(mensaje) from e
        print(f"   ⚠️ {mensaje}")
 
    return nombres
 
 
def verificar_evidencias_en_carpeta(service, folder_id: str,
                                    lista_evidencias: list[str],
                                    debug: bool = True) -> dict:
    """
    Verifica qué evidencias requeridas están presentes en la carpeta del aprendiz.
    Busca recursivamente en subcarpetas y usa comparación flexible.
    """
    resultados = {ev: False for ev in lista_evidencias}
 
    archivos_en_drive = _listar_archivos_recursivo(service, folder_id)
 
    if not archivos_en_drive:
        print(f"   ⚠️ No se encontraron archivos (carpeta vacía o sin acceso)")
        return resultados
 
    print(f"   📂 {len(archivos_en_drive)} archivo(s) en Drive:")
    for a in archivos_en_drive:
        print(f"      • {a}")
 
    print(f"   🔍 Comparando evidencias...")
    for ev in lista_evidencias:
        for nombre_real in archivos_en_drive:
            if _coincide(ev, nombre_real, debug=debug):
                resultados[ev] = True
                print(f"   ✅ '{ev}' → '{nombre_real}'")
                break
 
    return resultados
