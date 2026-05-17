import argparse
import sys

from Core.drive_manager import (
    EXPECTED_CLIENT_EMAIL_ENV,
    conectar_drive,
    obtener_info_credenciales_drive,
    validar_info_credenciales_drive,
)
from bot.drive_adapter import (
    FOLDER_MIME,
    SHORTCUT_MIME,
    extraer_id_carpeta,
    inspeccionar_carpeta_service,
    listar_archivos_con_info_service,
)


def _resolver_entrada(args) -> tuple[str, str]:
    entrada = args.folder_id or args.url or args.url_o_id
    if not entrada:
        raise ValueError("Debes indicar una URL o un folder_id.")
    folder_id = args.folder_id or extraer_id_carpeta(entrada)
    if not folder_id:
        raise ValueError("No se pudo extraer folder_id del valor recibido.")
    return entrada, folder_id


def _imprimir_contexto_prueba(url_original: str, folder_id: str, cuenta: str, cred_info: dict) -> None:
    print("\nContexto de prueba:")
    print(f"URL original     : {url_original}")
    print(f"folder_id probado: {folder_id}")
    print(f"cuenta autenticada: {cuenta}")
    print(f"ruta JSON        : {cred_info.get('ruta')}")


def _parece_404_drive(exc: Exception) -> bool:
    texto = str(exc).lower()
    return "404" in texto or "file not found" in texto


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Diagnostica el listado real de archivos de un portafolio Drive."
    )
    parser.add_argument(
        "url_o_id",
        nargs="?",
        help="URL de Google Drive o folder_id del portafolio",
    )
    parser.add_argument("--url", help="URL de Google Drive del portafolio")
    parser.add_argument("--folder-id", help="folder_id puro de Google Drive")
    args = parser.parse_args()

    try:
        url_original, folder_id = _resolver_entrada(args)
    except ValueError as exc:
        print(exc)
        return 2

    cred_info = obtener_info_credenciales_drive()
    try:
        validar_info_credenciales_drive(cred_info)
    except Exception as exc:
        print(f"Error credenciales Drive: {exc}")
        return 2

    try:
        service = conectar_drive(debug=False)
    except Exception as exc:
        print(f"Error autenticando Drive: {exc}")
        return 2

    cuenta = cred_info.get("client_email") or "(cuenta no disponible)"
    esperado = cred_info.get("expected_client_email")

    print("\nCuenta Drive autenticada:")
    print(f"fuente       : {cred_info.get('fuente')}")
    print(f"ruta JSON    : {cred_info.get('ruta')}")
    print(f"client_email : {cuenta}")
    print(f"project_id   : {cred_info.get('project_id')}")
    if esperado:
        print(f"esperado     : {esperado} [{EXPECTED_CLIENT_EMAIL_ENV}]")

    _imprimir_contexto_prueba(url_original, folder_id, cuenta, cred_info)

    try:
        inspeccion = inspeccionar_carpeta_service(service, folder_id)
        meta = inspeccion["metadata"]
        hijos = inspeccion["hijos"]
        print("\nCarpeta raíz accesible:")
        print(f"nombre  : {meta.get('name')}")
        print(f"mimeType: {meta.get('mimeType')}")
        if inspeccion["folder_id_efectivo"] != folder_id:
            print(f"folder_id efectivo: {inspeccion['folder_id_efectivo']}")

        carpetas = [
            h for h in hijos
            if h.get("mimeType") == FOLDER_MIME
            or h.get("shortcutDetails", {}).get("targetMimeType") == FOLDER_MIME
        ]
        shortcuts = [h for h in hijos if h.get("mimeType") == SHORTCUT_MIME]
        archivos_raiz = [h for h in hijos if h not in carpetas]
        print(f"hijos directos: {len(hijos)}")
        print(f"subcarpetas   : {len(carpetas)}")
        for item in carpetas:
            tipo = "shortcut->folder" if item.get("mimeType") == SHORTCUT_MIME else "folder"
            target = item.get("shortcutDetails", {}).get("targetId")
            suffix = f" target={target}" if target else ""
            print(f"  - {item.get('name')} [{tipo}]{suffix}")
        print(f"shortcuts     : {len(shortcuts)}")
        for item in shortcuts:
            target_mime = item.get("shortcutDetails", {}).get("targetMimeType")
            print(f"  - {item.get('name')} -> {target_mime}")
        print(f"archivos raíz : {len(archivos_raiz)}")
        for item in archivos_raiz[:20]:
            print(f"  - {item.get('name')} [{item.get('mimeType')}]")

        archivos = listar_archivos_con_info_service(
            service,
            folder_id,
            debug=True,
            link_original=url_original,
        )
    except Exception as exc:
        print(f"Error Drive: {exc}")
        if _parece_404_drive(exc):
            print("\nDiagnóstico accionable:")
            print(f"La cuenta autenticada es {cuenta}.")
            print(f"El folder_id probado es {folder_id}.")
            print(
                "Esto suele significar que el ID no corresponde a la carpeta compartida "
                "o que esa carpeta exacta no fue compartida con esta cuenta."
            )
        else:
            print(f"La cuenta autenticada es {cuenta} y no pudo leer el folder {folder_id}.")
            print("Comparte la carpeta exacta con ese client_email o corrige la ruta del JSON activo.")
        return 1

    print("\nResumen:")
    print(f"folder_id: {folder_id}")
    print(f"archivos : {len(archivos)}")
    for archivo in archivos:
        ruta = archivo.get("folder_path") or "(raíz)"
        ext = archivo.get("extension") or archivo.get("mime_type") or "?"
        print(f"- {archivo['nombre']} [{ext}] /{ruta}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
