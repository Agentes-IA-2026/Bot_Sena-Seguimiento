import argparse
import sys

from Core.drive_manager import conectar_drive, obtener_info_credenciales_drive
from bot.drive_adapter import (
    FOLDER_MIME,
    SHORTCUT_MIME,
    extraer_id_carpeta,
    inspeccionar_carpeta_service,
    listar_archivos_con_info_service,
)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(
        description="Diagnostica el listado real de archivos de un portafolio Drive."
    )
    parser.add_argument("url_o_id", help="URL de Google Drive o folder_id del portafolio")
    args = parser.parse_args()

    folder_id = extraer_id_carpeta(args.url_o_id)
    if not folder_id:
        print("No se pudo extraer folder_id del valor recibido.")
        return 2

    service = conectar_drive()
    if service is None:
        return 2
    cred_info = obtener_info_credenciales_drive()
    cuenta = cred_info.get("client_email") or "(cuenta no disponible)"

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
            link_original=args.url_o_id,
        )
    except Exception as exc:
        print(f"Error Drive: {exc}")
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
