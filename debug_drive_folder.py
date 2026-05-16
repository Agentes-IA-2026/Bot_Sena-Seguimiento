import argparse
import sys

from Core.drive_manager import conectar_drive
from bot.drive_adapter import extraer_id_carpeta, listar_archivos_con_info_service


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

    try:
        archivos = listar_archivos_con_info_service(
            service,
            folder_id,
            debug=True,
            link_original=args.url_o_id,
        )
    except Exception as exc:
        print(f"Error Drive: {exc}")
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
