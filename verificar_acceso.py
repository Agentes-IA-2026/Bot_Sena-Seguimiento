"""
Verifica si el bot puede acceder a las carpetas de Drive
de los aprendices de la ficha 2 (CGS - 3414937 - 2026.xlsx)
"""
import re
import openpyxl
from Core.drive_manager import conectar_drive, _listar_archivos_recursivo
 
ARCHIVO = "assets/Listados/CGS - 3414937 - 2026.xlsx"
 
service = conectar_drive()
if not service:
    print("❌ No se pudo conectar con Drive")
    exit()
 
wb = openpyxl.load_workbook(ARCHIVO, data_only=True)
hoja = wb["BASE DE DATOS"]
 
print(f"\n{'='*55}")
print(f"VERIFICACIÓN DE ACCESO — FICHA 3414937")
print(f"{'='*55}\n")
 
con_acceso = 0
sin_acceso = 0
sin_link   = 0
 
for fila in hoja.iter_rows(min_row=10, max_row=100):
    nombre_val = fila[1].value
    celda_link = fila[10]
 
    nombre = str(nombre_val).strip() if nombre_val else None
    if not nombre or nombre in ["Nombre", "None", "nan", "", "TI", "PPT", "CC"]:
        continue
 
    if celda_link.hyperlink:
        link = celda_link.hyperlink.target
    elif celda_link.value:
        link = str(celda_link.value).strip()
    else:
        link = None
 
    if not link or "http" not in str(link):
        print(f"  ⚠️  {nombre}: sin link de Drive")
        sin_link += 1
        continue
 
    m = re.search(r'/folders/([a-zA-Z0-9_-]+)', link)
    if not m:
        print(f"  ❓ {nombre}: link con formato desconocido")
        continue
 
    folder_id = m.group(1)
    archivos  = _listar_archivos_recursivo(service, folder_id)
 
    if archivos:
        print(f"  ✅ {nombre}: {len(archivos)} archivo(s)")
        con_acceso += 1
    else:
        print(f"  ❌ {nombre}: sin acceso o carpeta vacía")
        sin_acceso += 1
 
print(f"\n{'='*55}")
print(f"  ✅ Con acceso : {con_acceso}")
print(f"  ❌ Sin acceso : {sin_acceso}")
print(f"  ⚠️  Sin link   : {sin_link}")
print(f"{'='*55}\n")