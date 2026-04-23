import os
import re
import glob
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime

from Core.drive_manager import (
    conectar_drive,
    verificar_evidencias_en_carpeta,
)
from Core.document_analyzer import extraer_nombres_evidencias_manual, extraer_nombres_evidencias

# ── SUPABASE ──────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from supabase import create_client

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://cfahgjytbpnmsogzryov.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNmYWhnanl0YnBubXNvZ3pyeW92Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NTYwODY2MSwiZXhwIjoyMDkxMTg0NjYxfQ.UXzWepH1HJ-KanXHoRZ3JgPK7Umt6WramF_fw26YNXM")
DOCENTE_ID   = os.environ.get("DOCENTE_ID", "d114d85c-fb7e-4a79-a446-fb90d652eec2")

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
CARPETA_LISTADOS = "assets/Listados"
CARPETA_GUIAS    = "assets/Guias_Referencia"
CARPETA_REPORTES = "logs/reportes"

MAPA_PROGRAMAS = {
    "asesoria comercial":       "Asistencia_comercial",
    "asistencia comercial":     "Asistencia_comercial",
    "comunicacion comercial":   "Comunicacion_y_marketing",
    "comercial y marketing":    "Comunicacion_y_marketing",
    "comunicacion y marketing": "Comunicacion_y_marketing",
    "comunicacion":             "Comunicacion_y_marketing",
    "marketing":                "Comunicacion_y_marketing",
    "ventas de productos":      "Ventas_de_productos_en_linea",
    "productos en linea":       "Ventas_de_productos_en_linea",
    "ventas":                   "Ventas_de_productos_en_linea",
}

COLOR_VERDE    = "C6EFCE"
COLOR_ROJO     = "FFC7CE"
COLOR_AMARILLO = "FFEB9C"
COLOR_GRIS     = "D9D9D9"


# ── SUPABASE: GUARDAR ─────────────────────────────────────────────────────────

def _guardar_en_supabase(resultados: list[dict], ficha: str, colegio: str, siglas: str):
    """
    Guarda resultados en Supabase con ficha, colegio y siglas.
    Borra primero los registros anteriores de esa ficha.
    """
    try:
        db    = create_client(SUPABASE_URL, SUPABASE_KEY)
        fecha = datetime.now().isoformat()

        # Borrar registros anteriores de esta ficha
        db.table("verificaciones")\
          .delete()\
          .eq("docente_id", DOCENTE_ID)\
          .eq("ficha", ficha)\
          .execute()

        registros = []
        for aprendiz in resultados:
            for nombre_guia, evidencias in aprendiz.get("guias", {}).items():
                # Nombre limpio de la guía (sin extensión .pdf)
                guia_limpia = re.sub(r'\.pdf$', '', nombre_guia, flags=re.IGNORECASE)
                guia_limpia = guia_limpia.replace('_', ' ')
                guia_limpia = re.sub(r' {2,}', ' ', guia_limpia).strip()

                for evidencia, entregado in evidencias.items():
                    # Nombre limpio de evidencia (sin extensión)
                    ev_limpia = re.sub(r'\.\w+$', '', evidencia).replace('_', ' ').strip()

                    registros.append({
                        "docente_id" : DOCENTE_ID,
                        "ficha"      : ficha,
                        "colegio"    : f"{siglas} — {colegio}",
                        "estudiante" : aprendiz["nombre"],
                        "guia"       : guia_limpia,
                        "evidencia"  : ev_limpia,
                        "entregado"  : bool(entregado),
                        "fecha"      : fecha,
                    })

        if not registros:
            print("   ⚠️  Sin registros para guardar")
            return

        # Insertar en lotes de 500
        for i in range(0, len(registros), 500):
            db.table("verificaciones").insert(registros[i:i+500]).execute()

        print(f"   ☁️  {len(registros)} registros guardados en Supabase ✅")

    except Exception as e:
        print(f"   ❌ Error guardando en Supabase: {e}")


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _leer_celda_texto(celda) -> str:
    """Lee una celda de Excel y la convierte siempre a texto limpio."""
    val = celda.value
    if val is None:
        return ""
    # Si es número (int o float) convertir directamente
    if isinstance(val, (int, float)):
        return str(int(val)).strip()
    return str(val).strip()


def _normalizar(texto: str) -> str:
    texto = texto.lower()
    for a, b in [('á','a'),('é','e'),('í','i'),('ó','o'),('ú','u'),('ñ','n')]:
        texto = texto.replace(a, b)
    return texto


def _resolver_carpeta_guias(nombre_programa: str) -> str | None:
    prog_norm = _normalizar(nombre_programa)
    for clave, carpeta in MAPA_PROGRAMAS.items():
        if clave in prog_norm:
            ruta = os.path.join(CARPETA_GUIAS, carpeta)
            return ruta if os.path.isdir(ruta) else None
    return None


def _extraer_id_carpeta(link: str) -> str | None:
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


def _leer_evidencias_de_guias(carpeta_guias: str) -> dict:
    guias = {}
    todos = []
    if os.path.isdir(carpeta_guias):
        for archivo in os.listdir(carpeta_guias):
            if archivo.lower().endswith('.pdf'):
                todos.append(os.path.join(carpeta_guias, archivo))
    todos = sorted(todos)

    if not todos:
        print(f"❌ No se encontraron guías PDF en: {carpeta_guias}")
        return guias

    for pdf in todos:
        nombre = os.path.basename(pdf)
        # Primero usar lista manual (limpia y confiable)
        evidencias = extraer_nombres_evidencias_manual(pdf)
        if not evidencias:
            # Solo si no hay lista manual, extraer automáticamente del PDF
            evidencias = extraer_nombres_evidencias(pdf, debug=False)
        if evidencias:
            guias[nombre] = evidencias
            print(f"   📋 {nombre}: {len(evidencias)} evidencia(s)")
        else:
            # Incluir la guía con marcador para que aparezca en el dashboard
            guias[nombre] = ["Evidencia pendiente de configurar"]
            print(f"   ⚠️  Sin evidencias configuradas para: {nombre}")
    return guias


def _leer_aprendices(hoja) -> list[dict]:
    aprendices = []
    for fila in hoja.iter_rows(min_row=10, max_row=200):
        nombre_val = fila[1].value
        celda_link = fila[10]

        nombre = str(nombre_val).strip() if nombre_val else None
        if not nombre or nombre in ["Nombre", "TI", "PPT", "CC", "None", "nan", ""]:
            continue

        if celda_link.hyperlink:
            link = celda_link.hyperlink.target
        elif celda_link.value:
            link = str(celda_link.value).strip()
        else:
            link = None

        aprendices.append({"nombre": nombre, "link": link})
    return aprendices


# ── REPORTE EXCEL LOCAL ───────────────────────────────────────────────────────

def _generar_reporte_excel(ficha: str, programa: str, resultados: list[dict],
                           guias: dict, ruta_salida: str):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    borde = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'),  bottom=Side(style='thin')
    )

    for nombre_guia, evidencias in guias.items():
        nombre_hoja = re.sub(r'\.pdf$', '', nombre_guia, flags=re.IGNORECASE)[:28]
        ws = wb.create_sheet(title=nombre_hoja)

        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=2 + len(evidencias))
        ct = ws.cell(row=1, column=1)
        ct.value = f"REPORTE FICHA {ficha} | {programa} | {nombre_guia}"
        ct.font = Font(bold=True, size=11, color="FFFFFF")
        ct.fill = PatternFill("solid", fgColor="2E4057")
        ct.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        ws.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=2 + len(evidencias))
        cf = ws.cell(row=2, column=1)
        cf.value = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        cf.font = Font(italic=True, size=9)
        cf.alignment = Alignment(horizontal="right")

        encabezados = ["Aprendiz"] + [
            re.sub(r'\.\w+$', '', ev) for ev in evidencias
        ] + ["Total", "%"]

        for col, enc in enumerate(encabezados, 1):
            c = ws.cell(row=3, column=col, value=enc)
            c.font = Font(bold=True, size=9)
            c.fill = PatternFill("solid", fgColor=COLOR_GRIS)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde
        ws.row_dimensions[3].height = 40

        for fila_idx, aprendiz in enumerate(resultados, 4):
            reporte_guia = aprendiz.get("guias", {}).get(nombre_guia, {})
            sin_link = not aprendiz.get("link")

            c_nombre = ws.cell(row=fila_idx, column=1, value=aprendiz["nombre"])
            c_nombre.font = Font(size=9)
            c_nombre.border = borde
            c_nombre.alignment = Alignment(vertical="center")

            entregadas = 0
            for col_idx, ev in enumerate(evidencias, 2):
                c = ws.cell(row=fila_idx, column=col_idx)
                if sin_link:
                    c.value = "—"
                    c.fill = PatternFill("solid", fgColor=COLOR_AMARILLO)
                elif reporte_guia.get(ev):
                    c.value = "✓"
                    c.fill = PatternFill("solid", fgColor=COLOR_VERDE)
                    entregadas += 1
                else:
                    c.value = "✗"
                    c.fill = PatternFill("solid", fgColor=COLOR_ROJO)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = Font(size=9, bold=True)
                c.border = borde

            total_ev = len(evidencias)
            pct = (entregadas / total_ev * 100) if total_ev and not sin_link else 0

            c_total = ws.cell(row=fila_idx, column=2 + len(evidencias),
                              value=f"{entregadas}/{total_ev}" if not sin_link else "—")
            c_pct   = ws.cell(row=fila_idx, column=3 + len(evidencias),
                              value=f"{pct:.0f}%" if not sin_link else "—")

            for c in [c_total, c_pct]:
                c.font = Font(size=9, bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = borde
                if not sin_link:
                    color = COLOR_VERDE if pct == 100 else (COLOR_ROJO if pct < 50 else COLOR_AMARILLO)
                    c.fill = PatternFill("solid", fgColor=color)

        ws.column_dimensions["A"].width = 30
        for col in range(2, 2 + len(evidencias)):
            letra = openpyxl.utils.get_column_letter(col)
            ws.column_dimensions[letra].width = 14
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + len(evidencias))].width = 10
        ws.column_dimensions[openpyxl.utils.get_column_letter(3 + len(evidencias))].width = 8
        ws.freeze_panes = "B4"

    os.makedirs(os.path.dirname(ruta_salida), exist_ok=True)
    wb.save(ruta_salida)
    print(f"   💾 Reporte Excel: {ruta_salida}")


# ── FLUJO PRINCIPAL ───────────────────────────────────────────────────────────

def auditar_ficha(ruta_excel: str, service):
    nombre_excel = os.path.basename(ruta_excel)
    print(f"\n{'='*60}")
    print(f"📋 FICHA: {nombre_excel}")
    print(f"{'='*60}")

    try:
        wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    except Exception as e:
        print(f"❌ No se pudo abrir el Excel: {e}")
        return

    hoja_bd = wb["BASE DE DATOS"] if "BASE DE DATOS" in wb.sheetnames else None
    if not hoja_bd:
        print("❌ No se encontró la hoja 'BASE DE DATOS'")
        return

    # Leer datos con conversión robusta
    programa_raw = _leer_celda_texto(hoja_bd["C3"])
    ficha        = _leer_celda_texto(hoja_bd["C4"])   # ← ahora convierte int a str
    colegio      = _leer_celda_texto(hoja_bd["C7"])

    # Extraer siglas del nombre del archivo (ej: CGS, CJLL, CRFK, CTSM)
    siglas_m = re.match(r'^([A-Z]+)', nombre_excel)
    siglas   = siglas_m.group(1) if siglas_m else "GRP"

    if not programa_raw:
        print("❌ No se encontró el programa en celda C3")
        return

    if not ficha:
        # Si C4 está vacío, extraer número del nombre del archivo
        nums = re.findall(r'\d{7}', nombre_excel)
        ficha = nums[0] if nums else nombre_excel

    print(f"   Programa : {programa_raw}")
    print(f"   Ficha    : {ficha}")
    print(f"   Colegio  : {colegio}")
    print(f"   Siglas   : {siglas}")

    carpeta_guias = _resolver_carpeta_guias(programa_raw)
    if not carpeta_guias:
        print(f"❌ No se encontró carpeta de guías para: '{programa_raw}'")
        print(f"   Agrega una entrada en MAPA_PROGRAMAS dentro de main.py")
        return

    guias = _leer_evidencias_de_guias(carpeta_guias)
    if not guias:
        return

    aprendices = _leer_aprendices(hoja_bd)
    print(f"\n   👥 Aprendices encontrados: {len(aprendices)}")

    print(f"\n   0. Auditar TODOS")
    for i, a in enumerate(aprendices, 1):
        print(f"   {i}. {a['nombre']}")

    sel = input("\n   ¿Qué aprendiz auditar? (número o 0 para todos): ").strip()

    if sel == "0":
        aprendices_a_auditar = aprendices
    elif sel.isdigit() and 1 <= int(sel) <= len(aprendices):
        aprendices_a_auditar = [aprendices[int(sel) - 1]]
        print(f"\n   🎯 Auditando solo: {aprendices_a_auditar[0]['nombre']}")
    else:
        print("❌ Selección inválida")
        return

    print(f"   🔍 Iniciando auditoría...\n")
    resultados = []

    for aprendiz in aprendices_a_auditar:
        nombre = aprendiz["nombre"]
        link   = aprendiz["link"]
        print(f"   🧐 {nombre}")

        resultado_aprendiz = {"nombre": nombre, "link": link, "guias": {}}

        if not link or "http" not in str(link):
            print(f"      ⚠️  Sin link de Drive")
            resultados.append(resultado_aprendiz)
            continue

        folder_id = _extraer_id_carpeta(link)
        if not folder_id:
            print(f"      ❌ No se pudo extraer el ID del link")
            resultados.append(resultado_aprendiz)
            continue

        for nombre_guia, evidencias in guias.items():
            try:
                reporte = verificar_evidencias_en_carpeta(
                    service, folder_id, evidencias, debug=False
                )
                entregadas = sum(reporte.values())
                total      = len(evidencias)
                print(f"      📂 {nombre_guia[:40]}: {entregadas}/{total}")
                resultado_aprendiz["guias"][nombre_guia] = reporte
            except Exception as e:
                print(f"      ❌ Error en {nombre_guia}: {e}")

        resultados.append(resultado_aprendiz)

    # Reporte Excel local
    fecha_hoy   = datetime.now().strftime("%Y%m%d_%H%M")
    nombre_base = re.sub(r'\.xlsx$', '', nombre_excel, flags=re.IGNORECASE)
    ruta_salida = os.path.join(CARPETA_REPORTES, f"Reporte_{nombre_base}_{fecha_hoy}.xlsx")
    _generar_reporte_excel(ficha, programa_raw, resultados, guias, ruta_salida)

    # Guardar en Supabase
    _guardar_en_supabase(resultados, ficha, colegio, siglas)

    print(f"\n✅ Ficha {ficha} — {siglas} completada.")


def ejecutar_todas_las_fichas():
    print("\n🤖 GUÍABOT — AUDITORÍA DE EVIDENCIAS")
    print(f"   Buscando fichas en: {CARPETA_LISTADOS}\n")

    excels = sorted(glob.glob(os.path.join(CARPETA_LISTADOS, "*.xlsx")))

    if not excels:
        print(f"❌ No se encontraron archivos Excel en {CARPETA_LISTADOS}")
        return

    print(f"   Fichas encontradas ({len(excels)}):")
    for i, e in enumerate(excels, 1):
        print(f"   {i}. {os.path.basename(e)}")

    print(f"\n   0. Auditar TODAS")
    seleccion = input("\n   ¿Qué ficha auditar? (número o 0 para todas): ").strip()

    if seleccion == "0":
        fichas_a_auditar = excels
    elif seleccion.isdigit() and 1 <= int(seleccion) <= len(excels):
        fichas_a_auditar = [excels[int(seleccion) - 1]]
    else:
        print("❌ Selección inválida")
        return

    service = conectar_drive()
    if not service:
        return

    for ruta_excel in fichas_a_auditar:
        auditar_ficha(ruta_excel, service)

    print(f"\n🎉 Proceso completado. Reportes en: {CARPETA_REPORTES}/")


if __name__ == "__main__":
    ejecutar_todas_las_fichas()