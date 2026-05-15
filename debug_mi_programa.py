"""
debug_mi_programa.py
Diagnóstico estático para la evidencia "Mi programa de Formación".

Ejecutar desde la raíz del proyecto:
    python debug_mi_programa.py

No requiere credenciales de Drive ni conexión a Supabase.
Muestra la normalización y el score contra nombres de archivo típicos.
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from bot.auditor import normalizar, _tokens, _score_patron, _norm

TARGET = "Mi programa de Formación"

ARCHIVOS_TIPICOS = [
    "Mi programa de Formación.pdf",
    "Mi Programa de Formacion.pdf",
    "Mi programa de formacion.docx",
    "mi_programa_de_formacion.pdf",
    "Mi-Programa-de-Formacion.pdf",
    "PROGRAMA DE FORMACION.pdf",
    "Programa Formacion.pdf",
    "Mi programa Formacion.pdf",
    "programa_formacion.pdf",
    # variantes con tildes distintas
    "Mi programa de Formación.docx",
    "Mi Programa de Formación.xlsx",
]

print("=" * 70)
print(f"PATRÓN BUSCADO : {repr(TARGET)}")
target_norm = normalizar(TARGET)
target_toks = _tokens(target_norm)
print(f"  normalizado  : {repr(target_norm)}")
print(f"  tokens       : {sorted(target_toks)}")
print("=" * 70)
print()
print(f"{'SCORE':>5}  {'NOMBRE EN DRIVE':<45}  NORM")
print("-" * 70)

for nombre_drive in ARCHIVOS_TIPICOS:
    drive_norm = normalizar(nombre_drive)
    drive_toks = _tokens(drive_norm)
    score      = _score_patron(target_norm, drive_norm)
    marca      = "  ** CANDIDATO" if score >= 50 else ""
    if score >= 75: marca = "  ** OK"
    print(f"{score:>5}  {nombre_drive:<45}  {drive_norm}{marca}")

print()
print("Umbrales: CANDIDATO >= 50  |  OK >= 75  (** = supera umbral)")
print()
print("Si el archivo real en Drive NO aparece arriba, pégalo a continuación")
print("como variable NOMBRE_REAL y re-ejecuta.")
print()

NOMBRE_REAL = None  # ← reemplazar con el nombre exacto del archivo en Drive si se conoce
if NOMBRE_REAL:
    n = normalizar(NOMBRE_REAL)
    t = _tokens(n)
    s = _score_patron(target_norm, n)
    print(f"NOMBRE REAL: {repr(NOMBRE_REAL)}")
    print(f"  norm  : {repr(n)}")
    print(f"  tokens: {sorted(t)}")
    print(f"  score : {s}  ({'CANDIDATO' if s >= 50 else 'NO CANDIDATO'})")
