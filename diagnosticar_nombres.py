#!/usr/bin/env python3
"""
diagnosticar_nombres.py  —  Diagnostica el estado de nombre_esperado en Supabase

Uso:
    python diagnosticar_nombres.py
    python diagnosticar_nombres.py --guia "Guía_01_Diagnóstico_Empresarial"
    python diagnosticar_nombres.py --programa Asistencia_comercial
    python diagnosticar_nombres.py --guia "Guía_01_..." --verbose

Muestra por cada actividad:
    actividad_id | nombre_esperado | actividad_resumen | patrón final | estado
"""

import sys
import os
import argparse

# Garantizar que SSL funciona antes de cualquier conexión
try:
    from ssl_config import configure_environment
    configure_environment()
except ImportError:
    pass

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from sincronizador_tabla import cargar_actividades


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _ok(msg): print(f"  ✅  {msg}")
def _warn(msg): print(f"  ⚠️   {msg}")
def _fail(msg): print(f"  ❌  {msg}")
def _info(msg): print(f"  ℹ️   {msg}")


# ── DIAGNÓSTICO ───────────────────────────────────────────────────────────────

def diagnosticar(
    programa: str | None = None,
    guia: str | None = None,
    verbose: bool = False,
):
    print(f"\n{'═'*70}")
    print(f"  DIAGNÓSTICO DE NOMBRE_ESPERADO")
    filtros = []
    if programa: filtros.append(f"programa={programa}")
    if guia:     filtros.append(f"guia={guia}")
    print(f"  Filtros: {' | '.join(filtros) or 'ninguno (todas las actividades)'}")
    print(f"{'═'*70}\n")

    actividades = cargar_actividades(programa=programa, guia=guia)

    if not actividades:
        _fail("No se encontraron actividades con esos filtros.")
        _info("Verifica que el programa y guía existen en actividades_parametrizadas.")
        print(f"\nProgramas disponibles:")
        todos = cargar_actividades()
        programas = sorted({a.get("programa","") for a in todos if a.get("programa")})
        for p in programas:
            print(f"   • {p}")
        return

    # Agrupar por guía
    from collections import defaultdict
    por_guia: dict[str, list[dict]] = defaultdict(list)
    for a in actividades:
        por_guia[a.get("guia", "(sin guía)")].append(a)

    total_vacios = 0
    total_acts   = 0

    for nombre_guia, acts in sorted(por_guia.items()):
        vacios = [a for a in acts if not (a.get("nombre_esperado") or "").strip()]
        pct    = len(vacios) / len(acts) * 100 if acts else 0
        total_vacios += len(vacios)
        total_acts   += len(acts)

        estado_guia = "🚨 CRÍTICO" if pct >= 20 else ("⚠️  PARCIAL" if vacios else "✅ OK")
        print(f"{'─'*70}")
        print(f"  Guía: {nombre_guia}")
        print(f"  Actividades: {len(acts)}  |  Sin nombre_esperado: {len(vacios)} ({pct:.0f}%)  {estado_guia}")
        print()

        # Encabezado de tabla
        print(f"  {'ID':<16}  {'nombre_esperado':<35}  {'resumen':<30}  patrón_final")
        print(f"  {'─'*16}  {'─'*35}  {'─'*30}  {'─'*20}")

        for a in sorted(acts, key=lambda x: x.get("actividad_id","") or ""):
            a_id     = (a.get("actividad_id") or "?")[:15]
            n_esp    = (a.get("nombre_esperado") or "").strip()
            n_res    = (a.get("actividad_resumen") or "").strip()
            n_vars   = len(a.get("variantes_permitidas") or [])
            carpeta  = a.get("carpeta_drive") or ""
            tipo     = a.get("tipo_archivo") or "CUALQUIER_FORMATO"

            if n_esp:
                patron_final = n_esp
                patron_src   = "nombre_esperado"
                flag         = ""
            elif n_res:
                patron_final = f"[fallback] {n_res}"
                patron_src   = "actividad_resumen"
                flag         = " ⚠️"
            else:
                patron_final = "(sin patrón)"
                patron_src   = "VACÍO"
                flag         = " ❌"

            row_esp = (n_esp[:33] + "…") if len(n_esp) > 35 else n_esp
            row_res = (n_res[:28] + "…") if len(n_res) > 30 else n_res
            row_pat = (patron_final[:18] + "…") if len(patron_final) > 20 else patron_final

            print(f"  {a_id:<16}  {row_esp or '(vacío)':<35}  {row_res:<30}  {row_pat}{flag}")

            if verbose:
                if n_vars:
                    print(f"  {'':>16}  variantes ({n_vars}): {a.get('variantes_permitidas','')}")
                if carpeta:
                    print(f"  {'':>16}  carpeta: {carpeta}  tipo: {tipo}")
        print()

    # Resumen global
    pct_global = total_vacios / total_acts * 100 if total_acts else 0
    print(f"{'═'*70}")
    print(f"  RESUMEN GLOBAL")
    print(f"  Total actividades : {total_acts}")
    print(f"  Con nombre_esperado vacío : {total_vacios} ({pct_global:.0f}%)")

    if total_vacios == 0:
        _ok("Todos los nombre_esperado están rellenos. El matching debería funcionar.")
    elif pct_global >= 50:
        _fail(f"El {pct_global:.0f}% está vacío. El bot usará actividad_resumen como fallback.")
        _info("Solución definitiva: rellenar la columna nombre_esperado en el Excel maestro")
        _info("y ejecutar: python sincronizador_tabla.py" + (f" --guia \"{guia}\"" if guia else ""))
    elif total_vacios > 0:
        _warn(f"{total_vacios} actividades sin nombre_esperado usarán fallback de actividad_resumen.")
        _info("Recomendado: rellenar nombre_esperado para mejorar precisión del matching.")
    print(f"{'═'*70}\n")


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Diagnostica el estado de nombre_esperado en actividades_parametrizadas.",
    )
    parser.add_argument("--guia",     help="Filtrar por guía (parcial)")
    parser.add_argument("--programa", help="Filtrar por programa canónico")
    parser.add_argument("--verbose",  action="store_true",
                        help="Mostrar variantes y carpeta_drive por actividad")

    args = parser.parse_args()
    diagnosticar(programa=args.programa, guia=args.guia, verbose=args.verbose)
