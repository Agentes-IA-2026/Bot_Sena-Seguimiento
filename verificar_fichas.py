#!/usr/bin/env python3
"""
verificar_fichas.py
Valida la alineación entre el Excel madre (actividades_parametrizadas en Supabase)
y los registros reales de auditoría (verificaciones) para las fichas indicadas.

Uso:
    python verificar_fichas.py
    python verificar_fichas.py --fichas 3414936 3414937
    python verificar_fichas.py --fichas 3414936 --guia "Guía_01_Diagnóstico_Empresarial"

Criterio de éxito:
    Para cada (ficha, guía):
      - Actividades esperadas (Excel madre) == Actividades auditadas (verificaciones)
      → se imprime OK si coinciden, DESALINEADO si no.
"""

import os
import sys
import argparse
import unicodedata as _ud
import re as _re
from collections import defaultdict

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from supabase_client import get_supabase
from sincronizador_tabla import cargar_actividades

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────

DOCENTE_ID = os.environ.get("DOCENTE_ID", "d114d85c-fb7e-4a79-a446-fb90d652eec2")

FICHA_A_PROGRAMA = {
    "3414936": "Asistencia_comercial",
    "3414937": "Asistencia_comercial",
    "3415417": "Comunicacion_y_marketing",
    "3415087": "Comunicacion_y_marketing",
    "3458742": "Ventas_de_productos_en_linea",
}


# ── HELPERS ───────────────────────────────────────────────────────────────────

def _norm_guia_py(g: str) -> str:
    if not g: return ""
    g = g.lower().replace('_', ' ')
    g = _ud.normalize('NFD', g)
    g = ''.join(c for c in g if _ud.category(c) != 'Mn')
    return _re.sub(r'\s+', ' ', g).strip()

def _norm_evidencia_py(s: str) -> str:
    _STOP = frozenset(['de','del','la','el','los','las','un','una','y','o','a','en',
                       'con','por','al','que','se','su','sus','lo'])
    if not s: return ""
    s = s.lower()
    s = _re.sub(r'\.[a-z0-9]{1,5}(\s|$)', ' ', s)
    s = _ud.normalize('NFD', s)
    s = ''.join(c for c in s if _ud.category(c) != 'Mn')
    s = _re.sub(r'[_\-]', ' ', s)
    s = _re.sub(r'[^a-z0-9\s]', '', s)
    s = _re.sub(r'\s+', ' ', s).strip()
    tokens = [w for w in s.split() if len(w) > 1 and w not in _STOP]
    return ' '.join(tokens)

def _sep(char="─", n=60):
    print(char * n)

def _ok(msg):   print(f"  [OK]  {msg}")
def _warn(msg): print(f"  [!!]  {msg}")
def _info(msg): print(f"  [..] {msg}")
def _fail(msg): print(f"  [XX]  {msg}")


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────

def cargar_verificaciones(ficha: str, guia: str | None = None) -> list[dict]:
    """Carga registros de verificaciones de Supabase para una ficha (y guía opcional).
    El filtro de guía se aplica en Python con _norm_guia_py para tolerar diferencias
    de mayúsculas, tildes y guiones bajos entre lo almacenado y lo consultado."""
    try:
        db = get_supabase()
        resp = (db.table("verificaciones")
                  .select("*")
                  .eq("docente_id", DOCENTE_ID)
                  .eq("ficha", ficha)
                  .execute())
        rows = resp.data or []
        if guia:
            guia_norm = _norm_guia_py(guia)
            rows = [r for r in rows if _norm_guia_py(r.get("guia", "")) == guia_norm]
        return rows
    except Exception as e:
        _fail(f"Error cargando verificaciones para ficha {ficha}: {e}")
        return []


def guias_auditadas_en_supabase(ficha: str) -> list[str]:
    """Retorna las guías que ya tienen registros en verificaciones para esa ficha."""
    rows = cargar_verificaciones(ficha)
    return sorted({r["guia"] for r in rows if r.get("guia")})


# ── VALIDACIÓN POR FICHA Y GUÍA ───────────────────────────────────────────────

def validar_ficha_guia(ficha: str, guia: str, programa: str):
    """
    Compara actividades esperadas (Excel madre) con las auditadas (verificaciones).
    Imprime el resultado de alineación.
    """
    _sep("─")
    print(f"  Ficha: {ficha}  |  Programa: {programa}  |  Guía: {guia}")
    _sep("─")

    # ── Actividades esperadas desde el Excel madre (vía Supabase) ─────────────
    acts = cargar_actividades(guia=guia)
    n_esperadas = len(acts)

    if not acts:
        _fail(f"Sin actividades en actividades_parametrizadas para guía '{guia}'")
        _info("Ejecuta: python sincronizador_tabla.py para sincronizar el Excel maestro")
        return

    _info(f"Actividades en Excel madre    : {n_esperadas}")
    if n_esperadas > 0:
        for a in acts[:3]:
            print(f"       • [{a['actividad_id']}] '{a.get('nombre_esperado','')[:50]}'")
        if n_esperadas > 3:
            print(f"       … y {n_esperadas - 3} más")

    # ── Registros en verificaciones (ya auditados) ────────────────────────────
    rows = cargar_verificaciones(ficha, guia)
    n_auditadas = len(rows)
    estudiantes = sorted({r["estudiante"] for r in rows})
    n_aprendices = len(estudiantes)

    _info(f"Registros en verificaciones   : {n_auditadas}")
    _info(f"Aprendices con datos          : {n_aprendices}")

    if n_auditadas == 0:
        _warn(f"No hay datos auditados aún para ficha {ficha} + guía {guia}")
        _info("Ejecuta: python main.py → selecciona esta ficha y guía")
        return

    # ── Coherencia: actividades por aprendiz ──────────────────────────────────
    by_student: dict[str, dict] = defaultdict(dict)
    for r in rows:
        k = _norm_evidencia_py(r.get("evidencia", ""))
        if not k:
            continue
        # OR logic: entregada si al menos una versión del nombre lo registra como True
        by_student[r["estudiante"]][k] = by_student[r["estudiante"]].get(k, False) or r.get("entregado", False)

    print()
    issues = []
    for est in sorted(by_student.keys()):
        evs = by_student[est]
        n_est = len(evs)
        ok_est = sum(1 for v in evs.values() if v)
        pct = round(ok_est / n_est * 100) if n_est else 0

        if n_est != n_esperadas:
            issues.append(f"{est}: {n_est} actividades en verificaciones, se esperaban {n_esperadas}")
        status = "OK" if n_est == n_esperadas else "DESALINEADO"
        print(f"    [{status:<11}] {est[:45]:<45}  {ok_est}/{n_est}  ({pct}%)")

    print()
    if issues:
        _warn(f"{len(issues)} aprendice(s) con número de actividades distinto al Excel madre:")
        for iss in issues[:5]:
            print(f"       • {iss}")
        if len(issues) > 5:
            print(f"       … y {len(issues) - 5} más")
    else:
        _ok(f"Todos los aprendices tienen exactamente {n_esperadas} actividades registradas")


# ── RESUMEN POR FICHA ─────────────────────────────────────────────────────────

def resumen_ficha(ficha: str, programa: str, guia_filtro: str | None = None):
    _sep("═")
    print(f"FICHA {ficha}  —  Programa: {programa}")
    _sep("═")

    guias = guias_auditadas_en_supabase(ficha)

    if guia_filtro:
        guias = [g for g in guias if guia_filtro.lower() in g.lower()]
        if not guias:
            _warn(f"Ninguna guía auditada coincide con '{guia_filtro}' para ficha {ficha}")
            _info("Guías disponibles en verificaciones para esta ficha:")
            todas = guias_auditadas_en_supabase(ficha)
            for g in todas:
                print(f"       • {g}")
            return

    if not guias:
        _warn(f"No hay guías auditadas en Supabase para ficha {ficha}")
        _info("Ejecuta python main.py y audita esta ficha primero")

        # Mostrar qué guías DEBERÍA tener según el Excel madre
        acts = cargar_actividades()
        acts_programa = [a for a in acts if a.get("programa") == programa]
        guias_excel = sorted({a["guia"] for a in acts_programa if a.get("guia")})
        if guias_excel:
            _info(f"Guías del programa en Excel madre ({len(guias_excel)}):")
            for g in guias_excel:
                cnt = sum(1 for a in acts_programa if a.get("guia") == g)
                print(f"       • {g}  ({cnt} actividades)")
        return

    print(f"  Guías auditadas: {len(guias)}")
    for guia in guias:
        validar_ficha_guia(ficha, guia, programa)

    print()


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Valida alineación Excel madre ↔ verificaciones por ficha y guía.",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--fichas", nargs="+", default=["3414936", "3414937"],
        help="Fichas a verificar (default: 3414936 3414937)",
    )
    parser.add_argument(
        "--guia", default=None,
        help="Filtrar por guía específica (parcial, sin mayúsculas)",
    )
    args = parser.parse_args()

    print()
    print("════════════════════════════════════════════════════════════")
    print("   VERIFICADOR DE ALINEACIÓN  —  GuíaBot")
    print("   Compara Excel madre (actividades_parametrizadas)")
    print("   con auditorías reales (verificaciones en Supabase)")
    print("════════════════════════════════════════════════════════════")
    print()

    fichas_desconocidas = [f for f in args.fichas if f not in FICHA_A_PROGRAMA]
    if fichas_desconocidas:
        _warn(f"Fichas no mapeadas en FICHA_A_PROGRAMA: {fichas_desconocidas}")
        _info("Agrégalas en main.py → FICHA_A_PROGRAMA")

    for ficha in args.fichas:
        programa = FICHA_A_PROGRAMA.get(ficha)
        if not programa:
            _fail(f"Ficha {ficha} sin programa asignado — saltando")
            continue
        resumen_ficha(ficha, programa, args.guia)

    print("════════════════════════════════════════════════════════════")
    print("   FIN DE VALIDACIÓN")
    print("════════════════════════════════════════════════════════════")
    print()


if __name__ == "__main__":
    main()
