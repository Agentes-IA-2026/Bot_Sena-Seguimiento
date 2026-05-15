"""
sincronizador_tabla.py
Lee "Documento Base de Guías.xlsx" y sincroniza la tabla
actividades_parametrizadas en Supabase (fuente única de verdad).

Uso:
    python sincronizador_tabla.py                          # todo
    python sincronizador_tabla.py --programa Asistencia_comercial
    python sincronizador_tabla.py --guia GUIA_00_INDUCCION
    python sincronizador_tabla.py --dry-run                # sin escribir
    python sincronizador_tabla.py --excel ruta/otro.xlsx

Importable desde otros módulos:
    from sincronizador_tabla import cargar_actividades
    actividades = cargar_actividades(programa="Asistencia_comercial",
                                     guia="GUIA_01_DIAGNOSTICO_EMPRESARIAL")
"""

import os
import sys
import argparse
from collections import Counter
from datetime import datetime, timezone

import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from supabase_client import get_supabase, SUPABASE_URL, SUPABASE_KEY

TABLA = "actividades_parametrizadas"

# Ruta por defecto — maneja el nombre con o sin tilde en "Guías"
_RUTA_CON_TILDE    = os.path.join("assets", "Guias_Referencia", "Documento Base de Guías.xlsx")
_RUTA_SIN_TILDE    = os.path.join("assets", "Guias_Referencia", "Documento Base de Guias.xlsx")
RUTA_EXCEL_DEFAULT = _RUTA_CON_TILDE if os.path.exists(_RUTA_CON_TILDE) else _RUTA_SIN_TILDE

COLUMNAS_REQUERIDAS = frozenset({
    "programa", "guia", "actividad_id", "actividad_resumen",
    "texto_fuente", "nombre_esperado", "variantes_permitidas",
    "tipo_archivo", "obligatoria", "carpeta_drive",
    "regla_validacion", "observaciones",
})

LOTE = 200  # registros por batch hacia Supabase


# ── FUNCIONES DE LIMPIEZA ─────────────────────────────────────────────────────

def _texto(valor) -> str:
    """Convierte cualquier valor de celda a str limpio; NaN → ''."""
    if valor is None:
        return ""
    s = str(valor).strip()
    return "" if s.lower() in ("nan", "none", "<na>") else s


def _booleano(valor) -> bool:
    """
    Parsea 'TRUE/FALSE', 'SÍ/NO', 1/0, True/False → bool.
    Default: True (actividades son obligatorias salvo indicación contraria).
    """
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return bool(valor) and not (isinstance(valor, float) and valor != valor)  # nan → False
    return _texto(valor).upper() in ("TRUE", "VERDADERO", "SÍ", "SI", "S", "1", "YES", "Y")


def _variantes(valor) -> list[str]:
    """
    Parsea variantes separadas por ';' a lista limpia.
    Ej: "var 1;Var 2; var3" → ["var 1", "Var 2", "var3"]
    """
    s = _texto(valor)
    if not s:
        return []
    return [v.strip() for v in s.split(";") if v.strip()]


# ── CLASE PRINCIPAL ───────────────────────────────────────────────────────────

class SincronizadorTabla:
    """
    Lee el Excel de guías y hace upsert en la tabla actividades_parametrizadas.
    La clave de conflicto es (programa, guia, actividad_id).
    """

    def __init__(
        self,
        ruta_excel: str = RUTA_EXCEL_DEFAULT,
        programa: str | None = None,
        guia: str | None = None,
        dry_run: bool = False,
    ):
        self.ruta_excel      = ruta_excel
        self.filtro_programa = programa
        self.filtro_guia     = guia
        self.dry_run         = dry_run
        self._db             = None

    def _supabase(self):
        if self._db is None:
            self._db = get_supabase()
        return self._db

    # ── CARGA ──────────────────────────────────────────────────────────────

    def cargar_excel(self) -> pd.DataFrame:
        """
        Lee el Excel con dtype=str para evitar conversiones automáticas de pandas.
        Normaliza los nombres de columna a minúsculas con guiones bajos.
        """
        if not os.path.exists(self.ruta_excel):
            raise FileNotFoundError(
                f"No se encontró: {self.ruta_excel}\n"
                "Verifica la ruta o pasa --excel <ruta>."
            )

        print(f"📂 Leyendo: {self.ruta_excel}")
        df = pd.read_excel(self.ruta_excel, dtype=str)

        # Normalizar encabezados: strip + minúsculas + guiones bajos
        df.columns = [
            col.strip().lower().replace(" ", "_")
            for col in df.columns
        ]

        faltantes = COLUMNAS_REQUERIDAS - set(df.columns)
        if faltantes:
            raise ValueError(
                f"Columnas faltantes en el Excel: {sorted(faltantes)}\n"
                f"Columnas encontradas: {sorted(df.columns.tolist())}"
            )

        print(f"   ✅ {len(df)} filas leídas, columnas validadas")
        return df

    # ── FILTRADO ───────────────────────────────────────────────────────────

    def filtrar(self, df: pd.DataFrame) -> pd.DataFrame:
        if self.filtro_programa:
            mask = df["programa"].str.strip().str.lower().str.contains(
                self.filtro_programa.lower(), na=False
            )
            df = df[mask].copy()
            print(f"   🔍 Filtro --programa '{self.filtro_programa}': {len(df)} filas")

        if self.filtro_guia:
            mask = df["guia"].str.strip().str.lower().str.contains(
                self.filtro_guia.lower(), na=False
            )
            df = df[mask].copy()
            print(f"   🔍 Filtro --guia '{self.filtro_guia}': {len(df)} filas")

        return df

    # ── TRANSFORMACIÓN ─────────────────────────────────────────────────────

    def transformar(self, df: pd.DataFrame) -> tuple[list[dict], list[str]]:
        """
        Convierte el DataFrame a lista de dicts listos para upsert.
        Retorna (registros_validos, lista_de_advertencias).
        """
        registros    = []
        advertencias = []
        ahora        = datetime.now(tz=timezone.utc).isoformat()

        for idx, row in df.iterrows():
            numero_fila  = idx + 2  # +2: encabezado en fila 1 + base 0
            programa     = _texto(row.get("programa"))
            guia         = _texto(row.get("guia"))
            actividad_id = _texto(row.get("actividad_id"))

            # Clave compuesta vacía → fila vacía o de encabezado repetido
            if not programa or not guia or not actividad_id:
                if programa or guia or actividad_id:
                    advertencias.append(
                        f"Fila {numero_fila}: clave incompleta "
                        f"(programa='{programa}', guia='{guia}', "
                        f"actividad_id='{actividad_id}') — omitida"
                    )
                continue

            nombre_esperado = _texto(row.get("nombre_esperado"))
            obligatoria     = _booleano(row.get("obligatoria", True))

            if not nombre_esperado:
                if obligatoria:
                    advertencias.append(
                        f"Fila {numero_fila}: ⚠️  {actividad_id} es obligatoria "
                        f"pero no tiene nombre_esperado — omitida"
                    )
                    continue
                else:
                    advertencias.append(
                        f"Fila {numero_fila}: ℹ️  {actividad_id} no tiene "
                        f"nombre_esperado (opcional) — incluida sin patrón"
                    )

            tipo_archivo = _texto(row.get("tipo_archivo")) or "CUALQUIER_FORMATO"
            carpeta      = _texto(row.get("carpeta_drive")) or None
            regla        = _texto(row.get("regla_validacion")) or None
            obs          = _texto(row.get("observaciones")) or None

            registros.append({
                "programa"            : programa,
                "guia"                : guia,
                "actividad_id"        : actividad_id,
                "actividad_resumen"   : _texto(row.get("actividad_resumen")) or None,
                "texto_fuente"        : _texto(row.get("texto_fuente")) or None,
                "nombre_esperado"     : nombre_esperado or None,
                "variantes_permitidas": _variantes(row.get("variantes_permitidas")),
                "tipo_archivo"        : tipo_archivo,
                "obligatoria"         : obligatoria,
                "carpeta_drive"       : carpeta,
                "regla_validacion"    : regla,
                "observaciones"       : obs,
                "activa"              : True,
                "updated_at"          : ahora,
            })

        return registros, advertencias

    # ── UPSERT ─────────────────────────────────────────────────────────────

    def sincronizar(self, registros: list[dict]) -> dict:
        """
        Upsert en lotes de LOTE registros.
        La restricción de conflicto (programa, guia, actividad_id) debe existir
        en la tabla antes de ejecutar (ver sql/crear_actividades_parametrizadas.sql).
        """
        if not registros:
            print("   ⚠️  Sin registros para sincronizar")
            return {"ok": 0, "errores": 0}

        db      = self._supabase()
        total_ok  = 0
        total_err = 0

        for i in range(0, len(registros), LOTE):
            lote     = registros[i : i + LOTE]
            num_lote = i // LOTE + 1
            try:
                db.table(TABLA).upsert(
                    lote,
                    on_conflict="programa,guia,actividad_id",
                ).execute()
                total_ok += len(lote)
                print(f"   ☁️  Lote {num_lote}: {len(lote)} registros ✅")
            except Exception as e:
                total_err += len(lote)
                print(f"   ❌ Error en lote {num_lote}: {e}")

        return {"ok": total_ok, "errores": total_err}

    # ── FLUJO COMPLETO ─────────────────────────────────────────────────────

    def ejecutar(self) -> dict:
        modo = "DRY-RUN (sin cambios en Supabase)" if self.dry_run else "PRODUCCIÓN"
        print(f"\n{'='*60}")
        print(f"🔄 SINCRONIZADOR — Documento Base de Guías")
        print(f"   Modo  : {modo}")
        print(f"   Tabla : {TABLA}")
        print(f"{'='*60}\n")

        df = self.cargar_excel()
        df = self.filtrar(df)

        if df.empty:
            print("⚠️  Sin filas tras aplicar filtros. Revisa --programa / --guia.")
            return {"filas_excel": 0, "registros_validos": 0}

        registros, advertencias = self.transformar(df)

        if advertencias:
            print(f"\n⚠️  {len(advertencias)} advertencia(s):")
            for av in advertencias:
                print(f"   {av}")

        descartadas = len(df) - len(registros)
        print(f"\n📊 Resumen:")
        print(f"   Filas en Excel      : {len(df)}")
        print(f"   Registros válidos   : {len(registros)}")
        print(f"   Descartadas/avisos  : {descartadas}")

        por_programa = Counter(r["programa"] for r in registros)
        print(f"\n   Distribución por programa:")
        for prog, cnt in sorted(por_programa.items()):
            print(f"   • {prog}: {cnt} actividades")

        if self.dry_run:
            print(f"\n✅ Dry-run completado — {len(registros)} registros preparados, "
                  f"nada escrito en Supabase.")
            return {
                "filas_excel"      : len(df),
                "registros_validos": len(registros),
                "dry_run"          : True,
            }

        print(f"\n⬆️  Sincronizando con Supabase...")
        resultado = self.sincronizar(registros)

        print(f"\n{'='*60}")
        print(f"✅ Sincronización completada")
        print(f"   OK     : {resultado['ok']}")
        print(f"   Errores: {resultado['errores']}")
        print(f"{'='*60}")

        return {
            "filas_excel"      : len(df),
            "registros_validos": len(registros),
            **resultado,
        }


# ── API PÚBLICA PARA OTROS MÓDULOS ────────────────────────────────────────────

def cargar_actividades(
    programa: str | None = None,
    guia: str | None = None,
    solo_obligatorias: bool = False,
) -> list[dict]:
    """
    Consulta actividades_parametrizadas desde Supabase.
    Usado por auditor.py (Fase 2) y core.py como reemplazo de LISTAS_MANUALES.

    Retorna lista de dicts con todos los campos de la tabla.
    Retorna [] si hay error de conexión (el llamador decide el fallback).
    """
    try:
        db    = get_supabase()
        query = db.table(TABLA).select("*").eq("activa", True)

        if programa:
            query = query.eq("programa", programa)
        if guia:
            query = query.eq("guia", guia)
        if solo_obligatorias:
            query = query.eq("obligatoria", True)

        resp = query.execute()
        return resp.data or []

    except Exception as e:
        print(f"⚠️  cargar_actividades: no se pudo conectar a Supabase: {e}")
        return []


# ── CLI ───────────────────────────────────────────────────────────────────────

def _cli():
    parser = argparse.ArgumentParser(
        description="Sincroniza el Excel de guías con Supabase (tabla actividades_parametrizadas).",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--excel",
        default=RUTA_EXCEL_DEFAULT,
        help=f"Ruta al Excel.\nPor defecto: {RUTA_EXCEL_DEFAULT}",
    )
    parser.add_argument(
        "--programa",
        default=None,
        help="Filtrar por programa (parcial, sin mayúsculas).\n"
             "Ej: --programa asistencia",
    )
    parser.add_argument(
        "--guia",
        default=None,
        help="Filtrar por guía (parcial).\n"
             "Ej: --guia GUIA_00",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Solo muestra qué haría; no escribe nada en Supabase.",
    )

    args = parser.parse_args()

    sync = SincronizadorTabla(
        ruta_excel = args.excel,
        programa   = args.programa,
        guia       = args.guia,
        dry_run    = args.dry_run,
    )

    try:
        resultado = sync.ejecutar()
        sys.exit(0 if resultado.get("errores", 0) == 0 else 1)
    except (FileNotFoundError, ValueError) as e:
        print(f"\n❌ {e}")
        sys.exit(1)


if __name__ == "__main__":
    _cli()
