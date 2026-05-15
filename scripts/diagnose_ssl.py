#!/usr/bin/env python3
"""
scripts/diagnose_ssl.py  —  Diagnóstico de conectividad SSL/TLS

Uso:
    python scripts/diagnose_ssl.py

Prueba:
  - Python y certifi instalados
  - Variables de entorno SSL
  - Conexión HTTPS a Supabase
  - Conexión HTTPS a Google APIs
  - Detección de proxy corporativo / certificado self-signed
"""

import sys
import os
import socket

# Añadir el directorio raíz al path para poder importar ssl_config
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _ok(msg: str):
    print(f"  ✅  {msg}")

def _fail(msg: str):
    print(f"  ❌  {msg}")

def _warn(msg: str):
    print(f"  ⚠️   {msg}")

def _info(msg: str):
    print(f"  ℹ️   {msg}")

def _header(title: str):
    print(f"\n{'─'*60}")
    print(f"  {title}")
    print(f"{'─'*60}")


# ── 1. PYTHON Y CERTIFI ───────────────────────────────────────────────────────

_header("1. Python y certifi")
_info(f"Python: {sys.version}")

try:
    import certifi
    _ok(f"certifi instalado: {certifi.__version__}")
    _info(f"Ruta CA bundle: {certifi.where()}")
except ImportError:
    _fail("certifi NO está instalado — ejecuta: pip install certifi")


# ── 2. VARIABLES DE ENTORNO SSL ───────────────────────────────────────────────

_header("2. Variables de entorno SSL")
_SSL_VARS = (
    "SSL_CA_BUNDLE",
    "REQUESTS_CA_BUNDLE",
    "CURL_CA_BUNDLE",
    "SSL_CERT_FILE",
    "SSL_CERT_DIR",
    "HTTPS_PROXY",
    "HTTP_PROXY",
    "NO_PROXY",
)
for var in _SSL_VARS:
    val = os.environ.get(var)
    if val:
        _info(f"{var} = {val}")
    else:
        print(f"        {var} = (no definida)")


# ── 3. SSL_CONFIG DEL PROYECTO ────────────────────────────────────────────────

_header("3. ssl_config del proyecto")
try:
    from ssl_config import get_ca_bundle
    bundle = get_ca_bundle()
    if bundle:
        _ok(f"get_ca_bundle() → {bundle}")
    else:
        _fail("get_ca_bundle() retorna None — instala certifi o define SSL_CA_BUNDLE")
except ImportError:
    _warn("ssl_config.py no encontrado (ejecuta desde la raíz del proyecto)")


# ── 4. PRUEBAS DE CONEXIÓN HTTPS ──────────────────────────────────────────────

_header("4. Pruebas de conexión HTTPS")

ENDPOINTS = {
    "Supabase"        : ("cfahgjytbpnmsogzryov.supabase.co", 443),
    "Google APIs"     : ("www.googleapis.com", 443),
    "Google OAuth"    : ("oauth2.googleapis.com", 443),
    "certifi CDN"     : ("pypi.org", 443),
}

import ssl, socket

for nombre, (host, puerto) in ENDPOINTS.items():
    try:
        # Prueba 1: resolución DNS
        socket.gethostbyname(host)
    except socket.gaierror:
        _fail(f"{nombre}: no se pudo resolver DNS para {host}")
        continue

    # Prueba 2: handshake SSL con verificación
    try:
        ctx = ssl.create_default_context()
        bundle = None
        try:
            from ssl_config import get_ca_bundle
            bundle = get_ca_bundle()
        except ImportError:
            pass
        if bundle:
            ctx.load_verify_locations(cafile=bundle)

        with socket.create_connection((host, puerto), timeout=8) as sock:
            with ctx.wrap_socket(sock, server_hostname=host) as ssock:
                cert = ssock.getpeercert()
                issuer = dict(x[0] for x in cert.get("issuer", []))
                cn = issuer.get("commonName", "desconocido")
                _ok(f"{nombre} ({host}): SSL OK — issuer CN={cn}")

    except ssl.SSLCertVerificationError as e:
        _fail(f"{nombre} ({host}): SSL CERTIFICATE ERROR — {e}")
        _warn("Probable causa: proxy corporativo con SSL inspection o CA bundle desactualizado")
    except ssl.SSLError as e:
        _fail(f"{nombre} ({host}): SSL ERROR — {e}")
    except OSError as e:
        _fail(f"{nombre} ({host}): conexión fallida — {e}")


# ── 5. DETECCIÓN DE PROXY ─────────────────────────────────────────────────────

_header("5. Detección de proxy")

PROXIES = {k: v for k, v in os.environ.items()
           if k.upper() in ("HTTP_PROXY", "HTTPS_PROXY", "NO_PROXY", "ALL_PROXY")}

if PROXIES:
    _warn("Variables de proxy detectadas en el entorno:")
    for k, v in PROXIES.items():
        _info(f"  {k} = {v}")
    _warn("Si hay un proxy corporativo que intercepta HTTPS, necesitas exportar su CA")
else:
    _ok("No se detectaron variables de proxy en el entorno")


# ── 6. DIAGNÓSTICO DE LIBRERÍAS ───────────────────────────────────────────────

_header("6. Librerías HTTP del proyecto")

_libs = {
    "httpx"                    : "httpx",
    "supabase"                 : "supabase",
    "google-api-python-client" : "googleapiclient",
    "google-auth"              : "google.auth",
    "google-auth-httplib2"     : "google_auth_httplib2",
    "urllib3"                  : "urllib3",
}
for nombre, modulo in _libs.items():
    try:
        m = __import__(modulo)
        ver = getattr(m, "__version__", "?")
        _ok(f"{nombre} {ver}")
    except ImportError:
        _warn(f"{nombre} no instalado")


# ── 7. RESUMEN ────────────────────────────────────────────────────────────────

_header("7. Resumen y próximos pasos")

try:
    import certifi
    bundle_ok = True
except ImportError:
    bundle_ok = False
    _fail("Acción requerida: pip install certifi")

if not bundle_ok:
    print("""
  Pasos para resolver SSL en Windows (PowerShell):

  1. Actualizar pip y certifi:
     pip install --upgrade pip certifi httpx supabase

  2. Verificar el bundle:
     python -c "import certifi; print(certifi.where())"

  3. Si hay proxy corporativo, exportar su certificado raíz:
     $env:REQUESTS_CA_BUNDLE = "C:\\ruta\\al\\certificado_raiz.pem"
     $env:SSL_CERT_FILE       = "C:\\ruta\\al\\certificado_raiz.pem"

  4. Ejecutar el diagnóstico de nuevo:
     python scripts/diagnose_ssl.py
    """)
else:
    print("\n  Ejecuta el proyecto:")
    print("  python main.py")
    print("\n  Si aún falla con SSL, revisa SSL_FIX.md")
