"""
ssl_config.py  —  Configuración SSL centralizada

Orden de resolución del CA bundle:
  1. SSL_CA_BUNDLE     (variable de entorno explícita del proyecto)
  2. REQUESTS_CA_BUNDLE (estándar de requests / urllib3)
  3. CURL_CA_BUNDLE     (estándar de curl)
  4. certifi.where()   (CA bundle actualizado incluido con Python)

Punto de entrada principal:
    from ssl_config import configure_environment
    configure_environment()   # llamar UNA VEZ al arrancar main.py

También provee:
    get_ca_bundle()    → ruta del bundle activo
    build_ssl_context()→ ssl.SSLContext seguro listo para usar
    get_httpx_client() → httpx.Client con verify correcto
"""

import os
import ssl
import logging

logger = logging.getLogger(__name__)

_ENV_VARS = ("SSL_CA_BUNDLE", "REQUESTS_CA_BUNDLE", "CURL_CA_BUNDLE")

try:
    import certifi as _certifi
    _CERTIFI_BUNDLE = _certifi.where()
except ImportError:
    _CERTIFI_BUNDLE = None


# ── API PÚBLICA ───────────────────────────────────────────────────────────────

def get_ca_bundle() -> str | None:
    """
    Retorna la ruta al CA bundle activo.
    Prioridad: env vars → certifi → None.
    """
    for var in _ENV_VARS:
        val = os.environ.get(var, "").strip()
        if val:
            return val
    return _CERTIFI_BUNDLE


def configure_environment() -> str | None:
    """
    Configura REQUESTS_CA_BUNDLE, SSL_CERT_FILE y CURL_CA_BUNDLE con el
    bundle resuelto. Esto permite que requests, urllib3, google-auth y
    httplib2 usen automáticamente el CA bundle correcto sin cambios en su
    código de inicialización.

    Debe llamarse UNA VEZ al inicio del proceso, antes de cualquier HTTPS.
    Retorna la ruta aplicada (None si no hay bundle disponible).
    """
    bundle = get_ca_bundle()

    if not bundle:
        logger.warning(
            "ssl_config: no se encontró CA bundle. "
            "Instala certifi (pip install certifi) o define SSL_CA_BUNDLE."
        )
        print("   ⚠️  SSL: no hay CA bundle disponible — instala certifi")
        return None

    # Setear variables que leen requests, urllib3, google-auth, httplib2
    for var in ("REQUESTS_CA_BUNDLE", "SSL_CERT_FILE", "CURL_CA_BUNDLE"):
        if not os.environ.get(var):
            os.environ[var] = bundle

    # Detectar origen del bundle para el log
    origen = next(
        (v for v in _ENV_VARS[:1] if os.environ.get(v) == bundle),
        "certifi" if bundle == _CERTIFI_BUNDLE else "configurado",
    )
    print(f"   🔒 SSL CA bundle: {bundle}  ({origen})")
    return bundle


def build_ssl_context() -> ssl.SSLContext:
    """
    Construye un ssl.SSLContext estándar (no deshabilitado) usando el
    CA bundle activo. Útil para websockets o conexiones directas con asyncio.
    """
    ctx = ssl.create_default_context()
    bundle = get_ca_bundle()
    if bundle:
        ctx.load_verify_locations(cafile=bundle)
    return ctx


def get_httpx_client(**kwargs):
    """
    Retorna un httpx.Client (síncrono) con verify apuntando al CA bundle.
    Úsalo en lugar de httpx.Client() directo para Supabase y descargas HTTP.

    Ejemplo:
        from ssl_config import get_httpx_client
        with get_httpx_client(timeout=30) as client:
            resp = client.get("https://...")
    """
    try:
        import httpx
    except ImportError:
        raise ImportError("Instala httpx: pip install httpx")

    bundle = get_ca_bundle()
    verify: str | bool = bundle if bundle else True
    return httpx.Client(verify=verify, **kwargs)
