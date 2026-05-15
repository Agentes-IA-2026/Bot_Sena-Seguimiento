"""
supabase_client.py  —  Fábrica de cliente Supabase con SSL correcto

Uso en cualquier módulo del proyecto:
    from supabase_client import get_supabase
    db = get_supabase()
    db.table("actividades_parametrizadas").select("*").execute()

Reemplaza todos los create_client(SUPABASE_URL, SUPABASE_KEY) dispersos.
"""

import os

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SUPABASE_URL = os.environ.get(
    "SUPABASE_URL",
    "https://cfahgjytbpnmsogzryov.supabase.co",
)
SUPABASE_KEY = os.environ.get(
    "SUPABASE_KEY",
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
    "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImNmYWhnanl0YnBubXNvZ3pyeW92Iiw"
    "icm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NTYwODY2MSwiZXhwIjoyMDk"
    "xMTg0NjYxfQ.UXzWepH1HJ-KanXHoRZ3JgPK7Umt6WramF_fw26YNXM",
)


def get_supabase():
    """
    Retorna un cliente Supabase configurado con el CA bundle correcto.

    Estrategia (en orden):
      1. supabase-py >= 2.x: inyecta httpx.Client(verify=ca_bundle) vía
         ClientOptions — la forma limpia y explícita.
      2. Fallback: create_client estándar. Las variables REQUESTS_CA_BUNDLE
         y SSL_CERT_FILE seteadas por ssl_config.configure_environment()
         garantizan que urllib3 / httplib2 usen el bundle correcto.
    """
    from supabase import create_client
    from ssl_config import get_ca_bundle

    ca_bundle = get_ca_bundle()
    verify: str | bool = ca_bundle if ca_bundle else True

    try:
        import httpx
        from supabase import ClientOptions

        options = ClientOptions(
            http_client=httpx.Client(verify=verify),
        )
        return create_client(SUPABASE_URL, SUPABASE_KEY, options=options)

    except (ImportError, TypeError, AttributeError):
        # supabase-py antiguo o ClientOptions sin soporte http_client.
        # Los env vars del configure_environment() aplican igual.
        return create_client(SUPABASE_URL, SUPABASE_KEY)
