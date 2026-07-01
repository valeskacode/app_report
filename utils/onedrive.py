# -*- coding: utf-8 -*-
"""
utils/onedrive.py
Integración con Microsoft Graph API para subir reportes (Word/PDF),
el historial Excel y las fotos de visita directamente a OneDrive,
sin necesitar el cliente de escritorio instalado en el servidor.

Requiere las credenciales de la app registrada en Azure AD:
  - CLIENT_ID   (ID de la aplicación)
  - CLIENT_SECRET (secreto de cliente)
  - TENANT_ID   (ID del directorio / inquilino)

Estas tres credenciales se configuran como variables de entorno en
Streamlit Cloud (Settings > Secrets), sin tocar el código.
"""
import io
import os

import requests

# --------------------------------------------------------------------------
# CONFIGURACIÓN — lee desde variables de entorno (Streamlit Secrets o .env)
# --------------------------------------------------------------------------
CLIENT_ID     = os.environ.get("GRAPH_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")
TENANT_ID     = os.environ.get("GRAPH_TENANT_ID", "")

# Correo / UPN del usuario de OneDrive donde se guardarán los archivos.
# Ej: "auditoria@cajaarequipa.com.pe"
ONEDRIVE_USER = os.environ.get("GRAPH_ONEDRIVE_USER", "")

# Ruta de la carpeta DENTRO de ese OneDrive donde caerán los reportes.
# Ej: "Auditoria/VisitaClientes/Reportes"
# Deja vacío ("") para que quede en la raíz del OneDrive.
ONEDRIVE_CARPETA = os.environ.get("GRAPH_ONEDRIVE_CARPETA", "Auditoria/VisitaClientes")

# URL base de Graph API
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# --------------------------------------------------------------------------
# TOKEN — Client Credentials Flow (app-only, sin login interactivo).
# Necesita permiso "Files.ReadWrite.All" en la app de Azure.
# --------------------------------------------------------------------------
_token_cache: dict = {}


def _obtener_token() -> str:
    """Obtiene (o reutiliza en caché) un token de acceso de Graph API
    usando el flujo Client Credentials (sin login del usuario)."""
    import time
    ahora = time.time()
    if _token_cache.get("expires_at", 0) > ahora + 60:
        return _token_cache["access_token"]

    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    }, timeout=15)
    resp.raise_for_status()
    data = resp.json()
    _token_cache["access_token"] = data["access_token"]
    _token_cache["expires_at"] = ahora + data.get("expires_in", 3600)
    return _token_cache["access_token"]


def _headers() -> dict:
    return {"Authorization": f"Bearer {_obtener_token()}"}


def credenciales_configuradas() -> bool:
    """True si las tres credenciales mínimas están presentes."""
    return bool(CLIENT_ID and CLIENT_SECRET and TENANT_ID and ONEDRIVE_USER)


# --------------------------------------------------------------------------
# OPERACIONES CON ONEDRIVE
# --------------------------------------------------------------------------

def _ruta_onedrive(nombre_archivo: str, subcarpeta: str = "") -> str:
    """Construye la ruta completa dentro del OneDrive del usuario."""
    base = ONEDRIVE_CARPETA.strip("/")
    if subcarpeta:
        base = f"{base}/{subcarpeta.strip('/')}"
    if base:
        return f"{base}/{nombre_archivo}"
    return nombre_archivo


def _upload_url(ruta_en_onedrive: str) -> str:
    """URL de Graph API para subir un archivo por ruta."""
    ruta_enc = requests.utils.quote(ruta_en_onedrive, safe="/")
    return f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive/root:/{ruta_enc}:/content"


def subir_archivo(nombre_archivo: str, contenido_bytes: bytes,
                  subcarpeta: str = "") -> tuple[bool, str]:
    """Sube un archivo a OneDrive.

    Parámetros:
        nombre_archivo: nombre del archivo (con extensión).
        contenido_bytes: contenido del archivo en bytes.
        subcarpeta: subcarpeta adicional dentro de ONEDRIVE_CARPETA.

    Retorna:
        (True, url_web_del_archivo) si tuvo éxito.
        (False, mensaje_de_error)   si falló.
    """
    if not credenciales_configuradas():
        return False, "Credenciales de Graph API no configuradas."
    try:
        ruta = _ruta_onedrive(nombre_archivo, subcarpeta)
        url  = _upload_url(ruta)
        # Graph API acepta archivos ≤ 4 MB con un PUT simple.
        # Para archivos más grandes habría que usar upload session.
        resp = requests.put(url, headers=_headers(), data=contenido_bytes,
                            timeout=60)
        resp.raise_for_status()
        web_url = resp.json().get("webUrl", "")
        return True, web_url
    except Exception as e:
        return False, str(e)


def subir_reporte(nombre_archivo: str, contenido_bytes: bytes) -> tuple[bool, str]:
    """Atajo para subir un reporte Word/PDF a la subcarpeta 'Reportes'."""
    return subir_archivo(nombre_archivo, contenido_bytes, subcarpeta="Reportes")


def subir_historial(contenido_bytes: bytes) -> tuple[bool, str]:
    """Sube el Excel de historial de visitas a OneDrive."""
    return subir_archivo("historial_visitas.xlsx", contenido_bytes, subcarpeta="Historial")


def subir_foto(nombre_archivo: str, foto_bytes: bytes,
               agencia: str = "") -> tuple[bool, str]:
    """Sube una foto de verificación a la subcarpeta Fotos/<agencia>."""
    sub = f"Fotos/{agencia}" if agencia else "Fotos"
    return subir_archivo(nombre_archivo, foto_bytes, subcarpeta=sub)


def listar_carpeta(subcarpeta: str = "") -> list[dict]:
    """Lista los archivos en la carpeta configurada (o subcarpeta).

    Retorna lista de dicts con: name, size, lastModifiedDateTime, webUrl.
    """
    if not credenciales_configuradas():
        return []
    try:
        base = ONEDRIVE_CARPETA.strip("/")
        if subcarpeta:
            base = f"{base}/{subcarpeta.strip('/')}"
        ruta_enc = requests.utils.quote(base, safe="/")
        url = f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive/root:/{ruta_enc}:/children"
        resp = requests.get(url, headers=_headers(), timeout=15)
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return [
            {
                "name": i.get("name"),
                "size": i.get("size"),
                "fecha": i.get("lastModifiedDateTime", "")[:10],
                "webUrl": i.get("webUrl", ""),
            }
            for i in items if not i.get("folder")  # solo archivos, no carpetas
        ]
    except Exception:
        return []


def test_conexion() -> tuple[bool, str]:
    """Prueba rápida de conexión — verifica el token y el acceso al drive."""
    if not credenciales_configuradas():
        return False, "Falta configurar CLIENT_ID, CLIENT_SECRET, TENANT_ID o ONEDRIVE_USER."
    try:
        _obtener_token()
        url  = f"{GRAPH_URL}/users/{ONEDRIVE_USER}/drive"
        resp = requests.get(url, headers=_headers(), timeout=10)
        resp.raise_for_status()
        nombre = resp.json().get("owner", {}).get("user", {}).get("displayName", ONEDRIVE_USER)
        return True, f"Conectado correctamente al OneDrive de: {nombre}"
    except requests.HTTPError as e:
        codigo = e.response.status_code if e.response else "?"
        if codigo == 401:
            return False, "Error 401 — Credenciales incorrectas o secreto vencido."
        if codigo == 403:
            return False, "Error 403 — La app no tiene permiso 'Files.ReadWrite.All' en Azure."
        if codigo == 404:
            return False, f"Error 404 — El usuario '{ONEDRIVE_USER}' no se encontró en este tenant."
        return False, f"Error HTTP {codigo}: {e}"
    except Exception as e:
        return False, f"Error de conexión: {e}"
