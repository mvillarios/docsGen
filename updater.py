"""
updater.py
──────────────────────────────────────────────────────────────────────────────
Módulo de auto-actualización.
Compara la versión local con el último release publicado en GitHub y,
si hay una versión más nueva, ofrece descargarla e instalarla.

Para modificar en el futuro:
  - Cambiar la fuente de actualizaciones: edita `_obtener_ultimo_release()`
  - Cambiar la lógica de reemplazo:       edita `_descargar_hilo()`
  - Cambiar cómo se comparan versiones:   edita `_hay_version_nueva()`
──────────────────────────────────────────────────────────────────────────────
"""

import os
import sys
import json
import shutil
import logging
import threading
import subprocess
import urllib.request
import urllib.error
from pathlib import Path

# ── Configuración ─────────────────────────────────────────────────────────────
GITHUB_USER  = "mvillarios"   # ← Cambia esto
GITHUB_REPO  = "docsGen"      # ← Cambia esto
TIMEOUT_SEG  = 5              # Segundos de espera para la petición HTTP

# Versión embebida — el workflow la reemplaza automáticamente antes de compilar
VERSION_APP  = "1.0.0"
# ─────────────────────────────────────────────────────────────────────────────

# ── Logger ────────────────────────────────────────────────────────────────────
def _iniciar_log():
    if getattr(sys, "frozen", False):
        base = os.path.join(os.environ.get("APPDATA", ""), "GeneradorDocumentos")
        os.makedirs(base, exist_ok=True)
        log_path = os.path.join(base, "updater.log")
    else:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "updater.log")

    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        encoding="utf-8",
    )

_iniciar_log()
log = logging.getLogger(__name__)
# ─────────────────────────────────────────────────────────────────────────────


def leer_version_local() -> str:
    """Retorna la versión embebida en el ejecutable."""
    return VERSION_APP


def verificar_actualizacion(on_update_available: callable, on_error: callable = None):
    """
    Lanza la verificación en un hilo secundario para no bloquear la UI.

    Parámetros:
        on_update_available(version, url_descarga): llamado si hay versión nueva.
        on_error(mensaje):                          llamado si hay un error (opcional).
    """
    hilo = threading.Thread(
        target=_verificar_hilo,
        args=(on_update_available, on_error),
        daemon=True,
    )
    hilo.start()


def descargar_e_instalar(url_descarga: str, on_progreso: callable = None, on_error: callable = None):
    """
    Descarga el nuevo .exe y lo reemplaza en un hilo secundario.

    Parámetros:
        url_descarga:         URL directa al .exe del release de GitHub.
        on_progreso(pct:int): llamado con porcentaje 0-100 durante la descarga.
        on_error(mensaje):    llamado si la descarga falla.
    """
    hilo = threading.Thread(
        target=_descargar_hilo,
        args=(url_descarga, on_progreso, on_error),
        daemon=True,
    )
    hilo.start()

def limpiar_archivos_viejos():
    """Elimina el .exe anterior si quedó de una actualización previa."""
    exe_viejo = _ruta_exe_actual() + ".old"
    if os.path.exists(exe_viejo):
        try:
            os.remove(exe_viejo)
            log.info("Archivo .old eliminado correctamente.")
        except Exception as e:
            log.warning(f"No se pudo eliminar el archivo .old: {e}")

# ── Funciones internas ────────────────────────────────────────────────────────

def _verificar_hilo(on_update_available, on_error):
    limpiar_archivos_viejos()
    version_local = leer_version_local()
    log.info(f"Verificando actualizaciones — versión local: {version_local}")
    try:
        release = _obtener_ultimo_release()
        if release is None:
            log.info("Sin releases disponibles o sin conexión a internet.")
            return

        version_remota = release["tag_name"].lstrip("v")
        log.info(f"Versión remota: {version_remota}")

        if _hay_version_nueva(version_local, version_remota):
            log.info(f"Nueva versión disponible: {version_remota}")
            url = _obtener_url_exe(release)
            if url:
                on_update_available(version_remota, url)
            else:
                log.warning("El release no tiene un .exe adjunto.")
        else:
            log.info("La aplicación está actualizada.")

    except Exception as e:
        log.error(f"Error al verificar actualizaciones: {e}", exc_info=True)
        if on_error:
            on_error(str(e))


def _obtener_ultimo_release() -> dict | None:
    """Consulta la API de GitHub y retorna el dict del último release."""
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "GeneradorDocumentos-Updater/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT_SEG) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None   # Repo sin releases todavía
        raise
    except urllib.error.URLError:
        return None       # Sin conexión a internet


def _obtener_url_exe(release: dict) -> str | None:
    """Extrae la URL de descarga del primer .exe en los assets del release."""
    for asset in release.get("assets", []):
        if asset["name"].endswith(".exe"):
            return asset["browser_download_url"]
    return None


def _hay_version_nueva(local: str, remota: str) -> bool:
    """
    Compara versiones semánticas (mayor.menor.parche).
    Retorna True si la versión remota es mayor que la local.
    """
    def _partes(v: str):
        try:
            return tuple(int(x) for x in v.split("."))
        except ValueError:
            return (0, 0, 0)

    return _partes(remota) > _partes(local)


def _descargar_hilo(url_descarga: str, on_progreso, on_error):
    exe_viejo = None
    exe_nuevo = None
    try:
        exe_actual = _ruta_exe_actual()
        exe_nuevo  = exe_actual + ".new"
        exe_viejo  = exe_actual + ".old"

        log.info(f"Iniciando descarga desde: {url_descarga}")
        _descargar_archivo(url_descarga, exe_nuevo, on_progreso)
        log.info("Descarga completada. Reemplazando ejecutable.")

        # Reemplazar: actual → .old, .new → actual
        if os.path.exists(exe_viejo):
            os.remove(exe_viejo)
        shutil.move(exe_actual, exe_viejo)
        shutil.move(exe_nuevo, exe_actual)

        log.info("Ejecutable reemplazado. Reiniciando aplicación.")
        subprocess.Popen([exe_actual])
        sys.exit(0)

    except Exception as e:
        log.error(f"Error durante la actualización: {e}", exc_info=True)
        # Intentar restaurar el original si quedó a medias
        if exe_viejo and os.path.exists(exe_viejo) and not os.path.exists(_ruta_exe_actual()):
            shutil.move(exe_viejo, _ruta_exe_actual())
            log.info("Ejecutable original restaurado tras el error.")
        if on_error:
            on_error(str(e))


def _descargar_archivo(url: str, destino: str, on_progreso):
    """Descarga `url` a `destino` reportando progreso si se provee callback."""
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "GeneradorDocumentos-Updater/1.0"},
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        total = int(resp.headers.get("Content-Length", 0))
        descargado = 0
        bloque = 8192

        with open(destino, "wb") as f:
            while True:
                chunk = resp.read(bloque)
                if not chunk:
                    break
                f.write(chunk)
                descargado += len(chunk)
                if on_progreso and total:
                    on_progreso(int(descargado / total * 100))

    if on_progreso:
        on_progreso(100)


def _ruta_exe_actual() -> str:
    """Retorna la ruta absoluta del ejecutable actual (funciona con PyInstaller y con .py)."""
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")