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

import hashlib
import json
import logging
import os
import re
import sys
import threading
import subprocess
import urllib.request
import urllib.error

# ── Configuración ─────────────────────────────────────────────────────────────
GITHUB_USER  = os.environ.get("UPDATER_GITHUB_USER", "mvillarios")
GITHUB_REPO  = os.environ.get("UPDATER_GITHUB_REPO", "docsGen")
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
        on_update_available(version, url_descarga, sha256): llamado si hay versión nueva.
        on_error(mensaje):                                  llamado si hay un error (opcional).
    """
    hilo = threading.Thread(
        target=_verificar_hilo,
        args=(on_update_available, on_error),
        daemon=True,
    )
    hilo.start()


def descargar_e_instalar(url_descarga: str, sha256_esperado: str = "",
                         on_progreso: callable = None, on_error: callable = None):
    """
    Descarga el nuevo .exe y lo reemplaza en un hilo secundario.

    Parámetros:
        url_descarga:         URL directa al .exe del release de GitHub.
        sha256_esperado:      Hash SHA256 esperado para verificar la descarga.
        on_progreso(pct:int): llamado con porcentaje 0-100 durante la descarga.
        on_error(mensaje):    llamado si la descarga falla.
    """
    hilo = threading.Thread(
        target=_descargar_hilo,
        args=(url_descarga, sha256_esperado, on_progreso, on_error),
        daemon=True,
    )
    hilo.start()

def limpiar_archivos_viejos():
    """Elimina residuos de actualizaciones previas (.old, .new, .bat)."""
    exe_base = _ruta_exe_actual() if getattr(sys, "frozen", False) else None
    if exe_base:
        for ext in (".old", ".new", ".bat"):
            ruta = exe_base + ext
            if os.path.exists(ruta):
                try:
                    os.remove(ruta)
                    log.info(f"Residuo eliminado: {ruta}")
                except Exception as e:
                    log.warning(f"No se pudo eliminar {ruta}: {e}")

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
                sha256 = _extraer_sha256(release)
                on_update_available(version_remota, url, sha256)
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


def _extraer_sha256(release: dict) -> str:
    body = release.get("body", "")
    for linea in body.splitlines():
        m = re.search(r"SHA256[:\s]+([a-fA-F0-9]{64})", linea)
        if m:
            return m.group(1)
    return ""


def _obtener_url_exe(release: dict) -> str | None:
    """Extrae la URL de descarga del primer .exe en los assets del release."""
    for asset in release.get("assets", []):
        if asset["name"].endswith(".exe"):
            return asset["browser_download_url"]
    return None


def _hay_version_nueva(local: str, remota: str) -> bool:
    def _partes(v: str):
        partes = []
        for x in v.split("."):
            digitos = ""
            for c in x:
                if c.isdigit():
                    digitos += c
                else:
                    break
            partes.append(int(digitos) if digitos else 0)
        while len(partes) < 3:
            partes.append(0)
        return tuple(partes[:3])
    return _partes(remota) > _partes(local)


def _escribir_script_reemplazo(exe_actual: str, exe_nuevo: str) -> str:
    """Escribe un .bat que espera a que el proceso termine y reemplaza el .exe."""
    nombre_exe = os.path.basename(exe_actual)
    bat_path = exe_actual + ".bat"
    bat_content = (
        '@echo off\r\n'
        ':: Generado por docsGen updater\r\n'
        ':wait\r\n'
        f'timeout /t 1 /nobreak >nul\r\n'
        f'tasklist /FI "IMAGENAME eq {nombre_exe}" 2>nul | find /I "{nombre_exe}" >nul\r\n'
        f'if not errorlevel 1 goto wait\r\n'
        f'del "{exe_actual}"\r\n'
        f'ren "{exe_nuevo}" "{nombre_exe}"\r\n'
        f'start "" "{exe_actual}"\r\n'
        f'del "%~f0"\r\n'
    )
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)
    return bat_path


def _descargar_hilo(url_descarga: str, sha256_esperado: str, on_progreso, on_error):
    exe_nuevo = None
    try:
        exe_actual = _ruta_exe_actual()
        exe_nuevo  = exe_actual + ".new"

        log.info(f"Iniciando descarga desde: {url_descarga}")
        _descargar_archivo(url_descarga, exe_nuevo, on_progreso)
        log.info("Descarga completada. Verificando integridad.")

        if sha256_esperado:
            sha256_real = hashlib.sha256()
            with open(exe_nuevo, "rb") as f:
                while True:
                    chunk = f.read(65536)
                    if not chunk:
                        break
                    sha256_real.update(chunk)
            hash_real = sha256_real.hexdigest()
            if hash_real.lower() != sha256_esperado.lower():
                raise IOError(
                    f"El hash SHA256 no coincide.\n"
                    f"Esperado: {sha256_esperado}\n"
                    f"Obtenido: {hash_real}"
                )
            log.info("Integridad verificada correctamente.")

        log.info("Creando script de reemplazo diferido.")
        bat_path = _escribir_script_reemplazo(exe_actual, exe_nuevo)

        log.info("Lanzando script y cerrando aplicación.")
        subprocess.Popen(
            [bat_path],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            close_fds=True,
        )
        logging.shutdown()
        os._exit(0)

    except Exception as e:
        log.error(f"Error durante la actualización: {e}", exc_info=True)
        if exe_nuevo and os.path.exists(exe_nuevo):
            try:
                os.remove(exe_nuevo)
                log.info("Archivo .new temporal eliminado tras el error.")
            except Exception:
                pass
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
    if getattr(sys, "frozen", False):
        return sys.executable
    raise RuntimeError("La actualización solo funciona en el ejecutable compilado (no en modo desarrollo).")