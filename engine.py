import os
import re
from typing import Any, Callable
from datetime import datetime
from docx import Document
from utils import copiar_a_temp, fecha_texto, limpiar_nombre

PDF_DISPONIBLE: bool = False
try:
    from docx2pdf import convert as _pdf_conv
    PDF_DISPONIBLE = True
except ImportError:
    try:
        import subprocess as _sp
        _sp.run(["libreoffice", "--version"], capture_output=True, check=True)
        PDF_DISPONIBLE = True
    except (FileNotFoundError, _sp.CalledProcessError):
        PDF_DISPONIBLE = False

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _convertir_a_pdf(ruta_docx: str) -> None:
    ruta_pdf = os.path.splitext(ruta_docx)[0] + ".pdf"
    try:
        from docx2pdf import convert as _conv
        _conv(ruta_docx, ruta_pdf)
    except ImportError:
        import subprocess as _sp
        _sp.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir",
             os.path.dirname(ruta_pdf), ruta_docx],
            capture_output=True, check=True, timeout=60,
        )


def reemplazar_runs_xml(runs: list, texto_completo: str) -> None:
    nuevo_texto = reemplazar_texto(texto_completo)
    if nuevo_texto == texto_completo:
        return
    t_elems = []
    for r in runs:
        t_elems.extend(r.iter(W_NS + "t"))
    for t in t_elems:
        t.text = ""
    if t_elems:
        t_elems[0].text = nuevo_texto


def generar_documentos(
    word_path: str,
    trabajadores: list[dict[str, Any]],
    columnas: list[str],
    formato_fecha: str,
    col_identificador: str,
    campos_extra_vars: dict[str, dict[str, Any]],
    formato_nombre: str,
    usar_guion_bajo: bool,
    carpeta_destino: str,
    generar_pdf: bool = False,
    on_progreso: Callable[[int, int], None] | None = None,
) -> tuple[int, list[str]]:
    errores: list[str] = []
    generados = 0
    total = len(trabajadores)

    def reemplazar_texto(texto: str) -> str:
        def sub(m: re.Match) -> str:
            campo = m.group(1).strip()
            for k, v in reemplazos.items():
                if k.lower() == campo.lower():
                    return v
            return m.group(0)
        return re.sub(r"\{\{([^}]+)\}\}", sub, texto)

    def procesar_parrafo(parrafo: Any) -> None:
        texto_completo = "".join(r.text for r in parrafo.runs)
        nuevo_texto = reemplazar_texto(texto_completo)
        if texto_completo == nuevo_texto:
            return

        mapa_formato: list[dict[str, Any]] = []
        for r in parrafo.runs:
            for _ in r.text:
                mapa_formato.append({
                    "bold":      r.bold,
                    "italic":    r.italic,
                    "underline": r.underline,
                    "size":      r.font.size,
                    "name":      r.font.name,
                })

        reemplazos_formato: dict[str, Any] = {}
        for m in re.finditer(r"\{\{([^}]+)\}\}", texto_completo):
            campo = m.group(1).strip()
            pos = m.start()
            fmt = mapa_formato[pos] if pos < len(mapa_formato) else None
            reemplazos_formato[campo.lower()] = fmt

        segmentos: list[tuple[str, Any]] = []
        cursor = 0
        for m in re.finditer(r"\{\{([^}]+)\}\}", texto_completo):
            campo = m.group(1).strip()
            if m.start() > cursor:
                segmentos.append((texto_completo[cursor:m.start()], mapa_formato[cursor] if cursor < len(mapa_formato) else None))
            valor = reemplazar_texto(m.group(0))
            fmt = reemplazos_formato.get(campo.lower())
            segmentos.append((valor, fmt))
            cursor = m.end()
        if cursor < len(texto_completo):
            segmentos.append((texto_completo[cursor:], mapa_formato[cursor] if cursor < len(mapa_formato) else None))

        run_base = parrafo.runs[0] if parrafo.runs else None
        for r in parrafo.runs:
            r.text = ""

        for texto_seg, fmt in segmentos:
            if not texto_seg:
                continue
            run = parrafo.add_run(texto_seg)
            if fmt:
                run.bold      = fmt["bold"]
                run.italic    = fmt["italic"]
                run.underline = fmt["underline"]
                run.font.size = fmt["size"]
                if fmt["name"]:
                    run.font.name = fmt["name"]
            elif run_base:
                run.bold      = run_base.bold
                run.italic    = run_base.italic
                run.underline = run_base.underline
                run.font.size = run_base.font.size

    for idx, persona in enumerate(trabajadores):
        ruta_lectura: str | None = None
        try:
            ruta_lectura = copiar_a_temp(word_path)
            doc = Document(ruta_lectura)

            reemplazos: dict[str, str] = {k.strip(): str(v) if v is not None else "" for k, v in persona.items()}
            for k, v in persona.items():
                if isinstance(v, datetime):
                    if formato_fecha == "texto":
                        reemplazos[k.strip()] = fecha_texto(v)
                    else:
                        reemplazos[k.strip()] = v.strftime(formato_fecha)

            nombre_persona = str(persona.get(col_identificador, str(list(persona.values())[0])))
            extras = campos_extra_vars.get(nombre_persona, {})
            for k, var in extras.items():
                reemplazos[k.strip()] = var.get()

            for p in doc.paragraphs:
                procesar_parrafo(p)
            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        for p in celda.paragraphs:
                            procesar_parrafo(p)
            for seccion in doc.sections:
                for p in seccion.header.paragraphs:
                    procesar_parrafo(p)
                for p in seccion.footer.paragraphs:
                    procesar_parrafo(p)
            for txbx in doc.element.body.iter(W_NS + "txbxContent"):
                for txbx_p in txbx.iter(W_NS + "p"):
                    runs = list(txbx_p.iter(W_NS + "r"))
                    texto_completo = "".join(
                        "".join(t.text or "" for t in r.iter(W_NS + "t"))
                        for r in runs
                    )
                    if not re.search(r"\{\{[^}]+\}\}", texto_completo):
                        continue
                    reemplazar_runs_xml(runs, texto_completo)

            nombre_archivo = formato_nombre
            for k, v in reemplazos.items():
                nombre_archivo = nombre_archivo.replace(f"{{{k}}}", str(v))
            nombre_archivo = nombre_archivo.replace("{fecha}", datetime.today().strftime("%Y%m%d"))
            nombre_archivo = limpiar_nombre(nombre_archivo, usar_guion_bajo)
            nombre_archivo = nombre_archivo + ".docx"

            ruta_final = os.path.join(carpeta_destino, nombre_archivo)
            doc.save(ruta_final)
            if generar_pdf:
                try:
                    _convertir_a_pdf(ruta_final)
                except Exception as e:
                    errores.append(f"{nombre_persona}: PDF - {e}")
            generados += 1

        except Exception as e:
            nombre = str(persona.get(col_identificador, str(list(persona.values())[0])))
            errores.append(f"{nombre}: {e}")
        finally:
            if ruta_lectura and os.path.exists(ruta_lectura):
                os.unlink(ruta_lectura)

        if on_progreso:
            on_progreso(idx + 1, total)

    return generados, errores
