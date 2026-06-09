import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from docx import Document
import os
import re
import threading
from datetime import datetime
import updater
import tempfile
import shutil

# ── Tema ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

AZUL       = "#1E3A5F"
AZUL_MED   = "#2E5896"
AZUL_CLARO = "#EBF2FF"
ACENTO     = "#3B82F6"
VERDE      = "#16A34A"
ROJO       = "#DC2626"
GRIS_BG    = "#F5F7FA"
GRIS_BORDE = "#D1D9E6"
BLANCO     = "#FFFFFF"
TEXTO      = "#1A1A2E"
TEXTO_SUAVE= "#64748B"

PASOS = ["archivos", "trabajadores", "datos_faltantes", "generar"]

MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
         "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

def _fecha_texto(dt) -> str:
    return f"{dt.day} de {MESES[dt.month - 1]} del {dt.year}"

def _copiar_a_temp(ruta_original: str) -> str:
    """Copia el archivo a un temporal para poder leerlo aunque este abierto en Office."""
    ext = os.path.splitext(ruta_original)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.close()
    shutil.copy2(ruta_original, tmp.name)
    return tmp.name

class AppContratos(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Documentos")
        self.geometry("780x700")
        self.minsize(720, 640)
        self.resizable(True, True)
        self.configure(fg_color=GRIS_BG)

        self.excel_path  = tk.StringVar(value="")
        self.word_path   = tk.StringVar(value="")
        self.personas    = []          # lista de dicts con datos del excel
        self.columnas    = []          # nombres de columnas (primera fila)
        self.persona_sel = None        # dict de la persona seleccionada
        self.campos_extra = {}         # campos manuales (fecha, etc.)

        self._build_ui()
        # Verificar actualizaciones en segundo plano al iniciar
        self.after(2000, self._verificar_actualizacion)

    def _verificar_actualizacion(self):
        updater.verificar_actualizacion(
            on_update_available=self._on_actualizacion_disponible,
            on_error=None,  # Silencioso si no hay internet
        )

    def _on_actualizacion_disponible(self, version, url):
        """Llamado desde hilo secundario — usar after() para tocar la UI."""
        self.after(0, lambda: self._mostrar_dialogo_actualizacion(version, url))

    def _mostrar_dialogo_actualizacion(self, version, url):
        respuesta = messagebox.askyesno(
            "Actualización disponible",
            f"Hay una nueva versión disponible: v{version}\n\n"
            f"¿Deseas actualizar ahora?\n"
            f"(La aplicación se reiniciará automáticamente)",
        )
        if respuesta:
            self._iniciar_descarga(url)

    def _iniciar_descarga(self, url):
        # Mostrar ventana de progreso
        ventana = ctk.CTkToplevel(self)
        ventana.title("Actualizando…")
        ventana.geometry("360x140")
        ventana.resizable(False, False)
        ventana.grab_set()

        ctk.CTkLabel(
            ventana,
            text="Descargando actualización…",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(pady=(20, 10))

        barra = ctk.CTkProgressBar(ventana, width=300)
        barra.set(0)
        barra.pack(pady=4)

        lbl_pct = ctk.CTkLabel(ventana, text="0%", font=ctk.CTkFont(size=12))
        lbl_pct.pack()

        def on_progreso(pct):
            self.after(0, lambda p=pct: (barra.set(p / 100), lbl_pct.configure(text=f"{p}%")))

        updater.descargar_e_instalar(url, on_progreso=on_progreso)

    # ─────────────────────────────────────────────────────────────────────────
    #  CONSTRUCCIÓN DE LA INTERFAZ
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Encabezado
        header = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="Generador de Documentos",
            font=ctk.CTkFont(family="Georgia", size=22, weight="bold"),
            text_color=BLANCO,
        ).pack(side="left", padx=28, pady=20)

        version_local = updater.leer_version_local()
        ctk.CTkLabel(
            header,
            text=f"Contratos · Finiquitos · Cartas   |   v{version_local}",
            font=ctk.CTkFont(size=12),
            text_color="#A8C4E8",
        ).pack(side="right", padx=28, pady=20)

        # Barra de pasos
        self.frame_barra = ctk.CTkFrame(self, fg_color=AZUL_CLARO, corner_radius=0, height=44)
        self.frame_barra.pack(fill="x")
        self.frame_barra.pack_propagate(False)
        self._build_barra_pasos()

        # Contenedor del paso actual
        self.frame_contenido = ctk.CTkScrollableFrame(self, fg_color=GRIS_BG, corner_radius=0)
        self.frame_contenido.pack(fill="both", expand=True, padx=24, pady=20)

        # Navegacion
        nav = ctk.CTkFrame(self, fg_color=BLANCO, corner_radius=0, height=60)
        nav.pack(fill="x", side="bottom")
        nav.pack_propagate(False)

        self.btn_inicio = ctk.CTkButton(
            nav,
            text="Reiniciar",
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            hover_color=AZUL_CLARO,
            text_color=TEXTO_SUAVE,
            corner_radius=8,
            height=38,
            width=100,
            command=self._reiniciar,
        )
        self.btn_inicio.pack(side="left", padx=(6, 0), pady=10)

        self.btn_atras = ctk.CTkButton(
            nav,
            text="Atras",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=GRIS_BORDE,
            hover_color="#B0BCC8",
            text_color=TEXTO,
            corner_radius=8,
            height=38,
            width=120,
            command=self._paso_atras,
        )
        self.btn_atras.pack(side="left", padx=20, pady=10)

        self.btn_siguiente = ctk.CTkButton(
            nav,
            text="Siguiente",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=AZUL_MED,
            hover_color=AZUL,
            text_color=BLANCO,
            corner_radius=8,
            height=38,
            width=120,
            command=self._paso_siguiente,
        )
        self.btn_siguiente.pack(side="right", padx=20, pady=10)

        self.paso_actual = 0
        self._renderizar_paso()
        self.after(2000, self._verificar_actualizacion)

    def _build_barra_pasos(self):
        titulos = ["1. Archivos", "2. Trabajadores", "3. Datos faltantes", "4. Generar"]
        for w in self.frame_barra.winfo_children():
            w.destroy()
        for i, titulo in enumerate(titulos):
            activo = i == getattr(self, "paso_actual", 0)
            ctk.CTkLabel(
                self.frame_barra,
                text=titulo,
                font=ctk.CTkFont(size=12, weight="bold" if activo else "normal"),
                text_color=AZUL if activo else TEXTO_SUAVE,
                fg_color=BLANCO if activo else "transparent",
                corner_radius=6,
                padx=12,
                pady=4,
            ).pack(side="left", padx=6, pady=8)

    def _renderizar_paso(self):
        self._limpiar_variables()
        for w in self.frame_contenido.winfo_children():
            w.destroy()
        self._build_barra_pasos()

        pasos_fn = [
            self._paso_archivos,
            self._paso_trabajadores,
            self._paso_datos_faltantes,
            self._paso_generar,
        ]
        pasos_fn[self.paso_actual]()

        self.btn_atras.configure(state="normal" if self.paso_actual > 0 else "disabled")
        ultimo = self.paso_actual == len(PASOS) - 1
        self.btn_siguiente.configure(
            text="Generar" if ultimo else "Siguiente",
            fg_color=VERDE if ultimo else AZUL_MED,
            hover_color="#15803D" if ultimo else AZUL,
            command=self._generar if ultimo else self._paso_siguiente,
        )

    def _paso_siguiente(self):
        if self.paso_actual == 0:
            if not self.excel_path.get() or not self.word_path.get():
                messagebox.showwarning("Atencion", "Selecciona ambos archivos antes de continuar.")
                return
        if self.paso_actual == 1:
            if not getattr(self, "trabajadores_sel", []):
                messagebox.showwarning("Atencion", "Selecciona al menos un trabajador.")
                return

            campos_word = self._detectar_campos_word()
            campos_excel = {k.lower() for k in self.columnas}
            campos_en_ambos = [c for c in campos_word if c.lower() in campos_excel]

            avisos = []
            for persona in self.trabajadores_sel:
                nombre = str(list(persona.values())[0])
                faltantes = [
                    c for c in campos_en_ambos
                    if not persona.get(c) and not persona.get(c.lower())
                ]
                if faltantes:
                    avisos.append(f"{nombre}: falta {', '.join(faltantes)}")

            if avisos:
                msg = "Los siguientes trabajadores tienen datos vacios:\n\n" + "\n".join(avisos)
                msg += "\n\nPuedes continuar de todas formas o volver a revisar."
                continuar = messagebox.askyesno("Datos incompletos", msg)
                if not continuar:
                    return
        if self.paso_actual < len(PASOS) - 1:
            self.paso_actual += 1
            self._renderizar_paso()

    def _paso_atras(self):
        if self.paso_actual > 0:
            self.paso_actual -= 1
            self._renderizar_paso()
    
    def _reiniciar(self):
        self.excel_path.set("")
        self.word_path.set("")
        self.personas = []
        self.columnas = []
        self.trabajadores_sel = []
        self.campos_extra_vars = {}
        self.paso_actual = 0
        self._renderizar_paso()

    def _limpiar_variables(self):
        if hasattr(self, "busqueda_var") and hasattr(self, "_trace_id_busqueda"):
            try:
                self.busqueda_var.trace_remove("write", self._trace_id_busqueda)
            except Exception:
                pass
            del self.busqueda_var
            del self._trace_id_busqueda

        if hasattr(self, "col_vars"):
            del self.col_vars

        if hasattr(self, "check_vars_trabajadores"):
            del self.check_vars_trabajadores
    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 1 – ARCHIVOS
    # ─────────────────────────────────────────────────────────────────────────
    def _paso_archivos(self):
        parent = self.frame_contenido
        ctk.CTkLabel(
            parent,
            text="Selecciona los archivos",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 16))

        self._file_row(
            parent,
            "Archivo Excel con personas:",
            "Seleccionar Excel (.xlsx)",
            self.excel_path,
            self._cargar_excel,
            [("Excel", "*.xlsx *.xls")],
        )
        self._file_row(
            parent,
            "Plantilla Word (base del documento):",
            "Seleccionar Word (.docx)",
            self.word_path,
            self._validar_word,
            [("Word", "*.docx")],
        )

    def _file_row(self, parent, label_txt, btn_txt, var, cmd, filetypes):
        ctk.CTkLabel(
            parent,
            text=label_txt,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXTO,
            anchor="w",
        ).pack(fill="x", pady=(6, 4))

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 10))

        entry = ctk.CTkEntry(
            row,
            textvariable=var,
            state="readonly",
            placeholder_text="Ningún archivo seleccionado",
            font=ctk.CTkFont(size=12),
            fg_color=GRIS_BG,
            border_color=GRIS_BORDE,
            text_color=TEXTO_SUAVE,
            height=38,
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            row,
            text=btn_txt,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=AZUL_MED,
            hover_color=AZUL,
            text_color=BLANCO,
            corner_radius=8,
            height=38,
            command=lambda ft=filetypes, c=cmd, v=var: self._browse(ft, c, v),
        ).pack(side="right")

    def _browse(self, filetypes, callback, var):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)
            callback(path)

    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 2 – PERSONA
    # ─────────────────────────────────────────────────────────────────────────
    def _paso_trabajadores(self):
        parent = self.frame_contenido

        ctk.CTkLabel(
            parent,
            text="Selecciona los trabajadores",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            parent,
            text="Puedes seleccionar uno o varios.",
            font=ctk.CTkFont(size=11),
            text_color=TEXTO_SUAVE,
        ).pack(anchor="w", pady=(0, 10))

        # Selector de columnas visibles
        ctk.CTkLabel(
            parent,
            text="Columnas a mostrar en la lista:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 4))

        fila_cols = ctk.CTkFrame(parent, fg_color="transparent")
        fila_cols.pack(fill="x", pady=(0, 10))

        if not hasattr(self, "col_vars"):
            self.col_vars = {
                col: tk.BooleanVar(value=any(k in col.lower() for k in ["nombre", "rut"]))
                for col in self.columnas
            }

        for col in self.columnas:
            ctk.CTkCheckBox(
                fila_cols,
                text=col,
                variable=self.col_vars[col],
                font=ctk.CTkFont(size=11),
                command=self._refrescar_lista_trabajadores,
            ).pack(side="left", padx=6)

        # Busqueda
        self.busqueda_var = tk.StringVar()
        self._trace_id_busqueda = self.busqueda_var.trace_add("write", lambda *_: self._refrescar_lista_trabajadores())

        ctk.CTkEntry(
            parent,
            textvariable=self.busqueda_var,
            placeholder_text="Buscar...",
            font=ctk.CTkFont(size=13),
            height=40,
            border_color=ACENTO,
        ).pack(fill="x", pady=(0, 8))

        # Botones seleccionar/deseleccionar todos
        fila_sel = ctk.CTkFrame(parent, fg_color="transparent")
        fila_sel.pack(fill="x", pady=(0, 6))

        ctk.CTkButton(
            fila_sel,
            text="Seleccionar todos",
            font=ctk.CTkFont(size=11),
            fg_color=AZUL_CLARO,
            text_color=AZUL,
            hover_color=GRIS_BORDE,
            height=28,
            corner_radius=6,
            command=lambda: self._sel_todos(True),
        ).pack(side="left", padx=(0, 6))

        ctk.CTkButton(
            fila_sel,
            text="Deseleccionar todos",
            font=ctk.CTkFont(size=11),
            fg_color=AZUL_CLARO,
            text_color=AZUL,
            hover_color=GRIS_BORDE,
            height=28,
            corner_radius=6,
            command=lambda: self._sel_todos(False),
        ).pack(side="left")

        # Lista
        self.lista_frame = ctk.CTkScrollableFrame(
            parent, fg_color=GRIS_BG, corner_radius=8, height=240,
        )
        self.lista_frame.pack(fill="x")

        if not hasattr(self, "trabajadores_sel"):
            self.trabajadores_sel = []

        self._refrescar_lista_trabajadores()

    def _refrescar_lista_trabajadores(self):
        for w in self.lista_frame.winfo_children():
            w.destroy()

        cols_visibles = [c for c, v in self.col_vars.items() if v.get()]
        if not cols_visibles:
            cols_visibles = self.columnas[:1]

        q = getattr(self, "busqueda_var", tk.StringVar()).get().lower().strip()
        personas = [
            p for p in self.personas
            if not q or any(q in str(v).lower() for v in p.values())
        ]

        if not hasattr(self, "check_vars_trabajadores"):
            self.check_vars_trabajadores = {}

        for p in personas:
            key = str(list(p.values()))
            if key not in self.check_vars_trabajadores:
                self.check_vars_trabajadores[key] = tk.BooleanVar(value=p in self.trabajadores_sel)

            partes = [str(p.get(c, "")) for c in cols_visibles if p.get(c)]
            display = "   |   ".join(partes) if partes else "Sin datos"

            def on_toggle(persona=p, k=key):
                if self.check_vars_trabajadores[k].get():
                    if persona not in self.trabajadores_sel:
                        self.trabajadores_sel.append(persona)
                else:
                    if persona in self.trabajadores_sel:
                        self.trabajadores_sel.remove(persona)

            ctk.CTkCheckBox(
                self.lista_frame,
                text=display,
                variable=self.check_vars_trabajadores[key],
                font=ctk.CTkFont(size=12),
                text_color=TEXTO,
                command=on_toggle,
            ).pack(anchor="w", pady=3, padx=6)

    def _sel_todos(self, valor: bool):
        for key, var in self.check_vars_trabajadores.items():
            var.set(valor)
        if valor:
            self.trabajadores_sel = list(self.personas)
        else:
            self.trabajadores_sel = []

    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 3 – CAMPOS EXTRA
    # ─────────────────────────────────────────────────────────────────────────
    def _paso_datos_faltantes(self):
        parent = self.frame_contenido

        ctk.CTkLabel(
            parent,
            text="Datos faltantes",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 4))

        campos_word = self._detectar_campos_word()
        campos_excel = {k.lower() for k in self.columnas}
        campos_faltantes_word = [c for c in campos_word if c.lower() not in campos_excel]

        # Campos que estan en el excel pero vacios para algún trabajador
        campos_en_ambos = [c for c in campos_word if c.lower() in campos_excel]

        # self.campos_extra_vars ahora es dict de {nombre_trabajador: {campo: StringVar}}
        self.campos_extra_vars = {}

        hay_algo = False

        for persona in self.trabajadores_sel:
            nombre_persona = str(list(persona.values())[0])
            self.campos_extra_vars[nombre_persona] = {}

            # Campos vacios en el excel para esta persona
            faltantes_excel = [
                c for c in campos_en_ambos
                if not persona.get(c) and not persona.get(c.lower())
            ]

            campos_este_trabajador = campos_faltantes_word + faltantes_excel

            if not campos_este_trabajador:
                continue

            hay_algo = True

            # Separador por trabajador
            sep = ctk.CTkFrame(parent, fg_color=AZUL_CLARO, corner_radius=8)
            sep.pack(fill="x", pady=(10, 6))
            ctk.CTkLabel(
                sep,
                text=nombre_persona,
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=AZUL,
            ).pack(anchor="w", padx=12, pady=6)

            for campo in campos_este_trabajador:
                fila = ctk.CTkFrame(parent, fg_color="transparent")
                fila.pack(fill="x", pady=4, padx=8)

                ctk.CTkLabel(
                    fila,
                    text=f"{campo}:",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=TEXTO,
                    width=180,
                    anchor="w",
                ).pack(side="left")

                if "fecha" in campo.lower():
                    var = tk.StringVar(value=datetime.today().strftime("%d/%m/%Y"))
                    self.campos_extra_vars[nombre_persona][campo] = var

                    contenedor = ctk.CTkFrame(fila, fg_color="transparent")
                    contenedor.pack(side="left", fill="x", expand=True)

                    ctk.CTkEntry(
                        contenedor,
                        textvariable=var,
                        font=ctk.CTkFont(size=12),
                        height=34,
                        border_color=GRIS_BORDE,
                    ).pack(fill="x", pady=(0, 4))

                    fila_formatos = ctk.CTkFrame(contenedor, fg_color="transparent")
                    fila_formatos.pack(fill="x")

                    formatos = [
                        ("13/05/2025",          lambda v: datetime.today().strftime("%d/%m/%Y")),
                        ("13-05-2025",          lambda v: datetime.today().strftime("%d-%m-%Y")),
                        ("13 de Mayo del 2025", lambda v: _fecha_texto(datetime.today())),
                        ("2025-05-13",          lambda v: datetime.today().strftime("%Y-%m-%d")),
                    ]
                    for etiqueta, fn in formatos:
                        ctk.CTkButton(
                            fila_formatos,
                            text=etiqueta,
                            font=ctk.CTkFont(size=10),
                            fg_color=AZUL_CLARO,
                            text_color=AZUL,
                            hover_color=GRIS_BORDE,
                            height=24,
                            corner_radius=4,
                            command=lambda f=fn, variable=var: variable.set(f(variable.get())),
                        ).pack(side="left", padx=2)
                else:
                    var = tk.StringVar(value="")
                    self.campos_extra_vars[nombre_persona][campo] = var
                    ctk.CTkEntry(
                        fila,
                        textvariable=var,
                        font=ctk.CTkFont(size=12),
                        height=34,
                        border_color=GRIS_BORDE,
                    ).pack(side="left", fill="x", expand=True)

        if not hay_algo:
            ctk.CTkLabel(
                parent,
                text="No hay datos faltantes. Todos los trabajadores tienen sus datos completos.",
                font=ctk.CTkFont(size=13),
                text_color=VERDE,
            ).pack(anchor="w", pady=10)

    def _detectar_campos_word(self):
        if not self.word_path.get():
            return []
        try:
            doc = Document(self.word_path.get())
            texto = " ".join(p.text for p in doc.paragraphs)
            # También revisar tablas
            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        texto += " " + celda.text
            patron = re.compile(r"\{\{([^}]+)\}\}")
            campos = list(dict.fromkeys(m.group(1).strip() for m in patron.finditer(texto)))
            return campos
        except Exception:
            return []

    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 4 – GENERAR
    # ─────────────────────────────────────────────────────────────────────────
    def _paso_generar(self):
        parent = self.frame_contenido

        ctk.CTkLabel(
            parent,
            text="Configurar y generar",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 16))

        # Formato del nombre de archivo
        ctk.CTkLabel(
            parent,
            text="Formato del nombre de archivo:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            parent,
            text="Usa {columna} para insertar datos del trabajador. Ejemplo: {nombre}_{cargo}_{fecha}",
            font=ctk.CTkFont(size=11),
            text_color=TEXTO_SUAVE,
        ).pack(anchor="w", pady=(0, 6))

        # Botones rapidos de columnas disponibles
        fila_cols = ctk.CTkFrame(parent, fg_color="transparent")
        fila_cols.pack(fill="x", pady=(0, 6))

        self.formato_nombre_var = tk.StringVar(
            value="{nombre}_{cargo}" if "cargo" in [c.lower() for c in self.columnas] else "{" + self.columnas[0] + "}"
        )

        for col in self.columnas[:6]:
            ctk.CTkButton(
                fila_cols,
                text=f"+{col}",
                font=ctk.CTkFont(size=10),
                fg_color=AZUL_CLARO,
                text_color=AZUL,
                hover_color=GRIS_BORDE,
                height=24,
                corner_radius=4,
                command=lambda c=col: self.formato_nombre_var.set(
                    self.formato_nombre_var.get() + f"_{{{c}}}"
                ),
            ).pack(side="left", padx=2)

        ctk.CTkEntry(
            parent,
            textvariable=self.formato_nombre_var,
            font=ctk.CTkFont(size=12),
            height=38,
            border_color=GRIS_BORDE,
        ).pack(fill="x", pady=(0, 16))

        # Resumen
        n = len(getattr(self, "trabajadores_sel", []))
        ctk.CTkLabel(
            parent,
            text=f"Se generaran {n} documento(s).",
            font=ctk.CTkFont(size=13),
            text_color=TEXTO_SUAVE,
        ).pack(anchor="w", pady=(0, 16))

        self.lbl_estado = ctk.CTkLabel(parent, text="", font=ctk.CTkFont(size=13), text_color=VERDE)
        self.lbl_estado.pack(anchor="w")
    # ─────────────────────────────────────────────────────────────────────────
    #  LÓGICA
    # ─────────────────────────────────────────────────────────────────────────
    def _cargar_excel(self, path):
        try:
            ruta_lectura = _copiar_a_temp(path)
            wb = openpyxl.load_workbook(ruta_lectura, read_only=True, data_only=True)
            ws = wb.active

            filas = list(ws.iter_rows(values_only=True))
            if not filas:
                messagebox.showerror("Error", "El Excel está vacío.")
                return

            self.columnas = [str(c).strip() if c is not None else None
                            for i, c in enumerate(filas[0])]

            self.columnas = [c for c in self.columnas if c is not None]
            self.personas = []
            for fila in filas[1:]:
                if any(c is not None for c in fila):
                    persona = {self.columnas[i]: (fila[i] if i < len(fila) else "")
                            for i in range(len(self.columnas))
                            if self.columnas[i] is not None}
                    self.personas.append(persona)

            wb.close()
            self._notificar(f"Excel cargado: {len(self.personas)} personas encontradas")

        except Exception as e:
            messagebox.showerror("Error al leer Excel", str(e))

    def _validar_word(self, path):
        try:
            Document(path)
            campos = self._detectar_campos_word()
            if campos:
                self._notificar(f"Word cargado. Campos detectados: {', '.join(campos)}")
            else:
                self._notificar("Word cargado. No se detectaron campos {{...}}. Asegúrate de usar {{nombre}}, {{rut}}, etc.")
        except Exception as e:
            messagebox.showerror("Error al leer Word", str(e))

    def _generar(self):
        if not getattr(self, "trabajadores_sel", []):
            messagebox.showwarning("Atencion", "No hay trabajadores seleccionados.")
            return

        carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los documentos")
        if not carpeta:
            return

        self.btn_siguiente.configure(state="disabled", text="Generando...")
        threading.Thread(target=self._generar_thread, args=(carpeta,), daemon=True).start()

    def _generar_thread(self, carpeta):
        errores = []
        generados = 0

        for persona in self.trabajadores_sel:
            try:
                ruta_lectura = _copiar_a_temp(self.word_path.get())
                doc = Document(ruta_lectura)

                reemplazos = {k.strip(): str(v) if v is not None else "" for k, v in persona.items()}
                nombre_persona = str(list(persona.values())[0])
                extras = self.campos_extra_vars.get(nombre_persona, {})
                for k, var in extras.items():
                    reemplazos[k.strip()] = var.get()

                def reemplazar_texto(texto):
                    def sub(m):
                        campo = m.group(1).strip()
                        for k, v in reemplazos.items():
                            if k.lower() == campo.lower():
                                return v
                        return m.group(0)
                    return re.sub(r"\{\{([^}]+)\}\}", sub, texto)

                def procesar_parrafo(parrafo):
                    texto_completo = parrafo.text
                    nuevo_texto = reemplazar_texto(texto_completo)
                    if texto_completo == nuevo_texto:
                        return
                    # Copiar formato del primer run que tenga contenido
                    run_base = next((r for r in parrafo.runs if r.text.strip()), None)
                    for i, run in enumerate(parrafo.runs):
                        run.text = nuevo_texto if i == 0 else ""
                    # Restaurar formato del run base en el primero
                    if run_base and parrafo.runs:
                        r = parrafo.runs[0]
                        r.bold      = run_base.bold
                        r.italic    = run_base.italic
                        r.underline = run_base.underline
                        r.font.size = run_base.font.size
                        r.font.name = run_base.font.name
                        r.font.color.rgb = run_base.font.color.rgb if run_base.font.color and run_base.font.color.type else r.font.color.rgb

                for p in doc.paragraphs:
                    procesar_parrafo(p)
                for tabla in doc.tables:
                    for fila in tabla.rows:
                        for celda in fila.cells:
                            for p in celda.paragraphs:
                                procesar_parrafo(p)

                # Generar nombre del archivo
                formato = self.formato_nombre_var.get()
                nombre_archivo = formato
                for k, v in reemplazos.items():
                    nombre_archivo = nombre_archivo.replace(f"{{{k}}}", str(v))
                hoy = datetime.today().strftime("%Y%m%d")
                nombre_archivo = nombre_archivo.replace("{fecha}", hoy)
                nombre_archivo = re.sub(r'[\\/*?:"<>|]', "_", nombre_archivo)
                nombre_archivo = nombre_archivo.strip("_ ") + ".docx"

                ruta_final = os.path.join(carpeta, nombre_archivo)
                doc.save(ruta_final)
                generados += 1

            except Exception as e:
                nombre = str(list(persona.values())[0])
                errores.append(f"{nombre}: {e}")

        self.after(0, lambda: self._generar_resultado(carpeta, generados, errores))

    def _generar_resultado(self, carpeta, generados, errores):
        self.btn_siguiente.configure(state="normal", text="Generar", fg_color=VERDE)
        self.lbl_estado.configure(text=f"{generados} documento(s) generado(s) correctamente.")

        if errores:
            messagebox.showwarning("Algunos errores", "\n".join(errores))

        if generados > 0:
            respuesta = messagebox.askyesno("Listo", f"Se generaron {generados} documento(s).\n¿Abrir la carpeta?")
            if respuesta:
                os.startfile(carpeta) if os.name == "nt" else os.system(f'open "{carpeta}"')

    def _notificar(self, msg):
        """Muestra un toast temporal en el título."""
        original = self.title()
        self.title(msg)
        self.after(3000, lambda: self.title(original))

# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AppContratos()
    app.mainloop()