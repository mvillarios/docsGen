import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from docx import Document
import os
import re
import threading
from datetime import datetime

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

    # ─────────────────────────────────────────────────────────────────────────
    #  CONSTRUCCIÓN DE LA INTERFAZ
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Encabezado ────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=80)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="Generador de Documentos",
            font=ctk.CTkFont(family="Georgia", size=22, weight="bold"),
            text_color=BLANCO,
        ).pack(side="left", padx=28, pady=20)

        ctk.CTkLabel(
            header,
            text="Contratos · Finiquitos · Cartas",
            font=ctk.CTkFont(size=12),
            text_color="#A8C4E8",
        ).pack(side="right", padx=28, pady=20)

        # ── Contenido principal ───────────────────────────────────────────────
        main = ctk.CTkScrollableFrame(self, fg_color=GRIS_BG, corner_radius=0)
        main.pack(fill="both", expand=True, padx=24, pady=20)

        # Paso 1: Archivos
        self._seccion(main, "1", "Cargar archivos", self._build_paso1)
        # Paso 2: Persona
        self.frame_paso2 = self._seccion(main, "2", "Seleccionar persona", self._build_paso2)
        # Paso 3: Campos extra
        self.frame_paso3 = self._seccion(main, "3", "Revisar y completar datos", self._build_paso3)
        # Paso 4: Generar
        self.frame_paso4 = self._seccion(main, "4", "Generar documento", self._build_paso4)

        self._set_pasos_bloqueados()

    def _seccion(self, parent, numero, titulo, builder_fn):
        """Crea una tarjeta de sección y llama a builder_fn para rellenarla."""
        card = ctk.CTkFrame(parent, fg_color=BLANCO, corner_radius=12,
                             border_width=1, border_color=GRIS_BORDE)
        card.pack(fill="x", pady=(0, 14))

        # Encabezado de sección
        top = ctk.CTkFrame(card, fg_color=AZUL_CLARO, corner_radius=0, height=44)
        top.pack(fill="x")
        top.pack_propagate(False)

        badge = ctk.CTkLabel(
            top,
            text=numero,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=BLANCO,
            fg_color=AZUL_MED,
            corner_radius=14,
            width=28, height=28,
        )
        badge.pack(side="left", padx=14, pady=8)

        ctk.CTkLabel(
            top,
            text=titulo,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=AZUL,
        ).pack(side="left", pady=8)

        # Cuerpo
        body = ctk.CTkFrame(card, fg_color=BLANCO, corner_radius=0)
        body.pack(fill="x", padx=20, pady=16)

        builder_fn(body)
        return body

    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 1 – ARCHIVOS
    # ─────────────────────────────────────────────────────────────────────────
    def _build_paso1(self, parent):
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
    def _build_paso2(self, parent):
        self.frame_paso2_inner = parent

        self.lbl_buscar = ctk.CTkLabel(
            parent,
            text="Primero carga el Excel (paso 1)",
            font=ctk.CTkFont(size=13),
            text_color=TEXTO_SUAVE,
        )
        self.lbl_buscar.pack(anchor="w")

    def _refrescar_paso2(self):
        for w in self.frame_paso2_inner.winfo_children():
            w.destroy()

        ctk.CTkLabel(
            self.frame_paso2_inner,
            text="Buscar por nombre o RUT:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXTO,
            anchor="w",
        ).pack(fill="x", pady=(0, 6))

        self.busqueda_var = tk.StringVar()
        self.busqueda_var.trace_add("write", self._filtrar_personas)

        ctk.CTkEntry(
            self.frame_paso2_inner,
            textvariable=self.busqueda_var,
            placeholder_text="Escriba para filtrar…",
            font=ctk.CTkFont(size=13),
            height=40,
            border_color=ACENTO,
        ).pack(fill="x", pady=(0, 10))

        # Lista con scrollbar
        lista_frame = ctk.CTkScrollableFrame(
            self.frame_paso2_inner,
            fg_color=GRIS_BG,
            corner_radius=8,
            height=180,
        )
        lista_frame.pack(fill="x")
        self.lista_frame = lista_frame
        self._mostrar_lista(self.personas)

    def _mostrar_lista(self, personas):
        for w in self.lista_frame.winfo_children():
            w.destroy()

        if not personas:
            ctk.CTkLabel(
                self.lista_frame,
                text="No se encontraron resultados",
                text_color=TEXTO_SUAVE,
                font=ctk.CTkFont(size=12),
            ).pack(pady=10)
            return

        primera_col = self.columnas[0] if self.columnas else "nombre"

        for p in personas:
            nombre = str(p.get(primera_col, str(p)))
            # Buscar columna de RUT
            rut_val = ""
            for k in p:
                if "rut" in k.lower():
                    rut_val = str(p[k])
                    break

            display = nombre + (f"  —  RUT {rut_val}" if rut_val else "")

            btn = ctk.CTkButton(
                self.lista_frame,
                text=display,
                font=ctk.CTkFont(size=12),
                fg_color="transparent",
                text_color=TEXTO,
                hover_color=AZUL_CLARO,
                anchor="w",
                height=36,
                corner_radius=6,
                border_width=0,
                command=lambda persona=p: self._seleccionar_persona(persona),
            )
            btn.pack(fill="x", pady=1)

    def _filtrar_personas(self, *_):
        q = self.busqueda_var.get().lower().strip()
        if not q:
            self._mostrar_lista(self.personas)
            return
        filtradas = [
            p for p in self.personas
            if any(q in str(v).lower() for v in p.values())
        ]
        self._mostrar_lista(filtradas)

    # ─────────────────────────────────────────────────────────────────────────
    #  PASO 3 – CAMPOS EXTRA
    # ─────────────────────────────────────────────────────────────────────────
    def _build_paso3(self, parent):
        self.frame_paso3_inner = parent
        ctk.CTkLabel(
            parent,
            text="Selecciona una persona primero",
            font=ctk.CTkFont(size=13),
            text_color=TEXTO_SUAVE,
        ).pack(anchor="w")

    def _refrescar_paso3(self):
        for w in self.frame_paso3_inner.winfo_children():
            w.destroy()

        if not self.persona_sel:
            return

        # ── Datos de la persona (solo lectura) ────────────────────────────
        ctk.CTkLabel(
            self.frame_paso3_inner,
            text="Datos de la persona seleccionada:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 8))

        grid = ctk.CTkFrame(self.frame_paso3_inner, fg_color=AZUL_CLARO, corner_radius=8)
        grid.pack(fill="x", pady=(0, 14))

        cols = 2
        items = list(self.persona_sel.items())
        for i, (k, v) in enumerate(items):
            row_f = ctk.CTkFrame(grid, fg_color="transparent")
            row_f.grid(row=i // cols, column=i % cols, sticky="ew", padx=12, pady=4)
            grid.grid_columnconfigure(i % cols, weight=1)

            ctk.CTkLabel(row_f, text=f"{k}:", font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=AZUL_MED, anchor="w").pack(anchor="w")
            ctk.CTkLabel(row_f, text=str(v) if v else "—",
                         font=ctk.CTkFont(size=12), text_color=TEXTO, anchor="w").pack(anchor="w")

        # ── Detectar campos del Word no cubiertos por Excel ───────────────
        campos_word = self._detectar_campos_word()
        campos_excel = {k.lower() for k in self.persona_sel.keys()}
        campos_faltantes = [c for c in campos_word if c.lower() not in campos_excel]

        self.campos_extra_vars = {}

        if campos_faltantes:
            ctk.CTkLabel(
                self.frame_paso3_inner,
                text="Campos adicionales del documento:",
                font=ctk.CTkFont(size=13, weight="bold"),
                text_color=TEXTO,
            ).pack(anchor="w", pady=(4, 8))

            ctk.CTkLabel(
                self.frame_paso3_inner,
                text="Estos campos aparecen en el Word y no están en el Excel.",
                font=ctk.CTkFont(size=11),
                text_color=TEXTO_SUAVE,
            ).pack(anchor="w", pady=(0, 8))

            for campo in campos_faltantes:
                row = ctk.CTkFrame(self.frame_paso3_inner, fg_color="transparent")
                row.pack(fill="x", pady=3)

                # Sugerir fecha de hoy para campos de fecha
                default = ""
                if "fecha" in campo.lower():
                    default = datetime.today().strftime("%d/%m/%Y")

                ctk.CTkLabel(
                    row,
                    text=f"{campo}:",
                    font=ctk.CTkFont(size=12, weight="bold"),
                    text_color=TEXTO,
                    width=180,
                    anchor="w",
                ).pack(side="left")

                var = tk.StringVar(value=default)
                self.campos_extra_vars[campo] = var
                ctk.CTkEntry(
                    row,
                    textvariable=var,
                    font=ctk.CTkFont(size=12),
                    height=34,
                    border_color=GRIS_BORDE,
                ).pack(side="left", fill="x", expand=True)

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
    def _build_paso4(self, parent):
        self.frame_paso4_inner = parent

        ctk.CTkLabel(
            parent,
            text="Completa los pasos anteriores para habilitar la generación.",
            font=ctk.CTkFont(size=13),
            text_color=TEXTO_SUAVE,
        ).pack(anchor="w")

    def _refrescar_paso4(self):
        for w in self.frame_paso4_inner.winfo_children():
            w.destroy()

        # Nombre sugerido
        ctk.CTkLabel(
            self.frame_paso4_inner,
            text="Nombre del archivo a guardar:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXTO,
        ).pack(anchor="w", pady=(0, 6))

        primera_col = self.columnas[0] if self.columnas else "nombre"
        nombre_base = str(self.persona_sel.get(primera_col, "documento"))
        template_base = os.path.splitext(os.path.basename(self.word_path.get()))[0]
        hoy = datetime.today().strftime("%Y%m%d")
        sugerido = f"{template_base}_{nombre_base}_{hoy}.docx"
        # Limpiar caracteres no válidos
        sugerido = re.sub(r'[\\/*?:"<>|]', "_", sugerido)

        self.nombre_salida_var = tk.StringVar(value=sugerido)
        ctk.CTkEntry(
            self.frame_paso4_inner,
            textvariable=self.nombre_salida_var,
            font=ctk.CTkFont(size=12),
            height=38,
            border_color=GRIS_BORDE,
        ).pack(fill="x", pady=(0, 16))

        # Botón grande
        self.btn_generar = ctk.CTkButton(
            self.frame_paso4_inner,
            text="Generar Documento",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=VERDE,
            hover_color="#15803D",
            text_color=BLANCO,
            corner_radius=10,
            height=52,
            command=self._generar,
        )
        self.btn_generar.pack(fill="x")

        self.lbl_estado = ctk.CTkLabel(
            self.frame_paso4_inner,
            text="",
            font=ctk.CTkFont(size=13),
            text_color=VERDE,
        )
        self.lbl_estado.pack(pady=(10, 0))

    # ─────────────────────────────────────────────────────────────────────────
    #  LÓGICA
    # ─────────────────────────────────────────────────────────────────────────
    def _cargar_excel(self, path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active

            filas = list(ws.iter_rows(values_only=True))
            if not filas:
                messagebox.showerror("Error", "El Excel está vacío.")
                return

            self.columnas = [str(c).strip() if c is not None else f"columna_{i}"
                             for i, c in enumerate(filas[0])]

            self.personas = []
            for fila in filas[1:]:
                if any(c is not None for c in fila):
                    persona = {self.columnas[i]: (fila[i] if i < len(fila) else "")
                               for i in range(len(self.columnas))}
                    self.personas.append(persona)

            wb.close()
            self._refrescar_paso2()
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
                self._notificar("⚠️  Word cargado. No se detectaron campos {{...}}. Asegúrate de usar {{nombre}}, {{rut}}, etc.")
        except Exception as e:
            messagebox.showerror("Error al leer Word", str(e))

    def _seleccionar_persona(self, persona):
        self.persona_sel = persona
        self._refrescar_paso3()
        if self.word_path.get():
            self._refrescar_paso4()

    def _generar(self):
        if not self.persona_sel:
            messagebox.showwarning("Atención", "Selecciona una persona primero.")
            return
        if not self.word_path.get():
            messagebox.showwarning("Atención", "Selecciona el Word base primero.")
            return

        # Carpeta de destino
        destino = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word", "*.docx")],
            initialfile=self.nombre_salida_var.get(),
            title="Guardar documento como…",
        )
        if not destino:
            return

        self.btn_generar.configure(state="disabled", text="Generando…")
        threading.Thread(target=self._generar_thread, args=(destino,), daemon=True).start()

    def _generar_thread(self, destino):
        try:
            doc = Document(self.word_path.get())

            # Construir mapa de reemplazos (case-insensitive)
            reemplazos = {}
            for k, v in self.persona_sel.items():
                reemplazos[k.strip()] = str(v) if v is not None else ""
            for k, var in getattr(self, "campos_extra_vars", {}).items():
                reemplazos[k.strip()] = var.get()

            def reemplazar_texto(texto):
                def sub(m):
                    campo = m.group(1).strip()
                    # Búsqueda case-insensitive
                    for k, v in reemplazos.items():
                        if k.lower() == campo.lower():
                            return v
                    return m.group(0)  # Dejar sin cambio si no se encontró
                return re.sub(r"\{\{([^}]+)\}\}", sub, texto)

            def procesar_parrafo(parrafo):
                # Preservar formato juntando runs y re-escribiendo
                texto_completo = parrafo.text
                nuevo_texto = reemplazar_texto(texto_completo)
                if texto_completo != nuevo_texto:
                    # Limpiar y reescribir en el primer run
                    for i, run in enumerate(parrafo.runs):
                        run.text = nuevo_texto if i == 0 else ""

            for p in doc.paragraphs:
                procesar_parrafo(p)

            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        for p in celda.paragraphs:
                            procesar_parrafo(p)

            doc.save(destino)

            self.after(0, self._generar_exito, destino)

        except Exception as e:
            self.after(0, self._generar_error, str(e))

    def _generar_exito(self, destino):
        self.btn_generar.configure(state="normal", text="Generar Documento")
        self.lbl_estado.configure(
            text=f"Documento guardado exitosamente",
            text_color=VERDE,
        )
        respuesta = messagebox.askyesno(
            "¡Listo!",
            f"El documento fue creado:\n\n{destino}\n\n¿Deseas abrirlo ahora?",
        )
        if respuesta:
            os.startfile(destino) if os.name == "nt" else os.system(f'open "{destino}"')

    def _generar_error(self, error):
        self.btn_generar.configure(state="normal", text="Generar Documento")
        self.lbl_estado.configure(text="❌  Error al generar", text_color=ROJO)
        messagebox.showerror("Error al generar documento", error)

    def _set_pasos_bloqueados(self):
        pass  # Los pasos muestran mensajes guía hasta que se carguen los datos

    def _notificar(self, msg):
        """Muestra un toast temporal en el título."""
        original = self.title()
        self.title(msg)
        self.after(3000, lambda: self.title(original))


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = AppContratos()
    app.mainloop()
