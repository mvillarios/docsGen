import customtkinter as ctk

ctk.set_default_color_theme("blue")

AZUL       = "#1E3A5F"
AZUL_MED   = "#2E5896"
AZUL_CLARO = "#EBF2FF"
ACENTO     = "#3B82F6"
VERDE      = "#16A34A"
ROJO       = "#DC2626"
GRIS_BORDE = "#D1D9E6"
GRIS_BG    = "#F5F7FA"
BLANCO     = "#FFFFFF"
TEXTO      = "#1A1A2E"
TEXTO_SUAVE= "#64748B"


def actualizar_para_tema() -> None:
    global AZUL_CLARO, GRIS_BG, TEXTO, TEXTO_SUAVE, GRIS_BORDE
    oscuro = ctk.get_appearance_mode() == "Dark"
    AZUL_CLARO = "#2A3A4A" if oscuro else "#EBF2FF"
    GRIS_BG    = "#2B2B2B" if oscuro else "#F5F7FA"
    TEXTO      = "#E0E0E0" if oscuro else "#1A1A2E"
    TEXTO_SUAVE= "#999999" if oscuro else "#64748B"
    GRIS_BORDE = "#555555" if oscuro else "#D1D9E6"


PASOS = ["archivos", "trabajadores", "datos_faltantes", "generar"]

MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
         "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
