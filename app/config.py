"""
config.py
=========
Configuración central del proyecto: rutas, URLs y constantes.
Importa este módulo en lugar de definir constantes locales en cada archivo.
"""

import os

# ---------------------------------------------------------------------------
# Directorios y archivos de salida
# ---------------------------------------------------------------------------
FOLDER_GENERATES: str = "dist"
FOLDER_DOCUMENTS: str = os.path.join(FOLDER_GENERATES, "documents")
FOLDER_SITI_DOCS: str = os.path.join(FOLDER_GENERATES, "siti_docs")

INPUT_FILE: str = "layout.xlsx"
OUTPUT_EXCEL: str = os.path.join(FOLDER_GENERATES, "layout_folio_sugo.xlsx")
OUTPUT_JSON: str = os.path.join(FOLDER_GENERATES, "extract_data.json")
OUTPUT_TEMPORAL: str = os.path.join(FOLDER_GENERATES, "resultados_folio_sugo_temp.csv")

# Archivo generado por la descarga CNBV (input para el enricher de SUGO)
ARCHIVO_LAYOUT_DESCARGAS: str = os.path.join(FOLDER_GENERATES, "layout_descargas.xlsx")

# ---------------------------------------------------------------------------
# Credenciales y resultados CNBV
# ---------------------------------------------------------------------------
ARCHIVO_CREDENCIALES: str = "credenciales.json"

# ---------------------------------------------------------------------------
# Columnas esperadas en el layout Excel
# ---------------------------------------------------------------------------
COLUMNAS_REQUERIDAS: list[str] = [
    "Oficio",
    "Expediente",
    "Publicacion",
    "Plazo",
    "Expediente Corto",
    "Caratula",
    "XML",
    "Archivo Doc",
    "Archivo XML",
]

COLUMNAS_CON_SUGO: list[str] = [
    "Oficio",
    "Folio Sugo",
    "Extract Folio Sugo",
    "Expediente",
    "Publicacion",
    "Plazo",
    "Expediente Corto",
    "Caratula",
    "XML",
    "Archivo Doc",
    "Archivo XML",
]

# ---------------------------------------------------------------------------
# Configuración del proceso de enriquecimiento SUGO
# ---------------------------------------------------------------------------
CHECKPOINT_EVERY: int = 25   # Guarda el CSV temporal cada N registros procesados
NUM_PESTANAS: int = 5        # Número de pestañas / workers paralelos en Playwright

# ---------------------------------------------------------------------------
# URLs del sistema interno
# ---------------------------------------------------------------------------
URL_LOGIN: str = (
    "https://acprod.intranet.com.mx/mbom_mx_ws/mbom_mx_web/PortalLogon"
)
URL_ESTATUS_FOLIO: str = (
    "https://acprod.intranet.com.mx:443/boixp_mx_web/boixp_mx_web/servlet/"
    "ServletOperacionWeb?OPERACION=VGOMX012&LOCALE=es_ES"
    "&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10090"
)
URL_CONSULTA_CEDULA: str = (
    "https://acprod.intranet.com.mx/boixp_mx_web/boixp_mx_web/servlet/"
    "ServletOperacionWeb?OPERACION=VGOMX042&LOCALE=es_ES"
    "&DATOS_ENTRADA.FLUJO_LANZAR=GOMXFL10190"
)
URL_SUGO: str = (
    "https://acprod.intranet.com.mx/mbom_mx_ws/mbom_mx_web/mbom_mx_web_jsp/portal3.jsp"
)

# ---------------------------------------------------------------------------
# URLs del portal público CNBV
# ---------------------------------------------------------------------------
URL_CNBV_LOGIN: str = "https://websitiaa.cnbv.gob.mx/logOn.aspx"
URL_CNBV_PUBLICADOS: str = "https://websitiaa.cnbv.gob.mx/Publicados.aspx"
