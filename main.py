"""
main.py
=======
Punto de entrada del proyecto cnbv-response-automation.

Recopila todas las entradas del usuario AL INICIO (sin interrupciones),
luego orquesta las tres etapas del flujo de forma desatendida:

    1. Descarga de archivos (carátulas + XMLs) del portal CNBV.
    2. Enriquecimiento del layout con el Folio SUGO (portal interno).
    3. Extracción de datos y generación de documentos Word SITI/BBVA.
"""

import asyncio
import getpass
import json
import platform
import sys
import time
from datetime import datetime

import pandas as pd

from app.config import (
    COLUMNAS_CON_SUGO,
    FOLDER_DOCUMENTS,
    FOLDER_SITI_DOCS,
    OUTPUT_EXCEL,
    OUTPUT_JSON,
    OUTPUT_TEMPORAL,
)
from app.utils import print_banner_ascii
from app.rpa.cnbv_downloader import download_files_cnbv
from app.rpa.sugo_enricher import enrich_layout_with_folio_sugo
from app.extractor.data_exporter import extract_data_folios_to_json
from app.document.word_generator import generar_documento_siti


# ---------------------------------------------------------------------------
# Helpers de logging
# ---------------------------------------------------------------------------

def _seccion(titulo: str) -> None:
    """Imprime un encabezado de sección con timestamp."""
    ts = datetime.now().strftime("%H:%M:%S")
    linea = "─" * 55
    print(f"\n┌{linea}┐")
    print(f"│  [{ts}]  {titulo:<42}│")
    print(f"└{linea}┘")


def _ok(msg: str) -> None:
    print(f"  ✅  {msg}")


def _info(msg: str) -> None:
    print(f"  ℹ️   {msg}")


def _warn(msg: str) -> None:
    print(f"  ⚠️   {msg}")


def _error(msg: str) -> None:
    print(f"  ❌  {msg}")


def _elapsed(segundos: float) -> str:
    m, s = divmod(int(segundos), 60)
    return f"{m}m {s}s" if m else f"{s}s"


# ---------------------------------------------------------------------------
# Recopilación de entradas al inicio
# ---------------------------------------------------------------------------

def _pedir_fecha(mensaje: str) -> tuple[datetime, str]:
    """Solicita una fecha con validación de formato."""
    while True:
        raw = input(f"  {mensaje}").strip()
        try:
            obj = datetime.strptime(raw, "%Y-%m-%d")
            return obj, raw
        except ValueError:
            _warn("Formato incorrecto. Usa AAAA-MM-DD  (ej: 2024-03-15)")


def recopilar_inputs() -> dict:
    """
    Solicita de forma interactiva todos los datos necesarios antes de comenzar
    el proceso automatizado. Una vez completado, no se requiere intervención.

    Returns:
        Diccionario con las claves:
            fecha_inicio, fecha_fin   → str  (rango de búsqueda CNBV)
            sugo_user, sugo_pass      → str  (credenciales portal SUGO)
    """
    print("\n  Ingresa los datos necesarios para el proceso completo.")
    print("  El sistema correrá sin interrupciones una vez que confirmes.\n")

    # ── Fechas CNBV ──────────────────────────────────────────────────────────
    print("  📅  Rango de fechas para búsqueda en el portal CNBV")
    print("      Formato: AAAA-MM-DD\n")

    fecha_inicio_obj, fecha_inicio_str = _pedir_fecha("Fecha INICIO: ")

    while True:
        fecha_fin_obj, fecha_fin_str = _pedir_fecha("Fecha FIN:    ")
        if fecha_fin_obj >= fecha_inicio_obj:
            break
        _warn("La fecha de fin debe ser igual o posterior a la de inicio.")

    # ── Credenciales SUGO ────────────────────────────────────────────────────
    print("\n  🔐  Credenciales del portal interno SUGO\n")
    sugo_user = input("  Usuario SUGO:     ").strip()
    sugo_pass = getpass.getpass("  Contraseña SUGO: ").strip()

    print()
    print("  ┌─────────────────────────────────────────────┐")
    print(f"  │  Período CNBV : {fecha_inicio_str}  →  {fecha_fin_str}          │")
    print(f"  │  Usuario SUGO : {sugo_user:<29}│")
    print("  └─────────────────────────────────────────────┘")

    confirmacion = input("\n  ¿Iniciar proceso? [S/n]: ").strip().lower()
    if confirmacion not in ("", "s", "si", "sí", "y", "yes"):
        print("\n  Proceso cancelado por el usuario.")
        sys.exit(0)

    return {
        "fecha_inicio": fecha_inicio_str,
        "fecha_fin":    fecha_fin_str,
        "sugo_user":    sugo_user,
        "sugo_pass":    sugo_pass,
    }


# ---------------------------------------------------------------------------
# Validación del layout enriquecido
# ---------------------------------------------------------------------------

def _validar_layout_sugo(excel_path: str) -> pd.DataFrame:
    """
    Lee y valida el Excel enriquecido con Folio SUGO.
    Termina el proceso si faltan columnas o si algún Folio SUGO está vacío.

    Args:
        excel_path: Ruta al archivo Excel a validar.

    Returns:
        DataFrame validado y listo para procesar.
    """
    try:
        df = pd.read_excel(excel_path, dtype=str).fillna("")
    except Exception as exc:
        _error(f"No se pudo leer '{excel_path}': {exc}")
        sys.exit(1)

    faltantes = [c for c in COLUMNAS_CON_SUGO if c not in df.columns]
    if faltantes:
        _error(f"El layout no tiene las columnas requeridas: {', '.join(faltantes)}")
        sys.exit(1)

    sin_folio = df["Folio Sugo"].isna() | (df["Folio Sugo"].astype(str).str.strip() == "")
    if sin_folio.any():
        filas = df[sin_folio].index.tolist()
        _error(f"{len(filas)} registro(s) sin 'Folio Sugo'  →  filas: {filas}")
        sys.exit("Proceso detenido: faltan datos en 'Folio Sugo'.")

    _ok(f"Layout validado — {len(df)} registros con Folio SUGO completo.")
    return df


# ---------------------------------------------------------------------------
# Generación de documentos Word
# ---------------------------------------------------------------------------

def _generar_documentos(json_path: str, dir_salida: str) -> None:
    """
    Lee el JSON de datos extraídos y genera un documento Word por cada oficio.

    Args:
        json_path:  Ruta al archivo JSON generado por el extractor.
        dir_salida: Directorio donde se guardarán los documentos Word.
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data: list[dict] = json.load(f)

    total    = len(data)
    generados = 0
    errores   = 0

    for idx, oficio in enumerate(data, start=1):
        xml         = oficio.get("xml_extraido", {}) or {}
        solicitudes = xml.get("SolicitudesEspecificas", [])

        nombre = "NO ENCONTRADO"
        rfc    = "SIN RFC"

        if solicitudes:
            personas = solicitudes[0].get("PersonasSolicitud", [])
            if personas:
                nombre = personas[0].get("Nombre", "NO ENCONTRADO")
                rfc    = personas[0].get("Rfc",    "SIN RFC")

        variables = {
            "oficio":           oficio.get("Oficio",           ""),
            "expediente":       "",
            "folio":            oficio.get("Expediente",       ""),
            "expediente_corto": oficio.get("Expediente_Corto", ""),
            "autoridad":        oficio.get("autoridad_extraida", ""),
            "tipo_respuesta":   "Total",
            "tipo_asunto":      "Aseguramiento",
            "nombre":           nombre,
            "rfc":              rfc,
            "variable_dof":     oficio.get("Folio_Sugo", ""),
        }

        oficio_num = variables["oficio"] or f"#{idx}"
        print(f"    [{idx:>3}/{total}]  Generando Word  →  {oficio_num}")

        ruta = generar_documento_siti(variables, dir_salida=dir_salida)
        if ruta:
            generados += 1
        else:
            errores += 1
            _warn(f"No se pudo generar el documento para oficio '{oficio_num}'.")

    _ok(f"Documentos generados: {generados}/{total}.")
    if errores:
        _warn(f"{errores} documento(s) con error — revisa los avisos anteriores.")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    print_banner_ascii()

    # ── Recopilar TODOS los inputs antes de empezar ─────────────────────────
    _seccion("Configuración inicial")
    inputs = recopilar_inputs()

    tiempo_total_inicio = time.time()

    # ── ETAPA 1: Descarga CNBV ───────────────────────────────────────────────
    _seccion("Etapa 1 / 3  —  Descarga portal CNBV")
    t0 = time.time()

    archivo_layout = download_files_cnbv(
        fecha_inicio = inputs["fecha_inicio"],
        fecha_fin    = inputs["fecha_fin"],
    )

    if not archivo_layout:
        _warn("La descarga no encontró registros en el período indicado. Proceso finalizado.")
        sys.exit(0)

    _ok(f"Layout de descargas guardado en: {archivo_layout}")
    _info(f"Tiempo etapa 1: {_elapsed(time.time() - t0)}")

    # ── ETAPA 2: Enriquecimiento SUGO ────────────────────────────────────────
    _seccion("Etapa 2 / 3  —  Extracción Folio SUGO")
    _info(f"Iniciando sesión en SUGO con usuario '{inputs['sugo_user']}'...")
    t0 = time.time()

    resultado_sugo = asyncio.run(
        enrich_layout_with_folio_sugo(
            excel_file          = archivo_layout,
            output_excel        = OUTPUT_EXCEL,
            output_csv_temporal = OUTPUT_TEMPORAL,
            user                = inputs["sugo_user"],
            password            = inputs["sugo_pass"],
        )
    )

    if resultado_sugo != "OK":
        _error("La extracción de Folio SUGO falló. Revisa las credenciales o la conexión.")
        sys.exit(1)

    _ok(f"Excel enriquecido guardado en: {OUTPUT_EXCEL}")
    _info(f"Tiempo etapa 2: {_elapsed(time.time() - t0)}")

    # ── Validación del layout ────────────────────────────────────────────────
    _seccion("Validación del layout enriquecido")
    df = _validar_layout_sugo(OUTPUT_EXCEL)

    # ── ETAPA 3a: Extracción Word + XML → JSON ───────────────────────────────
    _seccion("Etapa 3 / 3  —  Extracción de datos y documentos Word")
    t0 = time.time()

    _info(f"Extrayendo datos de {len(df)} oficios desde '{FOLDER_DOCUMENTS}'...")
    extract_data_folios_to_json(
        folios_dict      = df.to_dict("records"),
        folder_documents = FOLDER_DOCUMENTS,
        json_output      = OUTPUT_JSON,
    )
    _ok(f"JSON de datos generado: {OUTPUT_JSON}")

    # ── ETAPA 3b: Generación de documentos Word ──────────────────────────────
    _info(f"Generando documentos Word en '{FOLDER_SITI_DOCS}'...")
    _generar_documentos(json_path=OUTPUT_JSON, dir_salida=FOLDER_SITI_DOCS)
    _info(f"Tiempo etapa 3: {_elapsed(time.time() - t0)}")

    # ── Resumen final ─────────────────────────────────────────────────────────
    tiempo_total = _elapsed(time.time() - tiempo_total_inicio)
    print()
    print("  ╔═══════════════════════════════════════════════════════╗")
    print("  ║            ✅  PROCESO COMPLETADO                     ║")
    print("  ╠═══════════════════════════════════════════════════════╣")
    print(f"  ║  📊 Excel SUGO    →  {OUTPUT_EXCEL:<34}║")
    print(f"  ║  📄 JSON datos    →  {OUTPUT_JSON:<34}║")
    print(f"  ║  📁 Docs Word     →  {FOLDER_SITI_DOCS + '/':<34}║")
    print(f"  ║  ⏱️  Tiempo total  →  {tiempo_total:<34}║")
    print("  ╚═══════════════════════════════════════════════════════╝")
    print()


if __name__ == "__main__":
    # Guard para Windows: Playwright async requiere el policy correcto
    if platform.system() == "Windows":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    main()