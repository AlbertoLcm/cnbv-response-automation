"""
main.py
=======
Punto de entrada del proyecto cnbv-response-automation.

Orquesta las tres etapas del flujo en orden:
    1. Descarga de archivos (carátulas + XMLs) del portal CNBV.
    2. Enriquecimiento del layout con el Folio SUGO (portal interno).
    3. Generación de documentos Word con formato SITI/BBVA.
"""

import asyncio
import json
import sys
import platform

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


def _validar_layout_sugo(excel_path: str) -> pd.DataFrame:
    """
    Lee y valida el Excel enriquecido con Folio SUGO.
    Termina el proceso si faltan columnas o si algún Folio SUGO está vacío.

    Args:
        excel_path: Ruta al archivo Excel a validar.

    Returns:
        DataFrame validado y listo para procesar.
    """
    print("\n[INFO] Validando layout con Folio SUGO...")
    try:
        df = pd.read_excel(excel_path, dtype=str).fillna("")
    except Exception as exc:
        print(f"[ERROR] No se pudo leer '{excel_path}': {exc}")
        sys.exit(1)

    faltantes = [c for c in COLUMNAS_CON_SUGO if c not in df.columns]
    if faltantes:
        print(f"\n{'!' * 50}")
        print(f"[ERROR] El archivo '{excel_path}' no tiene las columnas requeridas.")
        print(f"  Faltan: {', '.join(faltantes)}")
        print(f"{'!' * 50}\n")
        sys.exit(1)

    sin_folio = df["Folio Sugo"].isna() | (df["Folio Sugo"].astype(str).str.strip() == "")
    if sin_folio.any():
        filas = df[sin_folio].index.tolist()
        print(f"\n[ERROR] {len(filas)} registro(s) sin 'Folio Sugo': filas {filas}")
        sys.exit("Ejecución detenida: Faltan datos en la columna 'Folio Sugo'.")

    print("[SUCCESS] Validación correcta. Todos los registros tienen Folio SUGO.")
    return df


def _generar_documentos(json_path: str, dir_salida: str) -> None:
    """
    Lee el JSON de datos extraídos y genera un documento Word por cada oficio.

    Args:
        json_path:  Ruta al archivo JSON generado por el extractor.
        dir_salida: Directorio donde se guardarán los documentos Word.
    """
    print(f"\n[INFO] Generando documentos Word en '{dir_salida}'...")

    with open(json_path, "r", encoding="utf-8") as f:
        data: list[dict] = json.load(f)

    total = len(data)
    generados = 0

    for idx, oficio in enumerate(data, start=1):
        # Navegación segura por la estructura JSON
        xml        = oficio.get("xml_extraido", {}) or {}
        solicitudes = xml.get("SolicitudesEspecificas", [])

        nombre = "NO ENCONTRADO"
        rfc    = "SIN RFC"

        if solicitudes:
            personas = solicitudes[0].get("PersonasSolicitud", [])
            if personas:
                nombre = personas[0].get("Nombre", "NO ENCONTRADO")
                rfc    = personas[0].get("Rfc",    "SIN RFC")

        variables = {
            "oficio":           oficio.get("Oficio",          ""),
            "expediente":       "",
            "folio":            oficio.get("Expediente",      ""),
            "expediente_corto": oficio.get("Expediente_Corto", ""),
            "autoridad":        oficio.get("autoridad_extraida", ""),
            "tipo_respuesta":   "Total",
            "tipo_asunto":      "Aseguramiento",
            "nombre":           nombre,
            "rfc":              rfc,
            "variable_dof":     oficio.get("Folio_Sugo",     ""),
        }

        print(f"[{idx}/{total}] Generando Word para oficio: {variables['oficio']}")
        ruta = generar_documento_siti(variables, dir_salida=dir_salida)

        if ruta:
            generados += 1
        else:
            print(f"  [WARN] No se pudo generar el documento para oficio {variables['oficio']}.")

    print(f"\n[SUCCESS] Documentos generados: {generados}/{total} en '{dir_salida}'")


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    print_banner_ascii()

    # ── ETAPA 1: Descarga CNBV ──────────────────────────────────────────────
    print("\n" + "=" * 55)
    print("  ETAPA 1: Descarga de archivos del portal CNBV")
    print("=" * 55)
    archivo_layout = download_files_cnbv()

    if not archivo_layout:
        print("[ERROR] La descarga CNBV no generó resultados. Abortando.")
        sys.exit(1)

    # ── ETAPA 2: Enriquecimiento con Folio SUGO ─────────────────────────────
    print("\n" + "=" * 55)
    print("  ETAPA 2: Extracción de Folio SUGO")
    print("=" * 55)

    # FIX: se pasa `archivo_layout` (el Excel real de la descarga) en lugar del
    # archivo de entrada original. Así el enricher siempre procesa datos frescos.
    resultado_sugo = asyncio.run(
        enrich_layout_with_folio_sugo(
            excel_file          = archivo_layout,
            output_excel        = OUTPUT_EXCEL,
            output_csv_temporal = OUTPUT_TEMPORAL,
        )
    )

    if resultado_sugo != "OK":
        print("[ERROR] La extracción de Folio SUGO falló. Abortando.")
        sys.exit(1)

    # ── Validación del layout enriquecido ───────────────────────────────────
    df = _validar_layout_sugo(OUTPUT_EXCEL)

    # ── ETAPA 3a: Extracción de datos (Word + XML → JSON) ───────────────────
    print("\n" + "=" * 55)
    print("  ETAPA 3: Extracción de datos y generación de documentos")
    print("=" * 55)

    extract_data_folios_to_json(
        folios_dict      = df.to_dict("records"),
        folder_documents = FOLDER_DOCUMENTS,
        json_output      = OUTPUT_JSON,
    )

    # ── ETAPA 3b: Generación de documentos Word ─────────────────────────────
    _generar_documentos(json_path=OUTPUT_JSON, dir_salida=FOLDER_SITI_DOCS)

    print("\n" + "=" * 55)
    print("  ✅ Proceso completado exitosamente.")
    print(f"  📁 Documentos Word:  {FOLDER_SITI_DOCS}/")
    print(f"  📄 JSON de datos:    {OUTPUT_JSON}")
    print(f"  📊 Excel SUGO:       {OUTPUT_EXCEL}")
    print("=" * 55)


if __name__ == "__main__":
    # Guard para Windows: Playwright async requiere el policy correcto
    if platform.system() == "Windows":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    main()