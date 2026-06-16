"""
data_exporter.py
================
Integra los datos del layout Excel con la información extraída de los archivos
Word y XML, y exporta el resultado final a un archivo JSON.
"""

import os
import json

from app.extractor.word_reader import extractor_text_document_word, search_autoridad
from app.extractor.xml_parser import parsear_expediente_cnbv, generar_estructura_vacia_xml


def extract_data_folios_to_json(
    folios_dict: list[dict],
    folder_documents: str,
    json_output: str,
) -> None:
    """
    Procesa cada registro del layout, extrae datos de los archivos Word y XML
    asociados, y guarda el resultado consolidado en un archivo JSON.

    Args:
        folios_dict:      Lista de diccionarios con los datos del layout Excel.
        folder_documents: Ruta al directorio que contiene los archivos .doc y .xml.
        json_output:      Ruta del archivo JSON de salida.
    """
    resultados: list[dict] = []

    total = len(folios_dict)
    for idx, folio in enumerate(folios_dict, start=1):
        oficio              = folio.get("Oficio", "")
        folio_sugo          = folio.get("Folio Sugo", "")
        expediente          = folio.get("Expediente", "")
        publicacion         = folio.get("Publicacion", "")
        plazo               = folio.get("Plazo", "")
        expediente_corto    = folio.get("Expediente Corto", "")
        file_exist_caratula = folio.get("Caratula", "")
        file_exist_xml      = folio.get("XML", "")
        name_file_doc       = folio.get("Archivo Doc", "")
        name_file_xml       = folio.get("Archivo XML", "")

        print(f"[{idx}/{total}] Extrayendo datos del oficio: {oficio}")

        # --- Extracción del nombre de autoridad desde el Word ---
        autoridad_extraida = "No encontrado"
        if name_file_doc:
            path_doc = os.path.join(folder_documents, name_file_doc)
            if os.path.exists(path_doc):
                texto = extractor_text_document_word(path_doc)
                autoridad_extraida = search_autoridad(texto)
            else:
                print(f"  [WARN] Archivo Word no encontrado: {path_doc}")

        # --- Extracción o estructura vacía del XML ---
        if file_exist_xml == "OK" and name_file_xml:
            path_xml = os.path.join(folder_documents, name_file_xml)
            if os.path.exists(path_xml):
                try:
                    datos_xml = parsear_expediente_cnbv(xml_source=path_xml, es_archivo=True)
                except Exception as exc:
                    print(f"  [WARN] Error al parsear XML '{path_xml}': {exc}. Usando estructura vacía.")
                    datos_xml = generar_estructura_vacia_xml()
            else:
                print(f"  [WARN] Archivo XML no encontrado: {path_xml}. Usando estructura vacía.")
                datos_xml = generar_estructura_vacia_xml()
        else:
            datos_xml = generar_estructura_vacia_xml()

        resultados.append(
            {
                "Oficio":           oficio,
                "Folio_Sugo":       folio_sugo,
                "Expediente":       expediente,
                "Publicacion":      publicacion,
                "Plazo":            plazo,
                "Expediente_Corto": expediente_corto,
                "Caratula":         file_exist_caratula,
                "XML_Status":       file_exist_xml,
                "Archivo_Doc":      name_file_doc,
                "Archivo_XML":      name_file_xml,
                # Datos extraídos
                "autoridad_extraida": autoridad_extraida,
                "xml_extraido":       datos_xml,
            }
        )

    # Aseguramos que el directorio de salida exista
    os.makedirs(os.path.dirname(json_output), exist_ok=True)

    with open(json_output, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=4)

    print(f"\n[SUCCESS] Extracción completada.")
    print(f"  -> JSON guardado en: {json_output}")
    print(f"  -> Total de registros exportados: {len(resultados)}")
