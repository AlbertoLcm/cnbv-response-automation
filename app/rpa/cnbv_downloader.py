"""
cnbv_downloader.py
==================
Automatización sincrónica con Playwright para descargar archivos (carátulas y
XMLs) publicados en el portal CNBV (websitiaa.cnbv.gob.mx).

Flujo:
    1. Carga credenciales del archivo JSON.
    2. Solicita el rango de fechas al usuario.
    3. Prepara el entorno de salida (dist/documents/).
    4. Navega por el portal CNBV y descarga los archivos.
    5. Exporta un Excel con el registro de descargas.
    6. Retorna la ruta del Excel generado.
"""

import os
import sys
import json
import shutil

import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from app.config import (
    ARCHIVO_CREDENCIALES,
    FOLDER_DOCUMENTS,
    FOLDER_GENERATES,
    ARCHIVO_LAYOUT_DESCARGAS,
    URL_CNBV_LOGIN,
    URL_CNBV_PUBLICADOS,
)


# ---------------------------------------------------------------------------
# Credenciales
# ---------------------------------------------------------------------------

def cargar_credenciales() -> tuple[str, str]:
    """
    Carga las credenciales desde el archivo JSON configurado.

    Returns:
        Tupla (usuario, contrasena).

    Raises:
        SystemExit: Si el archivo no existe, falta alguna clave o hay otro error.
    """
    try:
        with open(ARCHIVO_CREDENCIALES, "r", encoding="utf-8") as f:
            creds = json.load(f)
        return creds["usuario"], creds["contrasena"]
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{ARCHIVO_CREDENCIALES}'.")
        sys.exit(1)
    except KeyError as exc:
        print(f"❌ Error: Formato de credenciales inválido. Falta la clave {exc}.")
        sys.exit(1)
    except Exception as exc:
        print(f"❌ Error inesperado al cargar credenciales: {exc}")
        sys.exit(1)


# ---------------------------------------------------------------------------
# Entorno de salida
# ---------------------------------------------------------------------------

def preparar_entorno() -> None:
    """
    Limpia y recrea el directorio de documentos para una ejecución limpia.
    Solo borra `dist/documents/`, nunca todo `dist/`.
    """
    if os.path.exists(FOLDER_DOCUMENTS):
        shutil.rmtree(FOLDER_DOCUMENTS)
    os.makedirs(FOLDER_DOCUMENTS, exist_ok=True)
    os.makedirs(FOLDER_GENERATES, exist_ok=True)
    print("✅ Entorno preparado (dist/documents/ limpiado y recreado).")


# ---------------------------------------------------------------------------
# Entrada de fechas
# ---------------------------------------------------------------------------

def obtener_fechas_terminal() -> tuple[str, str]:
    """
    Solicita y valida interactivamente el rango de fechas para la búsqueda.

    Returns:
        Tupla (fecha_inicio_str, fecha_fin_str) en formato 'AAAA-MM-DD'.
    """
    print("\n--- Configuración de Fechas ---")
    print("Formato requerido: AAAA-MM-DD (ejemplo: 2023-01-31)\n")

    def _pedir_fecha(mensaje: str) -> tuple[datetime, str]:
        while True:
            fecha_str = input(mensaje).strip()
            try:
                fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
                return fecha_obj, fecha_str
            except ValueError:
                print("❌ Formato incorrecto. Por favor usa AAAA-MM-DD.")

    fecha_inicio_obj, fecha_inicio_str = _pedir_fecha("Ingresa la Fecha de INICIO: ")

    while True:
        fecha_fin_obj, fecha_fin_str = _pedir_fecha("Ingresa la Fecha de FIN: ")
        if fecha_fin_obj >= fecha_inicio_obj:
            break
        print("❌ La fecha de fin debe ser igual o posterior a la de inicio.")

    return fecha_inicio_str, fecha_fin_str


# ---------------------------------------------------------------------------
# Descarga individual
# ---------------------------------------------------------------------------

def intentar_descarga(page, locator, ruta_destino: str) -> str:
    """
    Descarga un archivo interceptando el evento de descarga de Playwright.

    Args:
        page:          Instancia de `Page` de Playwright.
        locator:       Locator del enlace a hacer clic.
        ruta_destino:  Ruta completa donde se guardará el archivo.

    Returns:
        'OK'                    → descarga exitosa.
        'NO DISPONIBLE'         → el locator no encontró ningún enlace.
        'ERROR (Tiempo de espera)' → timeout al esperar la descarga.
        'ERROR (<detalle>)'     → cualquier otro error.
    """
    try:
        if locator.count() == 0:
            return "NO DISPONIBLE"

        with page.expect_download(timeout=15_000) as download_info:
            locator.click()

        download_info.value.save_as(ruta_destino)
        return "OK"

    except PlaywrightTimeoutError:
        return "ERROR (Tiempo de espera)"
    except Exception as exc:
        return f"ERROR ({exc})"


# ---------------------------------------------------------------------------
# Automatización principal
# ---------------------------------------------------------------------------

def ejecutar_automatizacion(
    usuario: str,
    password: str,
    fecha_inicio: str,
    fecha_fin: str,
) -> list[dict]:
    """
    Navega por el portal CNBV, busca oficios en el rango de fechas dado
    y descarga la carátula (Word) y el XML de cada uno.

    Args:
        usuario:      Usuario del portal CNBV.
        password:     Contraseña del portal CNBV.
        fecha_inicio: Fecha de inicio en formato 'AAAA-MM-DD'.
        fecha_fin:    Fecha de fin en formato 'AAAA-MM-DD'.

    Returns:
        Lista de diccionarios con los resultados de cada oficio procesado.
    """
    resultados: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            print("\n🚀 Iniciando sesión en el portal CNBV...")
            page.goto(URL_CNBV_LOGIN)
            page.fill("#ctl00_DefaultPlaceholder_textBoxUser", usuario)
            page.fill("#ctl00_DefaultPlaceholder_textBoxPassword", password)
            page.click("#ButtonValidate")
            page.wait_for_load_state("networkidle")

            print(f"🔎 Buscando registros del {fecha_inicio} al {fecha_fin}...")
            page.goto(URL_CNBV_PUBLICADOS)
            page.select_option(
                "#ctl00_DefaultPlaceholder_ComboBoxAreas", label="Aseguramiento"
            )
            page.select_option(
                "#ctl00_DefaultPlaceholder_ComboBoxEstatusOficio", label="Todos"
            )
            page.fill("#ctl00_DefaultPlaceholder_TextFechaPublicacion1", fecha_inicio)
            page.fill("#ctl00_DefaultPlaceholder_TextFechaPublicacion2", fecha_fin)
            page.click("#ctl00_DefaultPlaceholder_ButtonQuery")

            try:
                page.wait_for_selector(
                    "#ctl00_DefaultPlaceholder_GridResult", timeout=10_000
                )
            except PlaywrightTimeoutError:
                print(
                    "⚠️ No se encontró la tabla de resultados. "
                    "Posiblemente no hay registros en esas fechas."
                )
                return resultados

            filas = page.locator("#ctl00_DefaultPlaceholder_GridResult tr").all()
            total_oficios = len(filas) - 1  # restamos fila de encabezado
            print(f"📊 Total de oficios encontrados: {total_oficios}")

            for i, fila in enumerate(filas[1:], start=1):
                columnas = fila.locator("td").all()
                if len(columnas) < 10:
                    continue

                oficio           = columnas[5].inner_text().strip()
                expediente       = columnas[6].inner_text().strip()
                publicacion      = columnas[7].inner_text().strip()
                plazo            = columnas[9].inner_text().strip()
                expediente_corto = expediente[-10:] if len(expediente) >= 10 else expediente

                archivo_doc = f"{oficio}.doc".replace("/", "-")
                archivo_xml = f"{oficio}.xml".replace("/", "-")
                ruta_doc    = os.path.join(FOLDER_DOCUMENTS, archivo_doc)
                ruta_xml    = os.path.join(FOLDER_DOCUMENTS, archivo_xml)

                print(f"[{i}/{total_oficios}] Procesando: {oficio}")

                estado_caratula = intentar_descarga(
                    page, columnas[-3].locator("a"), ruta_doc
                )
                estado_xml = intentar_descarga(
                    page, columnas[-1].locator("a"), ruta_xml
                )

                print(f"    ↳ Carátula: {estado_caratula} | XML: {estado_xml}")

                resultados.append(
                    {
                        "Oficio":        oficio,
                        "Expediente":    expediente,
                        "Publicacion":   publicacion,
                        "Plazo":         plazo,
                        "Expediente Corto": expediente_corto,
                        "Caratula":      estado_caratula,
                        "XML":           estado_xml,
                        "Archivo Doc":   archivo_doc if estado_caratula == "OK" else "",
                        "Archivo XML":   archivo_xml if estado_xml    == "OK" else "",
                    }
                )

        except Exception as exc:
            print(f"\n❌ Error crítico durante la automatización CNBV: {exc}")
        finally:
            context.close()
            browser.close()

    return resultados


# ---------------------------------------------------------------------------
# Entry point del módulo
# ---------------------------------------------------------------------------

def download_files_cnbv() -> str | None:
    """
    Orquesta el proceso completo de descarga de archivos del portal CNBV.

    Returns:
        Ruta del archivo Excel de resultados, o None si no hubo datos que exportar.
    """
    print("=========================================")
    print("   Bot de Descarga de Acuses (CNBV)      ")
    print("=========================================\n")

    usuario, password      = cargar_credenciales()
    fecha_inicio, fecha_fin = obtener_fechas_terminal()
    preparar_entorno()

    datos = ejecutar_automatizacion(usuario, password, fecha_inicio, fecha_fin)

    if not datos:
        print("\n⚠️ El proceso finalizó sin resultados para exportar.")
        return None

    df = pd.DataFrame(datos)
    df.to_excel(ARCHIVO_LAYOUT_DESCARGAS, index=False)
    print(f"\n✅ {len(datos)} registros guardados en '{ARCHIVO_LAYOUT_DESCARGAS}'.")
    print(f"📁 Archivos descargados en '{FOLDER_DOCUMENTS}'.")
    return ARCHIVO_LAYOUT_DESCARGAS
