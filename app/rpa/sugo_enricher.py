"""
sugo_enricher.py
================
Automatización asincrónica con Playwright para extraer el Folio SUGO de cada
oficio y enriquecer el Excel de layout con esa información.

Flujo:
    1. Lee el Excel generado por la descarga CNBV.
    2. Lanza sesiones en paralelo (múltiples pestañas) para consultar SUGO.
    3. Guarda checkpoints en CSV para permitir reanudar si el proceso falla.
    4. Exporta el Excel enriquecido con la columna 'Folio Sugo'.
"""

import os
import getpass
import asyncio

import pandas as pd
from tqdm.asyncio import tqdm

# FIX: Una sola importación de PlaywrightTimeoutError desde async_api.
# Anteriormente se importaba dos veces (sync + async), generando una colisión
# de nombres que hacía que los bloques `except PlaywrightTimeoutError` en
# funciones async capturaran la clase incorrecta.
from playwright.async_api import (
    async_playwright,
    Page,
    TimeoutError as PlaywrightTimeoutError,
)

from app.config import (
    COLUMNAS_REQUERIDAS,
    CHECKPOINT_EVERY,
    NUM_PESTANAS,
    URL_LOGIN,
    URL_ESTATUS_FOLIO,
    URL_SUGO,
)
from app.utils import split_list


# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------

async def login_sugo(browser, user: str, password: str) -> dict | None:
    """
    Realiza el login en el portal interno SUGO y devuelve el storage state
    (cookies + localStorage) para reutilizarlo en los workers.

    Args:
        browser:  Instancia de navegador Playwright.
        user:     Usuario SUGO.
        password: Contraseña SUGO.

    Returns:
        Diccionario con el storage state, o None si el login falla.
    """
    context = await browser.new_context(
        ignore_https_errors=True, viewport={"width": 1280, "height": 800}
    )
    page = await context.new_page()

    try:
        await page.goto(URL_LOGIN, wait_until="domcontentloaded")
        await asyncio.sleep(2)
        await page.fill(".name", user)
        await page.fill(".pass", password)
        await asyncio.sleep(1)

        async with context.expect_page() as page_info:
            if await page.locator("//p[@onclick='validaCampos()']").is_visible():
                await page.evaluate("() => validaCampos()")
            else:
                await page.keyboard.press("Enter")

        popup = await page_info.value
        await popup.wait_for_load_state()
        await asyncio.sleep(3)

        storage = await context.storage_state()
        return storage

    except Exception as exc:
        print(f"[sugo_enricher] Error durante el login: {exc}")
        return None

    finally:
        await context.close()


# ---------------------------------------------------------------------------
# Consulta individual de folio
# ---------------------------------------------------------------------------

async def get_folio_sugo(oficio: str, page: Page) -> str:
    """
    Consulta el Folio SUGO de un oficio concreto en la página de estatus.

    Args:
        oficio: Número de oficio a consultar.
        page:   Página de Playwright ya autenticada.

    Returns:
        Folio SUGO encontrado o 'NO ENCONTRADO' si no hay resultado o hay error.
    """
    await page.goto(URL_ESTATUS_FOLIO, wait_until="domcontentloaded", timeout=50_000)
    await asyncio.sleep(0.3)

    try:
        checkbox = page.locator("#rOficio")
        await checkbox.wait_for(state="visible")
        await checkbox.focus()
        await page.keyboard.press("Space")
        await page.wait_for_selector("#fOficio.etiqueta2", state="attached", timeout=5_000)

        await page.fill("#fOficio", str(oficio))
        await page.evaluate("buscar();")
        await page.wait_for_selector("#panelDatos1", timeout=3_000)

        folio = await page.locator("#panelDatos1 #Cons1 td:nth-child(2)").inner_text()
        return folio.strip() if folio.strip() else "NO ENCONTRADO"

    except PlaywrightTimeoutError:
        # Puede aparecer un modal de alerta; intentamos aceptarlo y reintentar
        try:
            await page.wait_for_selector("#BTACEPTAR", timeout=40_000)
            await page.click("#BTACEPTAR")
            await asyncio.sleep(0.5)
        except Exception:
            await page.goto(URL_ESTATUS_FOLIO, wait_until="domcontentloaded")
            await asyncio.sleep(0.5)
        return "NO ENCONTRADO"

    except Exception as exc:
        print(f"[sugo_enricher] Error consultando oficio '{oficio}': {exc}")
        await page.goto(URL_ESTATUS_FOLIO, wait_until="domcontentloaded")
        await asyncio.sleep(0.5)
        return "NO ENCONTRADO"


# ---------------------------------------------------------------------------
# Worker de lote
# ---------------------------------------------------------------------------

async def _procesar_lote(
    id_worker: int,
    oficios_lote: list[int],
    context,
    senal_inicio: asyncio.Event,
    pbar,
    lock: asyncio.Lock,
    df: pd.DataFrame,
    output_csv: str,
) -> None:
    """
    Worker asincrónico que procesa un subconjunto de registros en su propia
    pestaña del navegador.

    Args:
        id_worker:    Identificador del worker (para logging).
        oficios_lote: Índices del DataFrame que le corresponden.
        context:      Contexto de Playwright compartido (ya autenticado).
        senal_inicio: Event que señaliza al coordinador que este worker cargó.
        pbar:         Barra de progreso compartida.
        lock:         Lock asincrónico para escrituras seguras al CSV compartido.
        df:           DataFrame compartido donde se guardan los resultados.
        output_csv:   Ruta del CSV de checkpoint.
    """
    if not oficios_lote:
        senal_inicio.set()
        return

    page = await context.new_page()

    async def _handle_dialog(dialog):
        print(f"\n[⚠️ Worker {id_worker}] Alerta del navegador: {dialog.message}")
        await dialog.dismiss()

    page.on("dialog", _handle_dialog)

    try:
        await page.goto(URL_SUGO, wait_until="domcontentloaded", timeout=40_000)
        await asyncio.sleep(2)
        senal_inicio.set()

        for i, idx in enumerate(oficios_lote):
            oficio = df.at[idx, "Oficio"]
            folio  = await get_folio_sugo(oficio=oficio, page=page)

            df.at[idx, "Folio Sugo"]          = folio
            df.at[idx, "Extract Folio Sugo"]  = "PROCESADO"
            pbar.update(1)

            # Checkpoint: guarda progreso cada N registros
            if i > 0 and i % CHECKPOINT_EVERY == 0:
                async with lock:
                    df.to_csv(output_csv, index=False)

    except Exception as exc:
        print(f"\n[!] Error en worker {id_worker}: {exc}")
    finally:
        # Checkpoint final del worker
        async with lock:
            df.to_csv(output_csv, index=False)
        if not senal_inicio.is_set():
            senal_inicio.set()
        await page.close()


# ---------------------------------------------------------------------------
# Orquestador principal
# ---------------------------------------------------------------------------

async def enrich_layout_with_folio_sugo(
    excel_file: str,
    output_excel: str,
    output_csv_temporal: str,
) -> str | None:
    """
    Lee el layout Excel, extrae el Folio SUGO de cada oficio usando Playwright
    en paralelo y guarda el resultado enriquecido.

    Args:
        excel_file:         Ruta al Excel de entrada (generado por download_files_cnbv).
        output_excel:       Ruta del Excel de salida con la columna 'Folio Sugo'.
        output_csv_temporal: Ruta del CSV de checkpoint para reanudar si falla.

    Returns:
        'OK' si el proceso finalizó correctamente, None en caso de error.
    """
    print("\n===============================================")
    print("        LOGIN EN SUGO (Para Extracción)      ")
    print("===============================================")
    user_input = input("Usuario SUGO: ").strip()
    pass_input = getpass.getpass("Contraseña SUGO: ").strip()

    # Carga o reanuda desde checkpoint
    if os.path.exists(output_csv_temporal):
        print("\n[INFO] Checkpoint encontrado. Retomando progreso...")
        df = pd.read_csv(output_csv_temporal)
    else:
        print("\n[INFO] Leyendo layout Excel...")
        try:
            df = pd.read_excel(excel_file, dtype=str).fillna("")
        except Exception as exc:
            print(f"[ERROR] No se pudo leer '{excel_file}': {exc}")
            return None

        columnas_faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
        if columnas_faltantes:
            print(f"\n{'!' * 50}")
            print(f"[ERROR] El archivo '{excel_file}' no contiene las columnas requeridas:")
            print(f"  Faltan: {', '.join(columnas_faltantes)}")
            print(f"{'!' * 50}\n")
            return None

        df["Extract Folio Sugo"] = "SIN PROCESAR"
        df["Folio Sugo"]         = ""

    pendientes = df[df["Extract Folio Sugo"] == "SIN PROCESAR"].index.tolist()

    if not pendientes:
        print("[INFO] No hay registros pendientes de procesar.")
        return "OK"

    print(f"[INFO] Procesando {len(pendientes)} registros...")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, channel="chrome")
        storage_state = await login_sugo(browser, user_input, pass_input)

        if not storage_state:
            print("[ERROR] Login fallido. Abortando.")
            await browser.close()
            return None

        context = await browser.new_context(
            storage_state=storage_state, ignore_https_errors=True
        )

        lotes = list(split_list(pendientes, NUM_PESTANAS))
        lock  = asyncio.Lock()

        bar_format = "{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]"
        pbar = tqdm(
            total=len(pendientes),
            desc="Inicializando workers...",
            unit="reg",
            bar_format=bar_format,
            colour="blue",
        )

        tasks = []
        for i, lote in enumerate(lotes):
            event_load = asyncio.Event()
            pbar.set_description(f"Cargando Worker {i + 1}/{NUM_PESTANAS}")

            task = asyncio.create_task(
                _procesar_lote(
                    id_worker    = i + 1,
                    oficios_lote = lote,
                    context      = context,
                    senal_inicio = event_load,
                    pbar         = pbar,
                    lock         = lock,
                    df           = df,
                    output_csv   = output_csv_temporal,
                )
            )
            tasks.append(task)
            await event_load.wait()

        pbar.set_description("Extrayendo Folio SUGO")
        await asyncio.gather(*tasks)
        pbar.close()

        await context.close()
        await browser.close()

    # Exporta resultado final y elimina checkpoint
    df.to_excel(output_excel, index=False)
    if os.path.exists(output_csv_temporal):
        os.remove(output_csv_temporal)

    print(f"[SUCCESS] Extracción SUGO completada → {output_excel}")
    return "OK"
