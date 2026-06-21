# -*- coding: utf-8 -*-
"""
Descarga automatizada de documentos del portal SMV (Bp_Memorias).

Flujo por empresa:
    1. Abre la URL de memorias del SMV.
    2. Escribe el nombre de la empresa en el buscador y dispara la busqueda.
    3. Espera la tabla dinamica de resultados (id = MainContent_grdMemorias).
    4. Recorre las filas (la tabla tiene 5 columnas):
         col 1 -> Anio
         col 2 -> Documento  (aqui se busca el patron "SECCIONES B Y C")
         col 5 -> enlace de descarga
       Descarga la fila SOLO si:
         - la col 2 contiene el PATRON, y
         - la col 1 (anio) esta en {anio actual, anio-1, anio-2}.
    5. Renombra cada archivo descargado y, opcionalmente, descarta los que
       tengan <= N "hojas" (paginas en PDF / worksheets en Excel).

Restricciones de diseno:
    - No usa xlwings ni win32com (cuelgan el kernel).
    - Es un archivo completo, ejecutable por celdas (# %%) en Spyder.
    - Comentarios en espaniol.

Dependencias (env fran36):
    pip install selenium pypdf openpyxl
    # Selenium >= 4.6 auto-gestiona el chromedriver (Selenium Manager).
    # Para .xls antiguos (opcional, conteo de hojas):  pip install xlrd
"""

# %% ===========================  IMPORTS  ====================================
import os
import re
import time
import shutil
import unicodedata
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException,
    WebDriverException,
)


# %% ===========================  PARAMETROS  =================================

URL_SMV = "https://www.smv.gob.pe/SIMV/Bp_Memorias?op=bq12"

# Patron a buscar dentro de la 2da columna (Documento). Sin tildes/espacios
# extra: la comparacion se hace normalizada (ver _norm()).
PATRON_DOCUMENTO = "SECCIONES B Y C"

# Anios objetivo: actual, pasado y antepasado (dinamico).
_ANIO_ACTUAL = datetime.now().year
ANIOS_OBJETIVO = {_ANIO_ACTUAL, _ANIO_ACTUAL - 1, _ANIO_ACTUAL - 2}

# Ruta base donde se guardan las descargas. Dentro se crea una subcarpeta por
# empresa. Ajustar si cambia la maquina.
RUTA_BASE = r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA\Docs privados\Otros\Remuneracion Directorio - Bottom Up\Remuneraciones"

# Filtro por numero de "hojas":
#   - PDF  -> numero de paginas.
#   - XLSX -> numero de worksheets (hojas) [interpretacion de "hojas"].
# Pon None para DESACTIVAR el filtro (descargar todo sin descartar nada).
# Con valor 4 -> se descartan los archivos con <= 4 hojas.
FILTRAR_MIN_HOJAS = 4

# Que hacer con los archivos descartados por el filtro:
#   "mover"    -> se mueven a la subcarpeta _descartados (reversible, recomendado).
#   "eliminar" -> se borran.
ACCION_DESCARTE = "mover"

# Ejecutar el navegador visible (False) o sin ventana (True).
# Para depurar la interaccion de busqueda conviene dejarlo visible.
HEADLESS = False

# Lista de empresas. El segundo valor es el nombre con el que se renombra la
# carpeta/archivos (util para FOSSAL -> HOCHSCHILD MINING PERU).
EMPRESAS = [
    ("COMPAÑIA DE MINAS BUENAVENTURA S.A.A.", "COMPAÑIA DE MINAS BUENAVENTURA S.A.A."),
    ("MINSUR S.A.", "MINSUR S.A."),
    ("VOLCAN COMPAÑIA MINERA S.A.A.", "VOLCAN COMPAÑIA MINERA S.A.A."),
    ("NEXA RESOURCES PERU S.A.A.", "NEXA RESOURCES PERU S.A.A."),
    ("ORYGEN PERU S.A.A.", "ORYGEN PERU S.A.A."),
    ("PLUZ ENERGIA PERU S.A.A.", "PLUZ ENERGIA PERU S.A.A."),
    ("ENGIE ENERGIA PERU S.A.A.", "ENGIE ENERGIA PERU S.A.A."),
    ("FERREYCORP S.A.A.", "FERREYCORP S.A.A."),
    ("AENZA S.A.A.", "AENZA S.A.A."),
    ("UNACEM CORP SOCIEDAD ANONIMA ABIERTA - UNACEM CORP S.A.A.", "UNACEM CORP S.A.A."),
    ("CEMENTOS PACASMAYO S.A.A.", "CEMENTOS PACASMAYO S.A.A."),
    ("INRETAIL PERU CORP.", "INRETAIL PERU CORP"),
    ("ALICORP S.A.A.", "ALICORP S.A.A."),
    ("CREDICORP LTD.", "CREDICORP LTD"),
    ("INTERCORP FINANCIAL SERVICES INC.", "INTERCORP FINANCIAL SERVICES INC"),
    ("BANCO BBVA PERU", "BANCO BBVA PERU"),
    ("EMPRESA EDITORA EL COMERCIO S.A.", "EMPRESA EDITORA EL COMERCIO S.A."),
    ("FOSSAL S.A.A.", "HOCHSCHILD MINING PERU"),   # se busca FOSSAL, se renombra
]

# Timeouts (segundos)
TIMEOUT_BUSQUEDA = 25     # espera a que cargue la tabla tras buscar
TIMEOUT_DESCARGA = 90     # espera maxima por archivo descargado


# %% ======================  UTILIDADES DE TEXTO  ============================

def _norm(texto):
    """Normaliza texto para comparar: mayusculas, sin tildes, sin espacios
    redundantes. Asi el patron 'SECCIONES B Y C' matchea aunque haya tildes
    o espacios de relleno en la celda."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = "".join(c for c in t if not unicodedata.combining(c))
    t = re.sub(r"\s+", " ", t).strip().upper()
    return t


def _nombre_seguro(texto):
    """Convierte un texto en un nombre de carpeta/archivo valido en Windows."""
    t = _norm(texto)
    t = re.sub(r'[<>:"/\\|?*]', "", t)        # caracteres prohibidos
    t = re.sub(r"\s+", "_", t).strip("_")
    return t[:120] if t else "SIN_NOMBRE"


def limpiar_url_documento(href):
    """Extrae el GUID 'vidDoc' del href (que viene con espacios de relleno)
    y reconstruye una URL limpia de descarga."""
    if not href:
        return None
    m = re.search(r"vidDoc=\{([0-9A-Fa-f\-]+)\}", href)
    if not m:
        return None
    guid = m.group(1).strip()
    return f"https://www.smv.gob.pe/ConsultasP8/documento.aspx?vidDoc={{{guid}}}"


# %% ======================  CONTEO DE HOJAS (FILTRO)  =======================

def detectar_tipo(ruta):
    """Detecta el tipo REAL del archivo por sus bytes magicos (no por la
    extension, que en el SMV puede mentir: rotula 'XLS' pero entrega PDF).
    Devuelve 'pdf', 'xlsx', 'xls' o None."""
    try:
        with open(ruta, "rb") as fh:
            firma = fh.read(8)
    except OSError:
        return None

    if firma[:4] == b"%PDF":                       # PDF
        return "pdf"
    if firma[:4] == b"PK\x03\x04":                 # ZIP -> xlsx/docx/...
        return "xlsx"
    if firma[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":  # OLE2 -> xls antiguo
        return "xls"
    return None


def contar_hojas(ruta):
    """Devuelve el numero de 'hojas' del archivo, segun su tipo REAL:
        - PDF  -> numero de paginas.
        - XLSX -> numero de worksheets.
        - XLS  -> numero de worksheets (si esta xlrd instalado).
        - Otro / error -> None ('no filtrable' => se conserva).
    """
    tipo = detectar_tipo(ruta)
    try:
        if tipo == "pdf":
            from pypdf import PdfReader
            return len(PdfReader(ruta).pages)

        if tipo == "xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
            n = len(wb.sheetnames)
            wb.close()
            return n

        if tipo == "xls":
            try:
                import xlrd
                return xlrd.open_workbook(ruta).nsheets
            except ImportError:
                return None  # sin xlrd no se puede contar -> se conserva

    except Exception as e:
        print(f"      [aviso] no se pudo contar hojas de {os.path.basename(ruta)}: {e}")
        return None

    return None


def aplicar_filtro_hojas(ruta, carpeta_empresa):
    """Aplica el filtro de minimo de hojas. Devuelve True si el archivo se
    CONSERVA, False si fue descartado."""
    if FILTRAR_MIN_HOJAS is None:
        return True

    n = contar_hojas(ruta)
    if n is None:
        # No se pudo determinar -> conservar (no perder archivos por el filtro).
        print(f"      hojas: ? -> se conserva (no filtrable)")
        return True

    if n > FILTRAR_MIN_HOJAS:
        print(f"      hojas: {n} -> se conserva")
        return True

    # Descartar
    print(f"      hojas: {n} (<= {FILTRAR_MIN_HOJAS}) -> descartado")
    if ACCION_DESCARTE == "eliminar":
        os.remove(ruta)
    else:
        destino = os.path.join(carpeta_empresa, "_descartados")
        os.makedirs(destino, exist_ok=True)
        shutil.move(ruta, os.path.join(destino, os.path.basename(ruta)))
    return False


# %% ======================  CONFIGURACION DEL DRIVER  ======================

def iniciar_driver(carpeta_descargas):
    """Crea un Chrome WebDriver configurado para descargar automaticamente a
    'carpeta_descargas' (sin abrir el visor de PDF)."""
    os.makedirs(carpeta_descargas, exist_ok=True)

    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    prefs = {
        "download.default_directory": carpeta_descargas,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,   # descargar PDF en vez de abrirlo
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    driver = webdriver.Chrome(options=options)

    # Por si headless: habilitar descargas explicitamente.
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": carpeta_descargas,
        })
    except Exception:
        pass

    return driver


def fijar_carpeta_descargas(driver, carpeta):
    """Reapunta la carpeta de descargas del driver (una por empresa)."""
    os.makedirs(carpeta, exist_ok=True)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": carpeta,
        })
    except Exception:
        pass


# %% ======================  BUSQUEDA DE EMPRESA  ===========================

def buscar_empresa(driver, wait, nombre):
    """Flujo real de navegacion del SMV:
        1. Abre la pagina.
        2. Escribe la razon social en el buscador de emisor (txtSearch).
        3. Click en 'Buscar Empresa' (ibtnB -> buscar1()).
        4. Espera la pagina del emisor y hace click en el boton 'Memoria'
           (btnMemorias).
        5. Espera la grilla de memorias con datos.
    Devuelve True si la tabla de resultados quedo presente con datos.
    """
    driver.get(URL_SMV)

    # 1) Input de busqueda de emisor (txtSearch). No tiene modal asociado.
    campo = wait.until(EC.presence_of_element_located((By.ID, "txtSearch")))
    driver.execute_script("arguments[0].value = arguments[1];", campo, nombre)

    # 2) Click en el boton-imagen 'Buscar Empresa' (ibtnB -> buscar1()).
    boton = wait.until(EC.element_to_be_clickable((By.ID, "ibtnB")))
    driver.execute_script("arguments[0].click();", boton)

    # 3) Esperar a que cargue la pagina del emisor y aparezca el boton Memoria.
    try:
        btn_mem = WebDriverWait(driver, TIMEOUT_BUSQUEDA).until(
            EC.element_to_be_clickable((By.ID, "btnMemorias")))
    except TimeoutException:
        return False  # no se encontro la empresa / no cargo la ficha del emisor

    # 4) Ir a la seccion 'Memoria'.
    driver.execute_script("arguments[0].click();", btn_mem)

    # 5) Esperar a que la grilla de memorias cargue con datos.
    try:
        WebDriverWait(driver, TIMEOUT_BUSQUEDA).until(_grilla_tiene_datos)
        return True
    except TimeoutException:
        return False


def _grilla_tiene_datos(driver):
    """True si la tabla grdMemorias existe y tiene al menos una fila con celdas
    de datos (descarta cabecera y filas de paginacion)."""
    try:
        tabla = driver.find_element(By.ID, "MainContent_grdMemorias")
    except NoSuchElementException:
        return False
    filas = tabla.find_elements(By.XPATH, ".//tr")
    for f in filas:
        celdas = f.find_elements(By.XPATH, "./td")
        if len(celdas) >= 5:
            return True
    return False


# ---------------------------------------------------------------------------
# NOTA sobre la busqueda:
#   El flujo usa el buscador de emisor (txtSearch -> ibtnB/buscar1()) y luego
#   el boton 'Memoria' (btnMemorias). Si alguna empresa sale 'SIN RESULTADOS',
#   las causas tipicas son:
#     - La razon social no coincide exacto con la registrada en el SMV
#       (revisar/ajustar el nombre en la lista EMPRESAS).
#     - buscar1() abrio una lista de coincidencias en vez de la ficha directa
#       (en ese caso habria que seleccionar la fila correcta antes de btnMemorias).
#   Subir TIMEOUT_BUSQUEDA tambien ayuda si el portal va lento.
# ---------------------------------------------------------------------------


# %% ======================  LECTURA DE LA TABLA  ===========================

def leer_filas_coincidentes(driver):
    """Lee la grilla actual y devuelve la lista de coincidencias (patron + anio)
    como dicts: {anio, documento, url_descarga}. Robusto ante re-render."""
    coincidencias = []
    try:
        tabla = driver.find_element(By.ID, "MainContent_grdMemorias")
        filas = tabla.find_elements(By.XPATH, ".//tr")
    except (NoSuchElementException, StaleElementReferenceException):
        return coincidencias

    for fila in filas:
        try:
            celdas = fila.find_elements(By.XPATH, "./td")
            if len(celdas) < 5:
                continue  # cabecera o fila de paginacion

            texto_anio = _norm(celdas[0].text)
            texto_doc = _norm(celdas[1].text)

            m = re.search(r"(\d{4})", texto_anio)
            if not m:
                continue
            anio = int(m.group(1))

            if anio not in ANIOS_OBJETIVO:
                continue
            if _norm(PATRON_DOCUMENTO) not in texto_doc:
                continue

            # Enlace de descarga (col 5).
            enlace = celdas[4].find_element(By.XPATH, ".//a[@href]")
            url = limpiar_url_documento(enlace.get_attribute("href"))
            if not url:
                continue

            coincidencias.append({
                "anio": anio,
                "documento": texto_doc,
                "url_descarga": url,
            })
        except (StaleElementReferenceException, NoSuchElementException):
            continue

    return coincidencias


def anio_minimo_en_pagina(driver):
    """Menor anio presente en la grilla actual (para decidir si vale la pena
    pasar a la siguiente pagina; la tabla viene ordenada desc por anio)."""
    anios = []
    try:
        tabla = driver.find_element(By.ID, "MainContent_grdMemorias")
        for fila in tabla.find_elements(By.XPATH, ".//tr"):
            celdas = fila.find_elements(By.XPATH, "./td")
            if len(celdas) >= 5:
                m = re.search(r"(\d{4})", celdas[0].text)
                if m:
                    anios.append(int(m.group(1)))
    except (NoSuchElementException, StaleElementReferenceException):
        pass
    return min(anios) if anios else None


def ir_a_siguiente_pagina(driver, num_pagina):
    """Hace click en el enlace de paginacion 'Page$N'. Devuelve True si pudo
    avanzar (y la grilla se re-renderizo)."""
    try:
        tabla = driver.find_element(By.ID, "MainContent_grdMemorias")
        enlaces = tabla.find_elements(
            By.XPATH, f".//a[contains(@href, \"Page${num_pagina}\")]")
        if not enlaces:
            return False
        primera_celda = tabla.find_element(By.XPATH, ".//tr[td][1]/td[1]")
        enlaces[0].click()
        # Esperar a que la grilla cambie (la primera celda deja de ser la misma).
        WebDriverWait(driver, TIMEOUT_BUSQUEDA).until(
            EC.staleness_of(primera_celda))
        WebDriverWait(driver, TIMEOUT_BUSQUEDA).until(_grilla_tiene_datos)
        return True
    except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
        return False


# %% ======================  DESCARGA DE ARCHIVOS  ==========================

def _archivos_actuales(carpeta):
    return set(os.listdir(carpeta)) if os.path.isdir(carpeta) else set()


def esperar_descarga(carpeta, antes, timeout=TIMEOUT_DESCARGA):
    """Espera a que aparezca un archivo nuevo y termine de descargarse
    (sin .crdownload / .tmp). Devuelve la ruta del archivo nuevo o None."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        ahora = _archivos_actuales(carpeta)
        nuevos = [f for f in (ahora - antes)
                  if not f.endswith((".crdownload", ".tmp"))]
        # Asegurar que no quede ninguna descarga en curso de los nuevos.
        en_curso = [f for f in (ahora - antes)
                    if f.endswith((".crdownload", ".tmp"))]
        if nuevos and not en_curso:
            # tomar el mas reciente
            rutas = [os.path.join(carpeta, f) for f in nuevos]
            return max(rutas, key=os.path.getmtime)
        time.sleep(0.5)
    return None


def descargar_documento(driver, carpeta, url):
    """Abre la URL de descarga en una pestania nueva, espera el archivo y
    cierra la pestania. Devuelve la ruta del archivo descargado o None."""
    antes = _archivos_actuales(carpeta)
    handles_antes = set(driver.window_handles)

    driver.execute_script("window.open(arguments[0], '_blank');", url)

    ruta = esperar_descarga(carpeta, antes)

    # Cerrar cualquier pestania nueva abierta por la descarga.
    handles_nuevos = set(driver.window_handles) - handles_antes
    for h in handles_nuevos:
        try:
            driver.switch_to.window(h)
            driver.close()
        except WebDriverException:
            pass
    driver.switch_to.window(list(handles_antes)[0])

    return ruta


def renombrar(ruta, nombre_empresa, anio, idx):
    """Renombra el archivo descargado a un nombre descriptivo y estable.
    Usa la extension REAL (bytes magicos): si el SMV rotula 'XLS' pero entrega
    un PDF, el archivo se guarda como .pdf."""
    tipo_real = detectar_tipo(ruta)          # 'pdf' / 'xlsx' / 'xls' / None
    if tipo_real:
        ext = "." + tipo_real
    else:
        ext = os.path.splitext(ruta)[1].lower() or ".bin"
    base = f"{_nombre_seguro(nombre_empresa)}_{anio}_SECCIONES_B_Y_C_{idx}{ext}"
    destino = os.path.join(os.path.dirname(ruta), base)
    if os.path.abspath(destino) != os.path.abspath(ruta):
        if os.path.exists(destino):
            os.remove(destino)
        os.rename(ruta, destino)
    return destino


# %% ======================  PROCESO POR EMPRESA  ===========================

def procesar_empresa(driver, nombre_busqueda, nombre_guardado):
    """Ejecuta todo el flujo para una empresa. Devuelve un dict de resumen."""
    print(f"\n=== {nombre_guardado}  (buscando: '{nombre_busqueda}') ===")
    carpeta_empresa = os.path.join(RUTA_BASE, _nombre_seguro(nombre_guardado))
    os.makedirs(carpeta_empresa, exist_ok=True)
    fijar_carpeta_descargas(driver, carpeta_empresa)

    wait = WebDriverWait(driver, TIMEOUT_BUSQUEDA)

    resumen = {"empresa": nombre_guardado, "descargados": 0,
               "descartados": 0, "errores": 0, "ok_busqueda": False}

    if not buscar_empresa(driver, wait, nombre_busqueda):
        print("   [!] La busqueda no devolvio resultados (revisar nombre o modal).")
        return resumen
    resumen["ok_busqueda"] = True

    # Recolectar coincidencias recorriendo paginas mientras tenga sentido
    # (la grilla viene ordenada por anio desc; paramos cuando el anio minimo
    #  de la pagina ya es menor al anio objetivo mas antiguo).
    min_objetivo = min(ANIOS_OBJETIVO)
    coincidencias = []
    pagina = 1
    while True:
        coincidencias.extend(leer_filas_coincidentes(driver))
        amin = anio_minimo_en_pagina(driver)
        if amin is None or amin <= min_objetivo:
            break  # ya no habra mas anios objetivo en paginas siguientes
        pagina += 1
        if not ir_a_siguiente_pagina(driver, pagina):
            break

    # Quitar duplicados por URL (a veces hay expedientes repetidos del mismo anio).
    vistos = set()
    unicas = []
    for c in coincidencias:
        if c["url_descarga"] not in vistos:
            vistos.add(c["url_descarga"])
            unicas.append(c)

    print(f"   Coincidencias '{PATRON_DOCUMENTO}' en {sorted(ANIOS_OBJETIVO)}: "
          f"{len(unicas)}")

    for idx, c in enumerate(unicas, start=1):
        print(f"   - {c['anio']} | {c['documento'][:60]}")
        try:
            ruta = descargar_documento(driver, carpeta_empresa, c["url_descarga"])
            if not ruta:
                print("      [!] no se detecto el archivo descargado (timeout).")
                resumen["errores"] += 1
                continue
            ruta = renombrar(ruta, nombre_guardado, c["anio"], idx)
            if aplicar_filtro_hojas(ruta, carpeta_empresa):
                resumen["descargados"] += 1
            else:
                resumen["descartados"] += 1
        except Exception as e:
            print(f"      [!] error al descargar: {e}")
            resumen["errores"] += 1

    return resumen


# %% ======================  EJECUCION PRINCIPAL  ===========================

def main():
    os.makedirs(RUTA_BASE, exist_ok=True)
    driver = iniciar_driver(RUTA_BASE)
    resultados = []
    try:
        for nombre_busqueda, nombre_guardado in EMPRESAS:
            try:
                resultados.append(
                    procesar_empresa(driver, nombre_busqueda, nombre_guardado))
            except Exception as e:
                print(f"   [!!] Fallo general en {nombre_guardado}: {e}")
                resultados.append({"empresa": nombre_guardado, "descargados": 0,
                                   "descartados": 0, "errores": 1,
                                   "ok_busqueda": False})
    finally:
        driver.quit()

    # Resumen final
    print("\n" + "=" * 70)
    print("RESUMEN")
    print("=" * 70)
    tot_d = tot_x = tot_e = 0
    for r in resultados:
        estado = "OK" if r["ok_busqueda"] else "SIN RESULTADOS"
        print(f"{r['empresa'][:45]:45s} | desc:{r['descargados']:2d} "
              f"descart:{r['descartados']:2d} err:{r['errores']:2d} | {estado}")
        tot_d += r["descargados"]; tot_x += r["descartados"]; tot_e += r["errores"]
    print("-" * 70)
    print(f"TOTALES -> descargados: {tot_d} | descartados: {tot_x} | errores: {tot_e}")
    print(f"Carpeta base: {RUTA_BASE}")


# %% ======================  PUNTO DE ENTRADA  ==============================
if __name__ == "__main__":
    main()