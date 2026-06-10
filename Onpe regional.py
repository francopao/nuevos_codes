

"""
ONPE Segunda Vuelta 2026 — Scraper DOM v4
Selectores confirmados:
  - S_VISTA:    mat-select[formcontrolname='region']
  - S_DEPTO:    mat-select[formcontrolname='department']
  - S_PROV:     mat-select[formcontrolname='province']
  - Tarjetas:   .tarjeta-candidato--izquierda / --derecha
  - Actas:      ul.leyenda.vertical li

v4 fixes:
  1. InvalidSessionIdException → reinicio automático del browser si Chrome
     muere a mitad del BLOQUE 2. El scraper reanuda desde la última provincia
     fallida sin perder las ya recolectadas.
  2. Corrección de nombres de provincia con tildes/grafías exactas de la web
     (confirmadas por el log de v3):
       HUÁNUCO (prov) → HUÁNUCO, HUAMALÍES → HUAMALIES(?), etc.
     Se usa un segundo intento normalizando tildes antes de rendirse.
  3. elegir() hace fallback normalizado: si el texto exacto falla, prueba
     sin tildes y con variantes comunes antes de rendirse.
"""

import time, re, logging, unicodedata
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    StaleElementReferenceException, ElementClickInterceptedException,
    InvalidSessionIdException, WebDriverException,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger("ONPE")

URL        = "https://resultadosegundavuelta.onpe.gob.pe/main/resumen"
S_VISTA    = "mat-select[formcontrolname='region']"
S_DEPTO    = "mat-select[formcontrolname='department']"
S_PROV     = "mat-select[formcontrolname='province']"
S_LIMPIAR  = "//button[contains(normalize-space(.),'LIMPIAR')]"
S_BACKDROP = "div.cdk-overlay-backdrop"

# ── Regiones exactamente como la página ONPE las muestra ──────────
REGIONES = [
    "AMAZONAS", "ÁNCASH", "APURÍMAC", "AREQUIPA", "AYACUCHO",
    "CAJAMARCA", "CALLAO", "CUSCO", "HUANCAVELICA", "HUÁNUCO",
    "ICA", "JUNÍN", "LA LIBERTAD", "LAMBAYEQUE", "LIMA",
    "LORETO", "MADRE DE DIOS", "MOQUEGUA", "PASCO", "PIURA",
    "PUNO", "SAN MARTÍN", "TACNA", "TUMBES", "UCAYALI"
]

# ── Equivalencias: nombre CSV/base-datos → nombre ONPE ────────────
EQUIV_CSV_A_ONPE = {
    "ANCASH":       "ÁNCASH",
    "APURIMAC":     "APURÍMAC",
    "HUANUCO":      "HUÁNUCO",
    "JUNIN":        "JUNÍN",
    "SAN MARTIN":   "SAN MARTÍN",
    "LIMA REGION":  "LIMA",
    "LIMA PROVINCIA": "LIMA",
    # provincias
    "BONGARA":           "BONGARÁ",
    "RODRIGUEZ DE MENDOZA": "RODRÍGUEZ DE MENDOZA",
    "CARLOS FERMIN FITZCARRALD": "CARLOS FERMÍN FITZCARRAL",
    "ASUNCION":          "ASUNCIÓN",
    "CARAVELI":          "CARAVELÍ",
    "CELENDIN":          "CELENDÍN",
    "JAEN":              "JAÉN",
    "LA CONVENCION":     "LA CONVENCIÓN",
    "SANCHEZ CARRION":   "SÁNCHEZ CARRIÓN",
    "CHEPEN":            "CHEPÉN",
    "JULCAN":            "JULCÁN",
    "GRAN CHIMU":        "GRAN CHIMÚ",
    "MARISCAL RAMON CASTILLA": "MARISCAL RAMÓN CASTILLA",
    "GENERAL SANCHEZ CERRO":   "GENERAL SÁNCHEZ CERRO",
    "DANIEL ALCIDES CARRION":  "DANIEL ALCIDES CARRIÓN",
    "AZANGARO":    "AZÁNGARO",
    "SAN ROMAN":   "SAN ROMÁN",
    "MARISCAL CACERES": "MARISCAL CÁCERES",
    "PURUS":       "PURÚS",
}

# ── Provincias por región  ─────────────────────────────────────────
# Nombres corregidos con lo aprendido del log v3:
#   - HUÁNUCO provincia: la web lo muestra con tilde → "HUÁNUCO"
#   - HUAMALÍES: la web NO tiene tilde → "HUAMALIES"  (WARNING v3)
#   - MARAÑON: la web usa → "MARAÑÓN" (con tilde en Ó)
#   - LEONCIO PRADO: la web usa → "LEONCIO PRADO" (igual — era crash de sesión)
#   - CONCEPCION → "CONCEPCIÓN"
#   - VICTOR FAJARDO → "VÍCTOR FAJARDO"
#   - VILCAS HUAMAN → "VILCAS HUAMÁN"
#   - PAUCAR DEL SARA SARA → "PÁUCAR DEL SARA SARA"
#   - LA UNION → "LA UNIÓN"
#   - ANTONIO RAYMONDI → "ANTONIO RAIMONDI" (nombre real ONPE usa Raimondi)
#   - NAZCA/PISCO: eran crashes de sesión, no nombres wrongos; se mantienen
PROVINCIAS_POR_REGION = {
    "AMAZONAS": [
        "CHACHAPOYAS", "BAGUA", "BONGARÁ", "LUYA",
        "RODRÍGUEZ DE MENDOZA", "CONDORCANQUI", "UTCUBAMBA",
    ],
    "ÁNCASH": [
        "HUARAZ", "AIJA", "BOLOGNESI", "CARHUAZ", "CASMA", "CORONGO",
        "HUAYLAS", "HUARI", "MARISCAL LUZURIAGA", "PALLASCA", "POMABAMBA",
        "RECUAY", "SANTA", "SIHUAS", "YUNGAY",
        "ANTONIO RAIMONDI",          # ← RAIMONDI no RAYMONDI
        "CARLOS FERMÍN FITZCARRAL",
        "ASUNCIÓN", "HUARMEY", "OCROS",
    ],
    "APURÍMAC": [
        "ABANCAY", "AYMARAES", "ANDAHUAYLAS", "ANTABAMBA",
        "COTABAMBAS", "GRAU", "CHINCHEROS",
    ],
    "AREQUIPA": [
        "AREQUIPA", "CAYLLOMA", "CAMANÁ", "CARAVELÍ",
        "CASTILLA", "CONDESUYOS", "ISLAY",
        "LA UNIÓN",                  # ← con tilde
    ],
    "AYACUCHO": [
        "HUAMANGA", "CANGALLO", "HUANTA", "LA MAR", "LUCANAS",
        "PARINACOCHAS",
        "VÍCTOR FAJARDO",            # ← con tilde
        "HUANCA SANCOS",
        "VILCAS HUAMÁN",             # ← con tilde
        "PÁUCAR DEL SARA SARA",      # ← con tilde inicial
        "SUCRE",
    ],
    "CAJAMARCA": [
        "CAJAMARCA", "CAJABAMBA", "CELENDÍN", "CONTUMAZA", "CUTERVO",
        "CHOTA", "HUALGAYOC", "JAÉN", "SANTA CRUZ",
        "SAN MIGUEL", "SAN IGNACIO", "SAN MARCOS", "SAN PABLO",
    ],
    "CALLAO": [
        "CALLAO",
    ],
    "CUSCO": [
        "CUSCO", "ACOMAYO", "ANTA", "CALCA", "CANAS", "CANCHIS",
        "CHUMBIVILCAS", "ESPINAR", "LA CONVENCIÓN", "PARURO",
        "PAUCARTAMBO", "QUISPICANCHI", "URUBAMBA",
    ],
    "HUANCAVELICA": [
        "HUANCAVELICA", "ACOBAMBA", "ANGARAES", "CASTROVIRREYNA",
        "TAYACAJA", "HUAYTARA", "CHURCAMPA",
    ],
    "HUÁNUCO": [
        "HUÁNUCO",                   # ← con tilde (igual que la región)
        "AMBO", "DOS DE MAYO",
        "HUAMALIES",                 # ← SIN tilde (como lo muestra la web)
        "MARAÑÓN",                   # ← con tilde en Ó
        "LEONCIO PRADO",
        "PACHITEA", "PUERTO INCA", "HUACAYBAMBA", "LAURICOCHA", "YAROWILCA",
    ],
    "ICA": [
        "ICA", "CHINCHA", "NASCA", "PISCO", "PALPA",  # ← NASCA no NAZCA
    ],
    "JUNÍN": [
        "HUANCAYO",
        "CONCEPCIÓN",                # ← con tilde
        "JAUJA", "JUNÍN", "TARMA",
        "YAULI", "SATIPO", "CHANCHAMAYO", "CHUPACA",
    ],
    "LA LIBERTAD": [
        "TRUJILLO", "BOLÍVAR", "SÁNCHEZ CARRIÓN", "OTUZCO",
        "PACASMAYO", "PATAZ", "SANTIAGO DE CHUCO", "ASCOPE",
        "CHEPÉN", "JULCÁN", "GRAN CHIMÚ", "VIRÚ",
    ],
    "LAMBAYEQUE": [
        "CHICLAYO", "FERREÑAFE", "LAMBAYEQUE",
    ],
    "LIMA": [
        "LIMA",
        "BARRANCA", "CAJATAMBO", "CANTA", "CAÑETE", "HUARAL",
        "HUAROCHIRÍ", "HUAURA", "OYÓN", "YAUYOS",
    ],
    "LORETO": [
        "MAYNAS", "ALTO AMAZONAS", "LORETO", "REQUENA", "UCAYALI",
        "MARISCAL RAMÓN CASTILLA", "DATEM DEL MARAÑÓN", "PUTUMAYO",
    ],
    "MADRE DE DIOS": [
        "TAMBOPATA", "MANU", "TAHUAMANU",
    ],
    "MOQUEGUA": [
        "MARISCAL NIETO", "GENERAL SÁNCHEZ CERRO", "ILO",
    ],
    "PASCO": [
        "PASCO", "DANIEL ALCIDES CARRIÓN", "OXAPAMPA",
    ],
    "PIURA": [
        "PIURA", "AYABACA", "HUANCABAMBA", "MORROPÓN",
        "PAITA", "SULLANA", "TALARA", "SECHURA",
    ],
    "PUNO": [
        "PUNO", "AZÁNGARO", "CARABAYA", "CHUCUITO", "HUANCANÉ",
        "LAMPA", "MELGAR", "SANDIA", "SAN ROMÁN", "YUNGUYO",
        "SAN ANTONIO DE PUTINA", "EL COLLAO", "MOHO",
    ],
    "SAN MARTÍN": [
        "MOYOBAMBA", "HUALLAGA", "LAMAS", "MARISCAL CÁCERES",
        "RIOJA", "SAN MARTÍN", "BELLAVISTA", "TOCACHE",
        "PICOTA", "EL DORADO",
    ],
    "TACNA": [
        "TACNA", "TARATA", "JORGE BASADRE", "CANDARAVE",
    ],
    "TUMBES": [
        "TUMBES", "CONTRALMIRANTE VILLAR", "ZARUMILLA",
    ],
    "UCAYALI": [
        "CORONEL PORTILLO", "PADRE ABAD", "ATALAYA", "PURÚS",
    ],
}


# ══════════════════════════════════════════════════════════════════
# UTILIDADES TEXTO
# ══════════════════════════════════════════════════════════════════

def _sin_tildes(s):
    """'HUÁNUCO' → 'HUANUCO'"""
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )

def _variantes(texto):
    """
    Genera lista de variantes a probar en el panel, de mayor a menor
    especificidad.  Esto absorbe diferencias de tilde entre nuestra
    lista y lo que la web renderiza.
    """
    v = [texto]
    sin = _sin_tildes(texto)
    if sin != texto:
        v.append(sin)
    # Capitalizado (por si la web usa Title Case)
    v.append(texto.title())
    v.append(sin.title())
    return list(dict.fromkeys(v))   # dedup preservando orden


# ══════════════════════════════════════════════════════════════════
# DRIVER
# ══════════════════════════════════════════════════════════════════

def make_driver(headless=True):
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1440,900")
    opts.add_argument("--lang=es-PE")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    return webdriver.Chrome(options=opts)

def driver_vivo(driver):
    """Devuelve True si la sesión de Chrome sigue activa."""
    try:
        _ = driver.current_url
        return True
    except (InvalidSessionIdException, WebDriverException):
        return False


# ══════════════════════════════════════════════════════════════════
# ESPERAS Y CLICKS
# ══════════════════════════════════════════════════════════════════

def wait_for(driver, css, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css)))
        return True
    except TimeoutException:
        return False

def wait_gone(driver, css, timeout=5):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, css)))
    except TimeoutException:
        pass

def esperar_sin_backdrop(driver, timeout=6):
    wait_gone(driver, S_BACKDROP, timeout)
    time.sleep(0.1)   # ← 0.2→0.1

def click_robusto(driver, element):
    try:
        element.click()
        return True
    except ElementClickInterceptedException:
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            return False
    except Exception:
        return False

def abrir_select(driver, css, timeout=12):
    esperar_sin_backdrop(driver)
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)   # ← 0.3→0.15
        if not click_robusto(driver, el):
            return False
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "mat-option")))
        time.sleep(0.15)   # ← 0.3→0.15
        return True
    except (TimeoutException, Exception):
        return False

def elegir(driver, texto, timeout=8):
    """
    Intenta elegir 'texto' en el panel abierto.
    Prueba múltiples variantes (con/sin tildes) antes de rendirse.
    En la primera variante usa timeout completo; las siguientes usan 3s
    (el panel ya está abierto, solo buscamos el elemento).
    """
    for idx, variante in enumerate(_variantes(texto)):
        t = timeout if idx == 0 else 3
        for xpath in [
            f"//mat-option[normalize-space(.)='{variante}']",
            f"//mat-option[contains(normalize-space(.),'{variante}')]",
        ]:
            try:
                opt = WebDriverWait(driver, t).until(
                    EC.presence_of_element_located((By.XPATH, xpath)))
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center'});", opt)
                time.sleep(0.1)   # ← 0.2→0.1
                if click_robusto(driver, opt):
                    if variante != texto:
                        log.debug(f"  elegir: '{texto}' encontrado como '{variante}'")
                    esperar_sin_backdrop(driver, timeout=5)
                    time.sleep(1.5)   # ← 3.0→1.5  Angular re-render
                    return True
            except TimeoutException:
                continue
            except Exception:
                continue
    log.warning(f"  Opción no encontrada: '{texto}'")
    return False

def limpiar(driver):
    esperar_sin_backdrop(driver)
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, S_LIMPIAR)))
        click_robusto(driver, btn)
        esperar_sin_backdrop(driver)
        time.sleep(1.0)   # ← 2.0→1.0
    except Exception:
        pass

def seleccionar_vista(driver, vista):
    els = driver.find_elements(By.CSS_SELECTOR, S_VISTA)
    if els and vista.upper() in els[0].text.upper():
        return True
    if abrir_select(driver, S_VISTA):
        return elegir(driver, vista)
    return False

def provincia_select_disponible(driver, timeout=8):
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, S_PROV)))
        return el.get_attribute("aria-disabled") != "true"
    except TimeoutException:
        return False


# ══════════════════════════════════════════════════════════════════
# LECTURA DOM
# ══════════════════════════════════════════════════════════════════

def leer_actas(driver):
    """
    Parsea: <li><div class="color validados"></div> Contabilizadas <b>(89,385)</b></li>
    Lee el número desde el tag <b> para evitar que li.text mezcle etiquetas.
    El número puede tener comas o puntos de miles → se limpian antes de int().
    """
    meta = {"actas_contabilizadas": None, "actas_jee": None,
            "actas_pendientes": None, "actas_total": None,
            "pct_contabilizadas": None}
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "ul.leyenda.vertical li")
        for li in items:
            # Etiqueta de texto del li (sin el <b>): "Contabilizadas", "Para envío al JEE", "Pendientes"
            li_text = li.text.strip().lower()
            # Número está en el <b> como "(89,385)" o "(1,543)"
            try:
                b_text = li.find_element(By.TAG_NAME, "b").text.strip()
                # Quitar paréntesis y separadores de miles
                num_str = re.sub(r"[^\d]", "", b_text)
                n = int(num_str) if num_str else 0
            except Exception:
                # fallback: buscar número en el texto completo del li
                m = re.search(r"\([\d,\.]+\)", li_text)
                num_str = re.sub(r"[^\d]", "", m.group(0)) if m else ""
                n = int(num_str) if num_str else 0

            if   "contabilizad" in li_text: meta["actas_contabilizadas"] = n
            elif "jee" in li_text or "envío" in li_text or "envio" in li_text:
                meta["actas_jee"] = n
            elif "pendiente" in li_text:    meta["actas_pendientes"] = n

        vals = [v for v in [meta["actas_contabilizadas"],
                             meta["actas_jee"],
                             meta["actas_pendientes"]] if v is not None]
        if vals:
            meta["actas_total"] = sum(vals)
            meta["pct_contabilizadas"] = round(
                100 * (meta["actas_contabilizadas"] or 0) / meta["actas_total"], 3
            ) if meta["actas_total"] else None
    except Exception:
        pass
    return meta

def _txt(parent, css):
    try:    return parent.find_element(By.CSS_SELECTOR, css).text.strip()
    except: return ""

def _pct(txt):
    c = re.sub(r"[^\d.,]", "", txt).replace(",", ".")
    try:    return float(c)
    except: return 0.0

def _votos(txt):
    limpio = txt.replace("'","").replace(".","").replace(",","")
    nums = re.findall(r"\d+", limpio)
    return int("".join(nums)) if nums else 0

def leer_tarjetas(driver, ubigeo, nivel, extra=None):
    rows  = []
    ts    = datetime.now().isoformat()
    extra = extra or {}
    actas = leer_actas(driver)
    for css in [".tarjeta-candidato--izquierda", ".tarjeta-candidato--derecha"]:
        try:
            card    = driver.find_element(By.CSS_SELECTOR, css)
            nombre  = _txt(card, ".tarjeta-candidato__nombre")
            pct_txt = _txt(card, ".tarjeta-candidato__porcentaje")
            partido = _txt(card, ".tarjeta-candidato__organizacion")
            vt      = (_txt(card, ".tarjeta-candidato__votos.d-none-movil")
                    or _txt(card, ".tarjeta-candidato__votos"))
            if not nombre:
                continue
            rows.append({
                "nivel":      nivel,
                "ubigeo":     ubigeo,
                "timestamp":  ts,
                "candidato":  nombre,
                "partido":    partido,
                "votos":      _votos(vt),
                "porcentaje": _pct(pct_txt),
                **actas,
                **extra,
            })
        except Exception:
            continue
    return rows


# ══════════════════════════════════════════════════════════════════
# REINICIO DE SESIÓN
# ══════════════════════════════════════════════════════════════════

def reiniciar_driver(driver, headless):
    """
    Cierra el driver actual (si puede) y abre uno nuevo ya en la URL de ONPE.
    Retorna el nuevo driver o None si falla.
    """
    try:
        driver.quit()
    except Exception:
        pass
    log.warning("  Reiniciando Chrome...")
    try:
        nuevo = make_driver(headless)
        nuevo.get(URL)
        cargado = False
        for css, t in [(".tarjeta-candidato", 60), ("mat-select", 60)]:
            if wait_for(nuevo, css, t):
                cargado = True
                break
        if not cargado:
            nuevo.quit()
            return None
        time.sleep(2)
        log.warning("  Chrome reiniciado ✓")
        return nuevo
    except Exception as e:
        log.error(f"  Fallo reinicio: {e}")
        return None


# ══════════════════════════════════════════════════════════════════
# SCRAPING PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def scrape(headless=True):
    log.info(f"\n{'═'*60}")
    log.info(f"ONPE — Regional + Provincial + Extranjero [{datetime.now().strftime('%H:%M:%S')}]")
    log.info(f"{'═'*60}")

    driver = make_driver(headless)
    rows_regional   = []
    rows_provincial = []
    rows_extranjero = []

    try:
        driver.get(URL)
        log.info("Esperando carga Angular...")
        cargado = False
        for css, t in [(".tarjeta-candidato", 60), ("mat-select", 60), ("app-root", 90)]:
            if wait_for(driver, css, t):
                log.info(f"  SPA lista [{css}]")
                cargado = True
                break
        if not cargado:
            log.error("SPA no cargó — verifica conexión")
            return _empty_result()
        time.sleep(2)

        # ── BLOQUE 1: 25 REGIONES ──────────────────────────────────
        log.info(f"\n{'─'*50}")
        log.info("BLOQUE 1: PERÚ → 25 Regiones")
        log.info(f"{'─'*50}")

        seleccionar_vista(driver, "PERÚ")
        if not wait_for(driver, S_DEPTO, timeout=15):
            log.error(f"  {S_DEPTO} no disponible")
        else:
            log.info(f"  {S_DEPTO} disponible ✓")

        for i, region in enumerate(REGIONES, 1):
            log.info(f"  [{i:02d}/25] {region}...")
            limpiar(driver)
            seleccionar_vista(driver, "PERÚ")
            if not wait_for(driver, S_DEPTO, timeout=12):
                log.warning(f"    {S_DEPTO} no disponible para {region}")
                continue
            if not abrir_select(driver, S_DEPTO):
                log.warning(f"    No se pudo abrir REGIÓN para {region}")
                continue
            if not elegir(driver, region):
                continue
            rows = leer_tarjetas(driver, region, "regional",
                                  {"departamento": region})
            if rows:
                rows_regional.extend(rows)
                for r in rows:
                    log.info(f"    {r['candidato'][:38]:38s} "
                             f"| {r['porcentaje']:6.3f}% "
                             f"| {r['votos']:>12,} votos")
                log.info(f"    Actas → contab={rows[0]['actas_contabilizadas']} "
                         f"JEE={rows[0]['actas_jee']} "
                         f"pend={rows[0]['actas_pendientes']} "
                         f"({rows[0]['pct_contabilizadas']}%)")
            else:
                log.warning(f"    Sin datos DOM para {region}")

        # ── BLOQUE 2: PROVINCIAS ───────────────────────────────────
        log.info(f"\n{'─'*50}")
        total_provs = sum(len(v) for v in PROVINCIAS_POR_REGION.values())
        log.info(f"BLOQUE 2: PERÚ → Región → Provincia ({total_provs} provincias)")
        log.info(f"{'─'*50}")

        prov_counter       = 0
        MAX_REINTENTOS_DRIVER = 3
        # ── Registro de fallos para reporte final ──────────────────
        # Estructura: {"razon": str, "region": str, "provincia": str}
        provincias_fallidas = []

        for region, provincias in PROVINCIAS_POR_REGION.items():
            log.info(f"\n  ── {region} ({len(provincias)} provincias) ──")

            for provincia in provincias:
                prov_counter += 1
                log.info(f"  [{prov_counter:03d}/{total_provs}] {region} → {provincia}...")

                # ── Comprobar sesión; reiniciar si murió ──
                reintentos_driver = 0
                while not driver_vivo(driver):
                    reintentos_driver += 1
                    if reintentos_driver > MAX_REINTENTOS_DRIVER:
                        log.error("Chrome no pudo reiniciarse — abortando bloque 2")
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "driver_muerto_sin_reinicio"})
                        raise RuntimeError("driver_muerto")
                    nuevo = reiniciar_driver(driver, headless)
                    if nuevo:
                        driver = nuevo
                    else:
                        time.sleep(5)

                # ── Flujo normal; cada rama de fallo registra la razón ──
                try:
                    limpiar(driver)
                    seleccionar_vista(driver, "PERÚ")

                    if not wait_for(driver, S_DEPTO, timeout=12):
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "S_DEPTO_no_disponible"})
                        continue

                    if not abrir_select(driver, S_DEPTO):
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "no_pudo_abrir_select_region"})
                        continue

                    if not elegir(driver, region):
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": f"region_no_encontrada:{region}"})
                        continue

                    if not provincia_select_disponible(driver, timeout=8):
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "select_provincia_no_disponible"})
                        continue

                    if not abrir_select(driver, S_PROV):
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "no_pudo_abrir_select_provincia"})
                        continue

                    if not elegir(driver, provincia):
                        # elegir() ya hizo WARNING; registramos para el reporte
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "provincia_no_encontrada_en_panel"})
                        continue

                    rows = leer_tarjetas(
                        driver,
                        ubigeo=f"{region}|{provincia}",
                        nivel="provincial",
                        extra={"departamento": region, "provincia": provincia},
                    )
                    if rows:
                        rows_provincial.extend(rows)
                        for r in rows:
                            log.info(f"    {r['candidato'][:35]:35s} "
                                     f"| {r['porcentaje']:6.3f}% "
                                     f"| {r['votos']:>10,} votos")
                        log.info(f"    Actas → contab={rows[0]['actas_contabilizadas']} "
                                 f"JEE={rows[0]['actas_jee']} "
                                 f"pend={rows[0]['actas_pendientes']} "
                                 f"({rows[0]['pct_contabilizadas']}%)")
                    else:
                        provincias_fallidas.append({
                            "region": region, "provincia": provincia,
                            "razon": "sin_datos_DOM"})
                        log.warning(f"    Sin datos DOM para {provincia}")

                except (InvalidSessionIdException, WebDriverException) as e:
                    provincias_fallidas.append({
                        "region": region, "provincia": provincia,
                        "razon": f"excepcion:{e.__class__.__name__}"})
                    log.warning(f"    Sesión Chrome perdida en {provincia}: {e.__class__.__name__}")
                    try: driver.quit()
                    except Exception: pass
                    # driver placeholder muerto → el while lo detectará
                    try:
                        driver = make_driver(headless)
                    except Exception:
                        pass
                    continue

        # ── BLOQUE 3: EXTRANJERO ───────────────────────────────────
        log.info(f"\n{'─'*50}")
        log.info("BLOQUE 3: EXTRANJERO")
        log.info(f"{'─'*50}")

        if not driver_vivo(driver):
            driver = reiniciar_driver(driver, headless) or make_driver(headless)
            driver.get(URL)
            wait_for(driver, ".tarjeta-candidato", 60)
            time.sleep(2)

        limpiar(driver)
        seleccionar_vista(driver, "EXTRANJERO")
        time.sleep(1.5)   # ← era 2.5

        rows = leer_tarjetas(driver, "EXTRANJERO", "extranjero")
        if rows:
            rows_extranjero.extend(rows)
            for r in rows:
                log.info(f"  {r['candidato'][:38]:38s} "
                         f"| {r['porcentaje']:6.3f}% "
                         f"| {r['votos']:>12,} votos")
        else:
            log.warning("  Sin datos DOM para EXTRANJERO")

    except RuntimeError as e:
        log.error(f"Scraping interrumpido: {e}")
    finally:
        try: driver.quit()
        except Exception: pass
        log.info("\nBrowser cerrado.")

    df_regional   = pd.DataFrame(rows_regional)
    df_provincial = pd.DataFrame(rows_provincial)
    df_extranjero = pd.DataFrame(rows_extranjero)
    partes = [df for df in [df_regional, df_provincial, df_extranjero] if not df.empty]
    df_consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

    # ── REPORTE FINAL DE PROVINCIAS FALLIDAS ──────────────────────
    log.info(f"\n{'═'*60}  RESUMEN:")
    log.info(f"  df_regional   : {len(df_regional):4d} filas | "
             f"{df_regional['ubigeo'].nunique() if not df_regional.empty else 0} regiones")
    log.info(f"  df_provincial : {len(df_provincial):4d} filas | "
             f"{df_provincial['ubigeo'].nunique() if not df_provincial.empty else 0} provincias OK")
    log.info(f"  df_extranjero : {len(df_extranjero):4d} filas")
    log.info(f"  df_consolidado: {len(df_consolidado):4d} filas")
    if provincias_fallidas:
        log.warning(f"\n  ⚠  PROVINCIAS SIN DATOS ({len(provincias_fallidas)}):")
        # Agrupar por razón para un reporte compacto
        from collections import defaultdict
        por_razon = defaultdict(list)
        for f in provincias_fallidas:
            por_razon[f["razon"]].append(f"{f['region']}|{f['provincia']}")
        for razon, lista in sorted(por_razon.items()):
            log.warning(f"    [{razon}]")
            for item in lista:
                log.warning(f"      · {item}")
    else:
        log.info("  ✓  Todas las provincias procesadas sin fallos.")
    log.info(f"{'═'*60}")

    return {
        "regional":         df_regional,
        "provincial":       df_provincial,
        "extranjero":       df_extranjero,
        "consolidado":      df_consolidado,
        "provincias_fallidas": provincias_fallidas,   # lista de dicts para análisis posterior
    }

def _empty_result():
    return {"regional": pd.DataFrame(), "provincial": pd.DataFrame(),
            "extranjero": pd.DataFrame(), "consolidado": pd.DataFrame(),
            "provincias_fallidas": []}


# ══════════════════════════════════════════════════════════════════
# POLLING
# ══════════════════════════════════════════════════════════════════

def polling(headless=True, intervalo_seg=900, max_iter=None):
    hist = {"regional": [], "provincial": [], "extranjero": []}
    it = 0
    log.info(f"Polling — intervalo={intervalo_seg}s")
    try:
        while True:
            if max_iter and it >= max_iter:
                break
            log.info(f"\n══ Snapshot #{it+1}  {datetime.now().strftime('%H:%M:%S')} ══")
            try:
                res = scrape(headless=headless)
                for k in ["regional", "provincial", "extranjero"]:
                    df = res.get(k, pd.DataFrame())
                    if not df.empty:
                        df = df.copy()
                        df["snapshot"]    = it + 1
                        df["snapshot_ts"] = datetime.now().isoformat()
                        hist[k].append(df)
                conso = res.get("consolidado", pd.DataFrame())
                if not conso.empty:
                    print(f"\n=== Snapshot #{it+1} ===")
                    print(conso.to_string(index=False))
            except Exception as e:
                log.error(f"Error snapshot #{it+1}: {e}")
            it += 1
            if not (max_iter and it >= max_iter):
                log.info(f"Próximo en {intervalo_seg}s... (Ctrl+C)")
                time.sleep(intervalo_seg)
    except KeyboardInterrupt:
        log.info("Polling detenido.")
    return {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
            for k, v in hist.items()}


# ══════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":

    HEADLESS      = False
    MODO_POLLING  = False
    INTERVALO_SEG = 900
    MAX_ITER      = None

    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    pd.set_option("display.max_rows", 200)
    pd.set_option("display.float_format", "{:,.3f}".format)

    if MODO_POLLING:
        hist = polling(headless=HEADLESS, intervalo_seg=INTERVALO_SEG,
                       max_iter=MAX_ITER)
        df_regional    = hist["regional"]
        df_provincial  = hist["provincial"]
        df_extranjero  = hist["extranjero"]
        partes = [df for df in [df_regional, df_provincial, df_extranjero] if not df.empty]
        df_consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    else:
        res                = scrape(headless=HEADLESS)
        df_regional        = res["regional"]
        df_provincial      = res["provincial"]
        df_extranjero      = res["extranjero"]
        df_consolidado     = res["consolidado"]
        provincias_fallidas = res["provincias_fallidas"]

    for nombre, df in [("REGIONAL",    df_regional),
                       ("PROVINCIAL",  df_provincial),
                       ("EXTRANJERO",  df_extranjero),
                       ("CONSOLIDADO", df_consolidado)]:
        print(f"\n{'═'*70}")
        print(f"  df_{nombre.lower()}  "
              f"({len(df)} filas × {len(df.columns) if not df.empty else 0} cols)")
        print(f"{'═'*70}")
        if not df.empty:
            print(df.to_string(index=False))
        else:
            print("  (vacío)")

    if provincias_fallidas:
        print(f"\n{'═'*70}")
        print(f"  PROVINCIAS FALLIDAS ({len(provincias_fallidas)} de {sum(len(v) for v in PROVINCIAS_POR_REGION.values())})")
        print(f"{'═'*70}")
        df_fail = pd.DataFrame(provincias_fallidas)[["region","provincia","razon"]]
        print(df_fail.to_string(index=False))

    # ── Guardar Excel ──────────────────────────────────────────────
    OUTPUT_PATH = (
        r"C:\Users\usuario\OneDrive\Desktop\AFP INTEGRA"
        r"\Modelo Franco\Elecciones\df_consolidado.xlsx"
    )
    if not df_consolidado.empty:
        df_consolidado["timestamp"] = (
            pd.to_datetime(df_consolidado["timestamp"]).dt.floor("s")
        )
        df_consolidado.to_excel(OUTPUT_PATH, index=False)
        print(f"\nGuardado: {len(df_consolidado)} filas → {OUTPUT_PATH}")
    else:
        print("\ndf_consolidado vacío — no se guardó Excel.")