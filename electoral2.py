

"""
ONPE Segunda Vuelta 2026 — Scraper DOM v3
Selectores confirmados funcionando:
  - S_VISTA:    mat-select[formcontrolname='region']
  - S_DEPTO:    mat-select[formcontrolname='department']
  - S_PROV:     mat-select[formcontrolname='province']
  - Tarjetas:   .tarjeta-candidato--izquierda / --derecha
  - Actas:      ul.leyenda.vertical li

Fix v2: ElementClickInterceptedException
  → cdk-overlay-backdrop intercepta clicks cuando el panel anterior
    no se cerró del todo. Solución: esperar a que desaparezca el
    backdrop antes de cada acción, y usar JS click como fallback.

v3: Agrega scraping por PROVINCIA dentro de cada REGIÓN.
  - Equivalencias CSV (sin tildes) ↔ nombre ONPE (con tildes).
  - LIMA se divide en LIMA REGIÓN (provincias) + LIMA PROVINCIA (solo Lima).
  - Mapa REGION → lista de provincias con los nombres exactos que la web acepta.
"""

import time, re, logging
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    StaleElementReferenceException, ElementClickInterceptedException
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
log = logging.getLogger("ONPE")

URL       = "https://resultadosegundavuelta.onpe.gob.pe/main/resumen"
S_VISTA   = "mat-select[formcontrolname='region']"
S_DEPTO   = "mat-select[formcontrolname='department']"
S_PROV    = "mat-select[formcontrolname='province']"
S_LIMPIAR = "//button[contains(normalize-space(.),'LIMPIAR')]"
S_BACKDROP = "div.cdk-overlay-backdrop"

# ── Regiones exactamente como la página ONPE las muestra ──────────
REGIONES = [
    "AMAZONAS", "ÁNCASH", "APURÍMAC", "AREQUIPA", "AYACUCHO",
    "CAJAMARCA", "CALLAO", "CUSCO", "HUANCAVELICA", "HUÁNUCO",
    "ICA", "JUNÍN", "LA LIBERTAD", "LAMBAYEQUE", "LIMA",
    "LORETO", "MADRE DE DIOS", "MOQUEGUA", "PASCO", "PIURA",
    "PUNO", "SAN MARTÍN", "TACNA", "TUMBES", "UCAYALI"
]

# ── Equivalencias: nombre en CSV/base-datos → nombre ONPE ─────────
# Útil si necesitas cruzar el CSV "Provincia - Ubigeos.csv" con los
# resultados del scraper.  CSV no tiene tildes en algunos campos;
# ONPE sí las usa.
EQUIV_CSV_A_ONPE = {
    # Regiones
    "ANCASH":          "ÁNCASH",
    "APURIMAC":        "APURÍMAC",
    "HUANUCO":         "HUÁNUCO",
    "JUNIN":           "JUNÍN",
    "SAN MARTIN":      "SAN MARTÍN",
    "LIMA REGION":     "LIMA",       # La web llama "LIMA" a lo que el CSV llama "LIMA REGION"
    "LIMA PROVINCIA":  "LIMA",       # idem: en ONPE todo es "LIMA"; las provincias diferencian
    # Provincias (casos con tildes o diferencias de escritura)
    "BONGARA":         "BONGARÁ",
    "RODRIGUEZ DE MENDOZA": "RODRÍGUEZ DE MENDOZA",
    "ANTONIO RAYMONDI":     "ANTONIO RAYMONDI",        # sin cambio
    "CARLOS FERMIN FITZCARRALD": "CARLOS FERMÍN FITZCARRAL",
    "ASUNCION":        "ASUNCIÓN",
    "CARAVELI":        "CARAVELÍ",
    "CELENDIN":        "CELENDÍN",
    "JAEN":            "JAÉN",
    "LA CONVENCION":   "LA CONVENCIÓN",
    "HUAMALIES":       "HUAMALÍES",
    "MARANON":         "MARAÑON",
    "SANCHEZ CARRION": "SÁNCHEZ CARRIÓN",
    "CHEPEN":          "CHEPÉN",
    "JULCAN":          "JULCÁN",
    "GRAN CHIMU":      "GRAN CHIMÚ",
    "MARISCAL RAMON CASTILLA": "MARISCAL RAMÓN CASTILLA",
    "GENERAL SANCHEZ CERRO":   "GENERAL SÁNCHEZ CERRO",
    "DANIEL ALCIDES CARRION":  "DANIEL ALCIDES CARRIÓN",
    "AZANGARO":        "AZÁNGARO",
    "SAN ROMAN":       "SAN ROMÁN",
    "PAUCAR DEL SARA SARA":    "PAUCAR DEL SARA SARA",  # sin cambio
    "MARISCAL CACERES":        "MARISCAL CÁCERES",
    "PURUS":           "PURÚS",
    "FERREÑAFE":       "FERREÑAFE",  # sin cambio (ya tiene tilde en CSV)
    "NAZCA":           "NAZCA",      # sin cambio (CSV tiene comillas extra, se limpian)
}

# ── Provincias por región (nombres tal como ONPE los reconoce) ────
# Fuente: ubigeo-reniec.json (RENIEC) + validación contra la web.
# Nota Lima: "LIMA" en ONPE es la región completa; la provincia
#   capital se llama "LIMA" también.  El CSV distingue LIMA REGION
#   (9 provincias sin Lima ciudad) y LIMA PROVINCIA (solo Lima).
#   Aquí las unimos bajo la key "LIMA".
PROVINCIAS_POR_REGION = {
    "AMAZONAS": [
        "CHACHAPOYAS", "BAGUA", "BONGARÁ", "LUYA",
        "RODRÍGUEZ DE MENDOZA", "CONDORCANQUI", "UTCUBAMBA",
    ],
    "ÁNCASH": [
        "HUARAZ", "AIJA", "BOLOGNESI", "CARHUAZ", "CASMA", "CORONGO",
        "HUAYLAS", "HUARI", "MARISCAL LUZURIAGA", "PALLASCA", "POMABAMBA",
        "RECUAY", "SANTA", "SIHUAS", "YUNGAY", "ANTONIO RAYMONDI",
        "CARLOS FERMÍN FITZCARRAL", "ASUNCIÓN", "HUARMEY", "OCROS",
    ],
    "APURÍMAC": [
        "ABANCAY", "AYMARAES", "ANDAHUAYLAS", "ANTABAMBA",
        "COTABAMBAS", "GRAU", "CHINCHEROS",
    ],
    "AREQUIPA": [
        "AREQUIPA", "CAYLLOMA", "CAMANÁ", "CARAVELÍ",
        "CASTILLA", "CONDESUYOS", "ISLAY", "LA UNION",
    ],
    "AYACUCHO": [
        "HUAMANGA", "CANGALLO", "HUANTA", "LA MAR", "LUCANAS",
        "PARINACOCHAS", "VICTOR FAJARDO", "HUANCA SANCOS",
        "VILCAS HUAMAN", "PAUCAR DEL SARA SARA", "SUCRE",
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
        "HUANUCO", "AMBO", "DOS DE MAYO", "HUAMALÍES", "MARAÑON",
        "LEONCIO PRADO", "PACHITEA", "PUERTO INCA",
        "HUACAYBAMBA", "LAURICOCHA", "YAROWILCA",
    ],
    "ICA": [
        "ICA", "CHINCHA", "NAZCA", "PISCO", "PALPA",
    ],
    "JUNÍN": [
        "HUANCAYO", "CONCEPCION", "JAUJA", "JUNÍN", "TARMA",
        "YAULI", "SATIPO", "CHANCHAMAYO", "CHUPACA",
    ],
    "LA LIBERTAD": [
        "TRUJILLO", "BOLIVAR", "SÁNCHEZ CARRIÓN", "OTUZCO",
        "PACASMAYO", "PATAZ", "SANTIAGO DE CHUCO", "ASCOPE",
        "CHEPÉN", "JULCÁN", "GRAN CHIMÚ", "VIRU",
    ],
    "LAMBAYEQUE": [
        "CHICLAYO", "FERREÑAFE", "LAMBAYEQUE",
    ],
    "LIMA": [
        # Lima provincia capital
        "LIMA",
        # Lima Región (9 provincias fuera de Lima ciudad)
        "BARRANCA", "CAJATAMBO", "CANTA", "CAÑETE", "HUARAL",
        "HUAROCHIRI", "HUAURA", "OYON", "YAUYOS",
    ],
    "LORETO": [
        "MAYNAS", "ALTO AMAZONAS", "LORETO", "REQUENA", "UCAYALI",
        "MARISCAL RAMÓN CASTILLA", "DATEM DEL MARAÑON", "PUTUMAYO",
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
        "PIURA", "AYABACA", "HUANCABAMBA", "MORROPON",
        "PAITA", "SULLANA", "TALARA", "SECHURA",
    ],
    "PUNO": [
        "PUNO", "AZÁNGARO", "CARABAYA", "CHUCUITO", "HUANCANE",
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


# ── Parsers ────────────────────────────────────────────────────────

def _pct(txt):
    c = re.sub(r"[^\d.,]", "", txt).replace(",", ".")
    try:    return float(c)
    except: return 0.0

def _votos(txt):
    limpio = txt.replace("'","").replace(".","").replace(",","")
    nums = re.findall(r"\d+", limpio)
    return int("".join(nums)) if nums else 0

def _num_par(txt):
    m = re.search(r"\((\d+)\)", txt)
    return int(m.group(1)) if m else 0

def _txt(parent, css):
    try:    return parent.find_element(By.CSS_SELECTOR, css).text.strip()
    except: return ""


# ── Driver ─────────────────────────────────────────────────────────

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


# ── Utilidades de espera ───────────────────────────────────────────

def wait_for(driver, css, timeout=20):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css)))
        return True
    except TimeoutException:
        return False

def wait_gone(driver, css, timeout=5):
    """Espera hasta que el elemento CSS desaparezca del DOM."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, css)))
    except TimeoutException:
        pass

def esperar_sin_backdrop(driver, timeout=6):
    """Espera hasta que el cdk-overlay-backdrop desaparezca."""
    wait_gone(driver, S_BACKDROP, timeout)
    time.sleep(0.2)


# ── Click robusto ──────────────────────────────────────────────────

def click_robusto(driver, element):
    try:
        element.click()
        return True
    except ElementClickInterceptedException:
        log.debug("  click interceptado → usando JS click")
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception as e:
            log.debug(f"  JS click falló: {e}")
            return False
    except Exception as e:
        log.debug(f"  click falló: {e}")
        return False


# ── Interacción mat-select ─────────────────────────────────────────

def abrir_select(driver, css, timeout=12):
    """Abre el mat-select y espera el panel de opciones."""
    esperar_sin_backdrop(driver)
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, css)))
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.3)
        if not click_robusto(driver, el):
            return False
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "mat-option")))
        time.sleep(0.3)
        return True
    except TimeoutException:
        log.debug(f"  abrir_select timeout: {css}")
        return False
    except Exception as e:
        log.debug(f"  abrir_select error: {e}")
        return False

def elegir(driver, texto, timeout=8):
    """
    Elige la opción por texto en el panel abierto.
    Prueba coincidencia exacta primero, luego parcial.
    Usa JS click si el click normal es interceptado.
    """
    for xpath in [
        f"//mat-option[normalize-space(.)='{texto}']",
        f"//mat-option[contains(normalize-space(.),'{texto}')]",
    ]:
        try:
            opt = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", opt)
            time.sleep(0.2)
            if click_robusto(driver, opt):
                esperar_sin_backdrop(driver, timeout=5)
                time.sleep(3.0)   # Angular re-renderiza
                return True
        except TimeoutException:
            continue
        except Exception as e:
            log.debug(f"  elegir '{texto}': {e}")
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
        time.sleep(2.0)
    except Exception:
        pass

def seleccionar_vista(driver, vista):
    """Selecciona TODOS | PERÚ | EXTRANJERO si no está ya seleccionado."""
    els = driver.find_elements(By.CSS_SELECTOR, S_VISTA)
    if els and vista.upper() in els[0].text.upper():
        return True
    if abrir_select(driver, S_VISTA):
        return elegir(driver, vista)
    return False


# ── Detectar si el select de PROVINCIA está disponible ────────────

def provincia_select_disponible(driver, timeout=5):
    """
    Retorna True si el select de PROVINCIA existe y está habilitado.
    Después de elegir REGIÓN, Angular puede tardar en renderizarlo.
    """
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, S_PROV)))
        # mat-select deshabilitado tiene aria-disabled="true"
        disabled = el.get_attribute("aria-disabled")
        return disabled != "true"
    except TimeoutException:
        return False


# ── Leer DOM ───────────────────────────────────────────────────────

def leer_actas(driver):
    """Lee ul.leyenda.vertical: Contabilizadas, JEE, Pendientes."""
    meta = {"actas_contabilizadas": None, "actas_jee": None,
            "actas_pendientes": None, "actas_total": None,
            "pct_contabilizadas": None}
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "ul.leyenda.vertical li")
        for li in items:
            txt = li.text.strip()
            n   = _num_par(txt)
            tl  = txt.lower()
            if   "contabilizad" in tl: meta["actas_contabilizadas"] = n
            elif "jee" in tl or "envío" in tl or "envio" in tl:
                meta["actas_jee"] = n
            elif "pendiente" in tl:    meta["actas_pendientes"] = n
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
# SCRAPING
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

        # Espera inicial
        log.info("Esperando carga Angular...")
        cargado = False
        for css, t in [(".tarjeta-candidato", 60), ("mat-select", 60), ("app-root", 90)]:
            if wait_for(driver, css, t):
                log.info(f"  SPA lista [{css}]")
                cargado = True
                break
        if not cargado:
            log.error("SPA no cargó — verifica conexión")
            return {"regional": pd.DataFrame(), "provincial": pd.DataFrame(),
                    "extranjero": pd.DataFrame(), "consolidado": pd.DataFrame()}
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
                log.warning(f"    No se pudo abrir select REGIÓN para {region}")
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
        log.info("BLOQUE 2: PERÚ → Región → Provincia (195 provincias)")
        log.info(f"{'─'*50}")

        total_provs = sum(len(v) for v in PROVINCIAS_POR_REGION.values())
        prov_counter = 0

        for region, provincias in PROVINCIAS_POR_REGION.items():
            log.info(f"\n  ── {region} ({len(provincias)} provincias) ──")

            for provincia in provincias:
                prov_counter += 1
                log.info(f"  [{prov_counter:03d}/{total_provs}] {region} → {provincia}...")

                # Reset completo
                limpiar(driver)
                seleccionar_vista(driver, "PERÚ")
                if not wait_for(driver, S_DEPTO, timeout=12):
                    log.warning(f"    S_DEPTO no disponible — saltando {provincia}")
                    continue

                # Seleccionar REGIÓN
                if not abrir_select(driver, S_DEPTO):
                    log.warning(f"    No se pudo abrir REGIÓN para {provincia}")
                    continue
                if not elegir(driver, region):
                    log.warning(f"    No se pudo elegir REGIÓN {region}")
                    continue

                # Esperar a que el select de PROVINCIA esté disponible
                if not provincia_select_disponible(driver, timeout=8):
                    log.warning(f"    Select PROVINCIA no disponible para {region}")
                    continue

                # Abrir select PROVINCIA
                if not abrir_select(driver, S_PROV):
                    log.warning(f"    No se pudo abrir select PROVINCIA")
                    continue

                # Elegir provincia
                if not elegir(driver, provincia):
                    log.warning(f"    No se encontró '{provincia}' en el panel")
                    continue

                # Leer datos
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
                             f"({rows[0]['pct_contabilizadas']}%)")
                else:
                    log.warning(f"    Sin datos DOM para {provincia}")

        # ── BLOQUE 3: EXTRANJERO ───────────────────────────────────
        log.info(f"\n{'─'*50}")
        log.info("BLOQUE 3: EXTRANJERO")
        log.info(f"{'─'*50}")

        limpiar(driver)
        seleccionar_vista(driver, "EXTRANJERO")
        time.sleep(2.5)

        rows = leer_tarjetas(driver, "EXTRANJERO", "extranjero")
        if rows:
            rows_extranjero.extend(rows)
            for r in rows:
                log.info(f"  {r['candidato'][:38]:38s} "
                         f"| {r['porcentaje']:6.3f}% "
                         f"| {r['votos']:>12,} votos")
        else:
            log.warning("  Sin datos DOM para EXTRANJERO")

    finally:
        driver.quit()
        log.info("\nBrowser cerrado.")

    df_regional   = pd.DataFrame(rows_regional)
    df_provincial = pd.DataFrame(rows_provincial)
    df_extranjero = pd.DataFrame(rows_extranjero)
    partes = [df for df in [df_regional, df_provincial, df_extranjero] if not df.empty]
    df_consolidado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

    log.info(f"\n{'═'*60}  RESUMEN:")
    log.info(f"  df_regional   : {len(df_regional):4d} filas | "
             f"{df_regional['ubigeo'].nunique() if not df_regional.empty else 0} regiones")
    log.info(f"  df_provincial : {len(df_provincial):4d} filas | "
             f"{df_provincial['ubigeo'].nunique() if not df_provincial.empty else 0} provincias")
    log.info(f"  df_extranjero : {len(df_extranjero):4d} filas")
    log.info(f"  df_consolidado: {len(df_consolidado):4d} filas")
    log.info(f"{'═'*60}")

    return {
        "regional":    df_regional,
        "provincial":  df_provincial,
        "extranjero":  df_extranjero,
        "consolidado": df_consolidado,
    }


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

    HEADLESS      = False   # True una vez que funcione bien
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
        res            = scrape(headless=HEADLESS)
        df_regional    = res["regional"]
        df_provincial  = res["provincial"]
        df_extranjero  = res["extranjero"]
        df_consolidado = res["consolidado"]

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