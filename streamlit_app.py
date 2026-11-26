from io import BytesIO
import concurrent.futures
import time
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

BASE = Path(__file__).parent
LOGO_PATH = BASE / "imagenes" / "jocs_logo.png"

st.set_page_config(
    page_title="Facturación Emitida y Recibida",
    page_icon=str(LOGO_PATH),
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Paleta de marca JOCS ────────────────────────────────────────────────────
JOCS_AZUL = "#244072"
JOCS_VERDE = "#507058"
C_MARGEN = "#c98a2e"
C_NEG = "#c0392b"

# ══════════════════════════════════════════════════════════════════════════════
# FUENTE DE DATOS (2026-10-28, reconstruida desde cero -- ver
# [[tablero-facturacion-reset-2026-10-28]]). Carpeta "Carga Aérea" en Drive,
# organizada por mes + una carpeta "General" para lo que no es mensual
# (AMEX, Presupuestos). IDs confirmados uno por uno con el usuario, no
# adivinados -- varios no estaban compartidos "cualquiera con el link" y
# hubo que pedirle que los habilitara.
# ══════════════════════════════════════════════════════════════════════════════
MES_COL_MAYUS = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
                 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE",
                 12: "DICIEMBRE"}
ANIO_ACTUAL = 2026

# Cada mes vive en su propia carpeta de Drive con 3 archivos propios
# (Emitidas y Recibidas, Costos de Referencia, Corresponsalías) -- IDs
# resueltos siguiendo los accesos directos de cada carpeta (2026-10-30,
# carpeta "Julio" agregada a pedido del usuario: "vamos a cargar julio con
# las mismas reglas que se uso para agosto"). 'emitidas_recibidas_siguiente'
# es el archivo del mes SIGUIENTE a este periodo -- de ahí se saca SOLO
# Imile AIFA en Emitidas, y el MONTO 2 de KYA TI/Yatec en Recibidas (ver
# load_emitidas_recibidas). El nombre de la hoja de Corresponsalías NO es
# consistente entre meses (Agosto usa el nombre del mes, Julio usa
# "CORRESPONSALÍAS" a secas) -- por eso se guarda explícito por periodo en
# vez de asumir que siempre es el mes en español.
PERIODOS = {
    "agosto": {
        "mes_num": 8, "mes_txt": "Agosto",
        "emitidas_recibidas": "11of64actB1Lubuz88Ai55UJzgnPz_N5QVlKyfpPUuRk",
        "emitidas_recibidas_siguiente": "1EmLMmxI_-u2DXsaXD2jZKjNKviTR6ovpIbk6SRG1R88",  # Septiembre
        "costos_de_referencia": "1KcqmhUnh1-UHGdoF53kuDujoBZtiB94L5620ZIjFq4U",
        "corresponsalias": "1cGGgOZ1eX_gsKHcCeHEiPjYju5L1hmjB-HVq2hpa7oo",
        "corresponsalias_hoja": "Agosto",
    },
    "julio": {
        "mes_num": 7, "mes_txt": "Julio",
        "emitidas_recibidas": "1XiYbkZ6AzZgyIy7mRKTt-KItpt8ULkXI792b4Sou8EU",
        "emitidas_recibidas_siguiente": "11of64actB1Lubuz88Ai55UJzgnPz_N5QVlKyfpPUuRk",  # Agosto
        "costos_de_referencia": "1xSSJaAsw4sH6Gle2qTb7heRAFNAaf9rsImYPBtKRpQs",
        "corresponsalias": "1AF0CtzN6YXlOfcYIYB3VbppnuR86z9EfXbIha4deOD8",
        "corresponsalias_hoja": "CORRESPONSALÍAS",
    },
    "junio": {
        "mes_num": 6, "mes_txt": "Junio",
        "emitidas_recibidas": "1n73ZvgejPFKctqHwmu5VD9YZeBCUwRttNP6HONeLycE",
        "emitidas_recibidas_siguiente": "1XiYbkZ6AzZgyIy7mRKTt-KItpt8ULkXI792b4Sou8EU",  # Julio
        "costos_de_referencia": "1LWp7qp82_IDR8T5HvKL6pHfbeM7K5NHu6_1syPLchhA",
        "corresponsalias": "1hgWla6AYYh881yrzOWwpkGBKVhRHO4OPl9YLCSJ6Z5Y",
        "corresponsalias_hoja": "Hoja 1",
    },
    "mayo": {
        "mes_num": 5, "mes_txt": "Mayo",
        "emitidas_recibidas": "1gcl0spdaFHTBRoLzxRERtOY2R3aEAb0QfbN8tD4DGW4",
        "emitidas_recibidas_siguiente": "1n73ZvgejPFKctqHwmu5VD9YZeBCUwRttNP6HONeLycE",  # Junio
        "costos_de_referencia": "1dhOKHzxHyx998mSn-Esmon1ak9emjohepExxMGe6fMY",
        "corresponsalias": "1gpUBWWfxbQxW7iKW-lyqaegpGO-Wlaj2bZCTfYs8LaQ",
        "corresponsalias_hoja": "Hoja 1",
    },
    "abril": {
        "mes_num": 4, "mes_txt": "Abril",
        "emitidas_recibidas": "1SZ9Cymny-V5AQSbk26xeGT299kcBmQ1TChS6s3CnfH4",
        "emitidas_recibidas_siguiente": "1gcl0spdaFHTBRoLzxRERtOY2R3aEAb0QfbN8tD4DGW4",  # Mayo
        "costos_de_referencia": "1SUw-lZ1HFKUTVbKVjSblC9cVQ5V-p0oTvYXqJNkd9Uw",
        "corresponsalias": "1xl9rrzWuITKvv9xZmJupN5LAVfE8-pijGgd7qukyqAY",
        "corresponsalias_hoja": "CORRECTA",
        # Las Emitidas de IMILE de Abril NO vienen ni en el archivo de Abril
        # ni en el de Mayo -- aparecen en el archivo de Junio (2 meses
        # después), en una columna aparte "Monto 2" (confirmado con el
        # usuario, 2026-11-18: "solamente de la sucursal IMILE debe sacarse
        # del archivo de emitidas y recibidas de Junio ... y son las que
        # tengan como aduana 'IMILE' propios y que esten en la columna
        # Monto 2"). Ver _leer_imile_emitidas_monto2 / load_emitidas_recibidas.
        "imile_monto2_desde": "1n73ZvgejPFKctqHwmu5VD9YZeBCUwRttNP6HONeLycE",  # Junio
    },
    "marzo": {
        "mes_num": 3, "mes_txt": "Marzo",
        "emitidas_recibidas": "1R3diXUYNwx8M1fOT5VH5jynTUzkgQW2aolRVKyL3104",
        "emitidas_recibidas_siguiente": "1SZ9Cymny-V5AQSbk26xeGT299kcBmQ1TChS6s3CnfH4",  # Abril
        "costos_de_referencia": "1Lk_3keBA0df0T_9ADlOZEnbOIY4Br7FHbGHS211I8p8",
        "corresponsalias": "11xFlZFbbpFP0SrbVsY1GVQKHgLHchGLnx1cdg2KGFnw",
        "corresponsalias_hoja": "Marzo",
        # Las Emitidas de IMILE de Marzo tampoco vienen en el archivo de
        # Marzo ni en el de Abril -- aparecen en el archivo de Mayo (2 meses
        # después), igual patrón que Abril/Junio pero aquí SÍ en la columna
        # normal "Monto 1" (confirmado con el usuario, 2026-11-18, tras
        # corregir que la regla de Mayo era en realidad para Marzo: "Mayo no
        # se tenia que modificar, la modificacion era para Marzo, pero
        # sacando de las emitidas de mayo"). Ver _leer_hoja_factura /
        # load_emitidas_recibidas.
        "imile_monto1_desde": "1gcl0spdaFHTBRoLzxRERtOY2R3aEAb0QfbN8tD4DGW4",  # Mayo
    },
    "febrero": {
        "mes_num": 2, "mes_txt": "Febrero",
        "emitidas_recibidas": "1qwZBMIHX908ShtRp7qKg-RQtPJUUTMWsnALy7zDsihY",
        "emitidas_recibidas_siguiente": "1R3diXUYNwx8M1fOT5VH5jynTUzkgQW2aolRVKyL3104",  # Marzo
        "costos_de_referencia": "1UT2vN_KhnTc-SG6I9KGmwVm90COpDDdi6lk0e7c8fNQ",
        "corresponsalias": "1R82f8gVSI8Pr3jCcjLocs4AwlHki4e6H6vbgEnRSLCs",
        "corresponsalias_hoja": "Feb",
    },
    "enero": {
        "mes_num": 1, "mes_txt": "Enero",
        "emitidas_recibidas": "1lz93J8ApZu2fkxkTVa3QbJxlEfjvZ2i-hl6DO2WSoV4",
        "emitidas_recibidas_siguiente": "1qwZBMIHX908ShtRp7qKg-RQtPJUUTMWsnALy7zDsihY",  # Febrero
        "costos_de_referencia": "1c08JsbVLpET-xmfqNEi0U6Wg0MOkYFVHTcWvDm-DvnU",
        "corresponsalias": "1etUvRfuKqhuBsUmBtowwPEk4CYoFzZYJqvIWkceqaEc",
        "corresponsalias_hoja": "ENERO",
    },
}
PERIODO_DEFAULT = "agosto"

DRIVE_SHEET_IDS = {
    # Fuentes compartidas entre todos los meses -- ya traen su propia
    # columna de mes/periodo (filtran por MES_ACTUAL adentro del loader).
    "sueldos": "11MqeqH1_vo7_qi1HWnMBDNMcGCvBe9rczwLJRFF340I",
    "honorarios_aa": "1mwdtAe4wt5hwDaSa-xnevCVvXJIaBQD4oIe2ypfZe7Q",
    "presupuestos_2026": "1eFuRgbRap7OL4uh6BY3cHLv6O1rHc3915pC6j13rfO4",
    "amex_alejandro": "1JNeNj6JERd13i5TZCCH1_SHY5vXv7dmUg0lUyRzS3ow",
    "amex_josefa": "1MQyjGwAHgRSMBJdZ_HFGskhplB2tk1JdDo2oY-6tzZ0",
    "amex_angelica": "1OewXOKinbq6cOxedrb6NIiaeF3385eCrbmzZFdtBM-0",
    "efectivale": "1NE0MG2qaY8k9uiSwRPviqXEfa9uIOezHsHGSg8KruNY",
    # Hoja "Personal por CC" (2026-11-25): reemplaza el Cierre Carga Aerea
    # como fuente del reparto de Sueldos por aduana, ver
    # _construir_sueldos_por_aduana. PINEDA BERNAL ZUGELY y GONZALEZ
    # SAMANTHA, antes excluidas de Sueldos por duplicar un monto aparte del
    # P&L de Imile Laredo (2026-11-19), ya vienen en esta hoja como
    # empleados normales -- el usuario confirmo usarla tal cual, sin esa
    # exclusion (2026-11-25).
    "personal_por_cc": "1C1hLZKi7dKO2OH40WKqjhWKKmHKIPQIirgzmql9TUJA",
}

CONCEPTO_PRESUPUESTO = {"INGRESOS": "ingresos", "CTO REF": "costos",
                        "SUELDOS": "sueldos", "GASTOS OP": "gastos"}

COLS_NOMINA = [
    ("Total sueldo (N)", "Sueldo Base"),
    ("Total prestaciones (N)", "Prestaciones Patr."),
    ("Total cuotas patrón (N)", "Cuotas Patr."),
    ("vales 1 (N)", "Vales F."),
    ("vales 2 (N)", "Vales V."),
    ("Puntualidad (lira)", "Puntualidad"),
    ("Comisión (lira)", "Comisión"),
    ("Excedente (lira)", "Excedente"),
    ("Provisiones", "Provisiones (+/-)"),
]

AMEX_COL_UUID = {"ALEJANDRO": "uuid", "JOSEFA": "Unnamed: 10", "ANGELICA": "UUID"}

PALETA_CAT = ["#244072", "#507058", "#8a6d4b", "#4a7fa8", "#7a9e6e", "#c98a2e",
              "#6f5c8a", "#9aa5b1", "#a45c6b", "#3f8f8a", "#c9994e", "#5d7a9e"]

def color_for(i):
    return PALETA_CAT[i % len(PALETA_CAT)]

def norm_txt(v):
    if not isinstance(v, str):
        return None
    v = v.strip()
    return v if v else None

def coalesce_col(df, *names):
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series([None] * len(df), index=df.index)

def fmt(v):
    if v is None or pd.isna(v):
        return "—"
    s = "-$" if v < 0 else "$"
    return s + f"{abs(v):,.0f}"

@st.cache_data(ttl=1800, show_spinner=False)
def _descargar_sheet_drive(file_id):
    """3 intentos con backoff -- al paralelizar las descargas
    (_prefetch_drive) unas pocas empiezan a fallar con error SSL
    ('Max retries exceeded'/'UNEXPECTED_EOF_WHILE_READING') si coinciden
    demasiadas al mismo tiempo contra docs.google.com; un reintento breve
    las resuelve sin tener que volver a bajar la concurrencia al punto de
    perder la ganancia de velocidad."""
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    ultimo_error = None
    for intento in range(3):
        try:
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()
            return resp.content
        except requests.exceptions.RequestException as e:
            ultimo_error = e
            time.sleep(0.6 * (intento + 1))
    raise ultimo_error

def _prefetch_drive(file_ids):
    """Dispara en paralelo las descargas de Drive que hagan falta para
    esta corrida -- cada una ya está cacheada por su cuenta
    (_descargar_sheet_drive, por file_id), pero antes se pedían una tras
    otra: en frío (cache vacío, o un rango/mes nuevo) la ESPERA real era
    la SUMA de las ~30-40 descargas secuenciales. Al lanzarlas a la vez,
    el resultado es idéntico (cada una se guarda en su propio lugar del
    caché) pero la espera baja a unas pocas tandas en paralelo en vez de
    la suma de todas (pedido del usuario, 2026-11-25: "quiero optimizar
    todos los tiempos de espera"). max_workers se queda deliberadamente
    moderado (6) -- con más de ~10 a la vez, docs.google.com empieza a
    cortar conexiones con error SSL en vez de responder más rápido."""
    ids_unicos = sorted({fid for fid in file_ids if fid})
    if not ids_unicos:
        return
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(6, len(ids_unicos))) as ex:
        list(ex.map(_descargar_sheet_drive, ids_unicos))

# Julio nombra esta aduana "IMILE" a secas en vez de "IMILE AIFA" (archivos
# de meses mas recientes ya usan el nombre completo) -- sin este alias, el
# desfase de un mes de Imile AIFA en Emitidas no reconoce esas filas y se
# cuentan dos veces (confirmado revisando el archivo de Julio, 2026-10-30).
ALIAS_ADUANA_FACTURA = {"IMILE": "IMILE AIFA"}

# Aduanas cuyas EMITIDAS vienen un mes atrasadas en el archivo mensual --
# se excluyen del mes propio y se toman del archivo del mes siguiente (ver
# load_emitidas_recibidas). TMM agregado 2026-11-05, misma regla que ya
# tenia Imile AIFA.
ADUANAS_DESFASE_EMITIDAS = {"IMILE AIFA", "TMM"}

# Excepciones puntuales al desfase de arriba, por periodo -- el desfase de
# IMILE AIFA NO es uniforme mes a mes (ver 'imile_monto2_desde' /
# 'imile_monto1_desde' en PERIODOS, para los meses donde viene desfasada 2
# meses en vez de 1). Vacío por ahora -- no asumir que algún mes necesita
# quedarse con su propio archivo sin que el usuario lo confirme
# explícitamente (2026-11-18: se probó con Mayo por un malentendido y el
# usuario corrigió que esa regla en realidad era para Marzo).
ADUANAS_SIN_DESFASE_POR_PERIODO = {}

import unicodedata

def _normaliza_aduana(s):
    """Además del alias de Imile AIFA, quita acentos -- algunas fuentes
    (ej. Efectivale) escriben 'TRÁNSITO' con acento mientras que
    Emitidas/Recibidas usa 'TRANSITO' sin acento, y sin esto aparecían como
    dos aduanas distintas en el filtro."""
    if not isinstance(s, str):
        return "SIN ADUANA"
    s = s.upper()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return ALIAS_ADUANA_FACTURA.get(s, s)

def _normaliza_nombre(s):
    if not isinstance(s, str):
        return None
    s = s.strip().upper()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))

def _leer_hoja_factura(path, hoja, tipo_factura):
    """Misma lógica de siempre para leer 'emitidas'/'recibidas': columna de
    monto oficial ya convertida a MXN ('monto 1' en recibidas, 'monto' en
    emitidas -- NUNCA se usa 'Importe' como respaldo, el usuario deja líneas
    sin monto a propósito). El concepto de cierre de recibidas en esta
    fuente trae el encabezado vacío (un espacio ' '), por eso el fallback.
    A propósito NUNCA se usa 'MONTO 2' aquí, para ningún proveedor (confirmado
    con el usuario, 2026-10-28: "de monto 2 lo único que tomes sea de kya y
    yatec, si hay de otros proveedores, no los agregues"). La única columna
    de monto es 'monto 1' (o 'monto' en Emitidas). El único lugar donde
    'MONTO 2' sí se usa es la función _leer_kya_yatec_monto2_mes_anterior,
    que jala SOLO KYA TI/Yatec desde el archivo del mes siguiente."""
    df = pd.read_excel(path, sheet_name=hoja, header=0)
    df.columns = [str(c).strip().lower() for c in df.columns]
    out = pd.DataFrame(index=df.index)
    out["contraparte"] = coalesce_col(
        df, "receptornombre" if tipo_factura == "EMITIDA" else "emisornombre").map(norm_txt)
    out["descripcion"] = coalesce_col(df, "descripcion").map(norm_txt)
    monto = pd.to_numeric(coalesce_col(df, "monto 1", "monto"), errors="coerce")
    out["importe"] = monto.fillna(0.0)
    # 'nombre' es fallback para meses donde Recibidas no trae columna
    # 'aduana' (ej. Abril/Mayo 2026) -- ahi el nombre de aduana viene en una
    # columna llamada 'Nombre' en su lugar, confirmado revisando la fuente.
    # 'cc' es el mismo caso para Recibidas de Enero 2026 (no trae ni
    # 'aduana' ni 'nombre', solo 'cc').
    aduana = coalesce_col(df, "aduana", "nombre", "cc").map(norm_txt)
    out["aduana"] = aduana.map(_normaliza_aduana)
    # 'clasificación' es el mismo concepto que 'tipo' (PROPIOS/TERCEROS/
    # DISCRECIONAL, mismos valores) con otro nombre -- Emitidas de Enero
    # 2026 no trae columna 'tipo' en absoluto, solo 'clasificación'. Sin
    # este fallback, el filtro de "solo PROPIOS" excluía el 100% de
    # Emitidas de Enero (Ingresos caía a $0, quedaba solo Honorarios AA).
    tipo_cc = coalesce_col(df, "tipo", "clasificación").map(norm_txt)
    out["tipo_cc"] = tipo_cc.map(lambda s: s.upper() if isinstance(s, str) else "SIN TIPO")
    # 'conceptos para el cierre' es el mismo encabezado con otra redacción
    # (Abril/Mayo 2026, en vez de 'concepto para cierre'). 'conceptos para
    # cierre' (plural, sin 'el') es una 3ra variante de Febrero 2026.
    out["concepto_cierre"] = coalesce_col(
        df, "concepto para cierre", "conceptos para el cierre",
        "conceptos para cierre", " ").map(norm_txt)
    out["cancelada"] = (out["aduana"] == "CANCELADA") | (out["tipo_cc"] == "CANCELADA")
    if tipo_factura == "RECIBIDA":
        # pandas convierte 'N/A' a NaN por default (igual que una celda vacia),
        # asi que para detectar el texto literal 'N/A' hay que releer sin ese
        # comportamiento. A pedido del usuario (2026-10-28): en Recibidas, las
        # filas marcadas 'N/A' (aduana/tipo) no se cuentan -- igual que CANCELADA.
        path.seek(0)
        crudo = pd.read_excel(path, sheet_name=hoja, header=0, keep_default_na=False, na_values=[])
        crudo.columns = [str(c).strip().lower() for c in crudo.columns]
        es_na = pd.Series(False, index=crudo.index)
        for col in ("aduana", "tipo"):
            if col in crudo.columns:
                es_na = es_na | (crudo[col].astype(str).str.strip().str.upper() == "N/A")
        out["cancelada"] = out["cancelada"] | es_na.reindex(out.index, fill_value=False)
    out["tipo_factura"] = tipo_factura
    return out[out["importe"] != 0]

KYA_YATEC = {"KYA TI", "YATEC"}

def _leer_kya_yatec_monto2_mes_anterior(path, hoja):
    """SOLO para KYA TI y Yatec: su 'MONTO 2' en el archivo de un mes (p.ej.
    Septiembre) le pertenece al mes ANTERIOR de ese archivo (Agosto) --
    confirmado con el usuario, 2026-10-28. Se usa para completar Agosto con
    lo que aparezca en Septiembre."""
    df = pd.read_excel(path, sheet_name=hoja, header=0)
    df.columns = [str(c).strip().lower() for c in df.columns]
    contraparte = coalesce_col(df, "emisornombre").map(norm_txt)
    df = df[contraparte.map(lambda s: isinstance(s, str) and s.upper() in KYA_YATEC)]
    out = pd.DataFrame(index=df.index)
    out["contraparte"] = coalesce_col(df, "emisornombre").map(norm_txt)
    out["descripcion"] = coalesce_col(df, "descripcion").map(norm_txt)
    out["importe"] = pd.to_numeric(coalesce_col(df, "monto 2"), errors="coerce").fillna(0.0)
    aduana = coalesce_col(df, "aduana").map(norm_txt)
    out["aduana"] = aduana.map(_normaliza_aduana)
    tipo_cc = coalesce_col(df, "tipo").map(norm_txt)
    out["tipo_cc"] = tipo_cc.map(lambda s: s.upper() if isinstance(s, str) else "SIN TIPO")
    out["concepto_cierre"] = coalesce_col(df, "concepto para cierre", " ").map(norm_txt)
    out["cancelada"] = (out["aduana"] == "CANCELADA") | (out["tipo_cc"] == "CANCELADA")
    out["tipo_factura"] = "RECIBIDA"
    return out[out["importe"] != 0]

def _leer_imile_emitidas_monto2(path, hoja):
    """SOLO para las Emitidas de IMILE de Abril 2026: a diferencia del resto
    de aduanas con desfase (que se toman del archivo del mes siguiente, en
    la columna normal de monto), las de IMILE de Abril aparecen hasta el
    archivo de Junio (2 meses después) y en una columna separada 'Monto 2'
    -- confirmado con el usuario, 2026-11-18. Ver PERIODOS['abril']
    ['imile_monto2_desde']."""
    df = pd.read_excel(path, sheet_name=hoja, header=0)
    df.columns = [str(c).strip().lower() for c in df.columns]
    out = pd.DataFrame(index=df.index)
    out["contraparte"] = coalesce_col(df, "receptornombre").map(norm_txt)
    out["descripcion"] = coalesce_col(df, "descripcion").map(norm_txt)
    out["importe"] = pd.to_numeric(coalesce_col(df, "monto 2"), errors="coerce").fillna(0.0)
    aduana = coalesce_col(df, "aduana").map(norm_txt)
    out["aduana"] = aduana.map(_normaliza_aduana)
    tipo_cc = coalesce_col(df, "tipo", "clasificación").map(norm_txt)
    out["tipo_cc"] = tipo_cc.map(lambda s: s.upper() if isinstance(s, str) else "SIN TIPO")
    out["concepto_cierre"] = coalesce_col(df, "concepto para cierre", " ").map(norm_txt)
    out["cancelada"] = (out["aduana"] == "CANCELADA") | (out["tipo_cc"] == "CANCELADA")
    out["tipo_factura"] = "EMITIDA"
    out = out[out["aduana"] == "IMILE AIFA"]  # "IMILE" ya normalizado al alias
    return out[out["importe"] != 0]

CORRESPONSALIAS_ADUANAS = {"VERACRUZ", "MANZANILLO", "LAZARO"}

def _leer_corresponsalias(path, hoja):
    """Facturas de agencias aduanales corresponsales de Veracruz/Manzanillo/
    Lazaro -- no vienen en el archivo de Emitidas/Recibidas pero si son gasto
    real de esas 3 aduanas (confirmado con el usuario, 2026-10-28). Se usa
    'Attribute:Total' (con IVA, Total = SubTotal x 1.16), igual que el resto
    de montos de Recibidas que ya son el total de la factura. El nombre de
    'hoja' varia por periodo -- no siempre es el nombre del mes (ver
    PERIODOS[...]['corresponsalias_hoja'])."""
    df = pd.read_excel(path, sheet_name=hoja, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    # Normalizar ANTES de filtrar: algunos meses (Abril/Junio 2026) escriben
    # la aduana en Mayúscula inicial o con acento ("Veracruz", "Lázaro") en
    # vez de mayúsculas sin acento -- sin esto el filtro isin() no
    # encontraba ninguna fila y Corresponsalías salía vacío para esos meses.
    # 'Aduana' (minúsculas excepto inicial) es la variante de Enero 2026.
    aduana_norm = coalesce_col(df, "ADUANA", "Aduana").map(norm_txt).map(_normaliza_aduana)
    coinciden = aduana_norm.isin(CORRESPONSALIAS_ADUANAS)
    df = df[coinciden]
    aduana_norm = aduana_norm[coinciden]
    out = pd.DataFrame(index=df.index)
    # 'NOMBRE'/'TOTAL' son fallback para Abril 2026. 'NombreEmisor'/'Monto'
    # son la variante de Febrero 2026. 'Content.Emisor.Attribute:Nombre'/
    # 'Content.Attribute:Total' son la variante de Marzo 2026 (mismo nombre
    # base que el estándar, con prefijo 'Content.').
    out["contraparte"] = coalesce_col(
        df, "Emisor.Attribute:Nombre", "NOMBRE", "NombreEmisor",
        "Content.Emisor.Attribute:Nombre").map(norm_txt)
    out["descripcion"] = "Corresponsalía"
    out["importe"] = pd.to_numeric(coalesce_col(
        df, "Attribute:Total", "TOTAL", "Monto", "Content.Attribute:Total"), errors="coerce").fillna(0.0)
    out["aduana"] = aduana_norm
    out["tipo_cc"] = "PROPIOS"
    out["concepto_cierre"] = None
    out["cancelada"] = False
    out["tipo_factura"] = "RECIBIDA"
    return out[out["importe"] != 0]

def _leer_efectivale(path, mes_txt):
    """Efectivale es una tarjeta prepagada corporativa -- un solo archivo
    con una hoja por mes (ENERO..DICIEMBRE). No se tomaba en cuenta para
    Gastos hasta ahora (confirmado con el usuario, 2026-10-30): "sumar los
    cargos a la aduana que dice y los que esten en abono los restas al
    gasto". 'Cargo' = dispersión de dinero a la tarjeta de una aduana (gasto
    real de esa aduana, columna 'Centro de costo'); 'Abono' con Centro de
    costo lleno = devolución desde esa tarjeta (se resta del gasto); 'Abono'
    SIN Centro de costo = carga de fondos al monedero general (no es gasto
    de ninguna aduana, se ignora). La hoja trae una fila de total embebida
    al final, sin fecha/concepto/centro de costo -- se excluye con el mismo
    criterio, mismo patrón que [[utilidad-cliente-suma-total-duplica]]."""
    try:
        df = pd.read_excel(path, sheet_name=mes_txt.upper(), header=0)
    except ValueError:
        return pd.DataFrame(columns=["contraparte", "descripcion", "importe", "aduana",
                                     "tipo_cc", "concepto_cierre", "cancelada", "tipo_factura"])
    df.columns = [str(c).strip() for c in df.columns]
    cc = coalesce_col(df, "Centro de costo", "CENTRO DE COSTO", "Centro de costos").map(norm_txt)
    cargo = pd.to_numeric(coalesce_col(df, "Cargo"), errors="coerce")
    abono = pd.to_numeric(coalesce_col(df, "Abono"), errors="coerce")
    concepto = coalesce_col(df, "Concepto", "Tarjeta").map(norm_txt)
    es_fila_total = cc.isna() & concepto.isna()
    cargos = pd.DataFrame({"aduana": cc, "importe": cargo})[cargo.notna() & cc.notna() & ~es_fila_total]
    devoluciones = pd.DataFrame({"aduana": cc, "importe": -abono})[abono.notna() & cc.notna() & ~es_fila_total]
    d = pd.concat([cargos, devoluciones], ignore_index=True)
    out = pd.DataFrame(index=d.index)
    out["contraparte"] = "Efectivale"
    out["descripcion"] = "Tarjeta prepagada"
    out["importe"] = d["importe"]
    out["aduana"] = d["aduana"].map(_normaliza_aduana)
    out["tipo_cc"] = "PROPIOS"
    out["concepto_cierre"] = None
    out["cancelada"] = False
    out["tipo_factura"] = "RECIBIDA"
    return out[out["importe"] != 0]

@st.cache_data(ttl=1800, show_spinner=False)
def _cargar_emitidas_periodo(periodo):
    """Solo la hoja EMITIDAS de un periodo, con sus reglas de desfase de
    IMILE AIFA/TMM ya aplicadas -- separado de load_emitidas_recibidas
    (2026-11-18) para poder reusarse en historial_ingresos sin tener que
    descargar también Corresponsalías/Efectivale (que solo aplican a
    Recibidas/Gastos)."""
    cfg = PERIODOS[periodo]
    path_mes = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas"]))
    path_sig = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas_siguiente"]))
    hoja, tipo = "EMITIDAS", "EMITIDA"
    d_mes = _leer_hoja_factura(path_mes, hoja, tipo)
    # Imile AIFA y TMM en EMITIDAS siguen viniendo un mes atrasadas -- se
    # excluyen del mes actual y se toman del mes siguiente (TMM confirmado
    # con el usuario, 2026-11-05: "las emitidas que estan en agosto le
    # pertenecen a julio y las de septiembre a agosto" -- misma regla que
    # ya tenia Imile AIFA). Imile Laredo YA NO tiene ese desfase, se toma
    # tal cual venga en el archivo del mes actual (confirmado con el
    # usuario, 2026-10-30). El desfase de Imile AIFA no es uniforme mes a
    # mes -- algunos periodos tienen su propia excepción (ver
    # ADUANAS_SIN_DESFASE_POR_PERIODO / imile_monto2_desde /
    # imile_monto1_desde).
    desfase_este_periodo = ADUANAS_DESFASE_EMITIDAS - ADUANAS_SIN_DESFASE_POR_PERIODO.get(periodo, set())
    d_mes = d_mes[~d_mes["aduana"].isin(desfase_este_periodo)]
    d_sig = _leer_hoja_factura(path_sig, hoja, tipo)
    # Si este periodo tiene su propia fuente de IMILE desfasada 2 meses
    # (Monto 2 o Monto 1, ver PERIODOS), IMILE ya NO se toma del desfase
    # normal de 1 mes -- se reemplaza por esa fuente más abajo, para no
    # duplicarlo.
    tiene_fuente_imile_2meses = bool(cfg.get("imile_monto2_desde") or cfg.get("imile_monto1_desde"))
    desfase_normal = desfase_este_periodo - (
        {"IMILE AIFA"} if tiene_fuente_imile_2meses else set())
    d_sig = d_sig[d_sig["aduana"].isin(desfase_normal)]
    partes = [d_mes, d_sig]
    if cfg.get("imile_monto2_desde"):
        path_imile2 = BytesIO(_descargar_sheet_drive(cfg["imile_monto2_desde"]))
        partes.append(_leer_imile_emitidas_monto2(path_imile2, hoja))
    if cfg.get("imile_monto1_desde"):
        path_imile1 = BytesIO(_descargar_sheet_drive(cfg["imile_monto1_desde"]))
        d_imile1 = _leer_hoja_factura(path_imile1, hoja, tipo)
        partes.append(d_imile1[d_imile1["aduana"] == "IMILE AIFA"])
    d = pd.concat(partes, ignore_index=True)
    d = d[~d["cancelada"]]
    # A pedido del usuario (2026-10-28): "solo PROPIOS".
    d = d[d["tipo_cc"] == "PROPIOS"]
    return d.reset_index(drop=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _cargar_recibidas_periodo(periodo):
    """Solo la hoja RECIBIDAS de un periodo -- separado de
    load_emitidas_recibidas (2026-11-18, mismo motivo que
    _cargar_emitidas_periodo) para poder reusarse en historial_gastos."""
    cfg = PERIODOS[periodo]
    hoja, tipo = "RECIBIDAS", "RECIBIDA"
    path_mes = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas"]))
    path_sig = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas_siguiente"]))
    d_mes = _leer_hoja_factura(path_mes, hoja, tipo)
    # Las RECIBIDAS de Imile AIFA/Imile Laredo SI vienen correctas en el
    # archivo del mes actual -- a diferencia de Emitidas, aqui NO se
    # sustituyen desde el mes siguiente (confirmado con el usuario). KYA TI
    # y Yatec si se completan con lo que caiga en el 'MONTO 2' del archivo
    # del mes siguiente, que les pertenece a este periodo.
    d_kya_sig = _leer_kya_yatec_monto2_mes_anterior(path_sig, hoja)
    path_corr = BytesIO(_descargar_sheet_drive(cfg["corresponsalias"]))
    d_corr = _leer_corresponsalias(path_corr, cfg["corresponsalias_hoja"])
    path_efect = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["efectivale"]))
    d_efect = _leer_efectivale(path_efect, cfg["mes_txt"])
    d = pd.concat([d_mes, d_kya_sig, d_corr, d_efect], ignore_index=True)
    d = d[~d["cancelada"]]
    # A pedido del usuario (2026-10-28): "solo PROPIOS" -- aplica tanto en
    # Emitidas como en Recibidas (corrigio la regla anterior, que dejaba
    # TERCEROS/ACTIVO FIJO en Recibidas).
    d = d[d["tipo_cc"] == "PROPIOS"]
    return d.reset_index(drop=True)

@st.cache_data(ttl=1800, show_spinner="Cargando facturación...")
def load_emitidas_recibidas(periodo):
    """'periodo' es una clave de PERIODOS (ej. 'agosto', 'julio') -- mismas
    reglas para cualquier mes, solo cambian los archivos de origen (ver
    PERIODOS). Parametrizado 2026-10-30 a pedido del usuario ("vamos a
    cargar julio con las mismas reglas que se uso para agosto")."""
    return _cargar_emitidas_periodo(periodo), _cargar_recibidas_periodo(periodo)

def _mes_txt_completo(mes_num):
    cfg = next((c for c in PERIODOS.values() if c["mes_num"] == mes_num), None)
    return cfg["mes_txt"] if cfg else MES_COL_MAYUS[mes_num].capitalize()

@st.cache_data(ttl=1800, show_spinner=False)
def historial_ingresos(aduanas_filtro=None):
    """Ingreso real (Emitidas + Honorarios AA) vs. Presupuesto de TODOS los
    periodos ya cargados (ver PERIODOS), no solo el mes seleccionado --
    pedido del usuario, 2026-11-18: 'sobre esta grafica de tendencia
    quiero que se vean los demas meses, independientemente del año por el
    que estamos filtrando'. Respeta el filtro de aduana(s) del header --
    'aduanas_filtro' es None/vacio (todas) o una tupla de una o varias
    aduanas (pedido del usuario, 2026-11-25: poder elegir varias a la
    vez). Reusa _cargar_emitidas_periodo (cacheada) y los archivos de
    Honorarios AA/Presupuesto (un solo archivo para todos los meses,
    también cacheado) -- no vuelve a descargar nada de Recibidas/Gastos.
    El Presupuesto SÍ llega hasta Diciembre aunque el mes todavía no tenga
    datos reales cargados (pedido del usuario, 2026-11-20: "el presupuesto
    ya puede llegar a diciembre") -- PRESUPUESTOS 2026 ya trae los 12
    meses reales del año, independiente de qué meses de Emitidas/Recibidas
    estén cargados. 'real' se deja en None para esos meses sin datos; el
    frontend ya corta la línea ahí (Recharts no conecta huecos null)."""
    mes_num_a_clave = {cfg["mes_num"]: clave for clave, cfg in PERIODOS.items()}
    filas = []
    for mes_num in range(1, 13):
        ppto_mensual, _, _ = load_presupuesto(mes_num)
        ppto = _ppto_total(ppto_mensual, "ingresos", aduanas_filtro)
        if mes_num in mes_num_a_clave:
            emitidas = _cargar_emitidas_periodo(mes_num_a_clave[mes_num])
            honor = load_honorarios_aa(mes_num)
            if aduanas_filtro:
                emitidas = emitidas[emitidas["aduana"].isin(aduanas_filtro)]
                honor = honor[honor["aduana"].isin(aduanas_filtro)]
            real = float(emitidas["importe"].sum()) + float(honor["monto"].sum())
        else:
            real = None
        filas.append({"mes": _mes_txt_completo(mes_num), "real": real, "ppto": ppto})
    return filas

@st.cache_data(ttl=1800, show_spinner=False)
def historial_gastos(aduanas_filtro=None):
    """Mismo concepto que historial_ingresos pero para 'Tendencia de
    gasto' (pedido del usuario, 2026-11-18, "igual con este" sobre la
    misma gráfica pero de Gastos): Gasto real (Recibidas + Costos de
    Referencia + Sueldos) vs. Presupuesto en pesos de TODOS los periodos
    ya cargados, respetando el filtro de una o varias aduanas
    ('aduanas_filtro', 2026-11-25). 'variacion' se deja calculada por si
    se necesita en otro lado, pero ya no es lo que se grafica -- el
    usuario pidió (2026-11-19) que el eje muestre cifras en pesos, igual
    que Ingresos, no porcentaje. El Presupuesto llega hasta Diciembre
    igual que en historial_ingresos (2026-11-20), con 'real'/'variacion'
    en None para meses sin datos."""
    mes_num_a_clave = {cfg["mes_num"]: clave for clave, cfg in PERIODOS.items()}
    filas = []
    for mes_num in range(1, 13):
        ppto_mensual, _, _ = load_presupuesto(mes_num)
        ppto = _ppto_total(ppto_mensual, "gastos", aduanas_filtro)
        if mes_num in mes_num_a_clave:
            clave = mes_num_a_clave[mes_num]
            recibidas = _cargar_recibidas_periodo(clave)
            # El gasto real no es solo Recibidas -- tambien Costos de
            # Referencia y Sueldos de ese mes (pedido del usuario,
            # 2026-11-24: "la pestana de gasto no esta sumando todo el
            # gasto"). Corresponsalias/Efectivale ya vienen dentro de
            # Recibidas (ver _cargar_recibidas_periodo).
            costos = load_costos_referencia(clave)
            sueldos = _construir_sueldos_por_aduana(mes_num)
            if aduanas_filtro:
                recibidas = recibidas[recibidas["aduana"].isin(aduanas_filtro)]
                costos = costos[costos["aduana"].isin(aduanas_filtro)]
                sueldos = sueldos[sueldos["aduana"].isin(aduanas_filtro)]
            real = (float(recibidas["importe"].sum()) + float(costos["monto"].sum())
                   + float(sueldos["total"].sum()))
            variacion = (real / ppto - 1) if ppto else 0.0
        else:
            real = None
            variacion = None
        filas.append({"mes": _mes_txt_completo(mes_num), "real": real, "ppto": ppto, "variacion": variacion})
    return filas

@st.cache_data(ttl=1800, show_spinner=False)
def _cargar_sueldos_acumulado_crudo():
    """Lee la pestaña 'ACUMULADO' (los 12 meses juntos) una sola vez --
    antes load_sueldos(mes_num) volvía a descargar y re-parsear esta
    MISMA pestaña por cada mes distinto que se le pidiera; con
    historial_ingresos/historial_gastos pidiendo los 12 meses en cada
    corrida, eso eran 12 parseos completos en vez de 1 (pedido del
    usuario, 2026-11-25: "quiero optimizar todos los tiempos de
    espera")."""
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["sueldos"]))
    df = pd.read_excel(path, sheet_name="ACUMULADO", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data(ttl=1800, show_spinner=False)
def load_sueldos(mes_num):
    df = _cargar_sueldos_acumulado_crudo()
    df = df[df["MES NO."] == mes_num].copy()
    out = pd.DataFrame(index=df.index)
    out["nombre"] = coalesce_col(df, "nombre corregido", "NOMBRE").map(norm_txt)
    out["aduana"] = coalesce_col(df, "Departamento").map(norm_txt).map(
        _normaliza_aduana)
    for col_src, _ in COLS_NOMINA:
        out[col_src] = pd.to_numeric(coalesce_col(df, col_src), errors="coerce").fillna(0.0)
    out["total"] = pd.to_numeric(coalesce_col(df, "TOTAL2", "Total"), errors="coerce").fillna(0.0)
    out = out[out["nombre"].notna() & (out["total"] != 0)]
    return out.reset_index(drop=True)

BUCKETS_PERSONAL_CC = ["IMILE AIFA", "IMILE LAREDO", "TMM", "MULTIADUANAS", "ESTAFETA", "OTROS"]

# "MULTIADUANAS" no es una aduana real -- el usuario confirmo (2026-11-25,
# tras ver que el Cierre viejo traia notas sueltas como "40% en toluca" /
# "Toluca" en varias personas con fraccion en este cubo) que se trate como
# sinonimo de la aduana real TOLUCA, para todos los casos.
ALIAS_BUCKET_PERSONAL_CC = {"MULTIADUANAS": "TOLUCA"}

@st.cache_data(ttl=1800, show_spinner=False)
def load_personal_cc():
    """Hoja 'Personal por CC' compartida por el usuario (2026-11-25):
    reemplaza al Cierre Carga Aerea como fuente del reparto de Sueldos por
    aduana (ver _construir_sueldos_por_aduana) -- una fila por empleado con
    su fraccion fija en cada uno de los 6 "cubos" (BUCKETS_PERSONAL_CC), o
    2 filas (mismo nombre) cuando el reparto depende de si la aduana TMM
    tuvo ingresos ese mes ('SI INGRESOS TMM > 0' / '<= 0', confirmado con
    el usuario que la fuente es Emitidas de TMM). Las filas en escenario
    'PENDIENTE' o 'NO APLICA' ya traen 0.0 en las 6 fracciones -- no hace
    falta un caso especial, _fracciones_personal_cc las deja sin repartir
    solas. Un escenario que rota (ej. 'PERIODO A/B') no se puede resolver
    sin un calendario que el usuario aun no tiene -- se deja pendiente
    tambien (2026-11-25: "dejalas como pendientes, y los voy
    actualizando")."""
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["personal_por_cc"]))
    df = pd.read_excel(path, sheet_name="Personal por CC", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    reglas = {}
    for _, r in df.iterrows():
        nombre_norm = _normaliza_nombre(r.get("EMPLEADO"))
        if not nombre_norm:
            continue
        escenario = str(r.get("ESCENARIO / CONDICION") or "").strip().upper()
        fracciones = {b: float(pd.to_numeric(r.get(b), errors="coerce") or 0.0) for b in BUCKETS_PERSONAL_CC}
        reglas.setdefault(nombre_norm, []).append({"escenario": escenario, "fracciones": fracciones})
    return reglas

def _fracciones_personal_cc(nombre_norm, reglas, tmm_tiene_ingresos):
    """Regresa el dict de fracciones a aplicar, o None si Personal por CC
    no trae una regla USABLE para esta persona: no esta en la hoja, su
    unica fila es Pendiente/No aplica (fracciones en 0), o es un escenario
    condicional que no sabemos resolver (ej. una rotacion sin calendario
    confirmado) -- en cualquiera de esos casos el llamador cae a usar el
    Departamento real como respaldo."""
    filas = reglas.get(nombre_norm)
    if not filas:
        return None
    if len(filas) == 1:
        fracciones = filas[0]["fracciones"]
    else:
        fracciones = None
        for fila in filas:
            if "TMM > 0" in fila["escenario"] and tmm_tiene_ingresos:
                fracciones = fila["fracciones"]
            elif "TMM <= 0" in fila["escenario"] and not tmm_tiene_ingresos:
                fracciones = fila["fracciones"]
    if not fracciones or not any(v > 0 for v in fracciones.values()):
        return None
    return fracciones

@st.cache_data(ttl=1800, show_spinner=False)
def _aduanas_reales_conocidas():
    """Union de aduanas que aparecen en Emitidas/Recibidas de TODOS los
    periodos ya cargados en PERIODOS -- se usan como RESPALDO del
    Departamento de Sueldos solo cuando Personal por CC no trae una regla
    usable para esa persona (el usuario confirmo que Personal por CC manda
    siempre que la tenga, 2026-11-25, tras ver que 17 de 65 personas
    revisadas tienen un Departamento real valido que no coincide con su
    regla -- ej. 21 personas marcadas 'TOLUCA' en nomina que Personal por
    CC reparte entre otras aduanas). Es un dato estable de la empresa (los
    nombres de aduana no cambian de un mes a otro) -- por eso NO hace falta
    que el mes que se esta calculando ya tenga su propio archivo de
    Emitidas/Recibidas en PERIODOS: esto deja que Sueldos calcule un mes
    que ACUMULADO ya trae pero que el tablero todavia no integra (ej.
    Septiembre antes de tener Corresponsalias listo, 2026-11-25: 'pero
    como tal ya me puedes dar numeros de sueldos de septiembre?')."""
    aduanas = set()
    for clave in PERIODOS:
        em = _cargar_emitidas_periodo(clave)
        rec = _cargar_recibidas_periodo(clave)
        aduanas |= set(em["aduana"].dropna()) | set(rec["aduana"].dropna())
    return aduanas

def _tmm_tiene_ingresos_del_mes(mes_num):
    """Si la aduana TMM tuvo Emitidas > 0 ese mes -- resuelve el escenario
    condicional 'SI INGRESOS TMM > 0/<= 0' de Personal por CC. A diferencia
    de _aduanas_reales_conocidas, esto SI es un dato propio de cada mes;
    si ese mes todavia no tiene su periodo en PERIODOS (ej. Septiembre
    recien agregado a Sueldos pero sin Emitidas propia todavia) se asume
    'sin ingresos' por defecto, a falta de un dato mejor -- se actualiza
    solo en cuanto el periodo se agregue a PERIODOS."""
    mes_num_a_clave = {cfg["mes_num"]: clave for clave, cfg in PERIODOS.items()}
    clave = mes_num_a_clave.get(mes_num)
    if not clave:
        return False
    em = _cargar_emitidas_periodo(clave)
    return float(em.loc[em["aduana"] == "TMM", "importe"].sum()) > 0

@st.cache_data(ttl=1800, show_spinner=False)
def _construir_sueldos_por_aduana(mes_num):
    """Reparte Sueldos por aduana sin depender de que alguien lo llene a
    mano en un Cierre (reemplaza _construir_sueldos_desde_cierre /
    load_cierre_sueldos, 2026-11-25). Prioridad confirmada con el usuario
    tras encontrar 17 casos de conflicto en agosto (ver
    _aduanas_reales_conocidas):
      1. Si Personal por CC trae una regla usable para la persona (fraccion
         > 0 en algun cubo), esa manda SIEMPRE -- aunque el Departamento de
         Sueldos ya sea una aduana real distinta.
      2. Si no, se usa el Departamento de Sueldos tal cual, siempre que ya
         sea una aduana real de Emitidas/Recibidas de ese mes (esto es lo
         que captura que alguien cambie de aduana de un mes a otro sin
         tocar Personal por CC).
      3. Si ninguna de las dos aplica (no esta en la hoja y su
         Departamento tampoco es una aduana real), se deja bajo "SIN
         ADUANA" -- no se pierde del Total de Nomina pero se nota en el
         treemap que le falta su regla de reparto.
    Los 9 componentes de percepcion se escalan proporcionalmente a cada
    fraccion aplicada. Cacheada por mes_num (no recibe DataFrames como
    argumento) para que cambiar el filtro de Aduana en el header no
    dispare de nuevo este cruce fila por fila en Python (reportado por el
    usuario, 2026-11-24: 'cuando cambio por aduana se tarda demasiado')."""
    sueldos_f = load_sueldos(mes_num)
    if sueldos_f.empty:
        return sueldos_f
    cols_num = [c for c, _ in COLS_NOMINA]
    base = sueldos_f.copy()
    base["nombre_norm"] = base["nombre"].map(_normaliza_nombre)
    agg = {c: "sum" for c in cols_num}
    agg["total"] = "sum"
    agg["aduana"] = "first"
    resumen = base.groupby("nombre_norm").agg(agg)

    aduanas_reales = _aduanas_reales_conocidas()
    tmm_tiene_ingresos = _tmm_tiene_ingresos_del_mes(mes_num)
    reglas_cc = load_personal_cc()

    filas = []
    for nombre_norm, fila_base in resumen.iterrows():
        fracciones = _fracciones_personal_cc(nombre_norm, reglas_cc, tmm_tiene_ingresos)
        if fracciones:
            for bucket, fraccion in fracciones.items():
                if fraccion <= 0:
                    continue
                aduana_destino = ALIAS_BUCKET_PERSONAL_CC.get(bucket, bucket)
                fila = {"nombre": nombre_norm, "aduana": aduana_destino, "total": fila_base["total"] * fraccion}
                for c in cols_num:
                    fila[c] = fila_base[c] * fraccion
                filas.append(fila)
            continue
        aduana = fila_base["aduana"] if fila_base["aduana"] in aduanas_reales else "SIN ADUANA"
        fila = {"nombre": nombre_norm, "aduana": aduana, "total": fila_base["total"]}
        for c in cols_num:
            fila[c] = fila_base[c]
        filas.append(fila)
    return pd.DataFrame(filas, columns=["nombre", "aduana", "total"] + cols_num).reset_index(drop=True)

@st.cache_data(ttl=1800, show_spinner=False)
def load_costos_referencia(periodo):
    """Pestaña real 'Reporte de Facturas'. La columna 'Aduana' trae un código
    numérico que ni siquiera es único por aduana (confirmado con el usuario)
    -- se usa la columna 'CC', que sí trae el nombre correcto por fila. El
    monto YA NO sale de 'Total Conceptos AA' (se quedó en $0 para la mayoría
    de las filas nuevas) -- el usuario confirmó (2026-10-30) que ahora hay
    que tomar la columna con encabezado '0.2' (letra de Excel BJ, junto a
    otras dos '0.4'/'0.3' que parecen un reparto en 3 partes de 'TOTAL
    BASE'). La hoja trae una fila de TOTAL GENERAL embebida al final
    ('Clave Cliente' = 'TOTAL GENERAL') que hay que excluir, si no duplica
    la suma (mismo patrón que [[utilidad-cliente-suma-total-duplica]]).
    Archivo propio por periodo (ver PERIODOS)."""
    path = BytesIO(_descargar_sheet_drive(PERIODOS[periodo]["costos_de_referencia"]))
    df = pd.read_excel(path, sheet_name="Reporte de Facturas", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    es_total = coalesce_col(df, "Clave Cliente").astype(str).str.upper().str.contains("TOTAL", na=False)
    df = df[~es_total]
    out = pd.DataFrame(index=df.index)
    out["aduana"] = coalesce_col(df, "CC").map(norm_txt).map(
        _normaliza_aduana)
    out["monto"] = pd.to_numeric(coalesce_col(df, "0.2"), errors="coerce").fillna(0.0)
    return out[out["monto"] != 0].reset_index(drop=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _cargar_honorarios_aa_crudo():
    """Lee la pestaña 'Detalle' completa (los 12 meses juntos) una sola
    vez -- mismo motivo que _cargar_sueldos_acumulado_crudo (pedido del
    usuario, 2026-11-25: optimizar tiempos de espera)."""
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["honorarios_aa"]))
    df = pd.read_excel(path, sheet_name="Detalle", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data(ttl=1800, show_spinner=False)
def load_honorarios_aa(mes_num):
    """Pestaña 'Detalle' -- ya trae el nombre de la aduana directo (columna
    'NOMBRE ADUANA', no un código). Se filtra a mes_num (columna 'Mes'
    viene mezclada int/str en el archivo, por eso el cast a str antes de
    comparar). Solo se cuentan facturas con Estatus 'Timbrada'; las demás
    (Emitida, Cancelada, Cobrada) no se consideran honorario devengado."""
    df = _cargar_honorarios_aa_crudo()
    df = df[df["Mes"].astype(str) == str(mes_num)].copy()
    estatus = coalesce_col(df, "Estatus").map(norm_txt)
    df = df[estatus.map(lambda s: isinstance(s, str) and s.strip().upper() == "TIMBRADA")]
    out = pd.DataFrame(index=df.index)
    out["aduana"] = coalesce_col(df, "NOMBRE ADUANA").map(norm_txt).map(
        _normaliza_aduana)
    out["monto"] = pd.to_numeric(coalesce_col(df, "HONORARIOS DE AGENTE ADUANAL"), errors="coerce").fillna(0.0)
    return out[out["monto"] != 0].reset_index(drop=True)

# PRESUPUESTOS 2026 nombra varias aduanas distinto a como las llama
# Emitidas/Recibidas (mismo alias que ya se usaba antes del reset para esta
# misma reconciliación) -- sin esto, Rentabilidad por Aduana no encuentra
# con qué cruzar la marca de presupuesto para esas aduanas.
ALIAS_ADUANA_PPTO = {
    "AMAZON 1 (JUSTO)": "AMAZON 1", "AMAZON 2 (CARGA)": "AMAZON 2",
    "LAZARO CARDENAS": "LAZARO", "TOLUCA BSO": "TOLUCA", "IMILE": "IMILE AIFA",
}

def _leer_bloques_presupuesto_anual(ws, meses_idx):
    """PRESUPUESTOS 2026 trae 8 bloques (4 conceptos x mensual/acumulado, una
    columna por mes ENERO..DICIEMBRE) en vez de un archivo por mes como
    antes. El 2o bloque de cada concepto es el acumulado, pero el archivo NO
    siempre lo marca con la palabra 'ACUMULADO' en la columna A (INGRESOS y
    GASTOS OP sí, CTO REF y SUELDOS repiten su propio nombre) -- se detecta
    por posición (2a aparición de cada concepto), no por el texto.
    'meses_idx' es una lista (se piden varios meses de una pasada, ej. el
    mes actual Y diciembre -- diciembre del bloque acumulado da el total
    real del presupuesto anual completo, sin necesidad de proyectarlo).
    Regresa {mes_idx: (mensual, acumulado)}, cada uno {concepto: {aduana:
    valor}} -- por aduana, no un total plano, para poder cruzarlo con el
    real en la gráfica de Rentabilidad por Aduana."""
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    resultado = {m: ({}, {}) for m in meses_idx}
    vistos = {}
    ultimo_concepto = None
    nombres_mes = {m: MES_COL_MAYUS[m] for m in meses_idx}
    i = 0
    while i < len(rows):
        r = rows[i]
        if not r or r[0] is None:
            i += 1
            continue
        a0 = str(r[0]).strip().upper()
        concepto = None
        if a0 in CONCEPTO_PRESUPUESTO:
            concepto = CONCEPTO_PRESUPUESTO[a0]
            ultimo_concepto = concepto
        elif a0 == "ACUMULADO" and ultimo_concepto is not None:
            concepto = ultimo_concepto
        if concepto is not None and len(r) > 1 and isinstance(r[1], str):
            veces = vistos.get(concepto, 0)
            es_acum = veces >= 1
            vistos[concepto] = veces + 1
            cols_mes = {m: next((j for j, v in enumerate(r)
                                 if isinstance(v, str) and v.strip().upper() == nombre), None)
                       for m, nombre in nombres_mes.items()}
            datos = {m: {} for m in meses_idx}
            i += 1
            while i < len(rows):
                r2 = rows[i]
                if not r2 or r2[0] is None:
                    break
                a = str(r2[0]).strip().upper()
                if a.startswith("TOTAL"):
                    i += 1
                    break
                aduana = ALIAS_ADUANA_PPTO.get(a, a)
                for m, col_mes in cols_mes.items():
                    val = r2[col_mes] if col_mes is not None and col_mes < len(r2) else None
                    if isinstance(val, (int, float)):
                        datos[m][aduana] = datos[m].get(aduana, 0.0) + val
                i += 1
            for m in meses_idx:
                (resultado[m][1] if es_acum else resultado[m][0])[concepto] = datos[m]
            continue
        i += 1
    return resultado

@st.cache_data(ttl=1800, show_spinner=False)
def _cargar_presupuesto_todos_los_meses():
    """Parsea el workbook de Presupuestos completo UNA sola vez, para los
    12 meses a la vez -- antes load_presupuesto(mes_num) reabría y
    reparseaba el workbook entero (con openpyxl, el motor más lento de
    los dos) por cada mes distinto que se le pidiera; historial_ingresos/
    historial_gastos piden los 12 meses en cada corrida, eran 12 parseos
    completos del mismo archivo en vez de 1 (pedido del usuario,
    2026-11-25: "quiero optimizar todos los tiempos de espera")."""
    import openpyxl
    contenido = _descargar_sheet_drive(DRIVE_SHEET_IDS["presupuestos_2026"])
    wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True)
    return _leer_bloques_presupuesto_anual(wb[wb.sheetnames[0]], list(range(1, 13)))

@st.cache_data(ttl=1800, show_spinner=False)
def load_presupuesto(mes_num):
    """Regresa (mensual, acumulado, ppto_anual) del mes dado -- cada uno
    {concepto: {aduana: valor}}. 'ppto_anual' es el acumulado a DICIEMBRE
    (el total real del año completo según el presupuesto, ya no una
    proyección como antes del reset)."""
    resultado = _cargar_presupuesto_todos_los_meses()
    mensual, acumulado = resultado[mes_num]
    _, ppto_anual = resultado[12]
    return mensual, acumulado, ppto_anual

def _ppto_total(d, concepto, aduanas=None):
    """Suma el total de un concepto del presupuesto. 'aduanas' puede ser
    None (todas), una sola aduana (str) o una lista/tupla de aduanas --
    desde que el filtro de aduana del header admite varias a la vez
    (pedido del usuario, 2026-11-25), tanto Utilidad Operativa como
    Ingresos/Gastos lo respetan."""
    valores = d.get(concepto, {})
    if aduanas is None:
        return float(sum(valores.values()))
    if isinstance(aduanas, str):
        return float(valores.get(aduanas, 0.0))
    return float(sum(valores.get(a, 0.0) for a in aduanas))

import re
UUID_RE = re.compile(r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")

def _extraer_uuids(celda):
    return UUID_RE.findall(celda) if isinstance(celda, str) else []

@st.cache_data(ttl=1800, show_spinner=False)
def load_amex(persona):
    col_uuid = AMEX_COL_UUID[persona]
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS[f"amex_{persona.lower()}"]))
    hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
    if "hoja 1" not in hojas:
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=hojas["hoja 1"], header=0)
    df.columns = [str(c).strip() for c in df.columns]
    out = pd.DataFrame(index=df.index)
    out["proveedor"] = coalesce_col(df, "Supplier Name").map(norm_txt).fillna("SIN PROVEEDOR")
    monto_neto = pd.to_numeric(coalesce_col(df, "Net Billed Amount"), errors="coerce")
    monto_cargo = pd.to_numeric(coalesce_col(df, "Charge Amount"), errors="coerce").fillna(0.0)
    out["monto"] = monto_neto.fillna(monto_cargo)
    out["uuids"] = coalesce_col(df, col_uuid).map(_extraer_uuids)
    out["facturado"] = out["uuids"].map(len) > 0
    return out.reset_index(drop=True)

# Agregadores de RANGO de periodos -- antes solo se cargaba un periodo a
# la vez; el usuario pidió (2026-11-25) poder elegir un rango (ej. "Mayo a
# Agosto") y que los totales de las 5 pestañas sean la SUMA de esos
# meses. Cada uno reusa los loaders de un solo periodo/mes (ya cacheados
# por su cuenta) y solo concatena/suma -- así cambiar el rango no vuelve a
# descargar nada que ya estuviera en caché. Se cachean también por su
# cuenta, por la tupla de periodos/meses, para no repetir el concat/suma
# en cada rerun si el rango no cambió.
@st.cache_data(ttl=1800, show_spinner="Cargando facturación...")
def _cargar_rango(claves_periodos):
    ems, recs = [], []
    for clave in claves_periodos:
        em, rec = load_emitidas_recibidas(clave)
        ems.append(em)
        recs.append(rec)
    return pd.concat(ems, ignore_index=True), pd.concat(recs, ignore_index=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _sueldos_rango(meses_num):
    return pd.concat([_construir_sueldos_por_aduana(m) for m in meses_num], ignore_index=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _costos_rango(claves_periodos):
    return pd.concat([load_costos_referencia(c) for c in claves_periodos], ignore_index=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _honorarios_rango(meses_num):
    return pd.concat([load_honorarios_aa(m) for m in meses_num], ignore_index=True)

@st.cache_data(ttl=1800, show_spinner=False)
def _presupuesto_rango(meses_num):
    """Suma el presupuesto MENSUAL de cada mes del rango, concepto por
    concepto y aduana por aduana. El 'acumulado' (desde Enero) y el
    'ppto_anual' (meta del año completo) NUNCA dependen del rango
    elegido -- se piden aparte con load_presupuesto(mes del 'hasta')."""
    total = {}
    for m in meses_num:
        mensual, _, _ = load_presupuesto(m)
        for concepto, por_aduana in mensual.items():
            destino = total.setdefault(concepto, {})
            for aduana, monto in por_aduana.items():
                destino[aduana] = destino.get(aduana, 0.0) + monto
    return total

# Selector de RANGO de periodos (antes un solo periodo -- pedido del
# usuario, 2026-10-30: "vamos a cargar julio con las mismas reglas que se
# uso para agosto"; ahora "Desde"/"Hasta", 2026-11-25) -- mismo mecanismo
# que el filtro de aduana (pills en el header + session_state).
CLAVES_PERIODOS_ORDENADAS = [k for k, _ in sorted(PERIODOS.items(), key=lambda kv: kv[1]["mes_num"])]
if "periodo_desde" not in st.session_state:
    st.session_state["periodo_desde"] = PERIODO_DEFAULT
if "periodo_hasta" not in st.session_state:
    st.session_state["periodo_hasta"] = PERIODO_DEFAULT
periodo_desde = st.session_state["periodo_desde"]
periodo_hasta = st.session_state["periodo_hasta"]
if periodo_desde not in PERIODOS:
    periodo_desde = st.session_state["periodo_desde"] = PERIODO_DEFAULT
if periodo_hasta not in PERIODOS:
    periodo_hasta = st.session_state["periodo_hasta"] = PERIODO_DEFAULT
if PERIODOS[periodo_desde]["mes_num"] > PERIODOS[periodo_hasta]["mes_num"]:
    # "Desde" quedó después de "Hasta" (ej. cambiaron "Desde" a Agosto con
    # "Hasta" todavía en Mayo) -- se empareja para que el rango siempre
    # sea válido en vez de quedar vacío.
    periodo_desde = st.session_state["periodo_desde"] = periodo_hasta

RANGO_PERIODOS = tuple(k for k in CLAVES_PERIODOS_ORDENADAS
                       if PERIODOS[periodo_desde]["mes_num"] <= PERIODOS[k]["mes_num"] <= PERIODOS[periodo_hasta]["mes_num"])
RANGO_MESES_NUM = tuple(PERIODOS[k]["mes_num"] for k in RANGO_PERIODOS)
MES_ACTUAL = PERIODOS[periodo_hasta]["mes_num"]
RANGO_TXT = (PERIODOS[periodo_desde]["mes_txt"] if periodo_desde == periodo_hasta
            else f'{PERIODOS[periodo_desde]["mes_txt"]} - {PERIODOS[periodo_hasta]["mes_txt"]}')
MES_TXT = RANGO_TXT

try:
    # Dispara en paralelo TODAS las descargas de Drive que va a necesitar
    # esta corrida -- de TODOS los periodos definidos (no solo el rango
    # activo), porque historial_ingresos/historial_gastos siempre recorren
    # los 12 meses para la gráfica de tendencia, sin importar qué rango
    # esté filtrado. Más los compartidos entre todos los meses (Sueldos,
    # Honorarios AA, Presupuesto, Cierre Carga Aérea, Efectivale, los 3
    # AMEX). Por su cuenta cada loader de abajo ya los pide tal cual (y ya
    # quedan en caché de _descargar_sheet_drive), así que esto no cambia
    # ningún resultado -- solo adelanta la espera de red para que ya estén
    # calientes cuando cada loader los pida.
    _archivos_por_periodo = []
    for _cfg in PERIODOS.values():
        _archivos_por_periodo += [
            _cfg.get("emitidas_recibidas"), _cfg.get("emitidas_recibidas_siguiente"),
            _cfg.get("corresponsalias"), _cfg.get("costos_de_referencia"),
            _cfg.get("imile_monto2_desde"), _cfg.get("imile_monto1_desde"),
        ]
    _archivos_compartidos = [DRIVE_SHEET_IDS[k] for k in (
        "sueldos", "honorarios_aa", "presupuestos_2026", "personal_por_cc",
        "efectivale", "amex_alejandro", "amex_josefa", "amex_angelica")]
    _prefetch_drive(_archivos_por_periodo + _archivos_compartidos)

    EM, REC = _cargar_rango(RANGO_PERIODOS)
    sueldos_df = _sueldos_rango(RANGO_MESES_NUM)
    costos_ref_df = _costos_rango(RANGO_PERIODOS)
    honorarios_aa_df = _honorarios_rango(RANGO_MESES_NUM)
    ppto_mensual = _presupuesto_rango(RANGO_MESES_NUM)
    _, ppto_acumulado, ppto_anual = load_presupuesto(MES_ACTUAL)
    amex_alejandro_df = load_amex("ALEJANDRO")
    amex_josefa_df = load_amex("JOSEFA")
    amex_angelica_df = load_amex("ANGELICA")
except Exception as e:
    st.error(f"Error al descargar/leer los datos de Drive:\n\n{e}")
    st.stop()

TOT_EM = float(EM["importe"].sum())
TOT_REC = float(REC["importe"].sum())
MARGEN = TOT_EM - TOT_REC
MARGEN_PCT = MARGEN / TOT_EM if TOT_EM else None
cobertura_concepto = (REC.loc[REC["concepto_cierre"].notna(), "importe"].sum() / TOT_REC) if TOT_REC else 0.0

# ══════════════════════════════════════════════════════════════════════════════
# PAYLOADS para el componente React (mismo shape que ya esperaba el
# diseño -- ver components/facturacion_dashboard/). Utilidad Operativa se
# deja en None a propósito: esa página necesitaba fórmulas compuestas
# (Ingreso = Emitidas + Honorarios AA, Egreso = Recibida + Sueldos + Costos
# de Referencia, etc.) que eran justo de las reglas que se borraron -- se
# reconstruye cuando el usuario las vuelva a dar. Las otras 4 páginas son
# lectura directa de la fuente, sin fórmulas inventadas.
# ══════════════════════════════════════════════════════════════════════════════
def _flujo_payload(es_ingreso, base_df, margen, margen_pct, cobertura, aduana_filtro=None,
                   honorarios_df=None, costos_df=None, sueldos_df=None):
    if base_df.empty:
        return None
    ppto_key = "ingresos" if es_ingreso else "gastos"
    tiene_ppto = bool(ppto_mensual)
    honor_total = float(honorarios_df["monto"].sum()) if honorarios_df is not None else 0.0
    honor_ad = (honorarios_df.groupby("aduana")["monto"].sum()
               if honorarios_df is not None and not honorarios_df.empty else pd.Series(dtype=float))
    # El gasto total del mes no es solo 'base_df' (Recibidas) -- tambien
    # Costos de Referencia y Sueldos (Corresponsalias/Efectivale ya vienen
    # dentro de Recibidas, ver _cargar_recibidas_periodo). Antes faltaban
    # los dos y el total de Gastos salia incompleto (pedido del usuario,
    # 2026-11-24).
    costos_total = float(costos_df["monto"].sum()) if costos_df is not None else 0.0
    sueldos_total = float(sueldos_df["total"].sum()) if sueldos_df is not None else 0.0
    costos_ad = (costos_df.groupby("aduana")["monto"].sum()
                if costos_df is not None and not costos_df.empty else pd.Series(dtype=float))
    sueldos_ad = (sueldos_df.groupby("aduana")["total"].sum()
                 if sueldos_df is not None and not sueldos_df.empty else pd.Series(dtype=float))
    tot_mes = float(base_df["importe"].sum()) + honor_total + costos_total + sueldos_total
    ppto_mes_tot = _ppto_total(ppto_mensual, ppto_key, aduana_filtro) if tiene_ppto else None
    ppto_acum_tot = _ppto_total(ppto_acumulado, ppto_key, aduana_filtro) if tiene_ppto else None
    n_contrapartes = int(base_df.loc[base_df["importe"] != 0, "contraparte"].nunique())

    campo_concepto = "descripcion" if es_ingreso else "concepto_cierre"
    conc = base_df.copy()
    conc[campo_concepto] = conc[campo_concepto].fillna("Sin categorizar")
    por_concepto_full = conc.groupby(campo_concepto)["importe"].sum()
    if honor_total:
        por_concepto_full = pd.concat([por_concepto_full, pd.Series({"Honorarios AA": honor_total})])
    por_concepto_full = por_concepto_full[por_concepto_full > 0].sort_values(ascending=False)
    por_concepto_dona = por_concepto_full.head(6)
    if len(por_concepto_full) > 6:
        por_concepto_dona = pd.concat([por_concepto_dona, pd.Series({"Otros": por_concepto_full.iloc[6:].sum()})])

    por_contraparte_full = base_df.groupby("contraparte")["importe"].sum()
    por_contraparte_full = por_contraparte_full[por_contraparte_full > 0].sort_values(ascending=False)
    top10 = por_contraparte_full.head(10)
    if len(por_contraparte_full) > 10:
        top10 = pd.concat([top10, pd.Series({"Otros": por_contraparte_full.iloc[10:].sum()})])

    por_aduana = []
    aduanas_todas = sorted(set(base_df["aduana"].dropna().unique().tolist()) | set(honor_ad.index)
                          | set(costos_ad.index) | set(sueldos_ad.index))
    for aduana in aduanas_todas:
        base_aduana = base_df[base_df["aduana"] == aduana]
        honor_aduana = float(honor_ad.get(aduana, 0.0))
        if es_ingreso:
            tot_aduana = float(base_aduana["importe"].sum()) + honor_aduana
        else:
            cr_aduana = float(costos_ad.get(aduana, 0.0))
            sueldos_aduana = float(sueldos_ad.get(aduana, 0.0))
            tot_aduana = float(base_aduana["importe"].sum()) + cr_aduana + sueldos_aduana
        entry = {"aduana": str(aduana), "total": tot_aduana, "n_operaciones": int(len(base_aduana))}
        if es_ingreso:
            entry["emitidas"] = float(base_aduana["importe"].sum())
            entry["honorarios"] = honor_aduana
            por_desc = base_aduana.groupby("descripcion")["importe"].sum().sort_values(ascending=False)
            top3 = por_desc.head(3)
            otros_val = float(por_desc.iloc[3:].sum()) if len(por_desc) > 3 else 0.0
            entry["conceptos"] = [{"nombre": str(c), "pct": float(v / tot_aduana) if tot_aduana else 0.0}
                                  for c, v in top3.items()]
            if otros_val:
                entry["conceptos"].append({"nombre": "Otros conceptos",
                                           "pct": float(otros_val / tot_aduana) if tot_aduana else 0.0})
            if honor_aduana:
                entry["conceptos"].append({"nombre": "Honorarios AA",
                                           "pct": float(honor_aduana / tot_aduana) if tot_aduana else 0.0})
        else:
            entry["costos_referencia"] = cr_aduana
            entry["costos_referencia_pct"] = float(cr_aduana / tot_aduana) if tot_aduana else 0.0
            entry["sueldos"] = sueldos_aduana
        por_aduana.append(entry)

    cumplimiento = tot_mes / ppto_mes_tot if ppto_mes_tot else None
    return {
        "tiene_ppto": tiene_ppto, "mes_txt": MES_TXT,
        # Último mes del rango seleccionado -- el frontend lo usa para
        # saber hasta dónde acumular en 'historial' (antes 'mes_txt' era
        # un solo mes y servía para eso directo; ahora puede ser un rango
        # como "Mayo - Agosto" que no empata con ningún mes de 'historial'
        # tal cual, pedido del usuario, 2026-11-25).
        "mes_fin_txt": PERIODOS[periodo_hasta]["mes_txt"],
        "kpis": {"tot_mes": tot_mes, "ppto_mes_tot": ppto_mes_tot, "ppto_acum_tot": ppto_acum_tot,
                "n_contrapartes": n_contrapartes, "cumplimiento": cumplimiento,
                "margen": margen, "margen_pct": margen_pct,
                # "Emitidas propias" (Ingresos) / "Recibidas propias" (Gastos)
                # -- el monto base sin Honorarios/Costos/Sueldos, para el 2o.
                # % del cuadro de Principales clientes/proveedores (pedido
                # del usuario, 2026-11-24).
                "tot_propias": float(base_df["importe"].sum())},
        "por_aduana": por_aduana,
        "por_concepto": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_dona.items()],
        "por_concepto_full": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_full.items()],
        "por_contraparte": [{"nombre": str(c), "monto": float(v)} for c, v in top10.items()],
        "cobertura_concepto": (float(cobertura) if not es_ingreso else None),
        "historial": historial_ingresos(aduana_filtro) if es_ingreso else historial_gastos(aduana_filtro),
    }

def _construir_arbol_nomina(sueldos_f):
    por_aduana = sueldos_f.groupby("aduana")["total"].sum().sort_values(ascending=False)
    arbol = []
    for aduana, total_aduana in por_aduana.items():
        if total_aduana <= 0:
            continue
        tabla_ad = sueldos_f[sueldos_f["aduana"] == aduana]
        por_emp = tabla_ad.groupby("nombre")["total"].sum().sort_values(ascending=False)
        empleados = []
        for empleado, total_emp in por_emp.items():
            if total_emp <= 0:
                continue
            fila_sum = tabla_ad.loc[tabla_ad["nombre"] == empleado, [c for c, _ in COLS_NOMINA]].sum()
            componentes = [
                {"name": etq, "size": float(abs(fila_sum[c])), "etiqueta": fmt(fila_sum[c]),
                 "color": JOCS_AZUL if fila_sum[c] >= 0 else C_NEG}
                for c, etq in COLS_NOMINA if abs(fila_sum[c]) > 0
            ]
            empleados.append({"name": str(empleado), "size": float(total_emp),
                              "etiqueta": fmt(total_emp), "color": color_for(len(empleados)),
                              "componentes": componentes})
        arbol.append({"name": str(aduana), "size": float(total_aduana),
                      "etiqueta": fmt(total_aduana), "color": color_for(len(arbol)),
                      "empleados": empleados})
    return arbol

def _uo_payload(em_f, rec_f, sueldos_f, costos_f, honor_f, ppto_mensual, ppto_anual, aduanas_sel, mes_txt):
    """Utilidad Operativa por aduana -- regla confirmada con el usuario
    (2026-10-28): Ingreso = Emitidas + Honorarios AA; Utilidad = Ingreso -
    Sueldos - Gastos (Recibidas) - Costos de Referencia. Imile Laredo sale
    incompleta aqui (no tiene la hoja de P&L aparte que compensa que no
    factura por el SAT -- pendiente de conectar) -- se deja tal cual, no se
    inventa un ajuste.
    Ahora SÍ respeta el filtro de varias aduanas y el rango de periodos del
    header (pedido del usuario, 2026-11-25; antes nunca se filtraba a
    propósito) -- por eso recibe EM_F/REC_F/SUELDOS_F/COSTOS_F/HONOR_F ya
    filtrados por el script principal (los mismos que usan Ingresos/
    Gastos) en vez de volver a descargar todo sin filtrar. No está
    decorada con @st.cache_data: lo pesado (descargas, merge de Sueldos
    con el Cierre) ya está cacheado en los loaders de más abajo; esto solo
    agrupa DataFrames ya chicos, y de cualquier forma debe recalcularse
    cuando cambia el filtro ya que ahora el resultado depende de él."""
    tot_em = float(em_f["importe"].sum())
    tot_rec = float(rec_f["importe"].sum())

    aduanas = set(em_f["aduana"]) | set(rec_f["aduana"]) | set(sueldos_f["aduana"]) \
            | set(costos_f["aduana"]) | set(honor_f["aduana"])
    em_ad = em_f.groupby("aduana")["importe"].sum()
    rec_ad = rec_f.groupby("aduana")["importe"].sum()
    sueldos_ad = sueldos_f.groupby("aduana")["total"].sum()
    costos_ad = costos_f.groupby("aduana")["monto"].sum()
    honor_ad = honor_f.groupby("aduana")["monto"].sum()

    por_aduana = []
    for aduana in sorted(aduanas):
        ingreso = float(em_ad.get(aduana, 0.0)) + float(honor_ad.get(aduana, 0.0))
        sueldos_v = float(sueldos_ad.get(aduana, 0.0))
        gastos_v = float(rec_ad.get(aduana, 0.0))
        costos_v = float(costos_ad.get(aduana, 0.0))
        por_aduana.append({
            "aduana": str(aduana), "ingreso": ingreso, "costos_referencia": costos_v,
            "gastos_operativos": gastos_v, "sueldos": sueldos_v,
            "utilidad": ingreso - sueldos_v - gastos_v - costos_v,
        })

    tiene_ppto = bool(ppto_mensual)
    def _ppto_utilidad_total(d):
        return (_ppto_total(d, "ingresos", aduanas_sel or None) - _ppto_total(d, "costos", aduanas_sel or None)
               - _ppto_total(d, "sueldos", aduanas_sel or None) - _ppto_total(d, "gastos", aduanas_sel or None))
    ppto_por_aduana = []
    if tiene_ppto:
        aduanas_ppto = set()
        for concepto in ("ingresos", "costos", "sueldos", "gastos"):
            aduanas_ppto |= set(ppto_mensual.get(concepto, {}).keys())
        if aduanas_sel:
            aduanas_ppto &= set(aduanas_sel)
        for aduana in sorted(aduanas_ppto):
            ing = float(ppto_mensual.get("ingresos", {}).get(aduana, 0.0))
            cos = float(ppto_mensual.get("costos", {}).get(aduana, 0.0))
            sue = float(ppto_mensual.get("sueldos", {}).get(aduana, 0.0))
            gas = float(ppto_mensual.get("gastos", {}).get(aduana, 0.0))
            ppto_por_aduana.append({"aduana": aduana, "ingresos": ing, "costos": cos,
                                    "sueldos": sue, "gastos": gas,
                                    "utilidad": ing - cos - sue - gas})

    ingreso_total = tot_em + float(honor_ad.sum())
    sueldos_total = float(sueldos_ad.sum())
    gastos_total = tot_rec
    costos_total = float(costos_ad.sum())
    utilidad_total = ingreso_total - sueldos_total - gastos_total - costos_total

    def _concepto(real, clave_ppto):
        return {"real": real,
               "ppto": _ppto_total(ppto_mensual, clave_ppto, aduanas_sel or None) if tiene_ppto and clave_ppto else None,
               "ppto_anual": _ppto_total(ppto_anual, clave_ppto, aduanas_sel or None) if tiene_ppto and clave_ppto else None}

    estado_resultados = {
        "ingresos": _concepto(ingreso_total, "ingresos"),
        "costos_referencia": _concepto(costos_total, "costos"),
        "sueldos": _concepto(sueldos_total, "sueldos"),
        "gastos_operativos": _concepto(gastos_total, "gastos"),
        "utilidad": {"real": utilidad_total,
                    "ppto": _ppto_utilidad_total(ppto_mensual) if tiene_ppto else None,
                    "ppto_anual": _ppto_utilidad_total(ppto_anual) if tiene_ppto else None},
    }
    return {"tiene_ppto": tiene_ppto, "mes_txt": mes_txt, "por_aduana": por_aduana,
           "ppto_por_aduana": ppto_por_aduana, "estado_resultados": estado_resultados}

def _nomina_payload(sueldos_f, tot_rec):
    if sueldos_f.empty:
        return None
    tot_nomina = float(sueldos_f["total"].sum())
    pct_gasto_op = (tot_nomina / tot_rec) if tot_rec else None
    return {"mes_txt": MES_TXT, "tot_nomina": tot_nomina, "pct_gasto_op": pct_gasto_op,
           "arbol": _construir_arbol_nomina(sueldos_f)}

def _tarjeta_payload():
    resultado = []
    for nombre, amex_df in [("Alejandro", amex_alejandro_df), ("Josefa", amex_josefa_df),
                            ("Angélica", amex_angelica_df)]:
        if amex_df.empty:
            resultado.append({"nombre": nombre, "tiene_datos": False})
            continue
        total_gastado = float(amex_df["monto"].sum())
        movimientos = int(len(amex_df))
        facturados = int(amex_df["facturado"].sum())
        monto_facturado = float(amex_df.loc[amex_df["facturado"], "monto"].sum())
        monto_no_facturado = total_gastado - monto_facturado
        es_pase = amex_df["proveedor"].str.contains("PASE", case=False, na=False)
        proveedor_agrupado = amex_df["proveedor"].where(~es_pase, "PASE")
        por_prov = amex_df.groupby(proveedor_agrupado)["monto"].sum().sort_values(ascending=False)
        resultado.append({
            "nombre": nombre, "tiene_datos": True, "total_gastado": total_gastado,
            "movimientos": movimientos, "facturados": facturados,
            "no_facturados": movimientos - facturados, "monto_facturado": monto_facturado,
            "monto_no_facturado": monto_no_facturado,
            "pct_facturado": (monto_facturado / total_gastado) if total_gastado else 0.0,
            "por_proveedor": [{"nombre": str(k), "monto": float(v)} for k, v in por_prov.items()],
        })
    return resultado

# ══════════════════════════════════════════════════════════════════════════════
# UI -- componente React único, sin nada de Streamlit visible (sidebar,
# toolbar, etc). Ver [[feedback-respetar-referencias-visuales]].
# ══════════════════════════════════════════════════════════════════════════════
_dashboard_component = components.declare_component(
    "facturacion_dashboard", path=str(BASE / "components" / "facturacion_dashboard"))

components.html(
    "<script>try{var d=window.parent.document.documentElement;"
    "d.setAttribute('translate','no');d.classList.add('notranslate');}catch(e){}</script>",
    height=0,
)

NAV_OPCIONES = ["Utilidad Operativa", "Ingresos", "Gastos", "Nómina", "Tarjeta de Crédito"]
if "nav" not in st.session_state:
    st.session_state["nav"] = "Utilidad Operativa"
pagina = st.session_state["nav"]

# Filtro de VARIAS aduanas a la vez (pedido del usuario, 2026-11-25: antes
# solo se podía elegir una) -- aplica a las 5 pestañas, incluida Utilidad
# Operativa (antes nunca se filtraba por aduana; cambio explícito del
# usuario, 2026-11-25). Lista vacía = "Todas".
ADUANAS_DISPONIBLES = sorted(
    set(EM["aduana"].dropna()) | set(REC["aduana"].dropna()) | set(sueldos_df["aduana"].dropna())
)
if "aduanas_filtro" not in st.session_state:
    st.session_state["aduanas_filtro"] = []
aduanas_sel = [a for a in st.session_state["aduanas_filtro"] if a in ADUANAS_DISPONIBLES]
if aduanas_sel != st.session_state["aduanas_filtro"]:
    st.session_state["aduanas_filtro"] = aduanas_sel
# Tupla (hashable, orden estable) para los loaders cacheados que reciben
# el filtro de aduana (historial_ingresos/historial_gastos via
# _flujo_payload) -- None cuando no hay filtro, igual que antes.
ADUANAS_FILTRO = tuple(sorted(aduanas_sel)) if aduanas_sel else None

if not aduanas_sel:
    EM_F, REC_F, SUELDOS_F, HONOR_F, COSTOS_F = EM, REC, sueldos_df, honorarios_aa_df, costos_ref_df
else:
    EM_F = EM[EM["aduana"].isin(aduanas_sel)]
    REC_F = REC[REC["aduana"].isin(aduanas_sel)]
    SUELDOS_F = sueldos_df[sueldos_df["aduana"].isin(aduanas_sel)]
    HONOR_F = honorarios_aa_df[honorarios_aa_df["aduana"].isin(aduanas_sel)]
    COSTOS_F = costos_ref_df[costos_ref_df["aduana"].isin(aduanas_sel)]

# Ingreso de la página "Ingresos" = Emitidas + Honorarios AA, misma fórmula
# que ya se usaba en Utilidad Operativa (pedido 2026-11-05: "en la pestana
# de ingresos debe mostrarme lo mismo, aun no toma honorarios aa").
TOT_EM_F = float(EM_F["importe"].sum()) + float(HONOR_F["monto"].sum())
TOT_REC_F = float(REC_F["importe"].sum())
# La tarjeta compartida "Utilidad operativa" (antes "Resultado operativo")
# debe mostrar la Utilidad Operativa real de la(s) aduana(s) filtrada(s) --
# misma fórmula que Utilidad Operativa (Ingreso - Sueldos - Gastos - Costos
# de Referencia), no solo Emitidas - Recibidas (antes le faltaban Sueldos y
# Costos de Referencia; pedido del usuario, 2026-11-24: "así debe llamarse,
# utilidad operativa").
MARGEN_F = TOT_EM_F - TOT_REC_F - float(SUELDOS_F["total"].sum()) - float(COSTOS_F["monto"].sum())
MARGEN_PCT_F = MARGEN_F / TOT_EM_F if TOT_EM_F else None
COBERTURA_F = (REC_F.loc[REC_F["concepto_cierre"].notna(), "importe"].sum() / TOT_REC_F) if TOT_REC_F else 0.0

st.markdown("""<style>
    [data-testid='stSidebar'] { display: none !important; }
    [data-testid='stToolbar'] { display: none !important; }
    [data-testid='stHeader'] { display: none !important; }
    [data-testid='stDecoration'] { display: none !important; }
    footer { display: none !important; }
    #MainMenu { display: none !important; }
    div[data-testid='stAppViewContainer'] { background: #F4F5F7 !important; }
    div[data-testid='stMain'] { background: #F4F5F7 !important; }
    div[data-testid='stMainBlockContainer'] {
        padding: 0 !important; max-width: 100% !important;
    }
    iframe[title='app_facturacion.facturacion_dashboard'] {
        width: 100% !important; min-height: 100vh !important;
    }
    </style>""", unsafe_allow_html=True)

PERIODOS_DISPONIBLES = [{"clave": k, "mes_txt": v["mes_txt"]}
                        for k, v in sorted(PERIODOS.items(), key=lambda kv: kv[1]["mes_num"])]

_uo_pl = _uo_payload(EM_F, REC_F, SUELDOS_F, COSTOS_F, HONOR_F, ppto_mensual, ppto_anual, aduanas_sel, RANGO_TXT)
_gastos_pl = _flujo_payload(False, REC_F, MARGEN_F, MARGEN_PCT_F, COBERTURA_F,
                            ADUANAS_FILTRO, costos_df=COSTOS_F, sueldos_df=SUELDOS_F)
_ingresos_pl = _flujo_payload(True, EM_F, MARGEN_F, MARGEN_PCT_F, COBERTURA_F,
                              ADUANAS_FILTRO, honorarios_df=HONOR_F)
_nomina_pl = _nomina_payload(SUELDOS_F, TOT_REC_F)
_tarjeta_pl = _tarjeta_payload()

_dash_payload = {
    "pagina_activa": pagina,
    "nav_opciones": NAV_OPCIONES,
    "aduanas_filtro": aduanas_sel,
    "aduanas_disponibles": ADUANAS_DISPONIBLES,
    "periodo_desde": periodo_desde,
    "periodo_hasta": periodo_hasta,
    "periodos_disponibles": PERIODOS_DISPONIBLES,
    "utilidad_operativa": _uo_pl,
    "gastos": _gastos_pl,
    "ingresos": _ingresos_pl,
    "nomina": _nomina_pl,
    "tarjeta_credito": _tarjeta_pl,
    "colores": {"azul": JOCS_AZUL, "verde": JOCS_VERDE, "neg": C_NEG, "margen": C_MARGEN, "ppto": "#9DBFE8"},
}
_dash_resultado = _dashboard_component(datos=_dash_payload, key="fac_dashboard", default=None)
if _dash_resultado:
    _cambio = False
    if _dash_resultado.get("nav_to") and _dash_resultado["nav_to"] != pagina:
        st.session_state["nav"] = _dash_resultado["nav_to"]
        _cambio = True
    if (_dash_resultado.get("aduanas_to") is not None
            and sorted(_dash_resultado["aduanas_to"]) != sorted(aduanas_sel)):
        st.session_state["aduanas_filtro"] = _dash_resultado["aduanas_to"]
        _cambio = True
    if _dash_resultado.get("periodo_desde_to") is not None and _dash_resultado["periodo_desde_to"] != periodo_desde:
        st.session_state["periodo_desde"] = _dash_resultado["periodo_desde_to"]
        # Al cambiar el rango de periodos, las aduanas disponibles pueden
        # cambiar -- se resetea a "Todas" para no quedar filtrado a una
        # aduana que no exista en el nuevo rango (se revalida de todas
        # formas arriba).
        st.session_state["aduanas_filtro"] = []
        _cambio = True
    if _dash_resultado.get("periodo_hasta_to") is not None and _dash_resultado["periodo_hasta_to"] != periodo_hasta:
        st.session_state["periodo_hasta"] = _dash_resultado["periodo_hasta_to"]
        st.session_state["aduanas_filtro"] = []
        _cambio = True
    if _cambio:
        st.rerun()
