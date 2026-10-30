from io import BytesIO
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

@st.cache_data(ttl=300, show_spinner=False)
def _descargar_sheet_drive(file_id):
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    return resp.content

# Julio nombra esta aduana "IMILE" a secas en vez de "IMILE AIFA" (archivos
# de meses mas recientes ya usan el nombre completo) -- sin este alias, el
# desfase de un mes de Imile AIFA en Emitidas no reconoce esas filas y se
# cuentan dos veces (confirmado revisando el archivo de Julio, 2026-10-30).
ALIAS_ADUANA_FACTURA = {"IMILE": "IMILE AIFA"}

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
    aduana = coalesce_col(df, "aduana").map(norm_txt)
    out["aduana"] = aduana.map(_normaliza_aduana)
    tipo_cc = coalesce_col(df, "tipo").map(norm_txt)
    out["tipo_cc"] = tipo_cc.map(lambda s: s.upper() if isinstance(s, str) else "SIN TIPO")
    out["concepto_cierre"] = coalesce_col(df, "concepto para cierre", " ").map(norm_txt)
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
    df = df[df["ADUANA"].isin(CORRESPONSALIAS_ADUANAS)]
    out = pd.DataFrame(index=df.index)
    out["contraparte"] = df["Emisor.Attribute:Nombre"].map(norm_txt)
    out["descripcion"] = "Corresponsalía"
    out["importe"] = pd.to_numeric(df["Attribute:Total"], errors="coerce").fillna(0.0)
    out["aduana"] = df["ADUANA"].map(_normaliza_aduana)
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

@st.cache_data(ttl=300, show_spinner="Cargando facturación...")
def load_emitidas_recibidas(periodo):
    """'periodo' es una clave de PERIODOS (ej. 'agosto', 'julio') -- mismas
    reglas para cualquier mes, solo cambian los archivos de origen (ver
    PERIODOS). Parametrizado 2026-10-30 a pedido del usuario ("vamos a
    cargar julio con las mismas reglas que se uso para agosto")."""
    cfg = PERIODOS[periodo]
    path_mes = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas"]))
    path_sig = BytesIO(_descargar_sheet_drive(cfg["emitidas_recibidas_siguiente"]))
    resultado = {}
    for hoja, tipo in [("EMITIDAS", "EMITIDA"), ("RECIBIDAS", "RECIBIDA")]:
        d_mes = _leer_hoja_factura(path_mes, hoja, tipo)
        if tipo == "EMITIDA":
            # Imile AIFA en EMITIDAS sigue viniendo un mes atrasada -- se
            # excluye del mes actual y se toma del mes siguiente. Imile
            # Laredo YA NO tiene ese desfase, se toma tal cual venga en el
            # archivo del mes actual (confirmado con el usuario, 2026-10-30).
            d_mes = d_mes[d_mes["aduana"] != "IMILE AIFA"]
            d_sig = _leer_hoja_factura(path_sig, hoja, tipo)
            d_sig = d_sig[d_sig["aduana"] == "IMILE AIFA"]
            d = pd.concat([d_mes, d_sig], ignore_index=True)
        else:
            # Las RECIBIDAS de Imile AIFA/Imile Laredo SI vienen correctas
            # en el archivo del mes actual -- a diferencia de Emitidas, aqui
            # NO se sustituyen desde el mes siguiente (confirmado con el
            # usuario). KYA TI y Yatec si se completan con lo que caiga en
            # el 'MONTO 2' del archivo del mes siguiente, que les pertenece
            # a este periodo.
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
        resultado[tipo] = d.reset_index(drop=True)
    return resultado["EMITIDA"], resultado["RECIBIDA"]

@st.cache_data(ttl=300, show_spinner=False)
def load_sueldos(mes_num):
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["sueldos"]))
    df = pd.read_excel(path, sheet_name="ACUMULADO", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["MES NO."] == mes_num].copy()
    out = pd.DataFrame(index=df.index)
    out["nombre"] = coalesce_col(df, "nombre corregido", "NOMBRE").map(norm_txt)
    out["aduana"] = coalesce_col(df, "Departamento").map(norm_txt).map(
        _normaliza_aduana)
    for col_src, _ in COLS_NOMINA:
        out[col_src] = pd.to_numeric(coalesce_col(df, col_src), errors="coerce").fillna(0.0)
    out["total"] = pd.to_numeric(coalesce_col(df, "TOTAL2", "Total"), errors="coerce").fillna(0.0)
    return out[out["nombre"].notna() & (out["total"] != 0)].reset_index(drop=True)

@st.cache_data(ttl=300, show_spinner=False)
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

@st.cache_data(ttl=300, show_spinner=False)
def load_honorarios_aa(mes_num):
    """Pestaña 'Detalle' -- ya trae el nombre de la aduana directo (columna
    'NOMBRE ADUANA', no un código). Se filtra a mes_num (columna 'Mes'
    viene mezclada int/str en el archivo, por eso el cast a str antes de
    comparar)."""
    path = BytesIO(_descargar_sheet_drive(DRIVE_SHEET_IDS["honorarios_aa"]))
    df = pd.read_excel(path, sheet_name="Detalle", header=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["Mes"].astype(str) == str(mes_num)].copy()
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

@st.cache_data(ttl=300, show_spinner=False)
def load_presupuesto(mes_num):
    """Regresa (mensual, acumulado, ppto_anual) del mes dado -- cada uno
    {concepto: {aduana: valor}}. 'ppto_anual' es el acumulado a DICIEMBRE
    (el total real del año completo según el presupuesto, ya no una
    proyección como antes del reset)."""
    import openpyxl
    contenido = _descargar_sheet_drive(DRIVE_SHEET_IDS["presupuestos_2026"])
    wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True)
    resultado = _leer_bloques_presupuesto_anual(wb[wb.sheetnames[0]], [mes_num, 12])
    mensual, acumulado = resultado[mes_num]
    _, ppto_anual = resultado[12]
    return mensual, acumulado, ppto_anual

def _ppto_total(d, concepto, aduana=None):
    """Suma el total de un concepto del presupuesto. Si se da 'aduana', solo
    esa aduana; si no (default), la suma entre todas. El argumento se pasa
    explicito desde _flujo_payload (que sí respeta el filtro de aduana del
    header) y se omite desde _uo_payload (que nunca debe verse afectado por
    ese filtro, igual que el resto de Utilidad Operativa)."""
    valores = d.get(concepto, {})
    if aduana is not None:
        return float(valores.get(aduana, 0.0))
    return float(sum(valores.values()))

import re
UUID_RE = re.compile(r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")

def _extraer_uuids(celda):
    return UUID_RE.findall(celda) if isinstance(celda, str) else []

@st.cache_data(ttl=300, show_spinner=False)
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

# Selector de periodo (pedido por el usuario, 2026-10-30: "vamos a cargar
# julio con las mismas reglas que se uso para agosto") -- mismo mecanismo
# que el filtro de aduana (pill en el header + session_state).
if "periodo_filtro" not in st.session_state:
    st.session_state["periodo_filtro"] = PERIODO_DEFAULT
periodo_sel = st.session_state["periodo_filtro"]
if periodo_sel not in PERIODOS:
    periodo_sel = st.session_state["periodo_filtro"] = PERIODO_DEFAULT
MES_ACTUAL = PERIODOS[periodo_sel]["mes_num"]
MES_TXT = PERIODOS[periodo_sel]["mes_txt"]

try:
    EM, REC = load_emitidas_recibidas(periodo_sel)
    sueldos_df = load_sueldos(MES_ACTUAL)
    costos_ref_df = load_costos_referencia(periodo_sel)
    honorarios_aa_df = load_honorarios_aa(MES_ACTUAL)
    ppto_mensual, ppto_acumulado, ppto_anual = load_presupuesto(MES_ACTUAL)
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
def _flujo_payload(es_ingreso, base_df, margen, margen_pct, cobertura, aduana_filtro=None):
    if base_df.empty:
        return None
    ppto_key = "ingresos" if es_ingreso else "gastos"
    tiene_ppto = bool(ppto_mensual)
    tot_mes = float(base_df["importe"].sum())
    ppto_mes_tot = _ppto_total(ppto_mensual, ppto_key, aduana_filtro) if tiene_ppto else None
    ppto_acum_tot = _ppto_total(ppto_acumulado, ppto_key, aduana_filtro) if tiene_ppto else None
    n_contrapartes = int(base_df.loc[base_df["importe"] != 0, "contraparte"].nunique())

    campo_concepto = "descripcion" if es_ingreso else "concepto_cierre"
    conc = base_df.copy()
    conc[campo_concepto] = conc[campo_concepto].fillna("Sin categorizar")
    por_concepto_full = conc.groupby(campo_concepto)["importe"].sum()
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
    for aduana in sorted(base_df["aduana"].dropna().unique().tolist()):
        base_aduana = base_df[base_df["aduana"] == aduana]
        tot_aduana = float(base_aduana["importe"].sum())
        entry = {"aduana": str(aduana), "total": tot_aduana, "n_operaciones": int(len(base_aduana))}
        if es_ingreso:
            por_desc = base_aduana.groupby("descripcion")["importe"].sum().sort_values(ascending=False)
            top3 = por_desc.head(3)
            otros_val = float(por_desc.iloc[3:].sum()) if len(por_desc) > 3 else 0.0
            entry["conceptos"] = [{"nombre": str(c), "pct": float(v / tot_aduana) if tot_aduana else 0.0}
                                  for c, v in top3.items()]
            if otros_val:
                entry["conceptos"].append({"nombre": "Otros conceptos",
                                           "pct": float(otros_val / tot_aduana) if tot_aduana else 0.0})
        else:
            cr_aduana = (float(costos_ref_df.loc[costos_ref_df["aduana"] == aduana, "monto"].sum())
                        if not costos_ref_df.empty else 0.0)
            entry["costos_referencia"] = cr_aduana
            entry["costos_referencia_pct"] = float(cr_aduana / tot_aduana) if tot_aduana else 0.0
        por_aduana.append(entry)

    cumplimiento = tot_mes / ppto_mes_tot if ppto_mes_tot else None
    return {
        "tiene_ppto": tiene_ppto, "mes_txt": MES_TXT,
        "kpis": {"tot_mes": tot_mes, "ppto_mes_tot": ppto_mes_tot, "ppto_acum_tot": ppto_acum_tot,
                "n_contrapartes": n_contrapartes, "cumplimiento": cumplimiento,
                "margen": margen, "margen_pct": margen_pct},
        "por_aduana": por_aduana,
        "por_concepto": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_dona.items()],
        "por_concepto_full": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_full.items()],
        "por_contraparte": [{"nombre": str(c), "monto": float(v)} for c, v in top10.items()],
        "cobertura_concepto": (float(cobertura) if not es_ingreso else None),
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

def _uo_payload():
    """Utilidad Operativa por aduana -- regla confirmada con el usuario
    (2026-10-28): Ingreso = Emitidas + Honorarios AA; Utilidad = Ingreso -
    Sueldos - Gastos (Recibidas) - Costos de Referencia. Imile Laredo sale
    incompleta aqui (no tiene la hoja de P&L aparte que compensa que no
    factura por el SAT -- pendiente de conectar) -- se deja tal cual, no se
    inventa un ajuste."""
    aduanas = set(EM["aduana"]) | set(REC["aduana"]) | set(sueldos_df["aduana"]) \
            | set(costos_ref_df["aduana"]) | set(honorarios_aa_df["aduana"])
    em_ad = EM.groupby("aduana")["importe"].sum()
    rec_ad = REC.groupby("aduana")["importe"].sum()
    sueldos_ad = sueldos_df.groupby("aduana")["total"].sum()
    costos_ad = costos_ref_df.groupby("aduana")["monto"].sum()
    honor_ad = honorarios_aa_df.groupby("aduana")["monto"].sum()

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
        return (_ppto_total(d, "ingresos") - _ppto_total(d, "costos")
               - _ppto_total(d, "sueldos") - _ppto_total(d, "gastos"))
    ppto_por_aduana = []
    if tiene_ppto:
        aduanas_ppto = set()
        for concepto in ("ingresos", "costos", "sueldos", "gastos"):
            aduanas_ppto |= set(ppto_mensual.get(concepto, {}).keys())
        for aduana in sorted(aduanas_ppto):
            ing = float(ppto_mensual.get("ingresos", {}).get(aduana, 0.0))
            cos = float(ppto_mensual.get("costos", {}).get(aduana, 0.0))
            sue = float(ppto_mensual.get("sueldos", {}).get(aduana, 0.0))
            gas = float(ppto_mensual.get("gastos", {}).get(aduana, 0.0))
            ppto_por_aduana.append({"aduana": aduana, "ingresos": ing, "costos": cos,
                                    "sueldos": sue, "gastos": gas,
                                    "utilidad": ing - cos - sue - gas})

    ingreso_total = float(TOT_EM) + float(honor_ad.sum())
    sueldos_total = float(sueldos_ad.sum())
    gastos_total = float(TOT_REC)
    costos_total = float(costos_ad.sum())
    utilidad_total = ingreso_total - sueldos_total - gastos_total - costos_total

    def _concepto(real, clave_ppto):
        return {"real": real,
               "ppto": _ppto_total(ppto_mensual, clave_ppto) if tiene_ppto and clave_ppto else None,
               "ppto_anual": _ppto_total(ppto_anual, clave_ppto) if tiene_ppto and clave_ppto else None}

    estado_resultados = {
        "ingresos": _concepto(ingreso_total, "ingresos"),
        "costos_referencia": _concepto(costos_total, "costos"),
        "sueldos": _concepto(sueldos_total, "sueldos"),
        "gastos_operativos": _concepto(gastos_total, "gastos"),
        "utilidad": {"real": utilidad_total,
                    "ppto": _ppto_utilidad_total(ppto_mensual) if tiene_ppto else None,
                    "ppto_anual": _ppto_utilidad_total(ppto_anual) if tiene_ppto else None},
    }
    return {"tiene_ppto": tiene_ppto, "mes_txt": MES_TXT, "por_aduana": por_aduana,
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

# Filtro de aduana (pedido por el usuario, 2026-10-30): aplica a Ingresos,
# Gastos, Nómina y Tarjeta de Crédito -- Utilidad Operativa NUNCA se filtra,
# sigue usando EM/REC/sueldos_df completos sin tocar.
ADUANAS_DISPONIBLES = sorted(
    set(EM["aduana"].dropna()) | set(REC["aduana"].dropna()) | set(sueldos_df["aduana"].dropna())
)
if "aduana_filtro" not in st.session_state:
    st.session_state["aduana_filtro"] = "TODAS"
aduana_sel = st.session_state["aduana_filtro"]
if aduana_sel != "TODAS" and aduana_sel not in ADUANAS_DISPONIBLES:
    aduana_sel = st.session_state["aduana_filtro"] = "TODAS"

if aduana_sel == "TODAS":
    EM_F, REC_F, SUELDOS_F = EM, REC, sueldos_df
else:
    EM_F = EM[EM["aduana"] == aduana_sel]
    REC_F = REC[REC["aduana"] == aduana_sel]
    SUELDOS_F = sueldos_df[sueldos_df["aduana"] == aduana_sel]

TOT_EM_F = float(EM_F["importe"].sum())
TOT_REC_F = float(REC_F["importe"].sum())
MARGEN_F = TOT_EM_F - TOT_REC_F
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

_dash_payload = {
    "pagina_activa": pagina,
    "nav_opciones": NAV_OPCIONES,
    "aduana_filtro": aduana_sel,
    "aduanas_disponibles": ADUANAS_DISPONIBLES,
    "periodo_filtro": periodo_sel,
    "periodos_disponibles": PERIODOS_DISPONIBLES,
    "utilidad_operativa": _uo_payload(),
    "gastos": _flujo_payload(False, REC_F, MARGEN_F, MARGEN_PCT_F, COBERTURA_F,
                              aduana_sel if aduana_sel != "TODAS" else None),
    "ingresos": _flujo_payload(True, EM_F, MARGEN_F, MARGEN_PCT_F, COBERTURA_F,
                                aduana_sel if aduana_sel != "TODAS" else None),
    "nomina": _nomina_payload(SUELDOS_F, TOT_REC_F),
    "tarjeta_credito": _tarjeta_payload(),
    "colores": {"azul": JOCS_AZUL, "verde": JOCS_VERDE, "neg": C_NEG, "margen": C_MARGEN, "ppto": "#9DBFE8"},
}
_dash_resultado = _dashboard_component(datos=_dash_payload, key="fac_dashboard", default=None)
if _dash_resultado:
    _cambio = False
    if _dash_resultado.get("nav_to") and _dash_resultado["nav_to"] != pagina:
        st.session_state["nav"] = _dash_resultado["nav_to"]
        _cambio = True
    if _dash_resultado.get("aduana_to") is not None and _dash_resultado["aduana_to"] != aduana_sel:
        st.session_state["aduana_filtro"] = _dash_resultado["aduana_to"]
        _cambio = True
    if _dash_resultado.get("periodo_to") is not None and _dash_resultado["periodo_to"] != periodo_sel:
        st.session_state["periodo_filtro"] = _dash_resultado["periodo_to"]
        # Al cambiar de periodo, las aduanas disponibles pueden cambiar --
        # se resetea a "TODAS" para no quedar filtrado a una aduana que no
        # exista en el nuevo periodo (se revalida de todas formas arriba).
        st.session_state["aduana_filtro"] = "TODAS"
        _cambio = True
    if _cambio:
        st.rerun()
