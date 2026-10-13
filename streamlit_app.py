import math
import unicodedata
from collections import namedtuple
from io import BytesIO
import altair as alt
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

alt.data_transformers.disable_max_rows()

st.set_page_config(
    page_title="Facturación Emitida y Recibida",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).parent
LOGO_PATH = BASE / "imagenes" / "jocs_logo.png"

# Fuente de datos: Google Sheets en vivo (antes: Excel local, que vivia fuera de
# Google Drive porque los .xlsx sincronizados se corrompian solos). La hoja esta
# compartida como "cualquiera con el link" (solo lectura), lo que permite
# descargarla como .xlsx sin API key ni cuenta de servicio.
SHEET_ID = "1F-b12XUlASAXNnhgYm5G-PRB46s2Ms9RjJ3j6Rj_QVU"
SHEET_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
# De cada libro SOLO se usan las hojas 'emitidas' y 'recibidas' (el resto se ignora
# a proposito; ver load_sueldos/load_honorarios_josefa/load_costos_referencia/
# load_utilidad_operativa mas abajo para las otras hojas que si se leen).
HOJAS_USADAS = {"emitidas": "EMITIDA", "recibidas": "RECIBIDA"}

MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
         7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}

# namedtuple de modulo (no el de itertuples(), que es una clase dinamica y
# st.cache_data no la puede serializar con pickle).
Periodo = namedtuple("Periodo", ["anio", "mes", "periodo", "periodo_label"])

# ── Paleta de marca JOCS (tablero propio, NO reusa la identidad de Ancora) ─────
# Tomada del logo real (imagenes/jocs_logo.png): azul marino + verde salvia,
# sobre fondo blanco solido (el logo no tiene version transparente/oscura, por
# lo que este tablero usa tema claro, no el navy oscuro de los tableros Ancora).
JOCS_AZUL = "#244072"
JOCS_VERDE = "#507058"
C_EMITIDA = JOCS_AZUL       # ingreso / facturación emitida
C_RECIBIDA = JOCS_VERDE     # gasto / facturación recibida
C_MARGEN = "#c98a2e"        # ambar, para que la linea de margen resalte sobre azul y verde
C_NEG = "#c0392b"
C_POS = JOCS_AZUL

# Por ahora el tablero muestra solo PROPIOS (excluye TERCEROS, ACTIVO FIJO y lineas sin tipo).
# Poner en False para volver a incluir todo.
SOLO_PROPIOS = True

PALETA_CAT = ["#244072", "#507058", "#8a6d4b", "#4a7fa8", "#7a9e6e", "#c98a2e",
              "#6f5c8a", "#9aa5b1", "#a45c6b", "#3f8f8a", "#c9994e", "#5d7a9e",
              "#7fa374", "#95677a", "#4b6b7a", "#b0a160", "#6f8f9e", "#8f6f4b", "#5a7a5c"]

def color_for(i):
    return PALETA_CAT[i % len(PALETA_CAT)]

# ── Helpers de formato ──────────────────────────────────────────────────────────
def fmt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"-${-v:,.0f}" if v < 0 else f"${v:,.0f}"

def fmt_m(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    signo, a = ("-" if v < 0 else ""), abs(v)
    if a >= 1e6:
        return f"{signo}${a/1e6:,.1f}M"
    if a >= 1e3:
        return f"{signo}${a/1e3:,.0f}k"
    return f"{signo}${a:,.0f}"

def pct(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{'+' if v >= 0 else ''}{v*100:.1f}%"

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

# ── Tema Altair (reemplaza Plotly; misma identidad visual JOCS clara) ──────────
FUENTE = "system-ui, 'Segoe UI', Roboto, sans-serif"

# Etiquetas de eje en dinero abreviado ($1.2M / $340k), igual que antes con "$~s" de Plotly.
EJE_DINERO_EXPR = (
    "datum.value == 0 ? '$0' : "
    "(abs(datum.value) >= 1000000 ? '$' + format(datum.value/1000000, ',.1f') + 'M' : "
    "(abs(datum.value) >= 1000 ? '$' + format(datum.value/1000, ',.0f') + 'k' : "
    "'$' + format(datum.value, ',.0f')))"
)

def eje_dinero(**kw):
    return alt.Axis(labelExpr=EJE_DINERO_EXPR, labelFont=FUENTE, titleFont=FUENTE, **kw)

def tema(chart):
    """Aplica la identidad visual JOCS (tema claro) a cualquier chart de Altair."""
    return (chart
        .configure_view(strokeOpacity=0)
        .configure_axis(labelColor="#54616f", titleColor=JOCS_AZUL, gridColor="#eceff2",
                        domainColor="#d7dce1", labelFont=FUENTE, titleFont=FUENTE,
                        titleFontWeight=600, labelFontSize=11, titleFontSize=12)
        .configure_legend(labelColor="#54616f", titleColor=JOCS_AZUL,
                          labelFont=FUENTE, titleFont=FUENTE)
        .configure_title(color=JOCS_AZUL, font=FUENTE, fontWeight=600)
    )

def barra_h(serie, color, height=None, formateador=fmt_m):
    """Barra horizontal simple + etiqueta afuera a la derecha (reemplaza el patrón
    repetido go.Bar(orientation='h', textposition='outside') de varias vistas)."""
    df = serie.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    df["etiqueta"] = df["valor"].map(formateador)
    orden = df["label"].tolist()
    h = height or max(280, 30 * len(df) + 60)
    maxv = float(df["valor"].max()) if len(df) else 1.0
    barras = alt.Chart(df).mark_bar(color=color, cornerRadiusTopRight=3,
                                    cornerRadiusBottomRight=3).encode(
        y=alt.Y("label:N", title=None, sort=orden),
        x=alt.X("valor:Q", title=None, axis=eje_dinero(), scale=alt.Scale(domain=[0, maxv * 1.2])),
        tooltip=[alt.Tooltip("label:N", title=""), alt.Tooltip("etiqueta:N", title="Monto")],
    )
    etiq = alt.Chart(df).mark_text(align="left", dx=5, fontSize=11, color="#3d4752").encode(
        y=alt.Y("label:N", sort=orden), x="valor:Q", text="etiqueta:N",
    )
    return (barras + etiq).properties(height=h)

def donut_chart(valores, height=360, colores=None, mostrar_leyenda=True, etiquetas=None):
    """Dona con % dentro de cada rebanada (reemplaza go.Pie textinfo='label+percent').
    `valores`: Series indexada por etiqueta. `etiquetas`: texto de leyenda ya armado
    (ej. 'CONCEPTO · $12,345 · 34.7%'), si se quiere distinto al label crudo."""
    df = valores.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    df["leyenda"] = etiquetas if etiquetas is not None else df["label"]
    tot = df["valor"].sum()
    df["pct"] = df["valor"] / tot if tot else 0.0
    if colores is None:
        colores = [color_for(i) for i in range(len(df))]
    r_out, r_in = height * 0.46, height * 0.25
    base = alt.Chart(df)
    arcos = base.mark_arc(innerRadius=r_in, outerRadius=r_out, stroke="#FFFFFF", strokeWidth=2).encode(
        theta=alt.Theta("valor:Q", stack=True),
        order=alt.Order("valor:Q", sort="descending"),
        color=alt.Color("leyenda:N", sort=df["leyenda"].tolist(),
                        scale=alt.Scale(domain=df["leyenda"].tolist(), range=colores),
                        legend=alt.Legend(title=None, orient="right", labelLimit=320,
                                          symbolType="circle") if mostrar_leyenda else None),
        tooltip=[alt.Tooltip("label:N", title="Concepto"),
                 alt.Tooltip("valor:Q", title="Monto", format=",.0f"),
                 alt.Tooltip("pct:Q", title="% del total", format=".1%")],
    )
    texto = base.mark_text(radius=(r_in + r_out) / 2, fontSize=12, fontWeight=600,
                           color="#FFFFFF").encode(
        theta=alt.Theta("valor:Q", stack=True),
        order=alt.Order("valor:Q", sort="descending"),
        text=alt.Text("pct:Q", format=".0%"),
    )
    return (arcos + texto).properties(height=height)

def _squarify(valores, x, y, w, h):
    """Layout de treemap 'squarified' (Bruls, Huizing, van Wijk 2000): reparte el
    rectangulo (x,y,w,h) entre 'valores' (positivos) en cajas lo mas cuadradas
    posible -- sin esto, un treemap con muchos elementos sale en tiras angostas
    e ilegibles. Sin libreria externa (es el mismo algoritmo que usa 'squarify' de
    PyPI, en ~25 lineas)."""
    valores = [v for v in valores if v > 0]
    if not valores:
        return []
    total = sum(valores)
    escala = (w * h) / total
    vals = [v * escala for v in valores]

    def peor_aspecto(fila, lado):
        s, mx, mn = sum(fila), max(fila), min(fila)
        return max((lado * lado * mx) / (s * s), (s * s) / (lado * lado * mn))

    rects, resto = [], vals[:]
    cx, cy, cw, ch = x, y, w, h
    while resto:
        lado = min(cw, ch)
        fila = [resto[0]]
        while (len(resto) > len(fila)
               and peor_aspecto(fila + [resto[len(fila)]], lado) <= peor_aspecto(fila, lado)):
            fila.append(resto[len(fila)])
        resto = resto[len(fila):]
        fila_total = sum(fila)
        # El lado fijo de la fila es el mas corto del rectangulo restante (= 'lado'
        # de arriba): si el rectangulo es mas ancho que alto, la fila es una
        # COLUMNA vertical (abarca todo el alto 'ch'); si es mas alto que ancho,
        # la fila es una banda horizontal (abarca todo el ancho 'cw'). Al reves
        # de esto, las cajas salen en tiras angostas de ancho/alto completo.
        if cw >= ch:
            ancho_fila = fila_total / ch if ch else 0
            ay = cy
            for v in fila:
                alto = v / ancho_fila if ancho_fila else 0
                rects.append((cx, ay, cx + ancho_fila, ay + alto))
                ay += alto
            cx, cw = cx + ancho_fila, cw - ancho_fila
        else:
            alto_fila = fila_total / cw if cw else 0
            ax = cx
            for v in fila:
                ancho = v / alto_fila if alto_fila else 0
                rects.append((ax, cy, ax + ancho, cy + alto_fila))
                ax += ancho
            cy, ch = cy + alto_fila, ch - alto_fila
    return rects

def treemap_chart(serie, height=360, w_layout=900.0, colores=None, formateador=fmt_m,
                  valores_reales=None):
    """Treemap: cada caja es proporcional al valor (layout '_squarify' de arriba).
    'serie' indexada por etiqueta, valores positivos; una sola caja por índice
    repetido se pierde (usar .groupby(...).sum() antes si hace falta).
    'valores_reales' (opcional, mismo índice que 'serie'): cuando 'serie' ya viene
    en valor absoluto porque el componente real puede ser negativo (el área de una
    caja no puede serlo), aquí se pasa el monto con signo real para la etiqueta y
    para colorear azul/rojo según el signo."""
    s = serie[serie > 0].sort_values(ascending=False)
    if s.empty:
        return None
    rects = _squarify(s.tolist(), 0.0, 0.0, w_layout, float(height))
    df = pd.DataFrame(rects, columns=["x0", "y0", "x1", "y1"])
    df["label"] = s.index.tolist()
    df["valor"] = (valores_reales.reindex(s.index).values if valores_reales is not None else s.values)
    df["etiqueta"] = df["valor"].map(formateador)
    df["xc"] = (df["x0"] + df["x1"]) / 2
    df["yc"] = (df["y0"] + df["y1"]) / 2
    df["w"], df["h"] = df["x1"] - df["x0"], df["y1"] - df["y0"]
    if colores is not None:
        df["color"] = colores
    elif valores_reales is not None:
        df["color"] = df["valor"].map(lambda v: JOCS_AZUL if v >= 0 else C_NEG)
    else:
        df["color"] = [color_for(i) for i in range(len(df))]
    df["texto"] = df.apply(
        lambda r: f"{r['label']}" if r["w"] > 90 and r["h"] > 34 else "", axis=1)
    df["texto_valor"] = df.apply(
        lambda r: r["etiqueta"] if r["w"] > 90 and r["h"] > 50 else "", axis=1)

    escala_xy = {"x": alt.Scale(domain=[0, w_layout]), "y": alt.Scale(domain=[height, 0])}
    cajas = alt.Chart(df).mark_rect(stroke="#FFFFFF", strokeWidth=2, cornerRadius=4).encode(
        x=alt.X("x0:Q", axis=None, scale=escala_xy["x"]), x2="x1:Q",
        y=alt.Y("y0:Q", axis=None, scale=escala_xy["y"]), y2="y1:Q",
        color=alt.Color("color:N", scale=None, legend=None),
        tooltip=[alt.Tooltip("label:N", title=""), alt.Tooltip("etiqueta:N", title="Monto")],
    )
    etiq_label = alt.Chart(df).mark_text(fontSize=11, fontWeight=600, color="#FFFFFF", dy=-7).encode(
        x=alt.X("xc:Q", scale=escala_xy["x"]), y=alt.Y("yc:Q", scale=escala_xy["y"]), text="texto:N",
    )
    etiq_valor = alt.Chart(df).mark_text(fontSize=10, color="#FFFFFF", dy=8, opacity=0.9).encode(
        x=alt.X("xc:Q", scale=escala_xy["x"]), y=alt.Y("yc:Q", scale=escala_xy["y"]),
        text="texto_valor:N",
    )
    return (cajas + etiq_label + etiq_valor).properties(height=height)

def waterfall_chart(df_in, label_col, valor_col, color_pos, color_neg, color_total,
                    total_label="TOTAL", height=380):
    """Grafico de cascada (waterfall): cada barra arranca donde termino la anterior;
    la ultima barra (total_label) es el acumulado completo, desde $0."""
    df = df_in[[label_col, valor_col]].copy()
    df.columns = ["label", "valor"]
    total = df["valor"].sum()
    df = pd.concat([df, pd.DataFrame({"label": [total_label], "valor": [total]})], ignore_index=True)
    acumulado, inicio = 0.0, []
    for i, v in enumerate(df["valor"]):
        if i == len(df) - 1:
            inicio.append(0.0)   # la barra total arranca siempre desde 0
        else:
            inicio.append(acumulado)
            acumulado += v
    df["inicio"] = inicio
    df["fin"] = df["inicio"] + df["valor"]
    df["tipo"] = ["Total" if l == total_label else ("Resta" if v < 0 else "Suma")
                 for l, v in zip(df["label"], df["valor"])]
    df["color"] = df["tipo"].map({"Suma": color_pos, "Resta": color_neg, "Total": color_total})
    df["etiqueta"] = df["valor"].map(fmt_m)
    orden = df["label"].tolist()

    barras = alt.Chart(df).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("label:N", title=None, sort=orden, axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("inicio:Q", title=None, axis=eje_dinero()),
        y2="fin:Q",
        color=alt.Color("color:N", scale=None, legend=None),
        tooltip=[alt.Tooltip("label:N", title="Concepto"), alt.Tooltip("etiqueta:N", title="Monto")],
    )
    etiq = alt.Chart(df).mark_text(dy=-8, fontSize=10, color="#3d4752").encode(
        x=alt.X("label:N", sort=orden), y=alt.Y("fin:Q"), text="etiqueta:N",
    )
    return (barras + etiq).properties(height=height)

# ── Descubrimiento de archivos ──────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner="Descargando datos de Google Sheets...")
def _descargar_libro_sheets():
    resp = requests.get(SHEET_EXPORT_URL, timeout=30)
    resp.raise_for_status()
    return resp.content

def _detectar_periodo(xl):
    """El libro de Sheets no trae nombre de archivo (ahi antes salia el mes/año);
    se detecta leyendo las propias hojas: mes de 'sueldos'/'MES NO.', año de la
    primera fecha real que se encuentre. Si no hay nada, cae en la fecha de hoy."""
    hojas = {h.strip().lower(): h for h in xl.sheet_names}
    mes = None
    if "sueldos" in hojas:
        df = xl.parse(hojas["sueldos"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        if "MES NO." in df.columns:
            vals = pd.to_numeric(df["MES NO."], errors="coerce").dropna()
            if len(vals):
                mes = int(vals.mode().iloc[0])
    anio = None
    for hoja_lc, col in [("corresponsalias", "Attribute:Fecha"),
                         ("honorarios josefa", "Fecha"),
                         ("costos de referencia", "Fecha Pago Pedimento")]:
        if hoja_lc in hojas:
            df = xl.parse(hojas[hoja_lc], header=0)
            df.columns = [str(c).strip() for c in df.columns]
            if col in df.columns:
                fechas = pd.to_datetime(df[col], errors="coerce").dropna()
                if len(fechas):
                    anio = int(fechas.dt.year.mode().iloc[0])
                    break
    hoy = pd.Timestamp.now()
    return anio or hoy.year, mes or hoy.month

def descubrir_archivos():
    contenido = _descargar_libro_sheets()
    anio, mes = _detectar_periodo(pd.ExcelFile(BytesIO(contenido)))
    return [(anio, mes, BytesIO(contenido))]

def leer_hoja(path, hoja, tipo_factura):
    df = pd.read_excel(path, sheet_name=hoja, header=0)
    df.columns = [str(c).strip().lower() for c in df.columns]

    out = pd.DataFrame(index=df.index)
    out["emisor"] = coalesce_col(df, "emisornombre").map(norm_txt)
    out["receptor"] = coalesce_col(df, "receptornombre").map(norm_txt)
    # contraparte = cliente (en emitidas) o proveedor (en recibidas)
    out["contraparte"] = out["receptor"] if tipo_factura == "EMITIDA" else out["emisor"]
    out["descripcion"] = coalesce_col(df, "descripcion").map(norm_txt)
    out["referencia"] = coalesce_col(df, "referencia").map(norm_txt)

    # El monto oficial es la columna que el usuario mantiene ya convertida a MXN
    # ('MONTO' en emitidas, 'monto 1' en recibidas). Las lineas sin monto NO se
    # rellenan con 'Importe' (incluye conceptos que el usuario deja fuera a
    # proposito, ej. IMSS/Infonavit/honorarios): se excluyen y se avisan.
    monto = pd.to_numeric(coalesce_col(df, "monto 1", "monto"), errors="coerce")
    importe_bruto = pd.to_numeric(coalesce_col(df, "importe"), errors="coerce")
    out["importe"] = monto.fillna(0.0)
    out["sin_monto"] = monto.isna() & importe_bruto.notna()
    out["importe_sin_monto"] = importe_bruto.where(out["sin_monto"], 0.0)

    # "if s else" (en vez de isinstance) se ve bien en pandas 2.x pero truena en
    # pandas 3.x: ahi un .map() encadenado puede devolver los faltantes como NaN
    # (float) en lugar de None, y NaN es "verdadero" en Python (no es falsy) —
    # el guard no lo detecta y revienta en .upper(). isinstance() no depende de
    # como pandas represente el faltante.
    aduana = coalesce_col(df, "aduana").map(norm_txt)
    out["aduana"] = aduana.map(lambda s: s.upper() if isinstance(s, str) else "SIN ADUANA")
    out["cc"] = coalesce_col(df, "cc")
    tipo_cc = coalesce_col(df, "tipo").map(norm_txt)
    out["tipo_cc"] = tipo_cc.map(lambda s: s.upper() if isinstance(s, str) else "SIN TIPO")
    out["concepto_cierre"] = coalesce_col(df, "concepto para cierre").map(norm_txt)

    out["tipo_factura"] = tipo_factura
    out["cancelada"] = (out["aduana"] == "CANCELADA") | (out["tipo_cc"] == "CANCELADA")
    # Filas totalmente vacias (sin monto ni importe) no aportan nada.
    return out[(out["importe"] != 0) | out["sin_monto"]]

@st.cache_data(show_spinner="Cargando facturación...")
def load_data():
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        for hoja_lc, tipo in HOJAS_USADAS.items():
            if hoja_lc not in hojas:
                continue
            d = leer_hoja(path, hojas[hoja_lc], tipo)
            d["anio"], d["mes"] = anio, mes
            d["periodo"] = f"{anio}-{mes:02d}"
            d["periodo_label"] = etiqueta
            partes.append(d)

    if not partes:
        return pd.DataFrame(), []

    df = pd.concat(partes, ignore_index=True)
    periodos_df = (df[["anio", "mes", "periodo", "periodo_label"]]
                   .drop_duplicates().sort_values(["anio", "mes"]))
    periodos = [Periodo(*row) for row in periodos_df.itertuples(index=False, name=None)]
    return df, periodos

# ── Modelo y Originación / Nómina: hojas que antes se ignoraban a propósito ────
# (ver EXCEL/Dashboard Carga-Aérea - Requerimientos y Simulación Visual.docx)
def _normaliza_aduana(s):
    """Como norm_txt + upper, pero tambien quita acentos (las hojas de nomina/
    corresponsalias/honorarios traen 'Tránsito', 'Cdmx', etc. con casing y
    acentos sueltos que no calzan con los nombres de aduana ya usados en
    emitidas/recibidas, que vienen siempre en MAYUSCULAS sin acentos)."""
    if not isinstance(s, str) or not s.strip():
        return None
    return unicodedata.normalize("NFKD", s.strip()).encode("ascii", "ignore").decode().upper()

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
# Variantes sin acento de las mismas columnas: la copia en Google Sheets las trae
# así (se corrigió ahí un acento que llegó dañado, pero sin volver a ponerle tilde).
ALIAS_NOMINA = {
    "Total cuotas patrón (N)": "Total cuotas patron (N)",
    "Comisión (lira)": "Comision (lira)",
}

def _hoja_generica(hojas_cache, path, nombre_hoja):
    """Busca una hoja por nombre (case-insensitive); None si el archivo no la trae."""
    hojas = hojas_cache.setdefault(path, {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names})
    return hojas.get(nombre_hoja)

@st.cache_data(show_spinner=False)
def load_sueldos():
    """Nomina real por empleado (hoja 'sueldos'): un renglon por persona/mes, con el
    desglose exacto que arma la cascada de costo de la Ficha Tecnica (Sueldo Base +
    Prestaciones + Cuotas + Vales + Puntualidad + Comision + Excedente + Provisiones
    = TOTAL2; verificado contra el archivo real, 207/207 filas cuadran)."""
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "sueldos" not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas["sueldos"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        out["nombre"] = coalesce_col(df, "NOMBRE BIEN", "NOMBRE").map(norm_txt)
        out["aduana"] = coalesce_col(df, "Departamento").map(_normaliza_aduana)
        for col_src, _ in COLS_NOMINA:
            candidatos = (col_src, ALIAS_NOMINA[col_src]) if col_src in ALIAS_NOMINA else (col_src,)
            out[col_src] = pd.to_numeric(coalesce_col(df, *candidatos), errors="coerce").fillna(0.0)
        out["total"] = pd.to_numeric(coalesce_col(df, "TOTAL2", "Total"), errors="coerce").fillna(0.0)
        out = out[out["nombre"].notna() & (out["total"] != 0)].copy()
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_honorarios_josefa():
    """Ingreso por honorarios de asesoría aduanal (hoja 'honorarios josefa')."""
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "honorarios josefa" not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas["honorarios josefa"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        out["cliente"] = coalesce_col(df, "Cliente").map(norm_txt)
        out["aduana"] = coalesce_col(df, "NOMBRE ADUANA").map(_normaliza_aduana)
        out["monto"] = pd.to_numeric(coalesce_col(df, "HONORARIOS AA"), errors="coerce").fillna(0.0)
        out = out[out["monto"] != 0].copy()
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_costos_referencia():
    """Costos de referencia por pedimento (hoja 'costos de referencia')."""
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "costos de referencia" not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas["costos de referencia"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        out["cliente"] = coalesce_col(df, "Cliente").map(norm_txt)
        out["aduana"] = coalesce_col(df, "Nombre aduana").map(_normaliza_aduana)
        out["monto"] = pd.to_numeric(coalesce_col(df, "Total Conceptos AA", "Sub Total"),
                                     errors="coerce").fillna(0.0)
        out = out[out["monto"] != 0].copy()
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_utilidad_operativa():
    """Ingresos, Honorarios AA, Sueldos y Gastos por aduana (hoja 'UTILIDAD OPERATIVA').
    Utilidad Operativa = Ingresos + Honorarios AA - Sueldos - Gastos (Honorarios AA se
    trata como ingreso adicional, confirmado con el usuario)."""
    cols_monto = ["INGRESOS", "HONORARIOS AA", "SUELDOS", "GASTOS"]
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "utilidad operativa" not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas["utilidad operativa"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        for col in cols_monto:
            out[col] = pd.to_numeric(coalesce_col(df, col), errors="coerce")
        # La hoja trae una fila basura ("INGRESOS"/"AGOSTO"...) sin montos numericos
        # y una fila TOTAL al final; ambas se excluyen (el total se calcula aqui).
        out = out[out["INGRESOS"].notna()].copy()
        out["aduana"] = coalesce_col(df, df.columns[0]).map(_normaliza_aduana)
        out = out[out["aduana"] != "TOTAL"]
        out[cols_monto] = out[cols_monto].fillna(0.0)
        out = out[(out[cols_monto] != 0).any(axis=1)]  # sin actividad en ningun rubro
        out["utilidad"] = out["INGRESOS"] + out["HONORARIOS AA"] - out["SUELDOS"] - out["GASTOS"]
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

try:
    df, periodos = load_data()
    sueldos_df = load_sueldos()
    josefa_df = load_honorarios_josefa()
    costos_ref_df = load_costos_referencia()
    utilidad_op_df = load_utilidad_operativa()
except Exception as e:
    st.error(f"Error al descargar/leer los datos de Google Sheets:\n\n`{SHEET_EXPORT_URL}`\n\n{e}")
    st.stop()

if df.empty:
    st.error(
        "El Google Sheets se descargó pero no se encontraron hojas `emitidas`/`recibidas` "
        f"con datos.\n\n`{SHEET_EXPORT_URL}`"
    )
    st.stop()

# ── Estilos propios de JOCS (tema claro; NO reusa el CSS oscuro de los Ancora) ──
# Misma familia tipografica que usa la app interna "Jocs Finanzas" (React), no
# la Montserrat de marca Ancora.
st.markdown(f"""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], .stApp,
.stApp *:not([data-testid="stIconMaterial"]) {{
    font-family: system-ui, 'Segoe UI', Roboto, sans-serif !important;
}}
[data-testid="stAppViewContainer"], [data-testid="stMarkdownContainer"],
p, span:not([data-testid="stIconMaterial"]), label, td, th, li,
div[data-testid="stMetricValue"], div[data-testid="stMetricDelta"], .stDataFrame {{ font-weight: 500; }}
h1, h2, h3, h4, h5, h6, div[data-testid="stMetricLabel"] p {{ font-weight: 600 !important; color: {JOCS_AZUL} !important; }}

section[data-testid="stSidebar"] img {{ pointer-events: none; }}
section[data-testid="stSidebar"] [data-testid="StyledFullScreenButton"] {{ display: none; }}
section[data-testid="stSidebar"] > div:first-child {{ padding-top: 0 !important; }}
section[data-testid="stSidebar"] .block-container {{ padding-top: 0 !important; }}
div[data-testid="stRadio"] > div {{ gap: 2px !important; }}
div[data-testid="stRadio"] label {{
    display: flex !important; align-items: center !important;
    padding: 10px 14px !important; border-radius: 8px !important;
    cursor: pointer !important; font-size: 0.9rem !important;
    margin-bottom: 2px !important; transition: background 0.15s !important; border: none !important;
}}
div[data-testid="stRadio"] label:hover {{ background: rgba(36,64,114,0.08) !important; }}
div[data-testid="stRadio"] label:has(input:checked) {{
    background: rgba(36,64,114,0.14) !important; font-weight: 600 !important;
}}
div[data-testid="stRadio"] label > div:first-child {{ display: none !important; }}
div[data-testid="stRadio"] {{ border: none !important; }}
div[data-testid="stMetricValue"] {{
    font-size: 1.35rem !important; line-height: 1.2 !important;
    white-space: nowrap !important; overflow: visible !important;
    color: {JOCS_AZUL} !important;
}}
</style>
""", unsafe_allow_html=True)

components.html(
    "<script>try{var d=window.parent.document.documentElement;"
    "d.setAttribute('translate','no');d.classList.add('notranslate');}catch(e){}</script>",
    height=0,
)

with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), use_container_width=True)
    st.markdown("---")
    st.caption("TABLEROS")
    pagina = st.radio(
        "",
        ["Utilidad Operativa", "Gastos por Concepto", "Nómina", "Tarjeta de Crédito"],
        label_visibility="collapsed",
        key="nav",
    )
    st.markdown("---")
    st.caption("PERIODO")
    labels_disp = [p.periodo_label for p in periodos]
    sel_labels = st.multiselect("", labels_disp, default=labels_disp,
                                label_visibility="collapsed", key="periodo_sel")
    if not sel_labels:
        sel_labels = labels_disp
    st.caption("ADUANA")
    aduanas_todas = sorted(df["aduana"].dropna().unique().tolist())
    aduanas_sel_global = st.multiselect("", aduanas_todas, default=[],
                                        label_visibility="collapsed", key="aduana_sel_global",
                                        placeholder="Todas")
    st.markdown("---")
    st.caption(f"Facturación · {len(periodos)} mes(es) cargado(s)")
    if SOLO_PROPIOS:
        st.caption("Mostrando solo PROPIOS")

def _filtra_periodo_aduana(data):
    """Aplica el filtro global de Periodo + Aduana del sidebar a cualquier
    dataframe que traiga columnas 'periodo_label' y 'aduana' (nomina,
    honorarios, costos de referencia, utilidad operativa)."""
    if data.empty:
        return data
    out = data[data["periodo_label"].isin(sel_labels)]
    if aduanas_sel_global:
        out = out[out["aduana"].isin(aduanas_sel_global)]
    return out

# ── Filtro global de periodo + aduana + exclusion de canceladas ────────────────
dff = df[df["periodo_label"].isin(sel_labels)]
if aduanas_sel_global:
    dff = dff[dff["aduana"].isin(aduanas_sel_global)]
# Vista sin el filtro de SOLO_PROPIOS, para la unica seccion (Detalle Dinamico de
# Conceptos, en Panorama) que necesita poder mostrar TERCEROS bajo su propio control.
dff_todos_tipos = dff[~dff["cancelada"] & ~dff["sin_monto"]].copy()
if SOLO_PROPIOS:
    dff = dff[dff["tipo_cc"] == "PROPIOS"]
n_cancel = int(dff["cancelada"].sum())
monto_cancel = dff.loc[dff["cancelada"], "importe"].sum()
d = dff[~dff["cancelada"] & ~dff["sin_monto"]].copy()

sueldos_f = _filtra_periodo_aduana(sueldos_df)
josefa_f = _filtra_periodo_aduana(josefa_df)
costos_ref_f = _filtra_periodo_aduana(costos_ref_df)
utilidad_op_f = _filtra_periodo_aduana(utilidad_op_df)

EM = d[d["tipo_factura"] == "EMITIDA"]
REC = d[d["tipo_factura"] == "RECIBIDA"]

TOT_EM = EM["importe"].sum()
TOT_REC = REC["importe"].sum()
MARGEN = TOT_EM - TOT_REC
MARGEN_PCT = MARGEN / TOT_EM if TOT_EM else None

cobertura_concepto = (REC.loc[REC["concepto_cierre"].notna(), "importe"].sum() / TOT_REC) if TOT_REC else 0.0

# Lineas con importe en el Excel pero sin 'monto' (el usuario las deja fuera del cierre).
sm = dff[~dff["cancelada"] & dff["sin_monto"]]
n_sin_monto = len(sm)
importe_sin_monto = sm["importe_sin_monto"].sum()

# ══════════════════════════════════════════════════════════════════════════════
# PANORAMA
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Utilidad Operativa":
    st.title("Utilidad Operativa")
    st.caption(f"{', '.join(sel_labels)} · Importes en MXN" + (" · Solo PROPIOS" if SOLO_PROPIOS else ""))

    if n_cancel:
        st.info(f"Se excluyeron **{n_cancel} líneas canceladas** ({fmt(monto_cancel)}) de todas las cifras.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Emitida", fmt_m(TOT_EM))
    c2.metric("Recibida", fmt_m(TOT_REC))
    c3.metric("Margen", fmt_m(MARGEN),
              delta=pct(MARGEN_PCT) if MARGEN_PCT is not None else None)
    c4.metric("Líneas (E+R)", f"{(EM['importe'] != 0).sum() + (REC['importe'] != 0).sum():,}")
    if n_sin_monto:
        st.caption(f"{n_sin_monto:,} líneas traen importe pero no monto en el Excel "
                   f"({fmt(importe_sin_monto)} de importe) y no se suman a las cifras.")

    if utilidad_op_f.empty:
        st.info("Sin datos de Utilidad Operativa para el periodo y aduana seleccionados.")
    else:
        st.divider()
        st.subheader("Utilidad Operativa")
        st.caption("Utilidad Operativa = Ingresos + Honorarios AA − Sueldos − Gastos "
                   "(hoja 'UTILIDAD OPERATIVA' del Sheets).")
        tot = utilidad_op_f[["INGRESOS", "HONORARIOS AA", "SUELDOS", "GASTOS"]].sum()
        df_uo = pd.DataFrame({
            "label": ["Ingresos", "Honorarios AA", "Sueldos", "Gastos"],
            "valor": [tot["INGRESOS"], tot["HONORARIOS AA"], -tot["SUELDOS"], -tot["GASTOS"]],
        })
        cascada_uo = waterfall_chart(df_uo, "label", "valor",
                                     color_pos=JOCS_AZUL, color_neg=C_NEG, color_total=C_MARGEN,
                                     total_label="Utilidad Operativa", height=400)
        st.altair_chart(tema(cascada_uo), use_container_width=True)

        st.divider()
        st.subheader("Utilidad Operativa por Aduana")
        uo_ad = utilidad_op_f.groupby("aduana")[["INGRESOS", "HONORARIOS AA", "SUELDOS",
                                                 "GASTOS", "utilidad"]].sum()
        em_por_ad = EM.groupby("aduana")["importe"].sum()
        rec_por_ad = REC.groupby("aduana")["importe"].sum()
        aduanas_uo = uo_ad.sort_values("utilidad", ascending=False).index.tolist()

        bc1, bc2 = st.columns([1.6, 1])
        with bc1:
            uo_bar = uo_ad.reset_index()
            uo_bar["color"] = uo_bar["utilidad"].map(lambda v: JOCS_AZUL if v >= 0 else C_NEG)
            uo_bar["etiqueta"] = uo_bar["utilidad"].map(fmt_m)
            mn_uo, mx_uo = float(uo_bar["utilidad"].min()), float(uo_bar["utilidad"].max())
            alto_uo = max(340, 30 * len(uo_bar) + 60)
            barras_uo = alt.Chart(uo_bar).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
                y=alt.Y("aduana:N", sort=aduanas_uo, title=None),
                x=alt.X("utilidad:Q", title=None, axis=eje_dinero(),
                        scale=alt.Scale(domain=[min(0, mn_uo * 1.3), max(0, mx_uo * 1.25)])),
                color=alt.Color("color:N", scale=None, legend=None),
                tooltip=[alt.Tooltip("aduana:N", title="Aduana"),
                         alt.Tooltip("etiqueta:N", title="Utilidad Operativa")],
            )
            etiq_uo_bar = alt.Chart(uo_bar).mark_text(align="left", dx=5, fontSize=11,
                                                       color="#3d4752").encode(
                y=alt.Y("aduana:N", sort=aduanas_uo), x="utilidad:Q", text="etiqueta:N",
            )
            st.altair_chart(tema((barras_uo + etiq_uo_bar).properties(height=alto_uo)),
                            use_container_width=True)
        with bc2:
            aduana_uo_sel = st.selectbox(
                "Seleccionar Aduana", aduanas_uo,
                format_func=lambda a: f"{a}  ·  {fmt_m(uo_ad.loc[a, 'utilidad'])}")
            fila = uo_ad.loc[aduana_uo_sel]
            emit_sel = float(em_por_ad.get(aduana_uo_sel, 0.0))
            rec_sel = float(rec_por_ad.get(aduana_uo_sel, 0.0))
            ingresos_sel = float(fila["INGRESOS"])
            pct_uo_sel = fila["utilidad"] / ingresos_sel if ingresos_sel else None
            tabla_uo = pd.DataFrame({
                "Concepto": ["Emitidas", "Honorarios AA", "Ingresos", "Recibidas", "Sueldos",
                            "Egresos", "Utilidad Operativa", "% Utilidad Operativa"],
                "Valor": [fmt_m(emit_sel), fmt_m(float(fila["HONORARIOS AA"])), fmt_m(ingresos_sel),
                         fmt_m(rec_sel), fmt_m(float(fila["SUELDOS"])), fmt_m(float(fila["GASTOS"])),
                         fmt_m(float(fila["utilidad"])),
                         pct(pct_uo_sel) if pct_uo_sel is not None else "—"],
            })
            st.dataframe(tabla_uo, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Pareto de Utilidad Operativa (todo CC)")
        st.caption("Ordenado por magnitud (positiva o negativa). Azul = utilidad positiva, "
                   "rojo = negativa. Línea ámbar: % acumulado del total en valor absoluto.")
        up = uo_ad[["utilidad"]].copy()
        up["abs"] = up["utilidad"].abs()
        up = up.sort_values("abs", ascending=False)
        tot_abs = up["abs"].sum() or 1.0
        up["acumulado"] = up["abs"].cumsum() / tot_abs
        up["color"] = up["utilidad"].map(lambda v: JOCS_AZUL if v >= 0 else C_NEG)
        up["etiqueta"] = up["utilidad"].map(fmt_m)
        up["acum_txt"] = (up["acumulado"] * 100).map(lambda v: f"{v:.0f}%")
        up = up.reset_index()
        orden_up = up["aduana"].tolist()
        barras_up = alt.Chart(up).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X("aduana:N", title=None, sort=orden_up, axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("utilidad:Q", title=None, axis=eje_dinero()),
            color=alt.Color("color:N", scale=None, legend=None),
            tooltip=[alt.Tooltip("aduana:N", title="Aduana"),
                     alt.Tooltip("etiqueta:N", title="Utilidad Operativa"),
                     alt.Tooltip("acum_txt:N", title="Acumulado (magnitud)")],
        )
        etiq_up = alt.Chart(up).mark_text(dy=-10, fontSize=10, color="#3d4752").encode(
            x=alt.X("aduana:N", sort=orden_up), y="utilidad:Q", text="etiqueta:N",
        )
        linea_up = alt.Chart(up).mark_line(
            color=C_MARGEN, strokeWidth=2.5,
            point=alt.OverlayMarkDef(filled=True, size=55, color=C_MARGEN)).encode(
            x=alt.X("aduana:N", sort=orden_up),
            y=alt.Y("acumulado:Q", title=None, axis=alt.Axis(format=".0%", orient="right"),
                    scale=alt.Scale(domain=[0, 1.05])),
        )
        etiq_acum_up = alt.Chart(up).mark_text(dy=-10, fontSize=10, fontWeight=600,
                                               color=C_MARGEN).encode(
            x=alt.X("aduana:N", sort=orden_up),
            y=alt.Y("acumulado:Q", scale=alt.Scale(domain=[0, 1.05])), text="acum_txt:N",
        )
        linea_umbral_up = alt.Chart(pd.DataFrame({"y": [0.8]})).mark_rule(
            color=C_NEG, strokeDash=[4, 3], strokeWidth=1.5).encode(
            y=alt.Y("y:Q", scale=alt.Scale(domain=[0, 1.05])),
        )
        barras_layer_up = alt.layer(barras_up, etiq_up).properties(height=420)
        linea_layer_up = alt.layer(linea_umbral_up, linea_up, etiq_acum_up).properties(height=420)
        st.altair_chart(tema(alt.layer(barras_layer_up, linea_layer_up).resolve_scale(y="independent")),
                        use_container_width=True)

    st.divider()
    st.subheader("Costos de Originación")
    col_izq, col_der = st.columns(2)
    with col_izq:
        st.markdown("**Honorarios de Josefa**")
        tot_josefa = josefa_f["monto"].sum() if not josefa_f.empty else 0.0
        st.metric("Total", fmt_m(tot_josefa))
        if not josefa_f.empty:
            por_ad_j = josefa_f.groupby("aduana")["monto"].sum().sort_values(ascending=False)
            st.altair_chart(tema(barra_h(por_ad_j, JOCS_AZUL, height=max(260, 30 * len(por_ad_j)))),
                            use_container_width=True)
        else:
            st.info("Sin datos.")
    with col_der:
        st.markdown("**Costos de Referencias**")
        tot_cr = costos_ref_f["monto"].sum() if not costos_ref_f.empty else 0.0
        st.metric("Total", fmt_m(tot_cr))
        if not costos_ref_f.empty:
            por_ad_cr = costos_ref_f.groupby("aduana")["monto"].sum().sort_values(ascending=False)
            st.altair_chart(tema(barra_h(por_ad_cr, JOCS_VERDE, height=max(260, 30 * len(por_ad_cr)))),
                            use_container_width=True)
        else:
            st.info("Sin datos.")

# ══════════════════════════════════════════════════════════════════════════════
# GASTOS POR CONCEPTO
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Gastos por Concepto":
    st.title("Gastos por Concepto")
    st.caption(f"{', '.join(sel_labels)} · Distribución propia por aduana, una aduana a la vez")

    modo = st.radio("Ver", ["Gasto (Recibida)", "Ingreso (Emitida)"], horizontal=True, key="gc_modo")
    es_ingreso = modo == "Ingreso (Emitida)"

    if es_ingreso:
        st.caption("Emitidas no trae una categoría de cierre como Recibidas — aquí se agrupa por "
                   "**Descripción** de la línea en su lugar.")
        campo_concepto, titulo_concepto = "descripcion", "Descripción"
        titulo_plural, articulo_concepto = "Descripciones", "una"
        r_all = EM.copy()
        color_base = C_EMITIDA
    else:
        campo_concepto, titulo_concepto = "concepto_cierre", "Concepto"
        titulo_plural, articulo_concepto = "Conceptos", "un"
        r_all = REC.copy()
        color_base = C_RECIBIDA
        if cobertura_concepto < 0.95:
            st.warning(
                f"Solo el **{cobertura_concepto*100:.0f}%** del gasto recibido trae un concepto de "
                f"cierre asignado en el Excel de origen; el resto aparece como **'Sin categorizar'**."
            )

    r_all[campo_concepto] = r_all[campo_concepto].fillna("Sin categorizar")
    tot_all = r_all["importe"].sum()
    gasto_ad = r_all.groupby("aduana")["importe"].sum().sort_values(ascending=False)
    if gasto_ad.empty:
        st.info("Sin gasto recibido en el periodo seleccionado.")
        st.stop()

    f1, f2 = st.columns([3, 1])
    with f1:
        aduana_sel = st.selectbox("Aduana", gasto_ad.index.tolist(), index=0,
                                  format_func=lambda a: f"{a}  ·  {fmt_m(gasto_ad[a])}")
    with f2:
        top_n = st.selectbox("Conceptos en la dona", [8, 12, 18], index=0)

    r = r_all[r_all["aduana"] == aduana_sel]
    tot_gasto = r["importe"].sum()
    por_c = r.groupby(campo_concepto)["importe"].sum().sort_values(ascending=False)
    etq_kpi = "Ingreso" if es_ingreso else "Gasto"

    c1, c2, c3 = st.columns(3)
    c1.metric(f"{etq_kpi} de {aduana_sel}", fmt_m(tot_gasto))
    c2.metric(f"% del {etq_kpi.lower()} total", f"{tot_gasto / tot_all * 100:.1f}%" if tot_all else "—")
    c3.metric(f"{titulo_plural} distintos", f"{len(por_c):,}")

    st.divider()
    st.subheader(f"{etq_kpi} por {titulo_concepto.lower()} · {aduana_sel}")

    # La dona solo admite valores positivos (notas de credito quedan solo en la tabla);
    # los conceptos fuera del top se agrupan en 'Otros conceptos'.
    pos = por_c[por_c > 0]
    dona = pos.head(top_n)
    if len(pos) > top_n:
        dona = pd.concat([dona, pd.Series({"Otros conceptos": pos.iloc[top_n:].sum()})])
    colores, k = [], 0
    for c in dona.index:
        if c == "Sin categorizar":
            colores.append(C_NEG)
        elif c == "Otros conceptos":
            colores.append("#c9cdd3")
        else:
            colores.append(color_for(k))
            k += 1

    # Etiqueta de leyenda con concepto, monto exacto y %, para leer todo sin tabla.
    base_pct = dona.sum()
    etiquetas = [f"{c}  ·  {fmt(v)}  ·  {v / base_pct * 100:.1f}%" for c, v in dona.items()]
    df_dona = pd.DataFrame({
        "concepto": dona.index, "valor": dona.values, "leyenda": etiquetas,
        "pct": dona.values / base_pct if base_pct else 0.0,
        "orden": range(len(dona)),
    })
    r_out, r_in = 300, 160
    arcos = alt.Chart(df_dona).mark_arc(innerRadius=r_in, outerRadius=r_out,
                                        stroke="#FFFFFF", strokeWidth=2).encode(
        theta=alt.Theta("valor:Q", stack=True),
        order=alt.Order("orden:Q"),
        color=alt.Color("leyenda:N", sort=etiquetas,
                        scale=alt.Scale(domain=etiquetas, range=colores),
                        legend=alt.Legend(title=None, orient="right", labelLimit=420,
                                          labelFontSize=13, symbolType="circle")),
        tooltip=[alt.Tooltip("concepto:N", title=titulo_concepto),
                 alt.Tooltip("valor:Q", title="Monto", format=",.0f"),
                 alt.Tooltip("pct:Q", title="%", format=".1%")],
    )
    pct_text = alt.Chart(df_dona).mark_text(radius=(r_in + r_out) / 2, fontSize=14,
                                             color="#FFFFFF").encode(
        theta=alt.Theta("valor:Q", stack=True), order=alt.Order("orden:Q"),
        text=alt.Text("pct:Q", format=".0%"),
    )
    centro1 = alt.Chart(pd.DataFrame({"t": [fmt_m(tot_gasto)]})).mark_text(
        fontSize=22, fontWeight=700, color=JOCS_AZUL, dy=-8).encode(text="t:N")
    centro2 = alt.Chart(pd.DataFrame({"t": [aduana_sel]})).mark_text(
        fontSize=14, color="#54616f", dy=14).encode(text="t:N")
    chart_dona = (arcos + pct_text + centro1 + centro2).properties(height=640)
    st.altair_chart(tema(chart_dona), use_container_width=True)

    st.divider()
    etq_contraparte = "Clientes" if es_ingreso else "Proveedores"
    st.subheader(f"{etq_contraparte} dentro de {articulo_concepto} {titulo_concepto.lower()}")
    concepto_sel = st.selectbox(titulo_concepto, por_c.index.tolist())
    prov = (r[r[campo_concepto] == concepto_sel]
            .groupby("contraparte")["importe"].sum().nlargest(15).sort_values(ascending=False))
    if len(prov):
        st.altair_chart(tema(barra_h(prov, color_base, height=max(280, 30 * len(prov)))),
                        use_container_width=True)
    else:
        demostrativo = "esta" if articulo_concepto == "una" else "este"
        st.info(f"Sin {etq_contraparte.lower()} identificados para {demostrativo} {titulo_concepto.lower()}.")


# ══════════════════════════════════════════════════════════════════════════════
# NÓMINA
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Nómina":
    st.title("Nómina (Capital Humano)")
    st.caption(f"{', '.join(sel_labels)} · Nómina real por empleado (hoja 'sueldos')")
    st.warning(
        "**Cuotas Patronales en revisión**: tras conectar esta pestaña a Google Sheets "
        "(2026-10-12) se detectó que la columna 'Cuotas Patronales' llega sistemáticamente "
        "más baja que en el Excel original (la proporción varía por empleado, así que no es "
        "un ajuste fijo que se pueda corregir aquí) — probablemente una fórmula que no se "
        "recalculó igual al pasar el archivo a Sheets. El resto de los componentes de la "
        "nómina sí coincide exacto. Esta cifra (y por lo tanto el Total Nómina) puede estar "
        "subestimada hasta que se revise la fórmula de esa columna en el Sheets."
    )

    if sueldos_f.empty:
        st.info("Sin datos de nómina para el periodo y aduana seleccionados.")
        st.stop()

    st.subheader("Resumen Ejecutivo de Nómina")
    tot_nomina = sueldos_f["total"].sum()
    pct_gasto_op = tot_nomina / TOT_REC if TOT_REC else None
    c1, c2 = st.columns(2)
    c1.metric("Total Nómina Neta", fmt_m(tot_nomina))
    c2.metric("% del Gasto Operativo", pct(pct_gasto_op) if pct_gasto_op is not None else "—",
              delta_color="off")
    st.caption("% del Gasto Operativo = Nómina Neta ÷ Total Recibida del periodo seleccionado.")

    st.divider()
    st.subheader("Desglose por Aduana y Lista de Personal")
    st.caption("Cada cuadro es una aduana; el tamaño es proporcional a su nómina total.")
    por_aduana_nom = sueldos_f.groupby("aduana")["total"].sum().sort_values(ascending=False)
    tm_aduana = treemap_chart(por_aduana_nom, height=280)
    if tm_aduana is not None:
        st.altair_chart(tema(tm_aduana), use_container_width=True)

    aduana_nom_sel = st.selectbox("Seleccionar Aduana", por_aduana_nom.index.tolist(),
                                  format_func=lambda a: f"{a}  ·  {fmt_m(por_aduana_nom[a])}")
    tabla_nom = sueldos_f[sueldos_f["aduana"] == aduana_nom_sel].copy()
    tot_aduana_nom = tabla_nom["total"].sum()
    tabla_nom["pct"] = tabla_nom["total"] / tot_aduana_nom if tot_aduana_nom else 0.0
    tabla_nom = tabla_nom.sort_values("total", ascending=False)
    st.caption(f"Total nómina {aduana_nom_sel}: {fmt_m(tot_aduana_nom)} · {len(tabla_nom)} empleado(s) "
               "· cada cuadro es un empleado de esta aduana.")
    por_empleado_nom = tabla_nom.groupby("nombre")["total"].sum().sort_values(ascending=False)
    tm_empleado = treemap_chart(por_empleado_nom, height=380)
    if tm_empleado is not None:
        st.altair_chart(tema(tm_empleado), use_container_width=True)
    st.dataframe(pd.DataFrame({
        "Empleado": tabla_nom["nombre"],
        "% Conformación Aduana": tabla_nom["pct"].map(lambda v: f"{v*100:.1f}%"),
        "Sueldo Total": tabla_nom["total"].map(fmt),
    }), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Ficha Técnica del Trabajador")
    empleados_disp = sueldos_f.sort_values("nombre")["nombre"].unique().tolist()
    empleado_sel = st.selectbox("Seleccionar Empleado", empleados_disp, key="emp_sel")
    fila_emp = sueldos_f[sueldos_f["nombre"] == empleado_sel]
    if len(fila_emp) > 1:
        st.caption(f"{len(fila_emp)} registros para este empleado en el periodo — se suman.")
    fila_emp_sum = fila_emp[[c for c, _ in COLS_NOMINA] + ["total"]].sum()

    c1, c2 = st.columns(2)
    c1.metric("Aduana", fila_emp["aduana"].iloc[0] if len(fila_emp) else "—")
    c2.metric("Costo Total Neto", fmt_m(fila_emp_sum["total"]))

    st.caption("Cada cuadro es un componente de su percepción; el tamaño es proporcional al "
               "monto (azul = positivo, rojo = negativo, ej. ajustes de Provisiones).")
    serie_comp_real = pd.Series({et: fila_emp_sum[c] for c, et in COLS_NOMINA})
    tm_comp = treemap_chart(serie_comp_real.abs(), height=360, valores_reales=serie_comp_real)
    if tm_comp is not None:
        st.altair_chart(tema(tm_comp), use_container_width=True)
    else:
        st.info("Sin componentes que mostrar.")
    # st.caption() renderiza markdown: dos o mas "$" sueltos en el mismo texto los
    # interpreta como un bloque de LaTeX y se come los signos de pesos. Se escapan
    # a "\$" para que se vean literales (aqui se juntan 10 montos en una sola linea).
    partes_txt = " + ".join(fmt_m(fila_emp_sum[c]).replace("$", "\\$") for c, _ in COLS_NOMINA)
    total_txt = fmt_m(fila_emp_sum["total"]).replace("$", "\\$")
    st.caption(f"{partes_txt} = {total_txt} (Costo Total Neto)")

# ══════════════════════════════════════════════════════════════════════════════
# TARJETA DE CRÉDITO
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Tarjeta de Crédito":
    st.title("Tarjeta de Crédito")
    st.info(
        "**Pestaña pendiente de datos.** Esta vista necesita los movimientos de las tarjetas "
        "AMEX de Alejandro y Josefa (fecha, monto, UUID de factura relacionado si existe) para "
        "calcular: total gastado, movimientos, UUID's encontrados/no encontrados, monto "
        "facturado/no facturado y el % del gasto ya documentado por aduana. Esas hojas todavía "
        "no están en el Google Sheets conectado — hay que agregarlas ahí primero."
    )

