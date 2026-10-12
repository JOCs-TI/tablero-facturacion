import math
import re
import unicodedata
from collections import namedtuple
import altair as alt
import pandas as pd
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

# Ruta relativa: en este repo de despliegue el Excel vive dentro de EXCEL/, a
# diferencia del desarrollo local que lee de C:\Reportes\EXCEL (fuera de Drive).
EXCEL_DIR = BASE / "EXCEL"
# Se leen los archivos cuyo nombre contiene "carga aerea" (ej. "base de prueba agosto
# carga aerea.xlsx"). El mes sale del nombre; de cada archivo SOLO se usan las hojas
# 'emitidas' y 'recibidas' (el resto de las hojas del libro se ignora a proposito).
ARCHIVO_CLAVE = "carga aerea"
HOJAS_USADAS = {"emitidas": "EMITIDA", "recibidas": "RECIBIDA"}

MESES = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
         7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
MESES_NOMBRE = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
                "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
                "noviembre": 11, "diciembre": 12}

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

def pareto_chart(serie, color, height=420, umbral=0.8):
    """Barras descendentes + linea de % acumulado (analisis Pareto 80/20)."""
    df = serie.sort_values(ascending=False).rename("valor").reset_index()
    df.columns = ["label", "valor"]
    tot = df["valor"].sum() or 1.0
    df["acumulado"] = df["valor"].cumsum() / tot
    df["etiqueta"] = df["valor"].map(fmt_m)
    df["acum_txt"] = (df["acumulado"] * 100).map(lambda v: f"{v:.0f}%")
    orden = df["label"].tolist()

    barras = alt.Chart(df).mark_bar(color=color, cornerRadiusTopLeft=3,
                                    cornerRadiusTopRight=3).encode(
        x=alt.X("label:N", title=None, sort=orden, axis=alt.Axis(labelAngle=-30)),
        y=alt.Y("valor:Q", title=None, axis=eje_dinero()),
        tooltip=[alt.Tooltip("label:N", title=""), alt.Tooltip("etiqueta:N", title="Monto"),
                 alt.Tooltip("acum_txt:N", title="Acumulado")],
    )
    etiq = alt.Chart(df).mark_text(dy=-10, fontSize=10, color="#3d4752").encode(
        x=alt.X("label:N", sort=orden), y="valor:Q", text="etiqueta:N",
    )
    linea = alt.Chart(df).mark_line(color=C_MARGEN, strokeWidth=2.5,
                                    point=alt.OverlayMarkDef(filled=True, size=55, color=C_MARGEN)).encode(
        x=alt.X("label:N", sort=orden),
        y=alt.Y("acumulado:Q", title=None, axis=alt.Axis(format=".0%", orient="right"),
                scale=alt.Scale(domain=[0, 1.05])),
    )
    etiq_acum = alt.Chart(df).mark_text(dy=-10, fontSize=10, fontWeight=600, color=C_MARGEN).encode(
        x=alt.X("label:N", sort=orden), y=alt.Y("acumulado:Q", scale=alt.Scale(domain=[0, 1.05])),
        text="acum_txt:N",
    )
    umbral_df = pd.DataFrame({"y": [umbral]})
    linea_umbral = alt.Chart(umbral_df).mark_rule(color=C_NEG, strokeDash=[4, 3], strokeWidth=1.5).encode(
        y=alt.Y("y:Q", scale=alt.Scale(domain=[0, 1.05])),
    )
    barras_layer = alt.layer(barras, etiq).properties(height=height)
    linea_layer = alt.layer(linea_umbral, linea, etiq_acum).properties(height=height)
    return alt.layer(barras_layer, linea_layer).resolve_scale(y="independent")

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
def _sin_acentos(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()

def descubrir_archivos():
    archivos = []
    if not EXCEL_DIR.exists():
        return archivos
    for p in EXCEL_DIR.glob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        nombre = _sin_acentos(p.stem)
        if ARCHIVO_CLAVE not in nombre:
            continue
        mes = next((n for nom, n in MESES_NOMBRE.items() if nom in nombre), None)
        m_anio = re.search(r"(20\d{2})", nombre)
        anio = int(m_anio.group(1)) if m_anio else pd.Timestamp(p.stat().st_mtime, unit="s").year
        archivos.append((anio, mes or 0, p))
    archivos.sort()
    return archivos

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
            out[col_src] = pd.to_numeric(coalesce_col(df, col_src), errors="coerce").fillna(0.0)
        out["total"] = pd.to_numeric(coalesce_col(df, "TOTAL2", "Total"), errors="coerce").fillna(0.0)
        out = out[out["nombre"].notna() & (out["total"] != 0)].copy()
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

@st.cache_data(show_spinner=False)
def load_corresponsalias():
    """Gasto de corresponsalías (hoja 'corresponsalias'): facturas recibidas de
    agencias aduanales corresponsales en otras aduanas."""
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "corresponsalias" not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas["corresponsalias"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        out["proveedor"] = coalesce_col(df, "Emisor.Attribute:Nombre").map(norm_txt)
        out["aduana"] = coalesce_col(df, "ADUANA").map(_normaliza_aduana)
        out["monto"] = pd.to_numeric(coalesce_col(df, "Attribute:Total"), errors="coerce").fillna(0.0)
        out = out[out["monto"] != 0].copy()
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

try:
    df, periodos = load_data()
    sueldos_df = load_sueldos()
    corres_df = load_corresponsalias()
    josefa_df = load_honorarios_josefa()
    costos_ref_df = load_costos_referencia()
except Exception as e:
    st.error(f"Error al abrir los archivos de facturación:\n\n`{EXCEL_DIR}`\n\n{e}")
    st.stop()

if df.empty:
    st.error(
        "No se encontraron archivos de facturación. Se esperan archivos cuyo nombre "
        "contenga `carga aerea` (ej. `base de prueba agosto carga aerea.xlsx`), con hojas "
        f"`emitidas` y `recibidas`, en:\n\n`{EXCEL_DIR}`"
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
        ["Panorama", "Gastos por Concepto", "Por Aduana (Rentabilidad)",
         "Modelo y Originación", "Nómina", "Detalle de Facturas"],
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
    corresponsalias, honorarios, costos de referencia)."""
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
corres_f = _filtra_periodo_aduana(corres_df)
josefa_f = _filtra_periodo_aduana(josefa_df)
costos_ref_f = _filtra_periodo_aduana(costos_ref_df)

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
if pagina == "Panorama":
    st.title("Facturación Emitida y Recibida")
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

    st.divider()
    st.subheader("Emitida vs Recibida por mes")
    g = d.groupby(["periodo", "periodo_label", "tipo_factura"], as_index=False)["importe"].sum()
    orden_periodos = [p.periodo_label for p in periodos if p.periodo_label in sel_labels]
    em_serie = [g.loc[(g.periodo_label == pl) & (g.tipo_factura == "EMITIDA"), "importe"].sum()
                for pl in orden_periodos]
    rec_serie = [g.loc[(g.periodo_label == pl) & (g.tipo_factura == "RECIBIDA"), "importe"].sum()
                 for pl in orden_periodos]
    margen_serie = [e - r for e, r in zip(em_serie, rec_serie)]

    df_mes = pd.DataFrame({
        "periodo": orden_periodos * 2,
        "serie": ["Emitida"] * len(orden_periodos) + ["Recibida"] * len(orden_periodos),
        "valor": em_serie + rec_serie,
    })
    df_mes["etiqueta"] = df_mes["valor"].map(fmt_m)
    df_margen = pd.DataFrame({"periodo": orden_periodos, "margen": margen_serie})
    df_margen["etiqueta"] = df_margen["margen"].map(fmt_m)

    barras = alt.Chart(df_mes).mark_bar(size=32, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("periodo:N", title=None, sort=orden_periodos, axis=alt.Axis(labelAngle=0)),
        xOffset=alt.XOffset("serie:N", sort=["Emitida", "Recibida"]),
        y=alt.Y("valor:Q", title=None, axis=eje_dinero()),
        color=alt.Color("serie:N", title=None,
                        scale=alt.Scale(domain=["Emitida", "Recibida"], range=[C_EMITIDA, C_RECIBIDA]),
                        legend=alt.Legend(orient="top", direction="horizontal")),
        tooltip=[alt.Tooltip("periodo:N", title="Mes"), alt.Tooltip("serie:N", title="Serie"),
                 alt.Tooltip("etiqueta:N", title="Monto")],
    )
    etiq_barras = alt.Chart(df_mes).mark_text(dy=-10, fontSize=11, color="#3d4752").encode(
        x=alt.X("periodo:N", sort=orden_periodos),
        xOffset=alt.XOffset("serie:N", sort=["Emitida", "Recibida"]),
        y="valor:Q", text="etiqueta:N",
    )
    linea_margen = alt.Chart(df_margen).mark_line(
        color=C_MARGEN, strokeWidth=3,
        point=alt.OverlayMarkDef(filled=True, size=80, color=C_MARGEN),
    ).encode(
        x=alt.X("periodo:N", sort=orden_periodos),
        y=alt.Y("margen:Q"),
        tooltip=[alt.Tooltip("periodo:N", title="Mes"), alt.Tooltip("etiqueta:N", title="Margen")],
    )
    etiq_margen = alt.Chart(df_margen).mark_text(dy=-12, fontSize=11, fontWeight=600, color=C_MARGEN).encode(
        x=alt.X("periodo:N", sort=orden_periodos), y="margen:Q", text="etiqueta:N",
    )
    chart = (barras + etiq_barras + linea_margen + etiq_margen).properties(height=420)
    st.altair_chart(tema(chart), use_container_width=True)

    if not SOLO_PROPIOS:
        st.divider()
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("Propios vs Terceros · Emitida")
            vals = EM.groupby("tipo_cc")["importe"].sum()
            vals = vals[vals > 0]
            if len(vals):
                st.altair_chart(tema(donut_chart(vals, height=360)), use_container_width=True)
            else:
                st.info("Sin datos.")
        with col_b:
            st.subheader("Propios vs Terceros · Recibida")
            vals = REC.groupby("tipo_cc")["importe"].sum()
            vals = vals[vals > 0]
            if len(vals):
                st.altair_chart(tema(donut_chart(vals, height=360)), use_container_width=True)
            else:
                st.info("Sin datos.")

    st.divider()
    st.subheader("Centros de Costo · Emitidas vs Recibidas")
    claves_cc_pan = ["cc", "aduana"]
    em_cc_pan = d[d["tipo_factura"] == "EMITIDA"].groupby(claves_cc_pan)["importe"].sum().rename("em")
    rec_cc_pan = d[d["tipo_factura"] == "RECIBIDA"].groupby(claves_cc_pan)["importe"].sum().rename("rec")
    cc_pan = pd.concat([em_cc_pan, rec_cc_pan], axis=1).fillna(0.0).reset_index()
    cc_pan["etiq"] = cc_pan.apply(
        lambda r: ("Sin CC" if pd.isna(r["cc"]) else str(int(r["cc"]))) + f" · {r['aduana']}", axis=1)
    cc_pan["total"] = cc_pan["em"] + cc_pan["rec"]
    cc_pan = cc_pan.nlargest(12, "total").sort_values("total", ascending=False)
    if len(cc_pan):
        df_cc_pan = pd.concat([
            pd.DataFrame({"cc": cc_pan["etiq"], "serie": "Emitida", "valor": cc_pan["em"]}),
            pd.DataFrame({"cc": cc_pan["etiq"], "serie": "Recibida", "valor": cc_pan["rec"]}),
        ], ignore_index=True)
        df_cc_pan["etiqueta"] = df_cc_pan["valor"].map(fmt_m)
        orden_cc_pan = cc_pan["etiq"].tolist()
        barras_cc_pan = alt.Chart(df_cc_pan).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X("cc:N", title=None, sort=orden_cc_pan, axis=alt.Axis(labelAngle=-35)),
            xOffset=alt.XOffset("serie:N", sort=["Emitida", "Recibida"]),
            y=alt.Y("valor:Q", title=None, axis=eje_dinero()),
            color=alt.Color("serie:N", title=None,
                            scale=alt.Scale(domain=["Emitida", "Recibida"], range=[C_EMITIDA, C_RECIBIDA]),
                            legend=alt.Legend(orient="top", direction="horizontal")),
            tooltip=[alt.Tooltip("cc:N", title="CC"), alt.Tooltip("serie:N", title="Serie"),
                     alt.Tooltip("etiqueta:N", title="Monto")],
        )
        st.altair_chart(tema(barras_cc_pan.properties(height=420)), use_container_width=True)
        st.caption("Top 12 centros de costo por volumen combinado (emitida + recibida).")
    else:
        st.info("Sin datos.")

    st.divider()
    st.subheader("Detalle Dinámico de Conceptos")
    st.caption("Facturación recibida por concepto de cierre · filtros propios de esta sección "
               "(no dependen del filtro global de Aduana ni de 'Solo PROPIOS').")
    fdc1, fdc2 = st.columns(2)
    with fdc1:
        tipo_dc = st.multiselect("Propios / Terceros", ["PROPIOS", "TERCEROS", "ACTIVO FIJO", "SIN TIPO"],
                                 default=["PROPIOS"], key="dc_tipo")
    with fdc2:
        aduanas_dc_disp = sorted(dff_todos_tipos["aduana"].dropna().unique().tolist())
        aduanas_dc = st.multiselect("Aduana", aduanas_dc_disp, default=[], key="dc_aduana",
                                    placeholder="Todas")
    rec_dc = dff_todos_tipos[dff_todos_tipos["tipo_factura"] == "RECIBIDA"].copy()
    if tipo_dc:
        rec_dc = rec_dc[rec_dc["tipo_cc"].isin(tipo_dc)]
    if aduanas_dc:
        rec_dc = rec_dc[rec_dc["aduana"].isin(aduanas_dc)]
    rec_dc["concepto_cierre"] = rec_dc["concepto_cierre"].fillna("Sin categorizar")
    conceptos_dc = rec_dc.groupby("concepto_cierre")["importe"].sum().nlargest(15).sort_values(ascending=False)
    if len(conceptos_dc):
        st.altair_chart(tema(barra_h(conceptos_dc, C_RECIBIDA, height=max(320, 30 * len(conceptos_dc)))),
                        use_container_width=True)
    else:
        st.info("Sin datos para esta combinación de filtros.")

    st.divider()
    st.subheader("Top 10 Aduanas por Ingreso Facturado (Pareto 80/20)")
    top_ad_pareto = EM.groupby("aduana")["importe"].sum().nlargest(10)
    if len(top_ad_pareto):
        st.altair_chart(tema(pareto_chart(top_ad_pareto, C_EMITIDA, height=420)),
                        use_container_width=True)
        st.caption("Barras: ingreso emitido por aduana (de mayor a menor). Línea ámbar: % acumulado "
                   "del ingreso total · línea punteada roja: umbral Pareto 80%.")
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
# POR ADUANA / CC
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Por Aduana (Rentabilidad)":
    st.title("Facturación por Aduana (Rentabilidad)")
    st.caption(f"{', '.join(sel_labels)}")

    em_ad = EM.groupby("aduana")["importe"].sum()
    rec_ad = REC.groupby("aduana")["importe"].sum()
    aduanas = sorted(set(em_ad.index) | set(rec_ad.index),
                     key=lambda a: -(em_ad.get(a, 0) + rec_ad.get(a, 0)))

    st.subheader("Matriz de Rentabilidad por Aduana")
    df_mat = pd.DataFrame({"aduana": aduanas,
                           "em": [em_ad.get(a, 0) for a in aduanas],
                           "rec": [rec_ad.get(a, 0) for a in aduanas]})
    df_mat = df_mat[df_mat["em"] >= 1].copy()   # % margen solo tiene sentido si hay emitida real
    df_mat["margen"] = df_mat["em"] - df_mat["rec"]
    df_mat["pct_margen"] = df_mat["margen"] / df_mat["em"]
    if len(df_mat):
        ref_y = float(df_mat["margen"].median())
        ref_x = MARGEN_PCT if MARGEN_PCT is not None else float(df_mat["pct_margen"].median())

        def _cuadrante(row):
            alto_y, alto_x = row["margen"] >= ref_y, row["pct_margen"] >= ref_x
            if alto_y and alto_x:
                return "Estrellas"
            if alto_y and not alto_x:
                return "Oportunidades"
            if not alto_y and alto_x:
                return "Caballos de Batalla"
            return "Bajo Desempeño"

        df_mat["cuadrante"] = df_mat.apply(_cuadrante, axis=1)
        colores_cuad = {"Estrellas": JOCS_AZUL, "Oportunidades": "#c98a2e",
                        "Caballos de Batalla": JOCS_VERDE, "Bajo Desempeño": C_NEG}

        puntos = alt.Chart(df_mat).mark_circle(opacity=0.85, stroke="#FFFFFF", strokeWidth=1).encode(
            x=alt.X("pct_margen:Q", title="% Margen", axis=alt.Axis(format=".0%")),
            y=alt.Y("margen:Q", title="Margen ($)", axis=eje_dinero()),
            size=alt.Size("em:Q", title="Emitida", legend=None, scale=alt.Scale(range=[150, 1600])),
            color=alt.Color("cuadrante:N", title=None,
                            scale=alt.Scale(domain=list(colores_cuad.keys()),
                                            range=list(colores_cuad.values())),
                            legend=alt.Legend(orient="bottom")),
            tooltip=[alt.Tooltip("aduana:N", title="Aduana"),
                     alt.Tooltip("em:Q", title="Emitida", format=",.0f"),
                     alt.Tooltip("rec:Q", title="Recibida", format=",.0f"),
                     alt.Tooltip("margen:Q", title="Margen", format=",.0f"),
                     alt.Tooltip("pct_margen:Q", title="% Margen", format=".1%"),
                     alt.Tooltip("cuadrante:N", title="Cuadrante")],
        )
        etiq_mat = alt.Chart(df_mat).mark_text(dy=-14, fontSize=10, color="#3d4752").encode(
            x="pct_margen:Q", y="margen:Q", text="aduana:N",
        )
        regla_x = alt.Chart(pd.DataFrame({"x": [ref_x]})).mark_rule(
            color="#9aa5b1", strokeDash=[4, 3]).encode(x="x:Q")
        regla_y = alt.Chart(pd.DataFrame({"y": [ref_y]})).mark_rule(
            color="#9aa5b1", strokeDash=[4, 3]).encode(y="y:Q")
        chart_mat = (regla_x + regla_y + puntos + etiq_mat).properties(height=460)
        st.altair_chart(tema(chart_mat), use_container_width=True)
        st.caption(f"Cuadrantes definidos por el margen mediano ({fmt_m(ref_y)}) y el % margen global "
                   f"del negocio ({pct(ref_x)}). El tamaño de cada burbuja es proporcional a lo emitido.")
    else:
        st.info("Sin datos suficientes para la matriz de rentabilidad.")

    st.divider()
    st.subheader("Emitida vs Recibida por aduana")
    escala = st.radio("Escala", ["Logarítmica", "Normal"], horizontal=True, key="escala_ad")
    es_log = escala == "Logarítmica"
    if es_log:
        st.caption("En escala logarítmica cada línea vale 10 veces la anterior: así se ven también "
                   "las aduanas chicas. El monto exacto va sobre cada barra.")
    # Montos menores a $1 (centavos de facturas de prueba) no se grafican: estiran el eje log.
    y_em = [float(v) if v >= 1 else 0.0 for v in (em_ad.get(a, 0) for a in aduanas)]
    y_rec = [float(v) if v >= 1 else 0.0 for v in (rec_ad.get(a, 0) for a in aduanas)]

    df_ad = pd.DataFrame({
        "aduana": aduanas * 2,
        "serie": ["Emitida"] * len(aduanas) + ["Recibida"] * len(aduanas),
        "valor": y_em + y_rec,
    })
    df_ad = df_ad[df_ad["valor"] > 0].copy()
    df_ad["etiqueta"] = df_ad["valor"].map(fmt_m)
    positivos = df_ad["valor"].tolist()
    y2_kw = {}
    if positivos and es_log:
        dom_min = 10 ** (math.log10(min(positivos)) - 0.8)
        escala_y = alt.Scale(type="log", base=10, nice=False,
                             domain=[dom_min, 10 ** (math.log10(max(positivos)) + 0.6)])
        # Vega-Lite no infiere una base para barras en escala log (log(0) no existe);
        # sin esto las barras quedan invisibles, con solo la etiqueta flotando.
        y2_kw = {"y2": alt.Y2(datum=dom_min)}
    elif positivos:
        escala_y = alt.Scale(domain=[0, max(positivos) * 1.25])
    else:
        escala_y = alt.Undefined

    barras = alt.Chart(df_ad).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
        x=alt.X("aduana:N", title=None, sort=aduanas, axis=alt.Axis(labelAngle=-35)),
        xOffset=alt.XOffset("serie:N", sort=["Emitida", "Recibida"]),
        y=alt.Y("valor:Q", title=None, axis=eje_dinero(), scale=escala_y),
        color=alt.Color("serie:N", title=None,
                        scale=alt.Scale(domain=["Emitida", "Recibida"], range=[C_EMITIDA, C_RECIBIDA]),
                        legend=alt.Legend(orient="top", direction="horizontal")),
        tooltip=[alt.Tooltip("aduana:N", title="Aduana"), alt.Tooltip("serie:N", title="Serie"),
                 alt.Tooltip("etiqueta:N", title="Monto")],
        **y2_kw,
    )
    etiq_ad = alt.Chart(df_ad).mark_text(angle=270, align="left", baseline="middle", dx=4,
                                         fontSize=10, color="#3d4752").encode(
        x=alt.X("aduana:N", sort=aduanas),
        xOffset=alt.XOffset("serie:N", sort=["Emitida", "Recibida"]),
        y="valor:Q", text="etiqueta:N",
    )
    chart_ad = (barras + etiq_ad).properties(height=580)
    st.altair_chart(tema(chart_ad), use_container_width=True)

    st.divider()
    st.subheader("Margen por aduana")
    st.caption("Izquierda: margen en pesos (emitida − recibida). Derecha: margen como % de lo emitido, "
               "para comparar aduanas de cualquier tamaño (sin barra de % si la aduana no tiene emitida). "
               "Azul = positivo, rojo = negativo.")
    mg = pd.DataFrame({"aduana": aduanas, "em": y_em, "rec": y_rec})
    mg = mg[(mg["em"] > 0) | (mg["rec"] > 0)].copy()
    mg["margen"] = mg["em"] - mg["rec"]
    mg["pct"] = mg.apply(lambda r: r["margen"] / r["em"] if r["em"] else None, axis=1)
    mg = mg.sort_values("margen", ascending=False)   # mayor margen arriba (orden Altair: primero = arriba)
    mg["color"] = mg["margen"].map(lambda m: JOCS_AZUL if m >= 0 else C_NEG)
    mg["etiqueta"] = mg["margen"].map(fmt_m)
    mg["pct_recortado"] = mg["pct"].map(lambda p: None if p is None else max(p, -1.0))
    mg["pct_etiqueta"] = mg["pct"].map(lambda p: "s/ emitida" if p is None else f"{p*100:.0f}%")
    mg["pct_color"] = mg["pct"].map(lambda p: JOCS_AZUL if (p is not None and p >= 0) else C_NEG)
    orden_mg = mg["aduana"].tolist()
    pcts = [p for p in mg["pct"] if p is not None]
    mn, mx = float(mg["margen"].min()), float(mg["margen"].max())
    alto_mg = max(400, 34 * len(mg) + 40)

    panel_pesos = alt.Chart(mg).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
        y=alt.Y("aduana:N", sort=orden_mg, title=None),
        x=alt.X("margen:Q", title=None, axis=eje_dinero(),
                scale=alt.Scale(domain=[min(0, mn * 1.9), mx * 1.3])),
        color=alt.Color("color:N", scale=None, legend=None),
        tooltip=[alt.Tooltip("aduana:N", title="Aduana"), alt.Tooltip("etiqueta:N", title="Margen"),
                 alt.Tooltip("em:Q", title="Emitida", format=",.0f"),
                 alt.Tooltip("rec:Q", title="Recibida", format=",.0f")],
    )
    etiq_pesos = alt.Chart(mg).mark_text(align="left", dx=5, fontSize=11, color="#3d4752").encode(
        y=alt.Y("aduana:N", sort=orden_mg), x="margen:Q", text="etiqueta:N",
    )
    panel_izq = (panel_pesos + etiq_pesos).properties(height=alto_mg, title="Margen ($)")

    panel_pct = alt.Chart(mg).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
        y=alt.Y("aduana:N", sort=orden_mg, title=None, axis=None),
        x=alt.X("pct_recortado:Q", title=None, axis=alt.Axis(format=".0%"),
                scale=alt.Scale(domain=[min(-0.1, max(-1.0, min(pcts, default=0)) * 1.3), 1.25])),
        color=alt.Color("pct_color:N", scale=None, legend=None),
        tooltip=[alt.Tooltip("aduana:N", title="Aduana"), alt.Tooltip("pct_etiqueta:N", title="Margen %")],
    )
    etiq_pct = alt.Chart(mg).mark_text(align="left", dx=5, fontSize=11, color="#3d4752").encode(
        y=alt.Y("aduana:N", sort=orden_mg), x="pct_recortado:Q", text="pct_etiqueta:N",
    )
    panel_der = (panel_pct + etiq_pct).properties(height=alto_mg, title="Margen (% de lo emitido)")

    chart_mg = alt.hconcat(panel_izq, panel_der).resolve_scale(color="independent")
    st.altair_chart(tema(chart_mg), use_container_width=True)

    st.divider()
    st.subheader("Centros de costo (CC)")
    v1, v2 = st.columns(2)
    with v1:
        tipo_vista = st.radio("Ver", ["Emitida", "Recibida", "Margen"], horizontal=True, key="cc_tipo")
    es_margen = tipo_vista == "Margen"
    with v2:
        if es_margen:
            st.caption("El margen puede ser negativo, por eso se muestra siempre en escala normal.")
            escala_cc = "Normal"
        else:
            escala_cc = st.radio("Escala ", ["Logarítmica", "Normal"], horizontal=True, key="escala_cc")

    claves = ["cc", "aduana", "tipo_cc"]
    em_cc = EM.groupby(claves, dropna=False)["importe"].sum().rename("em").reset_index()
    rec_cc = REC.groupby(claves, dropna=False)["importe"].sum().rename("rec").reset_index()
    cc_g = em_cc.merge(rec_cc, on=claves, how="outer").fillna({"em": 0.0, "rec": 0.0})
    if es_margen:
        cc_g["valor"] = cc_g["em"] - cc_g["rec"]
        cc_g = cc_g[cc_g["valor"].abs() >= 1]
    else:
        cc_g["valor"] = cc_g["em"] if tipo_vista == "Emitida" else cc_g["rec"]
        cc_g = cc_g[cc_g["valor"] >= 1]
    cc_g = cc_g.sort_values("valor", ascending=False)   # mayor arriba (orden Altair: primero = arriba)
    if cc_g.empty:
        st.info("Sin datos para mostrar.")
    else:
        cc_g = cc_g.copy()
        cc_g["etiq"] = [("Sin CC" if pd.isna(c) else str(int(c))) + f" · {a}"
                        + ("" if SOLO_PROPIOS else f" ({t})")
                        for c, a, t in zip(cc_g["cc"], cc_g["aduana"], cc_g["tipo_cc"])]
        orden_cc = cc_g["etiq"].tolist()
        if es_margen:
            cc_g["color"] = cc_g["valor"].map(lambda v: JOCS_AZUL if v >= 0 else C_NEG)
            # el % solo se muestra si la emitida es real (>= $1) y el % no es absurdo (emitida de centavos)
            cc_g["texto"] = [fmt_m(v) + (f"  ·  {v / e * 100:.0f}%" if e >= 1 and abs(v / e) <= 10 else "")
                             for v, e in zip(cc_g["valor"], cc_g["em"])]
        else:
            tot_cc = cc_g["valor"].sum()
            cc_g["color"] = C_EMITIDA if tipo_vista == "Emitida" else C_RECIBIDA
            cc_g["texto"] = [f"{fmt_m(v)}  ·  {v / tot_cc * 100:.1f}%" for v in cc_g["valor"]]

        lo, hi = float(cc_g["valor"].min()), float(cc_g["valor"].max())
        h_cc = max(400, 30 * len(cc_g) + 90)
        x2_kw = {}
        if escala_cc == "Logarítmica":
            dom_min_cc = 10 ** (math.log10(lo) - 0.8)
            escala_x = alt.Scale(type="log", base=10, nice=False,
                                 domain=[dom_min_cc, 10 ** (math.log10(hi) + 0.9)])
            # Igual que en la gráfica de arriba: sin x2 explícito, las barras en log no se ven.
            x2_kw = {"x2": alt.X2(datum=dom_min_cc)}
        elif es_margen:
            escala_x = alt.Scale(domain=[min(0.0, lo * 1.9), hi * 1.35])
        else:
            escala_x = alt.Scale(domain=[0, hi * 1.3])

        barras_cc = alt.Chart(cc_g).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
            y=alt.Y("etiq:N", title=None, sort=orden_cc),
            x=alt.X("valor:Q", title=None, axis=eje_dinero(), scale=escala_x),
            color=alt.Color("color:N", scale=None, legend=None),
            tooltip=[alt.Tooltip("etiq:N", title="Centro de costo"),
                     alt.Tooltip("em:Q", title="Emitida", format=",.0f"),
                     alt.Tooltip("rec:Q", title="Recibida", format=",.0f")],
            **x2_kw,
        )
        etiq_cc = alt.Chart(cc_g).mark_text(align="left", dx=5, fontSize=11, color="#3d4752").encode(
            y=alt.Y("etiq:N", sort=orden_cc), x="valor:Q", text="texto:N",
        )
        chart_cc = (barras_cc + etiq_cc).properties(height=h_cc)
        st.altair_chart(tema(chart_cc), use_container_width=True)
        st.caption("Margen: emitida − recibida de cada centro de costo; a la derecha del monto, "
                   "el margen como % de lo emitido." if es_margen else
                   "Sobre cada barra: monto y % del total de la vista (emitida o recibida).")

# ══════════════════════════════════════════════════════════════════════════════
# MODELO Y ORIGINACIÓN
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Modelo y Originación":
    st.title("Modelo y Originación")
    st.caption(f"{', '.join(sel_labels)} · Asesoría aduanal, corresponsalías y costos de originación")
    st.caption("Ingreso Asesoría = costos de referencia (hoja 'costos de referencia'); "
               "Gasto Corresponsalía = facturas de corresponsalías (hoja 'corresponsalias'). "
               "El documento de requerimientos no deja claro qué hoja alimenta 'Ingreso Asesoría' — "
               "se probó primero con 'honorarios josefa' y dio un margen neto muy negativo e "
               "implausible (\\$53k de ingreso contra \\$387k de gasto), así que se usó 'costos de "
               "referencia' en su lugar (monto del mismo orden que el gasto). Avisa si no es la "
               "lectura correcta.")

    st.subheader("Corresponsalías vs. Asesoría Aduanal")
    ing_asesoria = costos_ref_f["monto"].sum() if not costos_ref_f.empty else 0.0
    gasto_corres = corres_f["monto"].sum() if not corres_f.empty else 0.0
    margen_neto_mo = ing_asesoria - gasto_corres

    c1, c2, c3 = st.columns(3)
    c1.metric("Ingreso Asesoría", fmt_m(ing_asesoria))
    c2.metric("Gasto Corresponsalía", fmt_m(gasto_corres))
    c3.metric("Margen Neto", fmt_m(margen_neto_mo),
              delta=pct(margen_neto_mo / ing_asesoria) if ing_asesoria else None)

    if costos_ref_f.empty and corres_f.empty:
        st.info("Sin datos de corresponsalías/asesoría para el periodo y aduana seleccionados.")
    else:
        ing_ad = (costos_ref_f.groupby("aduana")["monto"].sum()
                 if not costos_ref_f.empty else pd.Series(dtype=float))
        gas_ad = corres_f.groupby("aduana")["monto"].sum() if not corres_f.empty else pd.Series(dtype=float)
        aduanas_mo = sorted(set(ing_ad.index) | set(gas_ad.index),
                            key=lambda a: -(ing_ad.get(a, 0) + gas_ad.get(a, 0)))
        df_mo = pd.DataFrame({
            "aduana": aduanas_mo,
            "ingreso": [ing_ad.get(a, 0.0) for a in aduanas_mo],
            "gasto": [gas_ad.get(a, 0.0) for a in aduanas_mo],
        })
        df_mo["margen"] = df_mo["ingreso"] - df_mo["gasto"]
        df_mo["etiqueta"] = df_mo["margen"].map(fmt_m)
        df_barras_mo = pd.concat([
            pd.DataFrame({"aduana": df_mo["aduana"], "serie": "Ingreso Asesoría", "valor": df_mo["ingreso"]}),
            pd.DataFrame({"aduana": df_mo["aduana"], "serie": "Gasto Corresponsalía", "valor": df_mo["gasto"]}),
        ], ignore_index=True)
        df_barras_mo["etq"] = df_barras_mo["valor"].map(fmt_m)
        orden_mo = df_mo["aduana"].tolist()

        barras_mo = alt.Chart(df_barras_mo).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X("aduana:N", title=None, sort=orden_mo, axis=alt.Axis(labelAngle=-30)),
            xOffset=alt.XOffset("serie:N", sort=["Ingreso Asesoría", "Gasto Corresponsalía"]),
            y=alt.Y("valor:Q", title=None, axis=eje_dinero()),
            color=alt.Color("serie:N", title=None,
                            scale=alt.Scale(domain=["Ingreso Asesoría", "Gasto Corresponsalía"],
                                            range=[JOCS_VERDE, C_NEG]),
                            legend=alt.Legend(orient="top", direction="horizontal")),
            tooltip=[alt.Tooltip("aduana:N", title="Aduana"), alt.Tooltip("serie:N", title="Serie"),
                     alt.Tooltip("etq:N", title="Monto")],
        )
        linea_mo = alt.Chart(df_mo).mark_line(
            color=C_MARGEN, strokeWidth=3,
            point=alt.OverlayMarkDef(filled=True, size=70, color=C_MARGEN)).encode(
            x=alt.X("aduana:N", sort=orden_mo), y=alt.Y("margen:Q"),
            tooltip=[alt.Tooltip("aduana:N", title="Aduana"),
                     alt.Tooltip("etiqueta:N", title="Margen Neto")],
        )
        etiq_mo = alt.Chart(df_mo).mark_text(dy=-10, fontSize=10, fontWeight=600, color=C_MARGEN).encode(
            x=alt.X("aduana:N", sort=orden_mo), y="margen:Q", text="etiqueta:N",
        )
        chart_mo = (barras_mo + linea_mo + etiq_mo).properties(height=420)
        st.altair_chart(tema(chart_mo), use_container_width=True)

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
# NÓMINA
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Nómina":
    st.title("Nómina (Capital Humano)")
    st.caption(f"{', '.join(sel_labels)} · Nómina real por empleado (hoja 'sueldos')")

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
    por_aduana_nom = sueldos_f.groupby("aduana")["total"].sum().sort_values(ascending=False)
    aduana_nom_sel = st.selectbox("Seleccionar Aduana", por_aduana_nom.index.tolist(),
                                  format_func=lambda a: f"{a}  ·  {fmt_m(por_aduana_nom[a])}")
    tabla_nom = sueldos_f[sueldos_f["aduana"] == aduana_nom_sel].copy()
    tot_aduana_nom = tabla_nom["total"].sum()
    tabla_nom["pct"] = tabla_nom["total"] / tot_aduana_nom if tot_aduana_nom else 0.0
    tabla_nom = tabla_nom.sort_values("total", ascending=False)
    st.caption(f"Total nómina {aduana_nom_sel}: {fmt_m(tot_aduana_nom)} · {len(tabla_nom)} empleado(s)")
    st.dataframe(pd.DataFrame({
        "Empleado": tabla_nom["nombre"],
        "Sueldo Total": tabla_nom["total"].map(fmt),
        "% Conformación Aduana": tabla_nom["pct"].map(lambda v: f"{v*100:.1f}%"),
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

    df_cascada = pd.DataFrame({
        "concepto": [et for _, et in COLS_NOMINA],
        "valor": [fila_emp_sum[c] for c, _ in COLS_NOMINA],
    })
    st.altair_chart(
        tema(waterfall_chart(df_cascada, "concepto", "valor",
                             color_pos=JOCS_AZUL, color_neg=C_NEG, color_total=JOCS_VERDE,
                             total_label="COSTO TOTAL NETO", height=420)),
        use_container_width=True)
    # st.caption() renderiza markdown: dos o mas "$" sueltos en el mismo texto los
    # interpreta como un bloque de LaTeX y se come los signos de pesos. Se escapan
    # a "\$" para que se vean literales (aqui se juntan 10 montos en una sola linea).
    partes_txt = " + ".join(fmt_m(fila_emp_sum[c]).replace("$", "\\$") for c, _ in COLS_NOMINA)
    total_txt = fmt_m(fila_emp_sum["total"]).replace("$", "\\$")
    st.caption(f"{partes_txt} = {total_txt} (Costo Total Neto)")

# ══════════════════════════════════════════════════════════════════════════════
# DETALLE DE FACTURAS
# ══════════════════════════════════════════════════════════════════════════════
if pagina == "Detalle de Facturas":
    st.title("Detalle de Facturas")
    st.caption(f"{', '.join(sel_labels)} · Emitidas y recibidas juntas, línea por línea")

    f1, f2, f3 = st.columns([1.2, 1.2, 2])
    with f1:
        tipos_sel = st.multiselect("Tipo", ["EMITIDA", "RECIBIDA"],
                                   default=["EMITIDA", "RECIBIDA"])
    with f2:
        aduanas_disp = sorted(dff["aduana"].dropna().unique().tolist())
        aduanas_sel = st.multiselect("Aduana", aduanas_disp, default=[])
    with f3:
        buscar = st.text_input("Buscar (cliente, proveedor, descripción o referencia)", "")

    f4, f5, f6 = st.columns([1, 1.4, 1.6])
    with f4:
        incluir_cancel = st.checkbox("Incluir canceladas", value=False)
    with f5:
        incluir_sin_monto = st.checkbox("Incluir líneas sin monto (muestra su importe del Excel)",
                                        value=False)
    with f6:
        min_importe = st.number_input("Importe mínimo", min_value=0, value=0, step=1000)

    base = dff if incluir_cancel else dff[~dff["cancelada"]]
    if not incluir_sin_monto:
        base = base[~base["sin_monto"]]
    base = base.assign(importe=base["importe"].where(~base["sin_monto"], base["importe_sin_monto"]))
    fd = base[base["tipo_factura"].isin(tipos_sel)] if tipos_sel else base.iloc[0:0]
    if aduanas_sel:
        fd = fd[fd["aduana"].isin(aduanas_sel)]
    if buscar.strip():
        s = buscar.strip().lower()
        mask = (fd["contraparte"].fillna("").str.lower().str.contains(s, regex=False) |
                fd["descripcion"].fillna("").str.lower().str.contains(s, regex=False) |
                fd["referencia"].fillna("").str.lower().str.contains(s, regex=False))
        fd = fd[mask]
    if min_importe > 0:
        fd = fd[fd["importe"] >= min_importe]

    fd = fd.sort_values("importe", ascending=False)
    st.caption(f"{len(fd):,} líneas · {fmt(fd['importe'].sum())} en total")

    MAX_FILAS = 500
    if len(fd) > MAX_FILAS:
        st.info(f"Mostrando las primeras {MAX_FILAS:,} líneas de {len(fd):,} (ordenadas por importe).")
    mostrar = fd.head(MAX_FILAS)

    tabla = pd.DataFrame({
        "Tipo": mostrar["tipo_factura"],
        "Cliente / Proveedor": mostrar["contraparte"].fillna(""),
        "Descripción": mostrar["descripcion"].fillna(""),
        "Importe": mostrar["importe"].map(fmt),
        "Aduana": mostrar["aduana"],
        "CC": mostrar["cc"].map(lambda c: "" if pd.isna(c) else str(int(c))),
        "Propios / Terceros": mostrar["tipo_cc"],
        "Concepto": mostrar["concepto_cierre"].fillna(""),
        "Referencia": mostrar["referencia"].fillna(""),
    })
    st.dataframe(tabla, use_container_width=True, hide_index=True, height=560)
