import json
import math
import re
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
    v = v + 0.0  # normaliza -0.0 a 0.0 (si no, ":,.0f" imprime "-0")
    return f"-${-v:,.0f}" if v < 0 else f"${v:,.0f}"

# A pedido del usuario (2026-10-14) ya no se abrevia ningún monto en el
# tablero ("quita el redondeo de cifras" en todo el tablero, no solo en una
# tabla) -- fmt_m() se deja como alias de fmt() en vez de duplicar su lógica,
# así que cualquier helper que use 'formateador=fmt_m' por default también
# queda exacto automáticamente.
fmt_m = fmt

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

# ── Altair (solo para la dona de Gastos por Concepto -- ver nota mas abajo) ────
FUENTE = "system-ui, 'Segoe UI', Roboto, sans-serif"

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

# ── Recharts (piloto 2026-10-13, a pedido del usuario) ─────────────────────────
# No hay paquete de Streamlit ya hecho para esto (Recharts es de React/JS); se
# incrusta a mano via st.components.v1.html() cargando React/ReactDOM/Recharts
# desde unpkg y armando el arbol con React.createElement (NO JSX -- JSX
# necesitaria compilar con Babel standalone en el navegador en cada rerun, mucho
# mas pesado que algo ya compilado). Cada llamada crea su propio iframe aislado
# (sin estado/seleccion compartido con Streamlit ni con otros charts).
RECHARTS_CDN = """
<script src="https://unpkg.com/react@18.3.1/umd/react.production.min.js"></script>
<script src="https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js"></script>
<script src="https://unpkg.com/prop-types@15.8.1/prop-types.min.js"></script>
<script src="https://unpkg.com/recharts@2.15.0/umd/Recharts.js"></script>
"""

RECHARTS_FMT_JS = """
function fmtMoney(v) {
  const s = v < 0 ? '-' : '';
  return s + '$' + Math.round(Math.abs(v)).toLocaleString('en-US');
}
"""

def recharts_component(render_js, data, height=400):
    """Incrusta un chart de Recharts. 'render_js' arma el arbol de elementos con
    React.createElement y lo monta con ReactDOM.createRoot(...).render(...); tiene
    disponibles 'data' (ya inyectada como JSON), 'Recharts', 'React', 'ReactDOM'
    y 'fmtMoney' (monto exacto con separador de miles, igual que fmt()/fmt_m()
    del lado Python -- sin abreviar a M/k)."""
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">{RECHARTS_CDN}
<style>body{{margin:0;font-family:system-ui,'Segoe UI',Roboto,sans-serif;}}</style>
</head><body>
<div id="root"></div>
<script>
{RECHARTS_FMT_JS}
const data = {json.dumps(data)};
{render_js}
</script>
</body></html>"""
    components.html(html, height=height, scrolling=False)

def gauge_svg(frac, height=200, bandas=None, etiqueta="del gasto facturado"):
    """Velocímetro (gauge) semicircular con aguja -- SVG puro (sin
    React/Recharts: la aguja es el único dato variable, no hace falta nada
    interactivo). 'frac' en [0,1] (se clampa). 'bandas' (opcional): lista de
    (fracción_de_ancho, color) que suman 1.0, de izquierda (bajo) a derecha
    (alto); por default 3 tercios rojo/ámbar/verde."""
    frac = max(0.0, min(1.0, frac))
    bandas = bandas or [(1 / 3, C_NEG), (1 / 3, "#e8a93c"), (1 / 3, JOCS_VERDE)]
    cx, cy, r_out, r_in = 120, 120, 105, 68

    def punto(r, ang_deg):
        rad = math.radians(ang_deg)
        return cx + r * math.cos(rad), cy - r * math.sin(rad)

    def arco(a_ini, a_fin):
        x1o, y1o = punto(r_out, a_ini)
        x2o, y2o = punto(r_out, a_fin)
        x1i, y1i = punto(r_in, a_fin)
        x2i, y2i = punto(r_in, a_ini)
        return (f"M {x1o:.2f} {y1o:.2f} A {r_out} {r_out} 0 0 1 {x2o:.2f} {y2o:.2f} "
                f"L {x1i:.2f} {y1i:.2f} A {r_in} {r_in} 0 0 0 {x2i:.2f} {y2i:.2f} Z")

    piezas = []
    a = 180.0
    for proporcion, color in bandas:
        a2 = a - proporcion * 180.0
        piezas.append(f'<path d="{arco(a, a2)}" fill="{color}" />')
        a = a2

    ang_aguja = 180.0 - frac * 180.0
    xa, ya = punto(r_in - 6, ang_aguja)
    fuente = "system-ui,'Segoe UI',Roboto,sans-serif"
    # viewBox alto 185 (no 150): el texto va en cy+34/cy+52 = 154/172, por debajo
    # de donde terminan los arcos (cy=120) -- con un viewBox mas bajo que eso el
    # texto queda recortado fuera del area visible. Ancho/alto del <svg> FIJOS
    # (no '100%'): si se deja responsivo al ancho de la columna, el alto
    # renderizado (que sigue el aspect ratio del viewBox) deja de coincidir con
    # el 'height' fijo en pixeles del iframe de components.html, y el iframe
    # recorta lo que sobra -- con tamaño fijo ambos quedan sincronizados siempre.
    svg = f"""<svg viewBox="0 0 240 185" width="260" height="200"
                   xmlns="http://www.w3.org/2000/svg" style="display:block;margin:0 auto;">
      {"".join(piezas)}
      <line x1="{cx}" y1="{cy}" x2="{xa:.2f}" y2="{ya:.2f}" stroke="{JOCS_AZUL}"
            stroke-width="4" stroke-linecap="round" />
      <circle cx="{cx}" cy="{cy}" r="7" fill="{JOCS_AZUL}" />
      <text x="{cx}" y="{cy + 34}" text-anchor="middle" font-size="26" font-weight="700"
            fill="{JOCS_AZUL}" font-family="{fuente}">{frac * 100:.0f}%</text>
      <text x="{cx}" y="{cy + 52}" text-anchor="middle" font-size="12" fill="#54616f"
            font-family="{fuente}">{etiqueta}</text>
    </svg>"""
    components.html(f'<div style="font-family:{fuente};">{svg}</div>', height=height)

def nested_donut_recharts(monto_facturado, monto_no_facturado, por_categoria, height=380,
                          color_fact=JOCS_AZUL, color_no_fact=C_NEG):
    """'Two Level Pie Chart' (ver TABLERO.pdf, pestaña Tarjeta de Crédito):
    anillo interior = Facturado vs No Facturado (2 rebanadas); anillo exterior
    = 'por_categoria' (Series ya agrupada/ordenada, ej. gasto por proveedor).
    Tooltip a la medida en ambos anillos; leyenda manual solo del anillo
    interior (el exterior puede traer muchas categorías, se explora con el
    tooltip en vez de una leyenda larga)."""
    data_int = [{"nombre": "Facturado", "valor": float(monto_facturado), "color": color_fact},
               {"nombre": "No Facturado", "valor": float(monto_no_facturado), "color": color_no_fact}]
    data_ext = [{"nombre": str(k), "valor": float(v), "color": color_for(i)}
               for i, (k, v) in enumerate(por_categoria.items())]
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, PieChart, Pie, Cell, Tooltip }} = Recharts;
const dataInt = {json.dumps(data_int)};
const dataExt = {json.dumps(data_ext)};

function TooltipDonut(props) {{
  const {{ active, payload }} = props;
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return h('div', {{ style: {{ background: '#fff', border: '1px solid #d7dce1',
                             borderRadius: 6, padding: '8px 12px', fontSize: 13 }} }},
    h('div', {{ style: {{ fontWeight: 700, color: '{JOCS_AZUL}' }} }}, d.nombre),
    h('div', {{}}, fmtMoney(d.valor))
  );
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h('div', {{}},
    h('div', {{ style: {{ display: 'flex', gap: 14, justifyContent: 'center', marginBottom: 6 }} }},
      h('div', {{ style: {{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12,
                         color: '#3d4752' }} }},
        h('span', {{ style: {{ width: 10, height: 10, borderRadius: 2, background: '{color_fact}',
                             display: 'inline-block' }} }}), 'Facturado'),
      h('div', {{ style: {{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12,
                         color: '#3d4752' }} }},
        h('span', {{ style: {{ width: 10, height: 10, borderRadius: 2, background: '{color_no_fact}',
                             display: 'inline-block' }} }}), 'No Facturado')
    ),
    h(ResponsiveContainer, {{ width: '100%', height: {height - 30} }},
      h(PieChart, {{}},
        h(Tooltip, {{ content: h(TooltipDonut) }}),
        h(Pie, {{ data: dataInt, dataKey: 'valor', nameKey: 'nombre', cx: '50%', cy: '50%',
                 innerRadius: 0, outerRadius: 70, isAnimationActive: false }},
          dataInt.map((d, i) => h(Cell, {{ key: i, fill: d.color }}))
        ),
        h(Pie, {{ data: dataExt, dataKey: 'valor', nameKey: 'nombre', cx: '50%', cy: '50%',
                 innerRadius: 85, outerRadius: 130, isAnimationActive: false }},
          dataExt.map((d, i) => h(Cell, {{ key: i, fill: d.color }}))
        )
      )
    )
  )
);
"""
    recharts_component(render_js, [], height=height)

def _ancho_eje_categoria(labels, min_w=90, max_w=240, por_char=6.5):
    return int(min(max_w, max(min_w, por_char * max((len(str(l)) for l in labels), default=10))))

def bar_h_recharts(serie, color, height=None, formateador=fmt_m):
    """Version Recharts de barra_h(): barra horizontal simple + etiqueta a la derecha."""
    df = serie.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    df["etiqueta"] = df["valor"].map(formateador)
    h_px = height or max(280, 30 * len(df) + 60)
    data = [{"label": str(r.label), "valor": float(r.valor), "etiqueta": r.etiqueta}
            for r in df.itertuples()]
    ancho_eje = _ancho_eje_categoria(df["label"])
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, LabelList }} = Recharts;
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {h_px - 10} }},
    h(BarChart, {{ data: data, layout: 'vertical', margin: {{ top: 10, right: 70, left: 10, bottom: 10 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', horizontal: false }}),
      h(XAxis, {{ type: 'number', tickFormatter: fmtMoney, tick: {{ fontSize: 11, fill: '#54616f' }},
                 axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ dataKey: 'label', type: 'category', width: {ancho_eje},
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(Tooltip, {{ formatter: (v, n, p) => [p.payload.etiqueta, p.payload.label] }}),
      h(Bar, {{ dataKey: 'valor', fill: '{color}', radius: [0, 3, 3, 0], isAnimationActive: false }},
        h(LabelList, {{ dataKey: 'etiqueta', position: 'right', fontSize: 11, fill: '#3d4752' }})
      )
    )
  )
);
"""
    recharts_component(render_js, data, height=h_px)

def pareto_recharts(serie, color, height=420, umbral=0.8, color_neg=None):
    """Version Recharts de pareto_chart(): barras descendentes + linea de %
    acumulado (ComposedChart con dos ejes Y). 'color_neg' (opcional): si se pasa,
    'serie' puede traer negativos -- se ordena por magnitud absoluta (para que
    las barras negativas no queden todas al final) y cada barra se colorea
    'color' o 'color_neg' segun su signo real (ej. Utilidad Operativa, que puede
    ser negativa por aduana). La línea de % acumulado, en ese caso, se arma SOLO
    con las barras positivas (color): las negativas no suman nada a la línea (se
    queda plana en ese punto) -- un acumulado sobre valor absoluto no tenía
    sentido de negocio (restar una pérdida la hacía "ayudar" a llegar al 100%)."""
    df = serie.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    orden_col = df["valor"].abs() if color_neg is not None else df["valor"]
    df = df.assign(_orden=orden_col).sort_values("_orden", ascending=False).drop(columns="_orden")
    valores_acum = df["valor"].clip(lower=0) if color_neg is not None else df["valor"]
    tot = valores_acum.sum() or 1.0
    df["acumulado_monto"] = valores_acum.cumsum()
    df["acumulado"] = df["acumulado_monto"] / tot
    df["etiqueta"] = df["valor"].map(fmt_m)
    df["color"] = color if color_neg is None else df["valor"].map(lambda v: color if v >= 0 else color_neg)
    data = [{"label": str(r.label), "valor": float(r.valor), "etiqueta": r.etiqueta,
             "acumulado": float(r.acumulado), "acumulado_monto": float(r.acumulado_monto),
             "color": r.color} for r in df.itertuples()]
    # Los dos ejes Y (montos a la izquierda, % acumulado a la derecha) no
    # comparten "0" solos -- cada uno escala a su propio rango completo, así que
    # sin esto el "0%" de la derecha cae en el PISO del plot, no a la altura del
    # "$0" de la izquierda (que puede estar a la mitad si hay barras negativas).
    # El punto ámbar de la primera aduana termina viéndose "abajo", lejos de su
    # propia barra azul. Fix: se calcula a mano el dominio del eje derecho para
    # que su "0" quede exactamente a la misma altura que el "$0" del eje
    # izquierdo (mismo truco que alinear dos ejes Y en cualquier combo chart).
    techo_pos = max(float(df["valor"].max()), 0.0) * 1.15 or 1.0
    techo_neg = min(float(df["valor"].min()), 0.0) * 1.15
    frac_cero = (0.0 - techo_neg) / (techo_pos - techo_neg) if techo_pos != techo_neg else 0.0
    r_max = 1.05
    r_min = (frac_cero * r_max / (frac_cero - 1.0)) if 0.0 <= frac_cero < 1.0 else 0.0
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip,
         LabelList, ReferenceLine, Cell }} = Recharts;

function TooltipPareto(props) {{
  const {{ active, payload, label }} = props;
  if (!active || !payload || !payload.length) return null;
  const fila = payload[0].payload;
  return h('div', {{ style: {{ background: '#fff', border: '1px solid #d7dce1',
                             borderRadius: 6, padding: '8px 12px', fontSize: 13 }} }},
    h('div', {{ style: {{ fontWeight: 700, marginBottom: 4, color: '{JOCS_AZUL}' }} }}, label),
    h('div', {{}}, 'Monto : ' + fmtMoney(fila.valor)),
    h('div', {{}}, 'Monto acumulado : ' + fmtMoney(fila.acumulado_monto)),
    h('div', {{}}, '% acumulado : ' + Math.round(fila.acumulado * 100) + '%')
  );
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(ComposedChart, {{ data: data, margin: {{ top: 28, right: 40, left: 10, bottom: 50 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
      h(XAxis, {{ dataKey: 'label', angle: -30, textAnchor: 'end', interval: 0,
                 tick: {{ fontSize: 10, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ yAxisId: 'izq', domain: [{techo_neg}, {techo_pos}], tickFormatter: fmtMoney,
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ yAxisId: 'der', orientation: 'right', domain: [{r_min}, {r_max}],
                 tickFormatter: v => Math.round(v * 100) + '%',
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(Tooltip, {{ content: h(TooltipPareto) }}),
      h(ReferenceLine, {{ yAxisId: 'der', y: {umbral}, stroke: '{C_NEG}', strokeDasharray: '4 3' }}),
      h(Bar, {{ yAxisId: 'izq', dataKey: 'valor', radius: [3, 3, 0, 0], isAnimationActive: false }},
        data.map((d, i) => h(Cell, {{ key: i, fill: d.color }})),
        h(LabelList, {{ dataKey: 'etiqueta', position: 'top', fontSize: 10, fill: '#3d4752' }})
      ),
      h(Line, {{ yAxisId: 'der', type: 'monotone', dataKey: 'acumulado', stroke: '{C_MARGEN}',
                strokeWidth: 2.5, dot: {{ r: 4, fill: '{C_MARGEN}' }}, isAnimationActive: false }})
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

def waterfall_recharts(df_in, label_col, valor_col, color_pos, color_neg, color_total,
                       total_label="TOTAL", height=380):
    """Version Recharts de waterfall_chart(): misma logica de inicio/fin acumulado,
    pero como BarChart de Recharts con barras 'flotantes' (dataKey=[low,high] +
    'shape' a la medida para colorear) -- la tecnica que documenta Recharts para
    graficas de cascada."""
    df = df_in[[label_col, valor_col]].copy()
    df.columns = ["label", "valor"]
    total = df["valor"].sum()
    df = pd.concat([df, pd.DataFrame({"label": [total_label], "valor": [total]})], ignore_index=True)
    acumulado, inicio = 0.0, []
    for i, v in enumerate(df["valor"]):
        if i == len(df) - 1:
            inicio.append(0.0)
        else:
            inicio.append(acumulado)
            acumulado += v
    df["inicio"] = inicio
    df["fin"] = df["inicio"] + df["valor"]
    df["tipo"] = ["Total" if l == total_label else ("Resta" if v < 0 else "Suma")
                 for l, v in zip(df["label"], df["valor"])]
    df["etiqueta"] = df["valor"].map(fmt_m)

    data = [
        {"label": r.label, "range": [min(r.inicio, r.fin), max(r.inicio, r.fin)],
         "valor": r.valor, "tipo": r.tipo, "etiqueta": r.etiqueta}
        for r in df.itertuples()
    ]
    colores = json.dumps({"Suma": color_pos, "Resta": color_neg, "Total": color_total})

    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, LabelList }} = Recharts;
const COLORES = {colores};

function CustomBar(props) {{
  const {{ x, y, width, height, payload }} = props;
  return h('rect', {{ x, y, width, height, fill: COLORES[payload.tipo], rx: 3, ry: 3 }});
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 20} }},
    h(BarChart, {{ data: data, margin: {{ top: 28, right: 20, left: 10, bottom: 10 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
      h(XAxis, {{ dataKey: 'label', tick: {{ fontSize: 11, fill: '#54616f' }},
                 axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ tickFormatter: fmtMoney, tick: {{ fontSize: 11, fill: '#54616f' }},
                 axisLine: {{ stroke: '#d7dce1' }} }}),
      h(Tooltip, {{ formatter: (v, n, p) => [p.payload.etiqueta, p.payload.label] }}),
      h(Bar, {{ dataKey: 'range', shape: CustomBar, isAnimationActive: false }},
        h(LabelList, {{ dataKey: 'etiqueta', position: 'top', fontSize: 11, fill: '#3d4752' }})
      )
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

def bullet_uo_recharts(df, categoria_col, total_col, segmentos, egreso_col, utilidad_col,
                       color_ingreso="#27ae60", color_utilidad="#f2c94c", height=420):
    """Por cada categoría, una barra apilada desde abajo con 'segmentos' (lista de
    (etiqueta, color, columna), ej. Sueldos/Costos de Referencia/Recibidas =
    Egresos) + un segmento final 'Utilidad Operativa' (color 'color_utilidad')
    que rellena lo que falta para llegar a 'total_col' (ej. Ingreso) -- si los
    segmentos de gasto ya se pasaron de 'total_col' (pérdida), ese último
    segmento no se dibuja. Una marca horizontal 'color_ingreso' al nivel de
    'total_col' en cada barra señala el Ingreso, se haya alcanzado o no.
    Tooltip a la medida: 'total_col', 'egreso_col' (suma de los segmentos) y
    'utilidad_col'.
    Escala SIEMPRE logarítmica (17 aduanas de magnitudes muy distintas: $8k a
    $6M -- en escala lineal casi todas quedan invisibles pegadas al eje). Los
    valores en 0 se "pisan" a un mínimo positivo (log(0) no existe, mismo
    patrón que radar_recharts). La conversión valor→píxel de cada segmento se
    deriva a mano de los 2 puntos que Recharts ya calculó para el Bar de
    'total_col' en esa fila (su 'y' y su base 'y+height') -- válida para
    cualquier valor intermedio porque los es la MISMA escala log compartida
    por toda la gráfica, no una propia de esa barra."""
    valores_todos = ([float(row[total_col]) for _, row in df.iterrows()]
                     + [float(row[egreso_col]) for _, row in df.iterrows()])
    positivos = [v for v in valores_todos if v > 0]
    piso = (min(positivos) / 3) if positivos else 0.01
    techo = (max(valores_todos) if valores_todos else 1.0) * 1.15

    data = []
    for _, row in df.iterrows():
        d = {"categoria": str(row[categoria_col]), "total": max(float(row[total_col]), piso),
             "egreso": float(row[egreso_col]), "utilidad": float(row[utilidad_col])}
        for _, _, c in segmentos:
            d[c] = float(row[c])
        data.append(d)

    pasos_segmento = "\n".join(f"""
  {{
    cum += Math.max(payload[{json.dumps(col)}], 0);
    const yTope = yParaValor(cum);
    elementos.push(h('rect', {{ key: {json.dumps(col)}, x, y: yTope, width,
                              height: Math.max(yAnterior - yTope, 0), fill: '{color}' }}));
    yAnterior = yTope;
  }}""" for _, color, col in segmentos)

    leyenda_items = ([(etq, color) for etq, color, _ in segmentos]
                     + [("Utilidad Operativa", color_utilidad), ("Ingreso", color_ingreso)])
    leyenda_js = ",".join(
        f"""h('div', {{ key: {json.dumps(nombre)}, style: {{ display: 'flex', alignItems: 'center',
                      gap: 6, fontSize: 12, color: '#3d4752' }} }},
          h('span', {{ style: {{ width: 10, height: 10, borderRadius: 2, background: '{color}',
                               display: 'inline-block' }} }}),
          {json.dumps(nombre)}
        )""" for nombre, color in leyenda_items)

    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip }} = Recharts;
const PISO = {piso};
const LOG_PISO = Math.log(PISO);
// 'y'/'height' que Recharts calcula por fila para dataKey='total' no sirven
// cuando esa fila esta pisada a PISO (Ingreso real = 0, ej. STAFF): el propio
// 'total' de esa fila queda degenerado (denom≈0) y la escala no se puede
// derivar DE ESA fila -- pero el eje Y es el MISMO para todas las barras, así
// que la escala (pixeles por unidad de log) y la base (pixel de PISO) son en
// realidad globales: se calculan una vez con la primera fila NO degenerada
// (data viene ordenada por 'total' descendente, así que la primera fila
// siempre sirve) y se reusan para las demás, incluida cualquier fila pisada.
let cacheYBase = null;
let cachePxPorLogUnit = null;

function BarraBullet(props) {{
  const {{ x, y, width, height, payload }} = props;
  const yBaseFila = y + height;
  const denom = Math.log(Math.max(payload.total, PISO)) - LOG_PISO;
  if (denom >= 1e-9) {{
    cacheYBase = yBaseFila;
    cachePxPorLogUnit = height / denom;
  }}
  const yBase = cacheYBase !== null ? cacheYBase : yBaseFila;
  const pxPorLogUnit = cachePxPorLogUnit !== null ? cachePxPorLogUnit : 0;
  function yParaValor(v) {{
    return yBase - (Math.log(Math.max(v, PISO)) - LOG_PISO) * pxPorLogUnit;
  }}
  const elementos = [];
  let cum = 0;
  let yAnterior = yBase;
  {pasos_segmento}
  const ingresoReal = payload.total;
  const yIngreso = yParaValor(ingresoReal);
  if (ingresoReal > cum) {{
    elementos.push(h('rect', {{ key: 'utilidad', x, y: yIngreso, width,
                              height: Math.max(yAnterior - yIngreso, 0), fill: '{color_utilidad}' }}));
  }}
  elementos.push(h('line', {{ key: 'tope', x1: x, x2: x + width, y1: yIngreso, y2: yIngreso,
                            stroke: '{color_ingreso}', strokeWidth: 2 }}));
  return h('g', {{}}, elementos);
}}

function TooltipBullet(props) {{
  const {{ active, payload, label }} = props;
  if (!active || !payload || !payload.length) return null;
  const fila = payload[0].payload;
  const colorUtilidad = fila.utilidad < 0 ? '{C_NEG}' : '{color_utilidad}';
  return h('div', {{ style: {{ background: '#fff', border: '1px solid #d7dce1',
                             borderRadius: 6, padding: '8px 12px', fontSize: 13 }} }},
    h('div', {{ style: {{ fontWeight: 700, marginBottom: 4, color: '{JOCS_AZUL}' }} }}, label),
    h('div', {{}}, 'Ingreso : ' + fmtMoney(fila.total)),
    h('div', {{}}, 'Egresos : ' + fmtMoney(fila.egreso)),
    h('div', {{ style: {{ color: colorUtilidad, fontWeight: 600 }} }},
      'Utilidad Operativa : ' + fmtMoney(fila.utilidad))
  );
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h('div', {{}},
    h('div', {{ style: {{ display: 'flex', flexWrap: 'wrap', gap: 14, justifyContent: 'center',
                       marginBottom: 6 }} }}, {leyenda_js}),
    h(ResponsiveContainer, {{ width: '100%', height: {height - 34} }},
      h(BarChart, {{ data: data, margin: {{ top: 10, right: 20, left: 10, bottom: 50 }} }},
        h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
        h(XAxis, {{ dataKey: 'categoria', angle: -30, textAnchor: 'end', interval: 0,
                   tick: {{ fontSize: 10, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
        h(YAxis, {{ scale: 'log', domain: [{piso}, {techo}], tickFormatter: fmtMoney,
                   tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
        h(Tooltip, {{ content: h(TooltipBullet) }}),
        h(Bar, {{ dataKey: 'total', shape: BarraBullet, isAnimationActive: false }})
      )
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

def radar_recharts(df, categoria_col, series_colores, height=420, formateador=fmt_m,
                   escala_log=False, serie_resaltar_negativo=None,
                   etiqueta_negativo="Pérdida Operativa"):
    """RadarChart (componente nativo de Recharts): un eje por cada valor de
    'categoria_col', un polígono por cada serie en 'series_colores'
    ({nombre_columna: color}). Pensado para comparar 2+ series no-negativas por
    categoría (ej. Emitidas vs Recibidas por aduana) -- un radar no representa
    bien valores negativos, para eso usar una barra con signo.
    'escala_log' (recomendado si hay categorías de ordenes de magnitud muy
    distintos -- si no, las chicas se ven aplastadas al centro, todas pegadas):
    el tooltip/etiqueta siempre muestra el monto real; los valores en 0 (o muy
    chicos) se "pisan" a un piso positivo solo para que el log tenga donde
    dibujarlos (log(0) no existe).
    'serie_resaltar_negativo' (opcional, nombre de una columna de
    'series_colores'): un polígono de Recharts es un solo color, no puede
    pintar un tramo de rojo -- en vez de eso, cada VÉRTICE de esa serie con
    valor real negativo se marca con un punto rojo, y el tooltip (recuadro al
    pasar el mouse) muestra esa línea en rojo con la etiqueta
    '{etiqueta_negativo}' en vez del nombre de la serie."""
    valores_todos = [float(row[s]) for _, row in df.iterrows() for s in series_colores]
    positivos = [v for v in valores_todos if v > 0]
    piso = (min(positivos) / 3) if escala_log and positivos else 0.0
    techo = max(valores_todos) if valores_todos else 1.0
    data = []
    for _, row in df.iterrows():
        d = {"categoria": str(row[categoria_col])}
        for serie in series_colores:
            v = float(row[serie])
            d[serie] = max(v, piso) if escala_log else v
            d[f"{serie}_etq"] = formateador(v)
            d[f"{serie}_neg"] = v < 0
        data.append(d)
    funciones_punto = []
    elementos_radar = []
    for serie, color in series_colores.items():
        if serie == serie_resaltar_negativo:
            nombre_fn = "Punto_" + "".join(c if c.isalnum() else "_" for c in serie)
            funciones_punto.append(f"""
function {nombre_fn}(props) {{
  // Recharts anida: props.payload = {{x,y,cx,cy,radius,angle,name,value, payload: <fila real>}}
  const {{ cx, cy }} = props;
  const fila = props.payload.payload;
  const neg = fila[{json.dumps(serie + "_neg")}];
  const color = neg ? '{C_NEG}' : '{color}';
  return h('circle', {{ cx, cy, r: 4, fill: color, stroke: '#FFFFFF', strokeWidth: 1 }});
}}""")
            elementos_radar.append(
                f"""h(Radar, {{ name: {json.dumps(serie)}, dataKey: {json.dumps(serie)},
                             stroke: '{color}', fill: '{color}', fillOpacity: 0.25,
                             dot: h({nombre_fn}), isAnimationActive: false }}),""")
        else:
            elementos_radar.append(
                f"""h(Radar, {{ name: {json.dumps(serie)}, dataKey: {json.dumps(serie)},
                             stroke: '{color}', fill: '{color}', fillOpacity: 0.25,
                             isAnimationActive: false }}),""")
    funciones_punto_js = "\n".join(funciones_punto)
    elementos_radar_js = "\n".join(elementos_radar)
    eje_radio_kw = (f"scale: 'log', domain: [{piso}, {techo * 1.15}],"
                    if escala_log else f"domain: [0, {techo * 1.15}],")
    # Tooltip a la medida: el de Recharts por default pinta cada linea del
    # color de su serie -- aqui, si la fila es la serie resaltada Y es
    # negativa, esa linea del recuadro se pinta de rojo con la etiqueta
    # "{etiqueta_negativo}" en vez del nombre real de la serie.
    tooltip_js = f"""
function TooltipRadar(props) {{
  const {{ active, payload, label }} = props;
  if (!active || !payload || !payload.length) return null;
  return h('div', {{ style: {{ background: '#fff', border: '1px solid #d7dce1',
                             borderRadius: 6, padding: '8px 12px', fontSize: 13 }} }},
    h('div', {{ style: {{ fontWeight: 700, marginBottom: 4, color: '{JOCS_AZUL}' }} }}, label),
    ...payload.map((entry, i) => {{
      const fila = entry.payload.payload || entry.payload;
      const neg = fila[entry.dataKey + '_neg'];
      const esSerieResaltada = entry.dataKey === {json.dumps(serie_resaltar_negativo)};
      const nombre = (neg && esSerieResaltada) ? {json.dumps(etiqueta_negativo)} : entry.name;
      const color = (neg && esSerieResaltada) ? '{C_NEG}' : entry.color;
      return h('div', {{ key: i, style: {{ color }} }},
        nombre + ' : ' + fila[entry.dataKey + '_etq']);
    }})
  );
}}"""
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
         Tooltip, Legend }} = Recharts;
{funciones_punto_js}
{tooltip_js}
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(RadarChart, {{ data: data, margin: {{ top: 10, right: 30, bottom: 10, left: 30 }} }},
      h(PolarGrid, {{}}),
      h(PolarAngleAxis, {{ dataKey: 'categoria', tick: {{ fontSize: 10, fill: '#54616f' }} }}),
      h(PolarRadiusAxis, {{ {eje_radio_kw} tick: false, axisLine: false }}),
      h(Tooltip, {{ content: h(TooltipRadar) }}),
      h(Legend, {{}}),
      {elementos_radar_js}
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

_treemap_click_component = components.declare_component(
    "treemap_click", path=str(BASE / "components" / "treemap_click"))

def treemap_click_recharts(serie, height=360, formateador=fmt_m, valores_reales=None,
                           color_pos=None, color_neg=None, key=None):
    """Treemap clicable (componente real de Streamlit, no st.components.v1.html():
    ese es de una sola via, Python -> HTML, y no puede devolver el clic a Python.
    Al hacer clic en una caja, Streamlit hace un rerun con el nombre de esa caja
    disponible aqui -- así se arma el drill-down aduana -> empleado -> percepción
    pedido en el PDF ("al picarle a una...") sin selectbox.
    'valores_reales'/'color_pos'/'color_neg': igual que antes, para componentes
    que pueden ser negativos (el área de una caja no puede serlo)."""
    s = serie[serie > 0]
    if s.empty:
        return None
    reales = valores_reales.reindex(s.index) if valores_reales is not None else s
    if valores_reales is not None:
        color_list = [color_pos if v >= 0 else color_neg for v in reales]
    else:
        color_list = [color_for(i) for i in range(len(s))]
    data = [{"name": str(label), "size": float(v), "etiqueta": formateador(real), "color": c}
            for label, v, real, c in zip(s.index, s.values, reales.values, color_list)]
    return _treemap_click_component(data=data, height=height, key=key, default=None)

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

# La hoja 'UTILIDAD OPERATIVA' nombra 5 aduanas distinto a como las llama la
# hoja 'emitidas' (no es acento/mayúsculas -- _normaliza_aduana() no lo arregla,
# son palabras de mas/abreviaturas distintas). Sin este mapeo, el cruce con
# Emitidas por aduana las deja en $0 aunque sí facturen. Confirmado con el
# usuario que STAFF e IMILE LAREDO sí son $0 de verdad (no tienen equivalente
# en 'emitidas'), no se agregan aqui.
ALIAS_ADUANA_UO = {
    "AMAZON 1 (JUSTO)": "AMAZON 1",
    "AMAZON 2 (CARGA)": "AMAZON 2",
    "LAZARO CARDENAS": "LAZARO",
    "TOLUCA BSO": "TOLUCA",
    "IMILE": "IMILE AIFA",
}

@st.cache_data(show_spinner=False)
def load_utilidad_operativa():
    """Honorarios AA, Sueldos y Gastos por aduana (hoja 'UTILIDAD OPERATIVA').
    La columna INGRESOS de la propia hoja NO se usa como ingreso real -- el
    usuario aclaró que el Ingreso correcto es Emitidas (hoja 'emitidas', dato
    real por factura) + Honorarios AA, no lo que trae esa columna de la hoja
    (se queda solo para detectar la fila basura/el renglon sin montos). La
    'utilidad' real se calcula mas abajo, en la pagina, una vez que se cruza
    con Emitidas por aduana."""
    cols_monto = ["INGRESOS", "HONORARIOS AA", "SUELDOS", "GASTOS", "CTOS DE REFERENCIA"]
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
        # y una fila TOTAL al final; ambas se excluyen (el total se calcula aparte).
        out = out[out["INGRESOS"].notna()].copy()
        out["aduana"] = coalesce_col(df, df.columns[0]).map(_normaliza_aduana)
        out["aduana"] = out["aduana"].replace(ALIAS_ADUANA_UO)
        out = out[out["aduana"] != "TOTAL"]
        out[cols_monto] = out[cols_monto].fillna(0.0)
        out = out[(out[cols_monto] != 0).any(axis=1)]  # sin actividad en ningun rubro
        out["aduana"] = out["aduana"].fillna("SIN ADUANA")
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

MES_COL_MAYUS = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
                 7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE",
                 12: "DICIEMBRE"}

@st.cache_data(show_spinner=False)
def load_imile_laredo_extra():
    """Hoja 'imile laredo': P&L mensual propio de esa aduana, con su propia
    columna por mes (ENERO..DICIEMBRE). El usuario pidió sumar el mes vigente
    a lo que ya se tiene para la aduana IMILE LAREDO -- esa aduana da $0 en
    'emitidas'/'recibidas' porque (nota suelta en la propia hoja, columna MAYO
    de la sección GASTO) "el problema [es que] todo lo demás de este centro de
    costos es de una empresa gringa, y la info la tenemos que alimentar
    nosotros porque no aparece en el sat" -- o sea, esos clientes no facturan
    por el SAT mexicano, por eso no aparecen en las hojas normales.
    Se suman Ingreso Neto y (Gasto Neto + Costo de Referencia); Sueldos NO se
    repite aquí (ya viene en 'UTILIDAD OPERATIVA' para esta misma aduana, con
    el mismo monto -- verificado: $201,374.64 en ambas hojas)."""
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if "imile laredo" not in hojas:
            continue
        col_mes = MES_COL_MAYUS.get(mes)
        if col_mes is None:
            continue
        df = pd.read_excel(path, sheet_name=hojas["imile laredo"], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        if col_mes not in df.columns:
            continue
        etiquetas_fila = coalesce_col(df, "INGRESOS").map(
            lambda s: s.strip().upper() if isinstance(s, str) else None)

        def _valor(nombre_fila, etiquetas_fila=etiquetas_fila, df=df, col_mes=col_mes):
            m = etiquetas_fila == nombre_fila
            if not m.any():
                return 0.0
            return float(pd.to_numeric(df.loc[m, col_mes], errors="coerce").fillna(0.0).sum())

        partes.append({
            "aduana": "IMILE LAREDO", "anio": anio, "mes": mes, "periodo_label": etiqueta,
            "ingreso_extra": _valor("INGRESO NETO"),
            "egreso_extra": _valor("GASTO NETO") + _valor("TOTAL CTO DE REFERENCIA"),
        })
    return pd.DataFrame(partes)

UUID_RE = re.compile(r"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}")

def _extraer_uuids(celda):
    return UUID_RE.findall(celda) if isinstance(celda, str) else []

# En Alejandro el UUID conciliado vive en su propia columna 'uuid'; en Josefa
# vive mezclado con notas de texto libre (ej. "11.96 DOLARES") en 'Unnamed: 10'
# -- se extrae con el regex de arriba (patron real de UUID), no por nombre de
# columna, para no confundir las notas con UUIDs.
AMEX_COL_UUID = {"ALEJANDRO": "uuid", "JOSEFA": "Unnamed: 10"}

@st.cache_data(show_spinner=False)
def load_amex(persona):
    """Movimientos de tarjeta de credito corporativa AMEX (hojas 'AMEX ALEJANDRO'/
    'AMEX JOSEFA'). 'facturado' = se le encontro al menos un UUID conciliado --
    no hay con que cruzarlo contra las facturas reales (emitidas/recibidas ya
    no traen columna UUID, ver nota en descubrir_archivos), asi que es
    simplemente si esa columna trae uno o no."""
    hoja_nombre = f"AMEX {persona}"
    col_uuid = AMEX_COL_UUID[persona]
    partes = []
    for anio, mes, path in descubrir_archivos():
        etiqueta = f"{MESES[mes]} {anio}" if mes else path.stem
        hojas = {h.strip().lower(): h for h in pd.ExcelFile(path).sheet_names}
        if hoja_nombre.lower() not in hojas:
            continue
        df = pd.read_excel(path, sheet_name=hojas[hoja_nombre.lower()], header=0)
        df.columns = [str(c).strip() for c in df.columns]
        out = pd.DataFrame(index=df.index)
        out["proveedor"] = coalesce_col(df, "Supplier Name").map(norm_txt).fillna("SIN PROVEEDOR")
        out["monto"] = pd.to_numeric(coalesce_col(df, "Net Billed Amount"), errors="coerce").fillna(0.0)
        out["uuids"] = coalesce_col(df, col_uuid).map(_extraer_uuids)
        out["facturado"] = out["uuids"].map(len) > 0
        out["anio"], out["mes"], out["periodo_label"] = anio, mes, etiqueta
        partes.append(out)
    return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()

try:
    df, periodos = load_data()
    sueldos_df = load_sueldos()
    josefa_df = load_honorarios_josefa()
    costos_ref_df = load_costos_referencia()
    utilidad_op_df = load_utilidad_operativa()
    imile_laredo_extra_df = load_imile_laredo_extra()
    amex_alejandro_df = load_amex("ALEJANDRO")
    amex_josefa_df = load_amex("JOSEFA")
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
imile_laredo_extra_f = _filtra_periodo_aduana(imile_laredo_extra_df)

EM = d[d["tipo_factura"] == "EMITIDA"]
REC = d[d["tipo_factura"] == "RECIBIDA"]

TOT_EM = EM["importe"].sum()
TOT_REC = REC["importe"].sum()
MARGEN = TOT_EM - TOT_REC
MARGEN_PCT = MARGEN / TOT_EM if TOT_EM else None

# Ingreso real por aduana = Emitidas (factura real, hoja 'emitidas') + Honorarios AA
# (hoja 'UTILIDAD OPERATIVA'); Egreso real = Recibida (factura real, hoja
# 'recibidas') + Sueldos. Ni la columna INGRESOS ni la columna GASTOS de la hoja
# 'UTILIDAD OPERATIVA' se usan para esto -- el usuario aclaró explícitamente que
# Ingreso = Emitidas + Honorarios, y que Egreso debía calcularse igual
# (Recibida + Sueldos) tras notar que GASTOS no cuadraba con Recibida + Sueldos
# (diferian ~6% para IMILE AIFA). Utilidad Operativa = Ingreso - Egreso, misma
# regla para todas las aduanas. Se cruza por (aduana, periodo_label) por si algún
# día hay más de un periodo cargado.
if not utilidad_op_f.empty:
    em_por_aduana_periodo = EM.groupby(["aduana", "periodo_label"])["importe"].sum()
    rec_por_aduana_periodo = REC.groupby(["aduana", "periodo_label"])["importe"].sum()
    utilidad_op_f = utilidad_op_f.copy()
    utilidad_op_f["emitidas"] = [
        em_por_aduana_periodo.get((a, p), 0.0)
        for a, p in zip(utilidad_op_f["aduana"], utilidad_op_f["periodo_label"])
    ]
    utilidad_op_f["recibida"] = [
        rec_por_aduana_periodo.get((a, p), 0.0)
        for a, p in zip(utilidad_op_f["aduana"], utilidad_op_f["periodo_label"])
    ]
    utilidad_op_f["ingreso"] = utilidad_op_f["emitidas"] + utilidad_op_f["HONORARIOS AA"]
    # 'CTOS DE REFERENCIA' (columna nueva de la hoja 'UTILIDAD OPERATIVA'): se resta
    # igual que Sueldos -- EXCEPTO en IMILE LAREDO, donde este mismo costo ya se
    # suma más abajo vía la pestaña dedicada 'imile laredo' (TOTAL CTO DE
    # REFERENCIA); sumarlo aquí también lo duplicaría (confirmado con el usuario).
    utilidad_op_f["ctos_referencia"] = utilidad_op_f["CTOS DE REFERENCIA"].where(
        utilidad_op_f["aduana"] != "IMILE LAREDO", 0.0)
    utilidad_op_f["egreso"] = (utilidad_op_f["recibida"] + utilidad_op_f["SUELDOS"]
                              + utilidad_op_f["ctos_referencia"])

    # IMILE LAREDO factura a una empresa extranjera que no pasa por el SAT, por
    # eso no aparece en 'emitidas'/'recibidas' -- su Ingreso/Gasto real vive en
    # su propia hoja ('imile laredo'), que se suma aquí encima de lo ya calculado
    # (a pedido del usuario). Sueldos no se repite (ya viene en la hoja principal).
    if not imile_laredo_extra_f.empty:
        extra_map = imile_laredo_extra_f.set_index(["aduana", "periodo_label"])
        for idx in utilidad_op_f.index:
            key = (utilidad_op_f.at[idx, "aduana"], utilidad_op_f.at[idx, "periodo_label"])
            if key in extra_map.index:
                utilidad_op_f.at[idx, "ingreso"] += extra_map.at[key, "ingreso_extra"]
                utilidad_op_f.at[idx, "egreso"] += extra_map.at[key, "egreso_extra"]

    utilidad_op_f["utilidad"] = utilidad_op_f["ingreso"] - utilidad_op_f["egreso"]

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
        uo_ad = utilidad_op_f.groupby("aduana")[["emitidas", "HONORARIOS AA", "SUELDOS",
                                                 "ctos_referencia", "recibida", "ingreso",
                                                 "egreso", "utilidad"]].sum()
        aduanas_uo = uo_ad.sort_values("utilidad", ascending=False).index.tolist()

        st.divider()
        st.subheader("Utilidad Operativa por Aduana")
        st.caption("Utilidad Operativa = Ingresos − Egresos, donde Ingresos = Emitidas "
                   "(hoja 'emitidas') + Honorarios AA, y Egresos = Recibida (hoja 'recibidas') "
                   "+ Sueldos + Costos de Referencia (de la hoja 'UTILIDAD OPERATIVA'). Rojos = "
                   "los 3 gastos apilados; amarillo = Utilidad Operativa; línea verde = Ingreso. "
                   "Escala logarítmica para que las aduanas chicas no desaparezcan.")
        df_bullet_uo = uo_ad.reset_index().sort_values("ingreso", ascending=False)
        # El segmento "Recibidas" se calcula como resto (egreso - Sueldos - Costos de
        # Referencia), no se toma la columna 'recibida' cruda directo -- para la
        # mayoría de las aduanas da lo mismo (así se define 'egreso'), pero en IMILE
        # LAREDO 'egreso' trae un ajuste extra (ver load_imile_laredo_extra) que no
        # vive en ninguna columna individual; sin este resto, la barra apilada se
        # quedaría corta y la Utilidad Operativa se vería inflada visualmente.
        df_bullet_uo["recibida_visual"] = (df_bullet_uo["egreso"] - df_bullet_uo["SUELDOS"]
                                           - df_bullet_uo["ctos_referencia"])
        bullet_uo_recharts(
            df_bullet_uo, "aduana", "ingreso",
            segmentos=[("Sueldos", "#e3a39c", "SUELDOS"),
                      ("Costos de Referencia", C_NEG, "ctos_referencia"),
                      ("Recibidas", "#7b241c", "recibida_visual")],
            egreso_col="egreso", utilidad_col="utilidad", height=440,
        )

        st.divider()
        st.subheader("Cascada de Utilidad Operativa")
        aduana_cascada_sel = st.selectbox(
            "Aduana", ["Todas"] + aduanas_uo, key="aduana_cascada_uo")
        if aduana_cascada_sel == "Todas":
            tot_cascada = uo_ad[["HONORARIOS AA", "SUELDOS", "ctos_referencia",
                                 "ingreso", "egreso"]].sum()
        else:
            tot_cascada = uo_ad.loc[aduana_cascada_sel, ["HONORARIOS AA", "SUELDOS",
                                                          "ctos_referencia", "ingreso", "egreso"]]
        # "Ingresos"/"Recibidas" se leen de ingreso/egreso (ya con el ajuste de
        # IMILE LAREDO sumado) menos su componente que ya va en otra barra -- así
        # el total de la cascada cuadra con la utilidad de la aduana (o de todas).
        df_cascada_uo = pd.DataFrame({
            "label": ["Ingresos", "Honorarios AA", "Sueldos", "Costos de Referencia", "Recibidas"],
            "valor": [tot_cascada["ingreso"] - tot_cascada["HONORARIOS AA"], tot_cascada["HONORARIOS AA"],
                     -tot_cascada["SUELDOS"], -tot_cascada["ctos_referencia"],
                     -(tot_cascada["egreso"] - tot_cascada["SUELDOS"] - tot_cascada["ctos_referencia"])],
        })
        waterfall_recharts(df_cascada_uo, "label", "valor",
                          color_pos=JOCS_AZUL, color_neg=C_NEG, color_total=C_MARGEN,
                          total_label="Utilidad Operativa", height=400)

        bc1, bc2 = st.columns([1.6, 1])
        with bc1:
            # Las 17 aduanas, ordenadas por volumen combinado (ingreso+egreso ya
            # ajustado, no emitida+recibida cruda, para que IMILE LAREDO entre
            # ordenada con su actividad real -- no factura por el SAT, así que su
            # emitida/recibida real es $0). Escala log para que un $100k no
            # desaparezca junto a un $1.7M.
            aduanas_radar = sorted(
                aduanas_uo,
                key=lambda a: -(uo_ad.loc[a, "ingreso"] + uo_ad.loc[a, "egreso"]),
            )
            df_radar_uo = pd.DataFrame({
                "aduana": aduanas_radar,
                "Ingreso": [float(uo_ad.loc[a, "ingreso"]) for a in aduanas_radar],
                "Utilidad Operativa": [float(uo_ad.loc[a, "utilidad"]) for a in aduanas_radar],
            })
            st.caption("Las 17 aduanas, por volumen combinado. Escala logarítmica. "
                       "Utilidad Operativa negativa se marca en rojo (punto y tooltip).")
            radar_recharts(df_radar_uo, "aduana",
                          {"Ingreso": C_EMITIDA, "Utilidad Operativa": C_MARGEN},
                          height=480, escala_log=True,
                          serie_resaltar_negativo="Utilidad Operativa")
        with bc2:
            aduana_uo_sel = st.selectbox(
                "Seleccionar Aduana", aduanas_uo,
                format_func=lambda a: f"{a}  ·  {fmt_m(uo_ad.loc[a, 'utilidad'])}")
            fila = uo_ad.loc[aduana_uo_sel]
            ingresos_sel = float(fila["ingreso"])
            # % Utilidad Operativa = cuánto representa esta aduana de la Utilidad
            # Operativa TOTAL (todas las aduanas), no el margen propio de la aduana
            # (utilidad/ingreso) -- a pedido del usuario.
            tot_utilidad_uo = float(uo_ad["utilidad"].sum())
            pct_uo_sel = fila["utilidad"] / tot_utilidad_uo if tot_utilidad_uo else None
            # IMILE LAREDO no factura por el SAT (ver load_imile_laredo_extra): su
            # 'emitidas' real es $0 aunque sí tiene actividad real -- en la ficha de
            # detalle se muestra el mismo monto que 'Ingresos' en vez de $0, solo
            # para esta aduana, para no sugerir que no factura nada.
            emitidas_sel = ingresos_sel if aduana_uo_sel == "IMILE LAREDO" else float(fila["emitidas"])
            # Montos exactos (sin redondear a $X.XM/$Xk) para poder auditarlos a detalle.
            tabla_uo = pd.DataFrame({
                "Concepto": ["Emitidas", "Honorarios AA", "Ingresos", "Recibidas", "Sueldos",
                            "Costos de Referencia", "Egresos", "Utilidad Operativa",
                            "% Utilidad Operativa"],
                "Valor": [fmt(emitidas_sel), fmt(float(fila["HONORARIOS AA"])),
                         fmt(ingresos_sel), fmt(float(fila["recibida"])),
                         fmt(float(fila["SUELDOS"])), fmt(float(fila["ctos_referencia"])),
                         fmt(float(fila["egreso"])), fmt(float(fila["utilidad"])),
                         pct(pct_uo_sel) if pct_uo_sel is not None else "—"],
            })
            st.dataframe(tabla_uo, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Pareto de Utilidad Operativa (todo CC)")
        st.caption("Ordenado por magnitud (positiva o negativa). Azul = utilidad positiva, "
                   "rojo = negativa. Línea ámbar: % acumulado del total en valor absoluto.")
        pareto_recharts(uo_ad["utilidad"], JOCS_AZUL, color_neg=C_NEG, height=420)

    st.divider()
    st.subheader("Honorarios AA y Referencias")
    col_izq, col_der = st.columns(2)
    with col_izq:
        st.markdown("**Honorarios de Josefa**")
        tot_josefa = josefa_f["monto"].sum() if not josefa_f.empty else 0.0
        st.metric("Total", fmt_m(tot_josefa))
        if not josefa_f.empty:
            por_ad_j = josefa_f.groupby("aduana")["monto"].sum().sort_values(ascending=False)
            bar_h_recharts(por_ad_j, JOCS_AZUL, height=max(260, 30 * len(por_ad_j)))
        else:
            st.info("Sin datos.")
    with col_der:
        st.markdown("**Costos de Referencias**")
        tot_cr = costos_ref_f["monto"].sum() if not costos_ref_f.empty else 0.0
        st.metric("Total", fmt_m(tot_cr))
        if not costos_ref_f.empty:
            por_ad_cr = costos_ref_f.groupby("aduana")["monto"].sum().sort_values(ascending=False)
            bar_h_recharts(por_ad_cr, JOCS_VERDE, height=max(260, 30 * len(por_ad_cr)))
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

    aduana_sel = st.selectbox("Aduana", gasto_ad.index.tolist(), index=0,
                              format_func=lambda a: f"{a}  ·  {fmt_m(gasto_ad[a])}")
    top_n = 8

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
    # Esta dona se quedo en Altair (no Recharts): la version Recharts salia mal
    # posicionada/recortada en el despliegue (el usuario pidio regresarla).
    df_dona = pd.DataFrame({
        "concepto": dona.index, "valor": dona.values, "leyenda": etiquetas,
        "pct": dona.values / base_pct if base_pct else 0.0,
        "orden": range(len(dona)),
    })
    # Rebanada minima: una rebanada de <1% practicamente no se ve (ni cabe su %
    # adentro) -- se le da un piso de angulo (aqui 3% del circulo) quitandoselo
    # proporcionalmente a las demas, nada mas para el DIBUJO (theta); el monto y
    # el % reales (tooltip, leyenda, texto) no cambian, solo el tamaño del gajo.
    PISO_FRAC = 0.03
    chicas = df_dona["pct"] < PISO_FRAC
    if chicas.any() and not chicas.all() and chicas.sum() * PISO_FRAC < 1:
        presupuesto_chicas = chicas.sum() * PISO_FRAC
        suma_grandes = df_dona.loc[~chicas, "pct"].sum()
        escala = (1 - presupuesto_chicas) / suma_grandes if suma_grandes else 1
        df_dona["theta_val"] = df_dona["pct"].where(chicas, df_dona["pct"] * escala)
        df_dona.loc[chicas, "theta_val"] = PISO_FRAC
    else:
        df_dona["theta_val"] = df_dona["pct"]
    # Con el piso de arriba, una rebanada de 0.4% ya se ve -- pero ".0%" la
    # redondearia a "0%", que contradice que se vea un gajo ahi. Un decimal
    # para las chicas, entero para el resto (como ya se via antes).
    df_dona["pct_etq"] = df_dona["pct"].map(lambda p: f"{p*100:.1f}%" if p < 0.01 else f"{p*100:.0f}%")
    r_out, r_in = 300, 160
    arcos = alt.Chart(df_dona).mark_arc(innerRadius=r_in, outerRadius=r_out,
                                        stroke="#FFFFFF", strokeWidth=2).encode(
        theta=alt.Theta("theta_val:Q", stack=True),
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
        theta=alt.Theta("theta_val:Q", stack=True), order=alt.Order("orden:Q"),
        text="pct_etq:N",
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
        bar_h_recharts(prov, color_base, height=max(280, 30 * len(prov)))
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
    st.subheader("Desglose por Aduana, Personal y Percepción")
    st.caption("Haz clic en un cuadro para profundizar: aduana → empleados de esa aduana → "
               "componentes de su percepción.")

    if "nom_aduana" not in st.session_state:
        st.session_state.nom_aduana = None
    if "nom_empleado" not in st.session_state:
        st.session_state.nom_empleado = None

    por_aduana_nom = sueldos_f.groupby("aduana")["total"].sum().sort_values(ascending=False)
    # Si cambiaron los filtros globales y la aduana/empleado seleccionado ya no
    # existe en los datos filtrados, se descarta la selección en vez de tronar.
    if st.session_state.nom_aduana not in por_aduana_nom.index:
        st.session_state.nom_aduana = None
        st.session_state.nom_empleado = None

    st.markdown("**1. Aduanas** · el tamaño de cada cuadro es su nómina total")
    clic_aduana = treemap_click_recharts(por_aduana_nom, height=240, key="tm_aduana")
    if clic_aduana in por_aduana_nom.index and clic_aduana != st.session_state.nom_aduana:
        st.session_state.nom_aduana = clic_aduana
        st.session_state.nom_empleado = None
        st.rerun()

    if st.session_state.nom_aduana is None:
        st.info("Haz clic en una aduana para ver sus empleados.")
    else:
        aduana_nom_sel = st.session_state.nom_aduana
        if st.button(f"← Quitar selección de aduana ({aduana_nom_sel})"):
            st.session_state.nom_aduana = None
            st.session_state.nom_empleado = None
            st.rerun()

        tabla_nom = sueldos_f[sueldos_f["aduana"] == aduana_nom_sel].copy()
        tot_aduana_nom = tabla_nom["total"].sum()
        tabla_nom["pct"] = tabla_nom["total"] / tot_aduana_nom if tot_aduana_nom else 0.0
        tabla_nom = tabla_nom.sort_values("total", ascending=False)
        st.caption(f"Total nómina {aduana_nom_sel}: {fmt_m(tot_aduana_nom)} · {len(tabla_nom)} empleado(s)")

        por_empleado_nom = tabla_nom.groupby("nombre")["total"].sum().sort_values(ascending=False)
        if st.session_state.nom_empleado not in por_empleado_nom.index:
            st.session_state.nom_empleado = None

        st.markdown(f"**2. Empleados de {aduana_nom_sel}** · tamaño = sueldo total")
        clic_emp = treemap_click_recharts(por_empleado_nom, height=320, key="tm_empleado")
        if clic_emp in por_empleado_nom.index and clic_emp != st.session_state.nom_empleado:
            st.session_state.nom_empleado = clic_emp
            st.rerun()

        st.dataframe(pd.DataFrame({
            "Empleado": tabla_nom["nombre"],
            "% Conformación Aduana": tabla_nom["pct"].map(lambda v: f"{v*100:.1f}%"),
            "Sueldo Total": tabla_nom["total"].map(fmt),
        }), use_container_width=True, hide_index=True)

        if st.session_state.nom_empleado is None:
            st.info("Haz clic en un empleado para ver el desglose de su percepción.")
        else:
            empleado_sel = st.session_state.nom_empleado
            if st.button(f"← Quitar selección de empleado ({empleado_sel})"):
                st.session_state.nom_empleado = None
                st.rerun()

            st.divider()
            st.subheader("Ficha Técnica del Trabajador")
            fila_emp = sueldos_f[sueldos_f["nombre"] == empleado_sel]
            if len(fila_emp) > 1:
                st.caption(f"{len(fila_emp)} registros para este empleado en el periodo — se suman.")
            fila_emp_sum = fila_emp[[c for c, _ in COLS_NOMINA] + ["total"]].sum()

            c1, c2 = st.columns(2)
            c1.metric("Aduana", fila_emp["aduana"].iloc[0] if len(fila_emp) else "—")
            c2.metric("Costo Total Neto", fmt_m(fila_emp_sum["total"]))

            st.markdown(f"**3. Percepción de {empleado_sel}** · azul = positivo, "
                       "rojo = negativo (ej. ajustes de Provisiones)")
            serie_comp_real = pd.Series({et: fila_emp_sum[c] for c, et in COLS_NOMINA})
            treemap_click_recharts(serie_comp_real.abs(), height=320, valores_reales=serie_comp_real,
                                   color_pos=JOCS_AZUL, color_neg=C_NEG, key="tm_comp")
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
    st.caption(f"{', '.join(sel_labels)} · Fuente: Movimientos AMEX Alejandro y Movimientos AMEX Josefa")
    st.caption("'Facturado' = el movimiento ya trae un UUID de factura conciliado en el Sheets "
               "(no se cruza contra 'emitidas'/'recibidas' -- esas hojas ya no traen columna UUID).")

    for nombre, amex_df in [("Alejandro", amex_alejandro_df), ("Josefa", amex_josefa_df)]:
        amex_f = amex_df[amex_df["periodo_label"].isin(sel_labels)] if not amex_df.empty else amex_df
        st.divider()
        st.subheader(nombre)
        if amex_f.empty:
            st.info(f"Sin movimientos de {nombre} para el periodo seleccionado.")
            continue

        total_gastado = float(amex_f["monto"].sum())
        movimientos = len(amex_f)
        facturados = int(amex_f["facturado"].sum())
        no_facturados = movimientos - facturados
        monto_facturado = float(amex_f.loc[amex_f["facturado"], "monto"].sum())
        monto_no_facturado = total_gastado - monto_facturado
        pct_facturado = (monto_facturado / total_gastado) if total_gastado else 0.0

        col_donut, col_kpi = st.columns([1.3, 1])
        with col_donut:
            # No se agrupan "los demas" en un cajon de Otros -- a pedido del
            # usuario, solo se junta lo que es caseta/peaje (nombres que traen
            # "PASE", ej. "PASE CARGO ACUMULADO PEAJE", "PASE MARQUESA TLV");
            # el resto de proveedores se queda cada uno como su propia rebanada.
            es_pase = amex_f["proveedor"].str.contains("PASE", case=False, na=False)
            proveedor_agrupado = amex_f["proveedor"].where(~es_pase, "PASE")
            por_prov = amex_f.groupby(proveedor_agrupado)["monto"].sum().sort_values(ascending=False)
            nested_donut_recharts(monto_facturado, monto_no_facturado, por_prov, height=380)
        with col_kpi:
            k1, k2 = st.columns(2)
            k1.metric("Total gastado", fmt(total_gastado))
            k2.metric("Movimientos", f"{movimientos:,}")
            k1.metric("UUID's encontrados", f"{facturados:,}")
            k2.metric("UUID's no encontrados", f"{no_facturados:,}")
            k1.metric("Monto facturado", fmt(monto_facturado))
            k2.metric("Monto no facturado", fmt(monto_no_facturado))
            gauge_svg(pct_facturado, height=210)

