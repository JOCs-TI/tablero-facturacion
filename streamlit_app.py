import json
import unicodedata
from collections import namedtuple
from io import BytesIO
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path

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
  const a = Math.abs(v);
  if (a >= 1e6) return s + '$' + (a / 1e6).toFixed(1) + 'M';
  if (a >= 1e3) return s + '$' + Math.round(a / 1e3) + 'k';
  return s + '$' + Math.round(a).toLocaleString('en-US');
}
"""

def recharts_component(render_js, data, height=400):
    """Incrusta un chart de Recharts. 'render_js' arma el arbol de elementos con
    React.createElement y lo monta con ReactDOM.createRoot(...).render(...); tiene
    disponibles 'data' (ya inyectada como JSON), 'Recharts', 'React', 'ReactDOM'
    y 'fmtMoney' (abrevia $1.2M / $340k, igual que fmt_m() del lado Python)."""
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

def donut_recharts(valores, height=360, colores=None, etiquetas=None,
                   centro_titulo=None, centro_sub=None):
    """Version Recharts de donut_chart(): PieChart con innerRadius + leyenda a la
    derecha (nombre completo) y % dentro de cada rebanada. 'centro_titulo'/
    'centro_sub' (opcionales): dos lineas de texto al centro del anillo (ej. monto
    total + nombre de la aduana, como en Gastos por Concepto)."""
    df = valores.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    df["leyenda"] = etiquetas if etiquetas is not None else df["label"]
    tot = df["valor"].sum() or 1.0
    df["pct"] = df["valor"] / tot
    if colores is None:
        colores = [color_for(i) for i in range(len(df))]
    data = [{"label": str(r.label), "leyenda": str(r.leyenda), "valor": float(r.valor),
             "pct": float(r.pct), "etiqueta": fmt_m(r.valor), "color": c}
            for r, c in zip(df.itertuples(), colores)]
    centro_js = ""
    if centro_titulo is not None:
        centro_js = f"""
      h('text', {{ x: '38%', y: '47%', textAnchor: 'middle', fontSize: 22, fontWeight: 700,
                  fill: '{JOCS_AZUL}' }}, {json.dumps(str(centro_titulo))}),
      h('text', {{ x: '38%', y: '54%', textAnchor: 'middle', fontSize: 14,
                  fill: '#54616f' }}, {json.dumps(str(centro_sub or ""))}),"""
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, PieChart, Pie, Cell, Tooltip, Legend }} = Recharts;

function PctLabel(props) {{
  const {{ cx, cy, midAngle, innerRadius, outerRadius, pct }} = props;
  if (pct < 0.025) return null;
  const RAD = Math.PI / 180;
  const r = (innerRadius + outerRadius) / 2;
  const x = cx + r * Math.cos(-midAngle * RAD);
  const y = cy + r * Math.sin(-midAngle * RAD);
  return h('text', {{ x, y, fill: '#fff', textAnchor: 'middle', dominantBaseline: 'central',
                     fontSize: 12, fontWeight: 600 }}, (pct * 100).toFixed(0) + '%');
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(PieChart, {{}},
      h(Pie, {{ data: data, dataKey: 'valor', nameKey: 'leyenda', cx: '38%', cy: '50%',
               innerRadius: '46%', outerRadius: '78%', paddingAngle: 1, isAnimationActive: false,
               label: PctLabel, labelLine: false }},
        data.map((d, i) => h(Cell, {{ key: i, fill: d.color, stroke: '#FFFFFF', strokeWidth: 2 }}))
      ),
      h(Tooltip, {{ formatter: (v, n, p) => [p.payload.etiqueta + ' · ' + (p.payload.pct * 100).toFixed(1) + '%', p.payload.label] }}),
      h(Legend, {{ layout: 'vertical', align: 'right', verticalAlign: 'middle',
                  wrapperStyle: {{ fontSize: 12, lineHeight: '20px' }} }}),{centro_js}
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

def pareto_recharts(serie, color, height=420, umbral=0.8, color_neg=None):
    """Version Recharts de pareto_chart(): barras descendentes + linea de %
    acumulado (ComposedChart con dos ejes Y). 'color_neg' (opcional): si se pasa,
    'serie' puede traer negativos -- se ordena/acumula por magnitud absoluta (ya
    que un % acumulado no tiene sentido sobre una suma que cruza por cero) y cada
    barra se colorea 'color' o 'color_neg' segun su signo real (ej. Utilidad
    Operativa, que puede ser negativa por aduana)."""
    df = serie.rename("valor").reset_index()
    df.columns = ["label", "valor"]
    orden_col = df["valor"].abs() if color_neg is not None else df["valor"]
    df = df.assign(_orden=orden_col).sort_values("_orden", ascending=False).drop(columns="_orden")
    tot = (df["valor"].abs() if color_neg is not None else df["valor"]).sum() or 1.0
    df["acumulado"] = (df["valor"].abs() if color_neg is not None else df["valor"]).cumsum() / tot
    df["etiqueta"] = df["valor"].map(fmt_m)
    df["color"] = color if color_neg is None else df["valor"].map(lambda v: color if v >= 0 else color_neg)
    data = [{"label": str(r.label), "valor": float(r.valor), "etiqueta": r.etiqueta,
             "acumulado": float(r.acumulado), "color": r.color} for r in df.itertuples()]
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip,
         LabelList, ReferenceLine, Cell }} = Recharts;
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(ComposedChart, {{ data: data, margin: {{ top: 28, right: 40, left: 10, bottom: 50 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
      h(XAxis, {{ dataKey: 'label', angle: -30, textAnchor: 'end', interval: 0,
                 tick: {{ fontSize: 10, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ yAxisId: 'izq', tickFormatter: fmtMoney, tick: {{ fontSize: 11, fill: '#54616f' }},
                 axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ yAxisId: 'der', orientation: 'right', domain: [0, 1.05],
                 tickFormatter: v => Math.round(v * 100) + '%',
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(Tooltip, {{}}),
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

def radar_recharts(df, categoria_col, series_colores, height=420, formateador=fmt_m):
    """RadarChart (componente nativo de Recharts): un eje por cada valor de
    'categoria_col', un polígono por cada serie en 'series_colores'
    ({nombre_columna: color}). Pensado para comparar 2+ series no-negativas por
    categoría (ej. Emitidas vs Recibidas por aduana) -- un radar no representa
    bien valores negativos, para eso usar una barra con signo."""
    data = []
    for _, row in df.iterrows():
        d = {"categoria": str(row[categoria_col])}
        for serie in series_colores:
            d[serie] = float(row[serie])
            d[f"{serie}_etq"] = formateador(row[serie])
        data.append(d)
    radars_js = "\n".join(
        f"""h(Radar, {{ name: {json.dumps(serie)}, dataKey: {json.dumps(serie)},
                     stroke: '{color}', fill: '{color}', fillOpacity: 0.25,
                     isAnimationActive: false }}),"""
        for serie, color in series_colores.items()
    )
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
         Tooltip, Legend }} = Recharts;
const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(RadarChart, {{ data: data, margin: {{ top: 10, right: 30, bottom: 10, left: 30 }} }},
      h(PolarGrid, {{}}),
      h(PolarAngleAxis, {{ dataKey: 'categoria', tick: {{ fontSize: 10, fill: '#54616f' }} }}),
      h(PolarRadiusAxis, {{ tick: {{ fontSize: 9, fill: '#9aa5b1' }}, tickFormatter: fmtMoney }}),
      h(Tooltip, {{ formatter: fmtMoney }}),
      h(Legend, {{}}),
      {radars_js}
    )
  )
);
"""
    recharts_component(render_js, data, height=height)

def treemap_recharts(serie, height=360, colores=None, formateador=fmt_m, valores_reales=None,
                     color_pos=None, color_neg=None):
    """Treemap con el componente nativo de Recharts (a diferencia de Altair, que no
    trae uno -- ahi se armo 'treemap_chart()'/'_squarify()' a mano). 'valores_reales'
    + 'color_pos'/'color_neg': igual que en treemap_chart(), para componentes que
    pueden ser negativos (el area de una caja no puede serlo, se usa abs() para el
    tamaño y el signo real solo para el color/etiqueta)."""
    s = serie[serie > 0]
    if s.empty:
        return False
    reales = valores_reales.reindex(s.index) if valores_reales is not None else s
    if colores is not None:
        color_list = colores
    elif valores_reales is not None:
        color_list = [color_pos if v >= 0 else color_neg for v in reales]
    else:
        color_list = [color_for(i) for i in range(len(s))]
    data = [{"name": str(label), "size": float(v), "etiqueta": formateador(real), "color": c}
            for label, v, real, c in zip(s.index, s.values, reales.values, color_list)]
    render_js = f"""
const h = React.createElement;
const {{ ResponsiveContainer, Treemap }} = Recharts;

function Nodo(props) {{
  const {{ x, y, width, height, color, name, etiqueta }} = props;
  const cabe = width > 90 && height > 34;
  return h('g', {{}},
    h('rect', {{ x, y, width, height, fill: color, stroke: '#FFFFFF', strokeWidth: 2, rx: 4 }}),
    cabe ? h('text', {{ x: x + width / 2, y: y + height / 2 - 7, textAnchor: 'middle',
                       fill: '#FFFFFF', fontSize: 11, fontWeight: 600 }}, name) : null,
    (cabe && height > 50) ? h('text', {{ x: x + width / 2, y: y + height / 2 + 8, textAnchor: 'middle',
                       fill: '#FFFFFF', fontSize: 10, opacity: 0.9 }}, etiqueta) : null
  );
}}

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: {height - 10} }},
    h(Treemap, {{ data: data, dataKey: 'size', isAnimationActive: false, content: h(Nodo) }})
  )
);
"""
    recharts_component(render_js, data, height=height)
    return True

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
        waterfall_recharts(df_uo, "label", "valor",
                          color_pos=JOCS_AZUL, color_neg=C_NEG, color_total=C_MARGEN,
                          total_label="Utilidad Operativa", height=400)

        st.divider()
        st.subheader("Utilidad Operativa por Aduana")
        uo_ad = utilidad_op_f.groupby("aduana")[["INGRESOS", "HONORARIOS AA", "SUELDOS",
                                                 "GASTOS", "utilidad"]].sum()
        em_por_ad = EM.groupby("aduana")["importe"].sum()
        rec_por_ad = REC.groupby("aduana")["importe"].sum()
        aduanas_uo = uo_ad.sort_values("utilidad", ascending=False).index.tolist()

        bc1, bc2 = st.columns([1.6, 1])
        with bc1:
            df_radar_uo = pd.DataFrame({
                "aduana": aduanas_uo,
                "Emitidas": [float(em_por_ad.get(a, 0.0)) for a in aduanas_uo],
                "Recibidas": [float(rec_por_ad.get(a, 0.0)) for a in aduanas_uo],
            })
            radar_recharts(df_radar_uo, "aduana", {"Emitidas": C_EMITIDA, "Recibidas": C_RECIBIDA},
                          height=480)
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
        pareto_recharts(uo_ad["utilidad"], JOCS_AZUL, color_neg=C_NEG, height=420)

    st.divider()
    st.subheader("Costos de Originación")
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
    donut_recharts(dona, height=640, colores=colores, etiquetas=etiquetas,
                   centro_titulo=fmt_m(tot_gasto), centro_sub=aduana_sel)

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
    st.subheader("Desglose por Aduana y Lista de Personal")
    st.caption("Cada cuadro es una aduana; el tamaño es proporcional a su nómina total.")
    por_aduana_nom = sueldos_f.groupby("aduana")["total"].sum().sort_values(ascending=False)
    treemap_recharts(por_aduana_nom, height=280)

    aduana_nom_sel = st.selectbox("Seleccionar Aduana", por_aduana_nom.index.tolist(),
                                  format_func=lambda a: f"{a}  ·  {fmt_m(por_aduana_nom[a])}")
    tabla_nom = sueldos_f[sueldos_f["aduana"] == aduana_nom_sel].copy()
    tot_aduana_nom = tabla_nom["total"].sum()
    tabla_nom["pct"] = tabla_nom["total"] / tot_aduana_nom if tot_aduana_nom else 0.0
    tabla_nom = tabla_nom.sort_values("total", ascending=False)
    st.caption(f"Total nómina {aduana_nom_sel}: {fmt_m(tot_aduana_nom)} · {len(tabla_nom)} empleado(s) "
               "· cada cuadro es un empleado de esta aduana.")
    por_empleado_nom = tabla_nom.groupby("nombre")["total"].sum().sort_values(ascending=False)
    treemap_recharts(por_empleado_nom, height=380)
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
    ok_tm = treemap_recharts(serie_comp_real.abs(), height=360, valores_reales=serie_comp_real,
                             color_pos=JOCS_AZUL, color_neg=C_NEG)
    if not ok_tm:
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

