import itertools
import json
import math
import re
import unicodedata
from collections import namedtuple
from io import BytesIO
import altair as alt
import openpyxl
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

# Presupuesto: TODAVIA no vive en el Sheet en vivo (el usuario lo subio como
# archivo local suelto, carpeta "AGOSTO") -- esto funciona en local pero NO en
# produccion (Streamlit Cloud no tiene acceso a este disco). Pendiente: subir
# una pestaña de presupuesto al mismo Google Sheet y leerla de ahi, igual que
# todo lo demas. Mientras tanto, load_presupuesto() devuelve vacio si no
# encuentra el archivo (no truena el tablero).
PRESUPUESTO_PATH = (BASE / "EXCEL" / "AGOSTO-20261026T223758Z-1-001" / "AGOSTO"
                    / "082026 PRESUPUESTO.xlsx")
CONCEPTO_PRESUPUESTO = {"INGRESOS": "ingresos", "CTO REF": "costos",
                        "SUELDOS": "sueldos", "GASTOS OP": "gastos"}

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
    (alto); por default rojo 0-50%, ámbar 50-80%, verde 80-100%."""
    frac = max(0.0, min(1.0, frac))
    bandas = bandas or [(0.50, C_NEG), (0.30, "#e8a93c"), (0.20, JOCS_VERDE)]
    # cx/cy/r_out se corrieron un poco mas chicos que antes (eran 120/120/105)
    # para dejarle hueco a las marcas de 10% en 10% que ahora se dibujan por
    # fuera del arco (r_out+2 a r_out+9, mas su etiqueta en r_out+20) sin que
    # se salgan del viewBox a los lados.
    cx, cy, r_out, r_in = 130, 122, 92, 60

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

    # Marcas de 10% en 10% (0 a 100) alrededor del arco, a pedido del usuario.
    for p in range(0, 101, 10):
        ang = 180.0 - (p / 100.0) * 180.0
        x1, y1 = punto(r_out + 2, ang)
        x2, y2 = punto(r_out + 8, ang)
        xt, yt = punto(r_out + 19, ang)
        piezas.append(f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
                      f'stroke="#8a95a1" stroke-width="1.5" />')
        piezas.append(f'<text x="{xt:.2f}" y="{yt:.2f}" text-anchor="middle" font-size="9" '
                      f'fill="#8a95a1">{p}%</text>')

    ang_aguja = 180.0 - frac * 180.0
    xa, ya = punto(r_in - 6, ang_aguja)
    fuente = "system-ui,'Segoe UI',Roboto,sans-serif"
    # viewBox alto 195 (no 150): el texto va en cy+34/cy+52 = 156/174, por debajo
    # de donde terminan los arcos (cy=122) -- con un viewBox mas bajo que eso el
    # texto queda recortado fuera del area visible. Ancho 260 (cx=130) para que
    # las etiquetas de 10% (hasta r_out+19=111 desde el centro) no se salgan a
    # los lados. Ancho/alto del <svg> FIJOS (no '100%'): si se deja responsivo
    # al ancho de la columna, el alto renderizado (que sigue el aspect ratio
    # del viewBox) deja de coincidir con el 'height' fijo en pixeles del
    # iframe de components.html, y el iframe recorta lo que sobra -- con
    # tamaño fijo ambos quedan sincronizados siempre.
    svg = f"""<svg viewBox="0 0 260 195" width="280" height="210"
                   xmlns="http://www.w3.org/2000/svg">
      {"".join(piezas)}
      <line x1="{cx}" y1="{cy}" x2="{xa:.2f}" y2="{ya:.2f}" stroke="{JOCS_AZUL}"
            stroke-width="4" stroke-linecap="round" />
      <circle cx="{cx}" cy="{cy}" r="7" fill="{JOCS_AZUL}" />
      <text x="{cx}" y="{cy + 34}" text-anchor="middle" font-size="26" font-weight="700"
            fill="{JOCS_AZUL}" font-family="{fuente}">{frac * 100:.0f}%</text>
      <text x="{cx}" y="{cy + 52}" text-anchor="middle" font-size="12" fill="#54616f"
            font-family="{fuente}">{etiqueta}</text>
    </svg>"""
    # 'margin:0 auto' en el propio <svg> no lo centraba de forma confiable
    # (un SVG no siempre se comporta como bloque para ese truco) -- flexbox
    # en el contenedor sí centra siempre, sea cual sea el ancho de la columna.
    components.html(
        f'<div style="font-family:{fuente};display:flex;justify-content:center;">{svg}</div>',
        height=height)

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
    # Piso visual (mismo patron que la dona de Gastos por Concepto, en Altair):
    # con muchos proveedores chicos, su rebanada real es una tira invisible --
    # se infla a un piso de angulo (3% del circulo) quitandoselo
    # proporcionalmente a las grandes, solo para el DIBUJO ('valorAngulo'); el
    # monto real ('valor', el que lee el tooltip) no cambia.
    PISO_FRAC_EXT = 0.03
    tot_ext = sum(d["valor"] for d in data_ext) or 1.0
    es_chica = [d["valor"] / tot_ext < PISO_FRAC_EXT for d in data_ext]
    n_chicas = sum(es_chica)
    if 0 < n_chicas < len(data_ext) and n_chicas * PISO_FRAC_EXT < 1:
        presupuesto_chicas = n_chicas * PISO_FRAC_EXT
        suma_grandes = sum(d["valor"] for d, chica in zip(data_ext, es_chica) if not chica)
        escala = (1 - presupuesto_chicas) * tot_ext / suma_grandes if suma_grandes else 1
        for d, chica in zip(data_ext, es_chica):
            d["valorAngulo"] = (PISO_FRAC_EXT * tot_ext) if chica else d["valor"] * escala
    else:
        for d in data_ext:
            d["valorAngulo"] = d["valor"]
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
        h(Pie, {{ data: dataExt, dataKey: 'valorAngulo', nameKey: 'nombre', cx: '50%', cy: '50%',
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

def barra_real_ppto_html(real, ppto, color):
    """Barrita HTML pura (sin iframe/Recharts, muy chica para justificar uno):
    relleno de color = Real, marca ambar = Ppto -- mismo lenguaje visual que el
    mockup ('Ingresos - Avance anual'), aplicado aqui a Real vs Ppto del mes."""
    if ppto is None or ppto <= 0:
        return ""
    techo = max(real, ppto) * 1.15 or 1.0
    frac_real = max(0.0, min(1.0, real / techo)) * 100
    frac_ppto = max(0.0, min(1.0, ppto / techo)) * 100
    return (
        f'<div style="position:relative;height:12px;background:#eceff2;'
        f'border-radius:3px;margin:8px 0 2px;">'
        f'<div style="position:absolute;left:0;top:0;height:100%;width:{frac_real:.1f}%;'
        f'background:{color};border-radius:3px;"></div>'
        f'<div style="position:absolute;left:{frac_ppto:.1f}%;top:-2px;height:16px;'
        f'width:2px;background:{C_MARGEN};"></div></div>'
    )

def ranking_aduana_card_frag(titulo, subtitulo, consolidado_html, df, color, height=None):
    """Tarjeta 'Rentabilidad por Aduana': barra horizontal Real (color) + marca
    vertical Ppto (ambar), soporta aduanas con perdida (utilidad negativa) --
    mismo mecanismo de BarraReal que ranking_concentracion_card_html pero con
    dominio [piso, techo] en vez de [0, techo] (ver nota abajo). Devuelve
    (html, script) para fundirse en un documento con render_pagina_html()."""
    h_px = height or max(280, 32 * len(df) + 60)
    ancho_eje = _ancho_eje_categoria(df["aduana"])
    div_id = _next_id("rank-aduana")
    # Piso puede ser negativo (aduanas con perdida operativa) -- sin esto el
    # dominio se queda en [0, techo] y Recharts intenta dibujar barras con
    # ancho negativo para esas filas (error de SVG, la barra no se ve).
    piso = min(0.0, float(df["real"].min()), float(df["ppto"].fillna(0).min())) * 1.15
    techo = float(max(df["real"].max(), df["ppto"].fillna(0).max())) * 1.15 or 1.0
    data = [{"aduana": str(r.aduana), "real": float(r.real),
            "ppto": (float(r.ppto) if pd.notna(r.ppto) else None),
            "real_txt": fmt(r.real), "ppto_txt": fmt(r.ppto) if pd.notna(r.ppto) else "—",
            "delta_txt": r.delta_txt}
            for r in df.itertuples()]
    script = f"""(function() {{
const h = React.createElement;
const {{ ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, LabelList, ReferenceLine }} = Recharts;
const data = {json.dumps(data)};

function BarraReal(props) {{
  const {{ x, y, width, height, payload }} = props;
  // 'width' viene negativo cuando 'real' queda del lado negativo del eje
  // (perdida operativa) -- SVG no acepta width<0, así que se normaliza el
  // origen en vez de dejar que el <rect> truene.
  const xReal = width < 0 ? x + width : x;
  const wReal = Math.abs(width);
  const pxPorUnidad = payload.real !== 0 ? width / payload.real : 0;
  const elementos = [h('rect', {{ key: 'real', x: xReal, y, width: wReal, height, fill: '{color}', rx: 3 }})];
  if (payload.ppto != null && pxPorUnidad !== 0) {{
    const xPpto = x + pxPorUnidad * payload.ppto;
    elementos.push(h('line', {{ key: 'ppto', x1: xPpto, x2: xPpto, y1: y - 4, y2: y + height + 4,
                              stroke: '{C_MARGEN}', strokeWidth: 3 }}));
  }}
  return h('g', {{}}, elementos);
}}

function TooltipRanking(props) {{
  const {{ active, payload, label }} = props;
  if (!active || !payload || !payload.length) return null;
  const fila = payload[0].payload;
  return h('div', {{ style: {{ background: '#fff', border: '1px solid #d7dce1',
                             borderRadius: 6, padding: '8px 12px', fontSize: 13 }} }},
    h('div', {{ style: {{ fontWeight: 700, marginBottom: 4, color: '{JOCS_AZUL}' }} }}, label),
    h('div', {{}}, 'Real: ' + fila.real_txt),
    h('div', {{}}, 'Presupuesto: ' + fila.ppto_txt),
    fila.delta_txt ? h('div', {{ style: {{ color: '{C_MARGEN}', fontWeight: 600 }} }}, fila.delta_txt) : null
  );
}}

const root = ReactDOM.createRoot(document.getElementById('{div_id}'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: '100%' }},
    h(BarChart, {{ data: data, layout: 'vertical', margin: {{ top: 10, right: 70, left: 10, bottom: 10 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', horizontal: false }}),
      h(XAxis, {{ type: 'number', domain: [{piso}, {techo}], allowDataOverflow: true, tickFormatter: fmtMoney,
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(YAxis, {{ dataKey: 'aduana', type: 'category', width: {ancho_eje},
                 tick: {{ fontSize: 11, fill: '#54616f' }}, axisLine: {{ stroke: '#d7dce1' }} }}),
      h(Tooltip, {{ content: h(TooltipRanking) }}),
      h(ReferenceLine, {{ x: 0, stroke: '#9aa5b1' }}),
      h(Bar, {{ dataKey: 'real', shape: BarraReal, isAnimationActive: false }},
        h(LabelList, {{ dataKey: 'delta_txt', position: 'right', fontSize: 11, fill: '#3d4752' }})
      )
    )
  )
);
}})();
"""
    body = f"""
      <div class="card">
        <div class="card-header">{_dash_icono('chart')}<p class="card-title">{titulo}</p></div>
        <p class="card-subtitle">{subtitulo}</p>
        {consolidado_html}
        <div id="{div_id}" style="height:{h_px - 90}px;"></div>
      </div>"""
    return body, script

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

_treemap_nomina_component = components.declare_component(
    "treemap_click", path=str(BASE / "components" / "treemap_click"))
_dashboard_component = components.declare_component(
    "facturacion_dashboard", path=str(BASE / "components" / "facturacion_dashboard"))

def _construir_arbol_nomina(sueldos_f):
    """Arma los 3 niveles del treemap de Nómina (aduana -> empleado ->
    componentes de percepción) de una sola pasada, para mandárselos COMPLETOS
    al componente de React (ver treemap_nomina_recharts) -- así el drill-down
    puede vivir y animarse del lado del navegador en vez de ir a buscar cada
    nivel a Streamlit con un rerun por clic."""
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
            fila_sum = tabla_ad.loc[tabla_ad["nombre"] == empleado,
                                    [c for c, _ in COLS_NOMINA]].sum()
            componentes = [
                {"name": etq, "size": float(abs(fila_sum[c])), "etiqueta": fmt(fila_sum[c]),
                 "color": JOCS_AZUL if fila_sum[c] >= 0 else C_NEG}
                for c, etq in COLS_NOMINA if abs(fila_sum[c]) > 0
            ]
            empleados.append({"name": str(empleado), "size": float(total_emp),
                              "etiqueta": fmt_m(total_emp), "color": color_for(len(empleados)),
                              "componentes": componentes})
        arbol.append({"name": str(aduana), "size": float(total_aduana),
                      "etiqueta": fmt_m(total_aduana), "color": color_for(len(arbol)),
                      "empleados": empleados})
    return arbol

def treemap_nomina_recharts(arbol, height=420, path=None, key=None):
    """Treemap jerárquico de 3 niveles con drill-down ANIMADO, enteramente del
    lado del navegador (componente real de Streamlit, no st.components.v1.html():
    ese es de una sola vía, Python -> HTML, no puede devolver el clic a
    Python). Se le manda el ÁRBOL COMPLETO de una sola vez (ver
    _construir_arbol_nomina) -- a diferencia del enfoque anterior (un
    componente nuevo por nivel: cada clic destruía y recreaba el iframe
    entero, sin nada que animar entre un estado y otro), aquí los 3 niveles
    viven en el MISMO componente de React, que trae su propio estado de
    navegación (una miga de pan clicable + Recharts animando el cambio de
    tamaño/posición de las cajas). Cada clic sigue avisando a Streamlit con
    el nuevo 'path' (lista de nombres: [] | [aduana] | [aduana, empleado])
    para que las secciones de Python de abajo (tabla, Ficha Técnica) se
    mantengan sincronizadas -- pero ese aviso ya no bloquea la animación."""
    return _treemap_nomina_component(arbol=arbol, height=height, path=path or [],
                                     key=key, default=None)

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
MES_NUM_POR_NOMBRE = {v: k for k, v in MES_COL_MAYUS.items()}

def _leer_bloques_presupuesto(ws):
    """El archivo de presupuesto apila 4 bloques (INGRESOS/CTO REF/SUELDOS/
    GASTOS OP), cada uno dos veces: primero el mes, luego (tras una fila
    'ACUMULADO') el acumulado del año a ese mes -- mismo "bloques apilados"
    que Town Hall, pero con header de 2 celdas (concepto + nombre del mes) en
    vez de una fila 'AREA'. Devuelve (mensual, acumulado, mes_num), cada uno
    {concepto: {aduana: valor}}."""
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    mensual, acumulado, mes_num = {}, {}, None
    i, es_acum = 0, False
    while i < len(rows):
        r = rows[i]
        if not r or r[0] is None:
            i += 1
            continue
        a0 = str(r[0]).strip().upper()
        if a0 == "ACUMULADO" and (len(r) < 2 or r[1] is None):
            es_acum = True
            i += 1
            continue
        if a0 in CONCEPTO_PRESUPUESTO and len(r) > 1 and isinstance(r[1], str):
            concepto = CONCEPTO_PRESUPUESTO[a0]
            mes_num = mes_num or MES_NUM_POR_NOMBRE.get(r[1].strip().upper())
            datos = {}
            i += 1
            while i < len(rows):
                r2 = rows[i]
                if not r2 or r2[0] is None:
                    break
                a = str(r2[0]).strip().upper()
                if a.startswith("TOTAL"):
                    i += 1
                    break
                aduana = ALIAS_ADUANA_UO.get(_normaliza_aduana(str(r2[0])), _normaliza_aduana(str(r2[0])))
                val = r2[1] if len(r2) > 1 else None
                datos[aduana] = float(val) if isinstance(val, (int, float)) else 0.0
                i += 1
            (acumulado if es_acum else mensual)[concepto] = datos
            es_acum = False
            continue
        i += 1
    return mensual, acumulado, mes_num

@st.cache_data(show_spinner=False)
def load_presupuesto():
    """Presupuesto Real por aduana (ver nota junto a PRESUPUESTO_PATH). Devuelve
    (mensual, acumulado, mes_num) -- vacio/None si el archivo no esta."""
    if not PRESUPUESTO_PATH.exists():
        return {}, {}, None
    wb = openpyxl.load_workbook(PRESUPUESTO_PATH, data_only=True)
    return _leer_bloques_presupuesto(wb[wb.sheetnames[0]])

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
        # 'Net Billed Amount' viene vacío cuando Charge Amount == Credit Amount
        # (cargo abonado completo, ej. TELMEX CARGO RECURRENTE $12 cargo + $12
        # credito) -- a pedido del usuario, ahí se usa el Charge Amount
        # original en vez de mostrar $0.
        monto_neto = pd.to_numeric(coalesce_col(df, "Net Billed Amount"), errors="coerce")
        monto_cargo = pd.to_numeric(coalesce_col(df, "Charge Amount"), errors="coerce").fillna(0.0)
        out["monto"] = monto_neto.fillna(monto_cargo)
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

NAV_OPCIONES = ["Utilidad Operativa", "Gastos", "Ingresos", "Nómina", "Tarjeta de Crédito"]
PAGINAS_REDISENADAS = {"Utilidad Operativa", "Gastos", "Ingresos"}
if "nav" not in st.session_state:
    st.session_state["nav"] = "Utilidad Operativa"
pagina = st.session_state["nav"]

labels_disp = [p.periodo_label for p in periodos]
aduanas_todas = sorted(df["aduana"].dropna().unique().tolist())

if pagina in PAGINAS_REDISENADAS:
    # Estas 3 paginas ya no usan el sidebar ni widgets nativos de Streamlit --
    # el usuario pidio explicitamente eliminar "todo lo de Streamlit" y
    # reconstruir identico a sus mockups con React/HTML/CSS (ver
    # [[feedback-respetar-referencias-visuales]]). La navegacion y los
    # filtros viven DENTRO del componente (ver dashboard_component mas abajo);
    # Periodo/Aduana no tienen un control global aqui porque hoy solo hay un
    # periodo cargado y cada pagina ya filtra por aduana a su manera (ver
    # mockup "Resumen de la aduana").
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
    sel_labels = labels_disp
    aduanas_sel_global = []
else:
    with st.sidebar:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), use_container_width=True)
        st.markdown("---")
        st.caption("TABLEROS")
        idx_actual = NAV_OPCIONES.index(pagina) if pagina in NAV_OPCIONES else 0
        pagina_radio = st.radio("", NAV_OPCIONES, index=idx_actual,
                                label_visibility="collapsed", key="nav_radio")
        if pagina_radio != pagina:
            st.session_state["nav"] = pagina_radio
            st.rerun()
        st.markdown("---")
        st.caption("PERIODO")
        sel_labels = st.multiselect("", labels_disp, default=labels_disp,
                                    label_visibility="collapsed", key="periodo_sel")
        if not sel_labels:
            sel_labels = labels_disp
        st.caption("ADUANA")
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

# ── Tarjetas HTML estilo "BI dashboard" (mockups 'Analisis de Gastos'/'Analisis
# de Ingresos', 2026-10-27) ─────────────────────────────────────────────────────
# El usuario pidio fidelidad visual exacta al mockup aunque haya que dejar de
# usar los widgets nativos de Streamlit ("usa el lenguaje necesario para que
# quede igual") -- cada tarjeta de abajo es su propio documento HTML/CSS (y
# Recharts cuando hace falta grafica) incrustado via components.html(), mismo
# mecanismo que ya se usaba para las graficas Recharts de este tablero.
DASH_BG, DASH_CARD_BG, DASH_BORDER = "#F4F5F7", "#FFFFFF", "#E3E6EA"
DASH_TEXT, DASH_SUBTEXT = "#1C2A3A", "#6B7280"
DASH_POS_BG, DASH_POS_TXT = "#E6F4EA", "#1E7E34"
DASH_NEG_BG, DASH_NEG_TXT = "#FDEAEA", "#C0392B"
DASH_ICONOS = {"money": "💵", "target": "🎯", "people": "👥", "chart": "📊",
               "building": "🏢", "trophy": "🏆", "box": "📦", "trend": "📈",
               "calendar": "📅", "year": "📆", "filtro": "🗂️", "pin": "📍",
               "stack": "🗃️", "gauge": "⏱️", "list": "📋", "doc": "📄",
               "hex": "⬡", "award": "🏅"}

DASH_CSS = f"""
* {{ box-sizing: border-box; }}
body {{ margin: 0; padding: 2px; font-family: 'Segoe UI', system-ui, Roboto, sans-serif;
        background: transparent; }}
.card {{ background: {DASH_CARD_BG}; border: 1px solid {DASH_BORDER}; border-radius: 12px;
         padding: 16px 18px; box-shadow: 0 1px 3px rgba(16,24,40,0.06); height: 100%; }}
.card-header {{ display:flex; align-items:center; gap:10px; margin-bottom: 2px; }}
.card-icon {{ width:32px; height:32px; border-radius:50%; background:#EAF1FB; display:flex;
              align-items:center; justify-content:center; font-size:15px; flex-shrink:0; }}
.card-title {{ font-size: 15px; font-weight:700; color:{DASH_TEXT}; margin:0; }}
.card-subtitle {{ font-size:12px; color:{DASH_SUBTEXT}; margin: 2px 0 14px 42px; }}
.stat-grid {{ display:flex; flex-wrap:wrap; gap:14px; }}
.stat-card {{ flex: 1 1 210px; background:{DASH_CARD_BG}; border:1px solid {DASH_BORDER};
              border-radius:12px; padding:14px 16px; box-shadow:0 1px 3px rgba(16,24,40,0.06); }}
.stat-top {{ display:flex; align-items:center; gap:8px; margin-bottom:8px; }}
.stat-icon {{ width:26px; height:26px; border-radius:50%; background:#EAF1FB; display:flex;
              align-items:center; justify-content:center; font-size:12px; flex-shrink:0; }}
.stat-label {{ font-size:12px; color:{DASH_SUBTEXT}; font-weight:600; }}
.stat-value {{ font-size:21px; font-weight:800; color:{DASH_TEXT}; margin:1px 0; line-height:1.2; }}
.stat-sub {{ font-size:11px; color:{DASH_SUBTEXT}; margin-top:4px; }}
.badge {{ display:inline-flex; align-items:center; gap:3px; font-size:11px; font-weight:700;
          padding:2px 8px; border-radius:999px; }}
.badge-pos {{ background:{DASH_POS_BG}; color:{DASH_POS_TXT}; }}
.badge-neg {{ background:{DASH_NEG_BG}; color:{DASH_NEG_TXT}; }}
.badge-off {{ background:#EEF0F3; color:{DASH_SUBTEXT}; }}
.dtable {{ width:100%; border-collapse:collapse; font-size:12.5px; }}
.dtable th {{ text-align:left; color:{DASH_SUBTEXT}; font-weight:600; padding:6px 8px;
              border-bottom:1px solid {DASH_BORDER}; }}
.dtable td {{ padding:7px 8px; border-bottom:1px solid #F0F1F3; color:{DASH_TEXT}; }}
.dtable tr:last-child td {{ border-bottom:none; font-weight:700; }}
.dtable td.num, .dtable th.num {{ text-align:right; }}
.rank-row {{ display:flex; align-items:center; gap:10px; margin-bottom:11px; }}
.rank-num {{ width:18px; font-size:12px; color:{DASH_SUBTEXT}; font-weight:700; flex-shrink:0; }}
.rank-name {{ width:170px; font-size:12px; color:{DASH_TEXT}; flex-shrink:0;
              overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
.rank-bar-track {{ flex:1; background:#EEF0F3; border-radius:4px; height:11px; }}
.rank-bar-fill {{ height:100%; border-radius:4px; }}
.rank-val {{ width:92px; text-align:right; font-size:11.5px; color:{DASH_TEXT}; font-weight:600;
             flex-shrink:0; }}
.rank-pct {{ width:46px; text-align:right; font-size:12px; color:{DASH_TEXT}; font-weight:700;
             flex-shrink:0; }}
.empty-state {{ display:flex; flex-direction:column; align-items:center; justify-content:center;
                color:{DASH_SUBTEXT}; font-size:13px; text-align:center; padding:30px 20px; }}
.conc-row {{ display:flex; gap:14px; margin-top:12px; }}
.conc-item {{ flex:1; text-align:center; }}
.conc-label {{ font-size:11.5px; color:{DASH_SUBTEXT}; margin-bottom:4px; }}
.conc-value {{ font-size:19px; font-weight:800; color:{JOCS_AZUL}; }}
.dash-header {{ display:flex; align-items:flex-start; justify-content:space-between;
                flex-wrap:wrap; gap:12px; }}
.dash-header-left {{ display:flex; align-items:center; gap:12px; }}
.dash-header-icon {{ width:40px; height:40px; border-radius:10px; background:{JOCS_AZUL};
                     display:flex; align-items:center; justify-content:center; font-size:19px;
                     flex-shrink:0; }}
.dash-header h1 {{ margin:0; font-size:24px; color:{DASH_TEXT}; }}
.dash-header p {{ margin:0; font-size:12.5px; color:{DASH_SUBTEXT}; }}
.filtro-pill {{ display:flex; align-items:center; gap:7px; background:{DASH_CARD_BG};
                border:1px solid {DASH_BORDER}; border-radius:10px; padding:6px 12px; }}
.filtro-pill .fp-icon {{ font-size:15px; }}
.filtro-pill .fp-label {{ font-size:10px; color:{DASH_SUBTEXT}; display:block; }}
.filtro-pill .fp-value {{ font-size:12px; color:{DASH_TEXT}; font-weight:700; display:block; }}
.mini-stat {{ display:flex; align-items:center; gap:8px; }}
.mini-stat .stat-icon {{ width:24px; height:24px; font-size:11px; }}
.mini-stat .mini-val {{ font-size:15px; font-weight:800; color:{DASH_TEXT}; line-height:1.15; }}
.mini-stat .mini-label {{ font-size:11px; color:{DASH_SUBTEXT}; }}
"""

def _dash_icono(nombre):
    return f'<div class="card-icon">{DASH_ICONOS.get(nombre, "•")}</div>'

def _dash_badge(delta_txt, positivo):
    if not delta_txt:
        return ""
    cls = "badge-pos" if positivo else ("badge-neg" if positivo is False else "badge-off")
    flecha = "▲" if positivo else ("▼" if positivo is False else "")
    return f'<span class="badge {cls}">{flecha} {delta_txt}</span>'

_dash_id_counter = itertools.count()

def _next_id(prefix):
    """IDs unicos para los <div> donde monta cada grafica Recharts -- hace
    falta porque ahora varias graficas conviven en UN SOLO documento HTML
    (ver render_pagina_html), no una por iframe como antes."""
    return f"{prefix}-{next(_dash_id_counter)}"

def html_card(body_html, height):
    """Unico uso que queda: paginas viejas (Utilidad Operativa) que siguen
    incrustando una tarjeta por iframe. Gastos/Ingresos ya NO llaman esto --
    usan render_pagina_html() para fundir todas las tarjetas en un solo
    documento (a pedido del usuario: nada de 'varios Streamlit pegados',
    un layout real de pagina)."""
    components.html(f"<!DOCTYPE html><html><head><meta charset='utf-8'>"
                    f"<style>{DASH_CSS}</style></head><body>{body_html}</body></html>",
                    height=height, scrolling=False)

def render_pagina_html(filas, extra_script=""):
    """Arma UNA sola pagina HTML/CSS (+Recharts) a partir de 'filas': lista de
    filas, cada una lista de (ancho_fr, fragmento_html, alto_px). Un solo
    components.html() para toda la pagina -- sin huecos/alineacion rara entre
    iframes sueltos ni de st.columns, que es justo lo que se pidio evitar."""
    scripts = [extra_script] if extra_script else []
    filas_html = []
    alto_total = 0
    for fila in filas:
        cols_css = " ".join(f"{ancho}fr" for ancho, _, _ in fila)
        alto_fila = max(alto for _, _, alto in fila)
        celdas = "".join(f'<div style="height:{alto}px;overflow:hidden;">{html}</div>'
                         for _, html, alto in fila)
        filas_html.append(f'<div style="display:grid;grid-template-columns:{cols_css};'
                          f'gap:16px;margin-bottom:16px;">{celdas}</div>')
        alto_total += alto_fila + 16
    doc = (f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{DASH_CSS}</style></head>"
          f"<body>{''.join(filas_html)}</body>{RECHARTS_CDN}"
          f"<script>{RECHARTS_FMT_JS}{''.join(scripts)}</script></html>")
    components.html(doc, height=alto_total + 10, scrolling=False)

def stat_grid_card_html(cards):
    """'cards': lista de dicts {icono, label, valor, sub (opcional), delta_txt
    (opcional), delta_pos (True/False/None -> verde/rojo/gris), barra_html
    (opcional, ej. salida de barra_real_ppto_html -- barrita Real vs Ppto
    embebida dentro de la tarjeta, mismo lenguaje visual del mockup)."""
    items = "".join(f"""
      <div class="stat-card">
        <div class="stat-top"><div class="stat-icon">{DASH_ICONOS.get(c.get('icono'), '•')}</div>
          <div class="stat-label">{c['label']}</div></div>
        <div class="stat-value">{c['valor']}</div>
        {_dash_badge(c.get('delta_txt'), c.get('delta_pos'))}
        {f'<div class="stat-sub">{c["sub"]}</div>' if c.get('sub') else ''}
        {c.get('barra_html') or ''}
      </div>""" for c in cards)
    return f'<div class="stat-grid" style="height:100%;">{items}</div>', ""

def header_bar_html(icono, titulo, subtitulo, periodo_txt, anio_txt, filtros_txt):
    """Encabezado de pagina: icono+titulo+subtitulo a la izquierda, 'pills' de
    Periodo/Año/Filtros a la derecha -- reflejan el filtro global real del
    sidebar (no son controles nuevos; evita duplicar filtros no funcionales)."""
    body = f"""
      <div class="dash-header">
        <div class="dash-header-left">
          <div class="dash-header-icon">{DASH_ICONOS.get(icono, '•')}</div>
          <div><h1>{titulo}</h1><p>{subtitulo}</p></div>
        </div>
        <div style="display:flex;gap:10px;">
          <div class="filtro-pill"><span class="fp-icon">{DASH_ICONOS['calendar']}</span>
            <span><span class="fp-label">Periodo</span><span class="fp-value">{periodo_txt}</span></span></div>
          <div class="filtro-pill"><span class="fp-icon">{DASH_ICONOS['year']}</span>
            <span><span class="fp-label">Año</span><span class="fp-value">{anio_txt}</span></span></div>
          <div class="filtro-pill"><span class="fp-icon">{DASH_ICONOS['filtro']}</span>
            <span><span class="fp-label">Filtros</span><span class="fp-value">{filtros_txt}</span></span></div>
        </div>
      </div>"""
    return body, ""

def _chart_card_shell(icono, titulo, subtitulo, chart_div_id, extra_note, height, chart_height):
    return f"""
      <div class="card">
        <div class="card-header">{_dash_icono(icono)}<p class="card-title">{titulo}</p></div>
        <p class="card-subtitle">{subtitulo}</p>
        <div id="{chart_div_id}" style="height:{chart_height}px;"></div>
        {f'<p style="margin:6px 0 0 0;font-size:11px;color:{DASH_SUBTEXT};text-align:center;">{extra_note}</p>' if extra_note else ''}
      </div>"""

def chart_bar_card_html(icono, titulo, subtitulo, categorias, serie_real, serie_ppto,
                        nombre_real, nombre_ppto, color_real, color_ppto, extra_note="",
                        height=360):
    """Grafica de barras agrupadas Real vs Presupuesto (Recharts) -- funciona
    igual con 1 categoria que con 12; hoy solo hay el mes real cargado, asi
    que se ve una sola categoria (nota debajo lo aclara). Devuelve (html,
    script): varias de estas conviven en el mismo documento (ver
    render_pagina_html), cada una en su propio IIFE para no chocar variables."""
    div_id = _next_id("bar")
    data = [{"cat": c, "real": float(r), "ppto": float(p)}
            for c, r, p in zip(categorias, serie_real, serie_ppto)]
    script = f"""(function() {{
const h = React.createElement;
const {{ ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend }} = Recharts;
const data = {json.dumps(data)};
const root = ReactDOM.createRoot(document.getElementById('{div_id}'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: '100%' }},
    h(BarChart, {{ data: data, margin: {{ top: 6, right: 10, left: 0, bottom: 0 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
      h(XAxis, {{ dataKey: 'cat', tick: {{ fontSize: 11, fill: '{DASH_SUBTEXT}' }},
                 axisLine: {{ stroke: '{DASH_BORDER}' }} }}),
      h(YAxis, {{ tickFormatter: fmtMoney, tick: {{ fontSize: 10, fill: '{DASH_SUBTEXT}' }},
                 axisLine: {{ stroke: '{DASH_BORDER}' }} }}),
      h(Tooltip, {{ formatter: (v) => fmtMoney(v) }}),
      h(Legend, {{ wrapperStyle: {{ fontSize: 12 }} }}),
      h(Bar, {{ dataKey: 'real', name: '{nombre_real}', fill: '{color_real}', radius: [3,3,0,0],
               maxBarSize: 46, isAnimationActive: false }}),
      h(Bar, {{ dataKey: 'ppto', name: '{nombre_ppto}', fill: '{color_ppto}', radius: [3,3,0,0],
               maxBarSize: 46, isAnimationActive: false }})
    )
  )
);
}})();
"""
    return _chart_card_shell(icono, titulo, subtitulo, div_id, extra_note, height, height - 90), script

def chart_line_card_html(icono, titulo, subtitulo, categorias, serie_real, serie_ppto,
                         nombre_real, nombre_ppto, color_real, color_ppto, extra_note="",
                         height=360, formato_pct=False):
    div_id = _next_id("line")
    data = [{"cat": c, "real": float(r), "ppto": float(p)}
            for c, r, p in zip(categorias, serie_real, serie_ppto)]
    formatter = "(v) => (v*100).toFixed(1) + '%'" if formato_pct else "fmtMoney"
    axis_fmt = "(v) => (v*100).toFixed(1) + '%'" if formato_pct else "fmtMoney"
    script = f"""(function() {{
const h = React.createElement;
const {{ ResponsiveContainer, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend }} = Recharts;
const data = {json.dumps(data)};
const root = ReactDOM.createRoot(document.getElementById('{div_id}'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: '100%' }},
    h(LineChart, {{ data: data, margin: {{ top: 6, right: 10, left: 0, bottom: 0 }} }},
      h(CartesianGrid, {{ strokeDasharray: '3 3', vertical: false }}),
      h(XAxis, {{ dataKey: 'cat', tick: {{ fontSize: 11, fill: '{DASH_SUBTEXT}' }},
                 axisLine: {{ stroke: '{DASH_BORDER}' }} }}),
      h(YAxis, {{ tickFormatter: {axis_fmt}, tick: {{ fontSize: 10, fill: '{DASH_SUBTEXT}' }},
                 axisLine: {{ stroke: '{DASH_BORDER}' }} }}),
      h(Tooltip, {{ formatter: {formatter} }}),
      h(Legend, {{ wrapperStyle: {{ fontSize: 12 }} }}),
      h(Line, {{ type: 'monotone', dataKey: 'real', name: '{nombre_real}', stroke: '{color_real}',
                strokeWidth: 3, dot: {{ r: 4 }}, isAnimationActive: false }}),
      h(Line, {{ type: 'monotone', dataKey: 'ppto', name: '{nombre_ppto}', stroke: '{color_ppto}',
                strokeWidth: 2, strokeDasharray: '5 3', dot: {{ r: 4 }}, isAnimationActive: false }})
    )
  )
);
}})();
"""
    return _chart_card_shell(icono, titulo, subtitulo, div_id, extra_note, height, height - 90), script

def empty_state_card_html(icono, titulo, subtitulo, mensaje):
    body = f"""
      <div class="card">
        <div class="card-header">{_dash_icono(icono)}<p class="card-title">{titulo}</p></div>
        <p class="card-subtitle">{subtitulo}</p>
        <div class="empty-state">{mensaje}</div>
      </div>"""
    return body, ""

def donut_detalle_card_html(titulo_dona, subtitulo_dona, dona_serie, total_txt, subtotal_txt,
                            titulo_tabla, titulo_col1, tabla_df):
    """Tarjeta 'Composicion ... por concepto': dona (Recharts PieChart) + mini
    tabla a un lado (Concepto/Gasto/%Part/Var. vs. ant.), y debajo la tabla de
    detalle completa (+Gasto anterior/Var.$/Var.%) -- replica las 2 tablas del
    mockup. 'Var. vs. periodo anterior' se muestra como "—": no hay un mes
    anterior real cargado todavia (ver nota en pagina_analisis_*), no se
    inventa la cifra."""
    colores = [color_for(k) if c != "Otros" else "#c9cdd3" for k, c in enumerate(dona_serie.index)]
    data = [{"nombre": str(c), "valor": float(v), "color": colores[k]}
            for k, (c, v) in enumerate(dona_serie.items())]
    tot = dona_serie.sum() or 1.0
    leyenda_rows = "".join(f"""
      <tr><td style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px;">
          <span style="display:inline-block;width:8px;height:8px;border-radius:50%;
          background:{colores[k]};margin-right:5px;"></span>{c}</td>
          <td class="num" style="white-space:nowrap;">{fmt(v)}</td>
          <td class="num" style="white-space:nowrap;">{v/tot*100:.1f}%</td>
          <td class="num" style="color:{DASH_SUBTEXT};white-space:nowrap;">—</td></tr>"""
        for k, (c, v) in enumerate(dona_serie.items()))

    filas_tabla = "".join(f"""
      <tr><td>{row[titulo_col1]}</td><td class="num">{fmt(row['Monto'])}</td>
          <td class="num">{pct(row['Monto']/tot)}</td>
          <td class="num" style="color:{DASH_SUBTEXT};">—</td>
          <td class="num" style="color:{DASH_SUBTEXT};">—</td>
          <td class="num" style="color:{DASH_SUBTEXT};">—</td></tr>""" for _, row in tabla_df.iterrows())
    fila_total = (f'<tr><td>Total</td><td class="num">{fmt(tabla_df["Monto"].sum())}</td>'
                 f'<td class="num">100%</td><td class="num">—</td><td class="num">—</td>'
                 f'<td class="num">—</td></tr>')

    div_id = _next_id("dona")
    script = f"""(function() {{
const h = React.createElement;
const {{ ResponsiveContainer, PieChart, Pie, Cell, Tooltip }} = Recharts;
const data = {json.dumps(data)};
const root = ReactDOM.createRoot(document.getElementById('{div_id}'));
root.render(
  h(ResponsiveContainer, {{ width: '100%', height: '100%' }},
    h(PieChart, {{}},
      h(Pie, {{ data: data, dataKey: 'valor', nameKey: 'nombre', innerRadius: 44, outerRadius: 75,
               paddingAngle: 1, isAnimationActive: false }},
        data.map((d, i) => h(Cell, {{ key: i, fill: d.color, stroke: '#fff', strokeWidth: 2 }}))
      ),
      h(Tooltip, {{ formatter: (v) => fmtMoney(v) }})
    )
  )
);
}})();
"""
    body = f"""
      <div class="card">
        <div class="card-header">{_dash_icono('chart')}<p class="card-title">{titulo_dona}</p></div>
        <p class="card-subtitle">{subtitulo_dona}</p>
        <div style="display:flex;gap:10px;align-items:center;">
          <div style="flex:0 0 36%;position:relative;">
            <div id="{div_id}" style="height:170px;"></div>
            <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
                 text-align:center;pointer-events:none;">
              <div style="font-size:13px;font-weight:800;color:{JOCS_AZUL};">{total_txt}</div>
              <div style="font-size:9.5px;color:{DASH_SUBTEXT};">{subtotal_txt}</div>
            </div>
          </div>
          <div style="flex:1;min-width:0;">
            <table class="dtable" style="font-size:10.5px;table-layout:fixed;width:100%;">
              <colgroup><col style="width:42%;"><col style="width:22%;"><col style="width:16%;">
                <col style="width:20%;"></colgroup>
              <tr><th>Concepto</th><th class="num">Gasto</th><th class="num">%</th>
                  <th class="num">Var.</th></tr>
              {leyenda_rows}
            </table>
          </div>
        </div>
        <p class="card-title" style="font-size:13px;margin:14px 0 8px;">{titulo_tabla}</p>
        <div style="max-height:220px;overflow-y:auto;">
          <table class="dtable">
            <tr><th>{titulo_col1}</th><th class="num">Monto actual</th>
                <th class="num">% Part.</th><th class="num">Monto anterior</th>
                <th class="num">Var. $</th><th class="num">Var. %</th></tr>
            {filas_tabla}
            {fila_total}
          </table>
        </div>
      </div>"""
    return body, script

def ranking_concentracion_card_html(titulo, subtitulo, badge_txt, serie_top, color, etq_contraparte,
                                    etq_contraparte_sing):
    """Tarjeta 'Principales proveedores/clientes' (ranking con barra) +
    'Concentracion' (3 metricas), mismo iframe -- replica bar_h_recharts en
    HTML/CSS puro (mas facil de alinear pixel a pixel con el mockup que
    Recharts para una lista simple de barras numeradas)."""
    tot = float(serie_top.sum()) or 1.0
    maxv = float(serie_top.max()) or 1.0
    filas = "".join(f"""
      <div class="rank-row">
        <div class="rank-num">{i+1}</div>
        <div class="rank-name" title="{nombre}">{nombre}</div>
        <div class="rank-bar-track"><div class="rank-bar-fill" style="width:{v/maxv*100:.1f}%;
             background:{color};"></div></div>
        <div class="rank-val">{fmt(v)}</div>
        <div class="rank-pct">{v/tot*100:.1f}%</div>
      </div>""" for i, (nombre, v) in enumerate(serie_top.items()))

    top5 = serie_top.sort_values(ascending=False).head(5).sum() / tot
    top10 = serie_top.sort_values(ascending=False).head(10).sum() / tot
    principal = serie_top.sort_values(ascending=False).iloc[0] / tot

    body = f"""
      <div class="card">
        <div class="card-header">{_dash_icono('people')}<p class="card-title">{titulo}</p>
          <span class="badge badge-off" style="margin-left:auto;">{badge_txt}</span></div>
        <p class="card-subtitle">{subtitulo}</p>
        {filas}
        <div style="border-top:1px solid {DASH_BORDER};margin-top:6px;padding-top:12px;">
          <div class="card-header" style="margin-bottom:0;">{_dash_icono('trophy')}
            <p class="card-title" style="font-size:13px;">Concentración</p></div>
          <div class="conc-row">
            <div class="conc-item"><div class="conc-label">Top 5 {etq_contraparte.lower()}</div>
              <div class="conc-value">{pct(top5)}</div></div>
            <div class="conc-item"><div class="conc-label">Top 10 {etq_contraparte.lower()}</div>
              <div class="conc-value">{pct(top10)}</div></div>
            <div class="conc-item"><div class="conc-label">{etq_contraparte_sing} principal</div>
              <div class="conc-value">{pct(principal)}</div></div>
          </div>
        </div>
      </div>"""
    return body, ""

def _analisis_flujo_datos(es_ingreso):
    """Prepara las variables compartidas por ambas paginas (mismos calculos de
    antes, solo separados de la presentacion)."""
    base_df = EM if es_ingreso else REC
    ppto_key = "ingresos" if es_ingreso else "gastos"
    tot_mes = float(base_df["importe"].sum()) if not base_df.empty else 0.0
    n_contrapartes = (base_df.loc[base_df["importe"] != 0, "contraparte"].nunique()
                      if not base_df.empty else 0)
    ppto_mensual, ppto_acumulado, mes_ppto = load_presupuesto()
    tiene_ppto = bool(ppto_mensual)
    ppto_mes_tot = float(sum(ppto_mensual.get(ppto_key, {}).values())) if tiene_ppto else None
    ppto_acum_tot = float(sum(ppto_acumulado.get(ppto_key, {}).values())) if tiene_ppto else None
    cumplimiento = tot_mes / ppto_mes_tot if ppto_mes_tot else None
    mes_txt = MES_COL_MAYUS.get(mes_ppto, "").title() if mes_ppto else (sel_labels[-1] if sel_labels else "—")
    return dict(base_df=base_df, tot_mes=tot_mes, n_contrapartes=n_contrapartes,
               tiene_ppto=tiene_ppto, ppto_mes_tot=ppto_mes_tot, ppto_acum_tot=ppto_acum_tot,
               cumplimiento=cumplimiento, mes_txt=mes_txt)

def _dash_header_frag(icono, titulo, subtitulo):
    anios = sorted({p.anio for p in periodos}) if periodos else []
    n_filtros = len(sel_labels) + (len(aduanas_sel_global) if aduanas_sel_global else 0)
    return header_bar_html(
        icono, titulo, subtitulo,
        ", ".join(sel_labels) if sel_labels else "—",
        ", ".join(str(a) for a in anios) if anios else "—",
        f"{n_filtros} seleccionados" if n_filtros else "Todos")

def _nota_pendientes():
    st.caption(("Solo PROPIOS" if SOLO_PROPIOS else "Todos los tipos") +
              " · Acumulado, comparativo contra el mes/año anterior y evolución mensual "
              "pendientes de cargar histórico real mes a mes (hoy solo hay un mes de datos "
              "reales en vivo); se muestran como \"No disponible\"/\"—\" en vez de inventarse.")

def _composicion_y_ranking_filas(base_df, etq_titulo, etq_sing, etq_contraparte, etq_contraparte_sing,
                                 color_real, tot_mes, campo_concepto):
    conc = base_df.copy()
    conc[campo_concepto] = conc[campo_concepto].fillna("Sin categorizar")
    por_concepto_pos = conc.groupby(campo_concepto)["importe"].sum()
    por_concepto_pos = por_concepto_pos[por_concepto_pos > 0].sort_values(ascending=False)
    por_contraparte = base_df.groupby("contraparte")["importe"].sum()
    por_contraparte = por_contraparte[por_contraparte > 0].sort_values(ascending=False)

    celda_izq = ("", "")
    if por_concepto_pos.empty:
        celda_izq = ('<div class="card empty-state">Sin conceptos para mostrar.</div>', "")
    else:
        tabla_concepto = por_concepto_pos.reset_index()
        tabla_concepto.columns = ["Concepto", "Monto"]
        dona_serie = por_concepto_pos.head(6)
        if len(por_concepto_pos) > 6:
            dona_serie = pd.concat([dona_serie, pd.Series({"Otros": por_concepto_pos.iloc[6:].sum()})])
        celda_izq = donut_detalle_card_html(
            f"Composición de {etq_titulo.lower()} por concepto",
            "Participación sobre el total y variación vs. periodo anterior", dona_serie, fmt_m(tot_mes),
            f"{etq_sing} total", f"Detalle de {etq_titulo.lower()} por concepto", "Concepto", tabla_concepto)

    celda_der = ("", "")
    if por_contraparte.empty:
        celda_der = (f'<div class="card empty-state">Sin {etq_contraparte.lower()} identificados.</div>', "")
    else:
        top10 = por_contraparte.head(10)
        celda_der = ranking_concentracion_card_html(
            f"Principales {etq_contraparte.lower()} por {etq_sing.lower()}",
            f"Top 10 {etq_contraparte.lower()} del periodo", "Top 10", top10, color_real,
            etq_contraparte, etq_contraparte_sing)
    return [(1, celda_izq[0], 600), (1, celda_der[0], 680)], celda_izq[1] + celda_der[1]

def _resumen_aduana_gastos_frag(d, base_df, aduana_sel):
    base_aduana = base_df[base_df["aduana"] == aduana_sel]
    tot_aduana = float(base_aduana["importe"].sum())
    cr_aduana = (costos_ref_f.loc[costos_ref_f["aduana"] == aduana_sel, "monto"].sum()
                if not costos_ref_f.empty else 0.0)
    return f"""
      <div class="card">
        <div class="card-header">{_dash_icono('pin')}<p class="card-title">Resumen de la aduana</p></div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:10px;">
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['doc']}</div>
            <div><div class="mini-val">{fmt(tot_aduana)}</div>
            <div class="mini-label">Gasto total del mes</div></div></div>
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['hex']}</div>
            <div><div class="mini-val">{fmt(cr_aduana)}</div>
            <div class="mini-label">Costos de referencia · {pct(cr_aduana/tot_aduana) if tot_aduana else '—'} del gasto</div></div></div>
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['list']}</div>
            <div><div class="mini-val">{len(base_aduana):,}</div>
            <div class="mini-label">Número de operaciones</div></div></div>
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['chart']}</div>
            <div><div class="mini-val">{pct(tot_aduana/d['tot_mes']) if d['tot_mes'] else '—'}</div>
            <div class="mini-label">% del gasto total</div></div></div>
        </div>
      </div>"""

def pagina_analisis_gastos():
    """'Análisis de gastos por aduana' -- clon de la estructura del mockup del
    usuario (2026-10-27): mismo orden/proporcion de tarjetas que la imagen,
    fundidas en UN documento HTML por bloque (no un iframe de Recharts por
    tarjeta ni st.columns de por medio) -- el usuario pidio explicitamente
    "no uses Streamlit, usa lo necesario para que quede igual". Lo unico que
    se queda nativo es el selector de Aduana (components.html no puede
    devolver un clic/seleccion a Python sin un componente bidireccional, ver
    [[feedback-respetar-referencias-visuales]])."""
    d = _analisis_flujo_datos(es_ingreso=False)
    base_df = d["base_df"]
    header_html, _ = _dash_header_frag("chart", "Análisis de gastos por aduana",
                                       "Comportamiento, composición y principales proveedores de gastos")
    if base_df.empty:
        render_pagina_html([[(1, header_html, 80)]])
        _nota_pendientes()
        st.info("Sin gastos para el periodo y aduana seleccionados.")
        return

    ev_html, ev_script = chart_bar_card_html(
        "chart", "Evolución de gastos", "Gasto real vs. presupuesto mensual",
        [d["mes_txt"]], [d["tot_mes"]], [d["ppto_mes_tot"] or 0], "Gasto real", "Presupuesto",
        JOCS_AZUL, "#9DBFE8", extra_note="1 de 12 meses cargados · se completa con más histórico",
        height=340)
    kpi_html, _ = stat_grid_card_html([
        {"icono": "money", "label": "Gasto del mes", "valor": fmt(d["tot_mes"]), "sub": d["mes_txt"]},
        {"icono": "stack", "label": "Gasto acumulado", "valor": "No disponible"},
        {"icono": "target", "label": "Presupuesto del mes",
         "valor": fmt(d["ppto_mes_tot"]) if d["ppto_mes_tot"] else "No disponible",
         "delta_txt": f"{d['cumplimiento']*100:.0f}% de cumplimiento" if d["cumplimiento"] is not None else None,
         "delta_pos": (d["cumplimiento"] >= 1.0) if d["cumplimiento"] is not None else None},
        {"icono": "trend", "label": "Presupuesto acumulado",
         "valor": fmt(d["ppto_acum_tot"]) if d["ppto_acum_tot"] else "No disponible"},
    ])
    fila2_html, _ = stat_grid_card_html([
        {"icono": "people", "label": "Proveedores activos", "valor": f"{d['n_contrapartes']:,}"},
        {"icono": "gauge", "label": "Resultado operativo", "valor": fmt(MARGEN),
         "delta_txt": f"Margen {MARGEN_PCT*100:.1f}%" if MARGEN_PCT is not None else None,
         "delta_pos": (MARGEN_PCT >= 0) if MARGEN_PCT is not None else None},
        {"icono": "trend", "label": "Variación vs. mes anterior", "valor": "No disponible"},
    ])
    variacion = (d["tot_mes"] / d["ppto_mes_tot"] - 1) if d["ppto_mes_tot"] else 0.0
    tend_html, tend_script = chart_line_card_html(
        "trend", "Tendencia de gasto", "Variación mensual del gasto real vs. presupuesto",
        [d["mes_txt"]], [variacion], [0.0], "Gasto real", "Presupuesto", JOCS_AZUL, "#9DBFE8",
        formato_pct=True, extra_note="1 de 12 meses cargados · se completa con más histórico",
        height=340)

    render_pagina_html([
        [(1, header_html, 80)],
        [(1.4, ev_html, 340), (1, kpi_html, 340)],
        [(1, fila2_html, 135)],
    ], extra_script=ev_script)
    _nota_pendientes()

    aduanas_disp = sorted(base_df["aduana"].dropna().unique().tolist())
    if not aduanas_disp:
        render_pagina_html([[(1, tend_html, 340)]], extra_script=tend_script)
        st.info("Sin aduanas con actividad.")
    else:
        aduana_sel = st.selectbox("Aduana", aduanas_disp, key="resumen_aduana_gastos")
        aduana_html = _resumen_aduana_gastos_frag(d, base_df, aduana_sel)
        filas_cr, script_cr = _composicion_y_ranking_filas(
            base_df, "Gastos", "Gasto", "Proveedores", "Proveedor", JOCS_AZUL, d["tot_mes"], "concepto_cierre")
        if cobertura_concepto < 0.95:
            st.caption(f"Solo el **{cobertura_concepto*100:.0f}%** del gasto trae un concepto de cierre "
                      "asignado en el Excel de origen; el resto aparece como 'Sin categorizar'.")
        render_pagina_html([
            [(1.4, tend_html, 340), (1, aduana_html, 340)],
            filas_cr,
        ], extra_script=tend_script + script_cr)

def _resumen_aduana_ingresos_frag(base_df, aduana_sel):
    base_aduana = base_df[base_df["aduana"] == aduana_sel]
    tot_aduana = float(base_aduana["importe"].sum())
    por_desc = (base_aduana.groupby("descripcion")["importe"].sum()
               .sort_values(ascending=False)) if not base_aduana.empty else pd.Series(dtype=float)
    top_conceptos = por_desc.head(3)
    otros_val = float(por_desc.iloc[3:].sum()) if len(por_desc) > 3 else 0.0
    filas_concepto = "".join(
        f'<div style="display:flex;justify-content:space-between;font-size:11.5px;'
        f'margin-bottom:5px;"><span style="color:{DASH_TEXT};overflow:hidden;'
        f'text-overflow:ellipsis;white-space:nowrap;max-width:160px;">{c}</span>'
        f'<span style="color:{DASH_SUBTEXT};font-weight:600;">{f"{v/tot_aduana*100:.1f}%" if tot_aduana else "—"}</span></div>'
        for c, v in top_conceptos.items())
    if otros_val:
        filas_concepto += (f'<div style="display:flex;justify-content:space-between;font-size:11.5px;">'
                           f'<span style="color:{DASH_TEXT};">Otros conceptos</span>'
                           f'<span style="color:{DASH_SUBTEXT};font-weight:600;">{f"{otros_val/tot_aduana*100:.1f}%" if tot_aduana else "—"}</span></div>')
    return f"""
      <div class="card">
        <div class="card-header">{_dash_icono('pin')}<p class="card-title">Resumen de la aduana</p></div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:10px 0 14px;">
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['doc']}</div>
            <div><div class="mini-val">{fmt(tot_aduana)}</div>
            <div class="mini-label">Ingreso total del mes</div></div></div>
          <div class="mini-stat"><div class="stat-icon">{DASH_ICONOS['list']}</div>
            <div><div class="mini-val">{len(base_aduana):,}</div>
            <div class="mini-label">Número de operaciones</div></div></div>
        </div>
        <p class="card-title" style="font-size:12.5px;margin:0 0 8px;">Ingresos por concepto (mes)</p>
        {filas_concepto if filas_concepto else '<div style="color:' + DASH_SUBTEXT + ';font-size:12px;">Sin datos.</div>'}
      </div>"""

def pagina_analisis_ingresos():
    """'Análisis de Ingresos' -- el mockup usa una disposición distinta a la de
    Gastos (4 tarjetas en una sola fila arriba, fila angosta-grafica-angosta
    despues), por eso es su propia funcion y no un parametro de la de Gastos.
    Igual que Gastos, todo vive en documentos HTML fundidos -- el selector de
    Aduana es el unico control nativo (ver nota en pagina_analisis_gastos)."""
    d = _analisis_flujo_datos(es_ingreso=True)
    base_df = d["base_df"]
    header_html, _ = _dash_header_frag("chart", "Análisis de Ingresos",
                                       "Comportamiento, composición y principales generadores de ingresos")
    if base_df.empty:
        render_pagina_html([[(1, header_html, 80)]])
        _nota_pendientes()
        st.info("Sin ingresos para el periodo y aduana seleccionados.")
        return

    kpi_html, _ = stat_grid_card_html([
        {"icono": "money", "label": "Ingreso del mes", "valor": fmt(d["tot_mes"]), "sub": d["mes_txt"]},
        {"icono": "stack", "label": "Ingreso acumulado", "valor": "No disponible"},
        {"icono": "target", "label": "Presupuesto del mes",
         "valor": fmt(d["ppto_mes_tot"]) if d["ppto_mes_tot"] else "No disponible",
         "delta_txt": f"{d['cumplimiento']*100:.0f}% de cumplimiento" if d["cumplimiento"] is not None else None,
         "delta_pos": (d["cumplimiento"] >= 1.0) if d["cumplimiento"] is not None else None},
        {"icono": "trend", "label": "Presupuesto acumulado",
         "valor": fmt(d["ppto_acum_tot"]) if d["ppto_acum_tot"] else "No disponible"},
    ])
    crec_html, _ = stat_grid_card_html([
        {"icono": "trend", "label": "Crecimiento vs. mes anterior", "valor": "No disponible"},
    ])
    ev_html, ev_script = chart_bar_card_html(
        "chart", "Evolución de ingresos", "Ingreso real vs. presupuesto mensual",
        [d["mes_txt"]], [d["tot_mes"]], [d["ppto_mes_tot"] or 0], "Ingreso real", "Presupuesto",
        JOCS_AZUL, "#9DBFE8", extra_note="1 de 12 meses cargados · se completa con más histórico",
        height=300)
    resop_html, _ = stat_grid_card_html([
        {"icono": "people", "label": "Clientes activos", "valor": f"{d['n_contrapartes']:,}"},
        {"icono": "gauge", "label": "Resultado operativo",
         "valor": ("Superávit" if MARGEN is not None and MARGEN >= 0 else "Déficit") if MARGEN is not None else "—",
         "delta_txt": f"Margen {MARGEN_PCT*100:.1f}%" if MARGEN_PCT is not None else None,
         "delta_pos": (MARGEN_PCT >= 0) if MARGEN_PCT is not None else None},
    ])
    variacion = (d["tot_mes"] / d["ppto_mes_tot"] - 1) if d["ppto_mes_tot"] else 0.0
    tend_html, tend_script = chart_line_card_html(
        "trend", "Tendencia de crecimiento", "Variación mensual de los ingresos",
        [d["mes_txt"]], [variacion], [0.0], "Ingreso real", "Presupuesto", JOCS_AZUL, "#9DBFE8",
        formato_pct=True, extra_note="1 de 12 meses cargados · se completa con más histórico",
        height=340)

    render_pagina_html([
        [(1, header_html, 80)],
        [(1, kpi_html, 135)],
        [(0.8, crec_html, 300), (1.6, ev_html, 300), (0.8, resop_html, 300)],
    ], extra_script=ev_script)
    _nota_pendientes()

    aduanas_disp = sorted(base_df["aduana"].dropna().unique().tolist())
    if not aduanas_disp:
        render_pagina_html([[(1, tend_html, 340)]], extra_script=tend_script)
        st.info("Sin aduanas con actividad.")
    else:
        aduana_sel = st.selectbox("Aduana", aduanas_disp, key="resumen_aduana_ingresos")
        aduana_html = _resumen_aduana_ingresos_frag(base_df, aduana_sel)
        st.caption("Emitidas no trae una categoría de cierre como Recibidas; se agrupa por "
                  "**Descripción** de la línea en su lugar.")
        filas_cr, script_cr = _composicion_y_ranking_filas(
            base_df, "Ingresos", "Ingreso", "Clientes", "Cliente", C_EMITIDA, d["tot_mes"], "descripcion")
        render_pagina_html([
            [(1.4, tend_html, 340), (1, aduana_html, 340)],
            filas_cr,
        ], extra_script=tend_script + script_cr)

def _uo_payload():
    """Datos de 'Utilidad Operativa' para el componente React -- mismos
    calculos que antes (ver historia de este archivo), solo que en vez de
    pintar tarjetas HTML, arma un dict para mandar completo al componente."""
    if utilidad_op_f.empty:
        return None
    uo_ad = utilidad_op_f.groupby("aduana")[["emitidas", "HONORARIOS AA", "SUELDOS",
                                             "ctos_referencia", "recibida", "ingreso",
                                             "egreso", "utilidad"]].sum()
    uo_ad["gastos_operativos"] = uo_ad["egreso"] - uo_ad["SUELDOS"] - uo_ad["ctos_referencia"]

    ppto_mensual, ppto_acumulado, mes_ppto = load_presupuesto()
    tiene_ppto = bool(ppto_mensual)
    ppto_df = pd.DataFrame()
    ppto_anual_df = pd.DataFrame()
    if tiene_ppto:
        ppto_df = pd.DataFrame({c: pd.Series(v) for c, v in ppto_mensual.items()}).fillna(0.0)
        ppto_df["utilidad"] = ppto_df["ingresos"] - ppto_df["costos"] - ppto_df["sueldos"] - ppto_df["gastos"]
        ppto_acum_df = pd.DataFrame({c: pd.Series(v) for c, v in ppto_acumulado.items()}).fillna(0.0)
        ppto_acum_df["utilidad"] = (ppto_acum_df["ingresos"] - ppto_acum_df["costos"]
                                    - ppto_acum_df["sueldos"] - ppto_acum_df["gastos"])
        ppto_anual_df = ppto_acum_df / (mes_ppto or 1) * 12

    por_aduana = [
        {"aduana": str(a), "ingreso": float(r["ingreso"]), "costos_referencia": float(r["ctos_referencia"]),
         "gastos_operativos": float(r["gastos_operativos"]), "sueldos": float(r["SUELDOS"]),
         "utilidad": float(r["utilidad"])}
        for a, r in uo_ad.iterrows()
    ]
    ppto_por_aduana = []
    if tiene_ppto:
        for a, r in ppto_df.iterrows():
            ppto_por_aduana.append({"aduana": str(a), "ingresos": float(r["ingresos"]), "costos": float(r["costos"]),
                                    "gastos": float(r["gastos"]), "sueldos": float(r["sueldos"]),
                                    "utilidad": float(r["utilidad"])})

    def _t(campo):
        return float(uo_ad[campo].sum())
    def _p(campo):
        return float(ppto_df[campo].sum()) if tiene_ppto else None
    def _pa(campo):
        return float(ppto_anual_df[campo].sum()) if tiene_ppto else None

    estado_resultados = {
        "ingresos": {"real": _t("ingreso"), "ppto": _p("ingresos"), "ppto_anual": _pa("ingresos")},
        "costos_referencia": {"real": _t("ctos_referencia"), "ppto": _p("costos"), "ppto_anual": _pa("costos")},
        "sueldos": {"real": _t("SUELDOS"), "ppto": _p("sueldos"), "ppto_anual": _pa("sueldos")},
        "gastos_operativos": {"real": _t("gastos_operativos"), "ppto": _p("gastos"), "ppto_anual": _pa("gastos")},
        "utilidad": {"real": _t("utilidad"), "ppto": _p("utilidad"), "ppto_anual": _pa("utilidad")},
    }
    mes_txt = MES_COL_MAYUS.get(mes_ppto, "").title() if mes_ppto else (sel_labels[-1] if sel_labels else "—")
    return {"tiene_ppto": tiene_ppto, "mes_txt": mes_txt, "por_aduana": por_aduana,
           "ppto_por_aduana": ppto_por_aduana, "estado_resultados": estado_resultados}

def _flujo_payload(es_ingreso):
    """Datos de 'Gastos'/'Ingresos' para el componente React -- misma logica
    que pagina_analisis_flujo()/_composicion_y_ranking_filas() de antes, pero
    TODAS las aduanas (no solo la seleccionada) para que el dropdown de
    'Resumen de la aduana' funcione del lado del cliente sin volver a
    Python."""
    base_df = EM if es_ingreso else REC
    if base_df.empty:
        return None
    ppto_key = "ingresos" if es_ingreso else "gastos"
    ppto_mensual, ppto_acumulado, mes_ppto = load_presupuesto()
    tiene_ppto = bool(ppto_mensual)
    tot_mes = float(base_df["importe"].sum())
    ppto_mes_tot = float(sum(ppto_mensual.get(ppto_key, {}).values())) if tiene_ppto else None
    ppto_acum_tot = float(sum(ppto_acumulado.get(ppto_key, {}).values())) if tiene_ppto else None
    n_contrapartes = int(base_df.loc[base_df["importe"] != 0, "contraparte"].nunique())
    mes_txt = MES_COL_MAYUS.get(mes_ppto, "").title() if mes_ppto else (sel_labels[-1] if sel_labels else "—")

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
            cr_aduana = (float(costos_ref_f.loc[costos_ref_f["aduana"] == aduana, "monto"].sum())
                        if not costos_ref_f.empty else 0.0)
            entry["costos_referencia"] = cr_aduana
            entry["costos_referencia_pct"] = float(cr_aduana / tot_aduana) if tot_aduana else 0.0
        por_aduana.append(entry)

    cumplimiento = tot_mes / ppto_mes_tot if ppto_mes_tot else None
    return {
        "tiene_ppto": tiene_ppto, "mes_txt": mes_txt,
        "kpis": {"tot_mes": tot_mes, "ppto_mes_tot": ppto_mes_tot, "ppto_acum_tot": ppto_acum_tot,
                "n_contrapartes": n_contrapartes, "cumplimiento": cumplimiento,
                "margen": float(MARGEN), "margen_pct": (float(MARGEN_PCT) if MARGEN_PCT is not None else None)},
        "por_aduana": por_aduana,
        "por_concepto": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_dona.items()],
        "por_concepto_full": [{"concepto": str(c), "monto": float(v)} for c, v in por_concepto_full.items()],
        "por_contraparte": [{"nombre": str(c), "monto": float(v)} for c, v in top10.items()],
        "cobertura_concepto": (float(cobertura_concepto) if not es_ingreso else None),
    }

# ══════════════════════════════════════════════════════════════════════════════
# PANORAMA
# ══════════════════════════════════════════════════════════════════════════════
if pagina in PAGINAS_REDISENADAS:
    _dash_payload = {
        "pagina_activa": pagina,
        "nav_opciones": NAV_OPCIONES,
        "utilidad_operativa": _uo_payload(),
        "gastos": _flujo_payload(es_ingreso=False),
        "ingresos": _flujo_payload(es_ingreso=True),
        "colores": {"azul": JOCS_AZUL, "verde": JOCS_VERDE, "neg": C_NEG, "margen": C_MARGEN, "ppto": "#9DBFE8"},
    }
    _dash_resultado = _dashboard_component(datos=_dash_payload, key="fac_dashboard", default=None)
    if _dash_resultado and _dash_resultado.get("nav_to"):
        st.session_state["nav"] = _dash_resultado["nav_to"]
        st.rerun()

if False:  # CODIGO MUERTO, pendiente de borrar: Utilidad Operativa ahora se
          # resuelve arriba, en "if pagina in PAGINAS_REDISENADAS:" (componente
          # React unico). Se deja sin quitar por ahora para no perder las formulas
          # si hace falta revisarlas; no se ejecuta nunca.
    st.title("Utilidad Operativa")
    st.caption(f"{', '.join(sel_labels)} · Importes en MXN" + (" · Solo PROPIOS" if SOLO_PROPIOS else ""))

    if n_cancel:
        st.info(f"Se excluyeron **{n_cancel} líneas canceladas** ({fmt(monto_cancel)}) de todas las cifras.")
    if n_sin_monto:
        st.caption(f"{n_sin_monto:,} líneas traen importe pero no monto en el Excel "
                   f"({fmt(importe_sin_monto)} de importe) y no se suman a las cifras.")

    if utilidad_op_f.empty:
        st.info("Sin datos de Utilidad Operativa para el periodo y aduana seleccionados.")
    else:
        uo_ad = utilidad_op_f.groupby("aduana")[["emitidas", "HONORARIOS AA", "SUELDOS",
                                                 "ctos_referencia", "recibida", "ingreso",
                                                 "egreso", "utilidad"]].sum()
        # "Gastos Operativos" = lo que sobra del egreso una vez quitados Sueldos y
        # Costos de Referencia -- mismo calculo que 'recibida_visual' del cascada
        # anterior, es el concepto real que se compara contra "GASTOS OP" del
        # presupuesto (ver load_presupuesto).
        uo_ad["gastos_operativos"] = uo_ad["egreso"] - uo_ad["SUELDOS"] - uo_ad["ctos_referencia"]
        aduanas_uo = uo_ad.sort_values("utilidad", ascending=False).index.tolist()

        # ── Presupuesto (ver nota junto a PRESUPUESTO_PATH: archivo local, aun
        # no migrado al Sheet en vivo) ──────────────────────────────────────────
        ppto_mensual, ppto_acumulado, mes_ppto = load_presupuesto()
        tiene_ppto = bool(ppto_mensual)
        ppto_df = pd.DataFrame()
        ppto_anual_df = pd.DataFrame()
        if tiene_ppto:
            ppto_df = pd.DataFrame({c: pd.Series(v) for c, v in ppto_mensual.items()}).fillna(0.0)
            ppto_df["utilidad"] = (ppto_df["ingresos"] - ppto_df["costos"]
                                   - ppto_df["sueldos"] - ppto_df["gastos"])
            ppto_acum_df = pd.DataFrame({c: pd.Series(v) for c, v in ppto_acumulado.items()}).fillna(0.0)
            ppto_acum_df["utilidad"] = (ppto_acum_df["ingresos"] - ppto_acum_df["costos"]
                                        - ppto_acum_df["sueldos"] - ppto_acum_df["gastos"])
            # Ppto anual = PROYECCION (acumulado a la fecha / meses transcurridos x
            # 12), no una meta anual real -- no existe ese dato todavia, a pedido
            # del usuario se muestra como proyeccion mientras tanto.
            ppto_anual_df = ppto_acum_df / (mes_ppto or 1) * 12
        else:
            st.warning("No se encontró el archivo de presupuesto local "
                      f"(`{PRESUPUESTO_PATH.name}`); se muestra solo lo Real.")

        # ══════════════════════════════════════════════════════════════════════
        # RESUMEN EJECUTIVO -- selects nativos (cambian el dato), tarjetas fundidas
        # ══════════════════════════════════════════════════════════════════════
        st.divider()
        st.subheader("Resumen Ejecutivo")
        fc1, fc2 = st.columns(2)
        with fc1:
            meses_resumen = st.multiselect("Mes", labels_disp, default=sel_labels, key="uo_resumen_mes")
        with fc2:
            aduanas_resumen = st.multiselect("Aduana", aduanas_uo, default=[], key="uo_resumen_aduana",
                                             placeholder="Todas las aduanas")
        uo_resumen = (utilidad_op_f[utilidad_op_f["periodo_label"].isin(meses_resumen)]
                     if meses_resumen else utilidad_op_f.iloc[0:0])
        if aduanas_resumen:
            uo_resumen = uo_resumen[uo_resumen["aduana"].isin(aduanas_resumen)]
        uo_resumen_ad = uo_resumen.groupby("aduana")[["SUELDOS", "ctos_referencia", "ingreso",
                                                       "egreso", "utilidad"]].sum()
        uo_resumen_ad["gastos_operativos"] = (uo_resumen_ad["egreso"] - uo_resumen_ad["SUELDOS"]
                                              - uo_resumen_ad["ctos_referencia"])
        aduanas_ppto_sel = aduanas_resumen or aduanas_uo

        conceptos_kpi = [("Ingresos", "money", "ingreso", "ingresos", C_EMITIDA),
                         ("Costos de Referencia", "target", "ctos_referencia", "costos", C_NEG),
                         ("Gastos Operativos", "stack", "gastos_operativos", "gastos", JOCS_VERDE),
                         ("Utilidad Operativa", "trend", "utilidad", "utilidad", C_MARGEN)]
        kpi_cards = []
        for nombre, icono, real_key, ppto_key, color in conceptos_kpi:
            real_val = float(uo_resumen_ad[real_key].sum()) if not uo_resumen_ad.empty else 0.0
            card = {"icono": icono, "label": nombre, "valor": fmt(real_val)}
            if tiene_ppto:
                ppto_val = float(ppto_df.loc[ppto_df.index.isin(aduanas_ppto_sel), ppto_key].sum())
                delta = real_val / ppto_val - 1 if ppto_val else None
                card["sub"] = f"Ppto: {fmt(ppto_val)}" + (f" · {pct(delta)} vs ppto" if delta is not None else "")
                card["barra_html"] = barra_real_ppto_html(real_val, ppto_val, color)
            else:
                card["sub"] = "Ppto no disponible"
            kpi_cards.append(card)
        kpi_html, _ = stat_grid_card_html(kpi_cards)
        render_pagina_html([[(1, kpi_html, 160)]])
        st.caption("Acumulado y avance contra presupuesto anual pendientes de cargar el histórico "
                  "real mes a mes (hoy solo hay un mes de datos reales en vivo).")

        # ══════════════════════════════════════════════════════════════════════
        # RENTABILIDAD POR ADUANA -- toggle nativo (cambia el dato), tarjeta fundida
        # ══════════════════════════════════════════════════════════════════════
        st.divider()
        modo_rank = st.radio("Ver", ["Top 5", "Todas"], horizontal=True, key="uo_rank_modo")
        orden_rank = uo_ad.sort_values("utilidad", ascending=False)
        n_total_aduanas = len(orden_rank)
        if modo_rank == "Top 5":
            orden_rank = orden_rank.head(5)
        df_rank_chart = orden_rank.reset_index()[["aduana", "utilidad"]].rename(columns={"utilidad": "real"})
        if tiene_ppto:
            df_rank_chart["ppto"] = df_rank_chart["aduana"].map(ppto_df["utilidad"])
            df_rank_chart["delta_txt"] = df_rank_chart.apply(
                lambda r: pct(r["real"] / r["ppto"] - 1) if pd.notna(r["ppto"]) and r["ppto"] else "", axis=1)
        else:
            df_rank_chart["ppto"] = None
            df_rank_chart["delta_txt"] = ""
        margen_cons = uo_ad["utilidad"].sum() / uo_ad["ingreso"].sum() if uo_ad["ingreso"].sum() else None
        consolidado_html = (f'<p style="margin:0 0 10px;font-size:13px;color:{DASH_TEXT};">'
                            f'<b>Utilidad consolidada: {fmt(uo_ad["utilidad"].sum())}</b>' +
                            (f' · Margen {pct(margen_cons)}' if margen_cons is not None else '') + '</p>')
        rank_h = max(280, 44 * len(df_rank_chart) + 60)
        rank_html, rank_script = ranking_aduana_card_frag(
            "Rentabilidad por Aduana", "Utilidad Operativa real (barra) vs. presupuesto del mes (marca ámbar).",
            consolidado_html, df_rank_chart, JOCS_AZUL, height=rank_h)
        render_pagina_html([[(1, rank_html, rank_h)]], extra_script=rank_script)
        if modo_rank == "Top 5" and n_total_aduanas > 5:
            st.caption(f"Se muestran 5 de {n_total_aduanas} aduanas. Cambia a \"Todas\" para verlas completas.")

        # ══════════════════════════════════════════════════════════════════════
        # ESTADO DE RESULTADOS
        # ══════════════════════════════════════════════════════════════════════
        if tiene_ppto:
            st.divider()
            mes_ppto_txt = MES_COL_MAYUS.get(mes_ppto, "").title() if mes_ppto else "—"

            ingresos_real, costos_real = float(uo_ad["ingreso"].sum()), float(uo_ad["ctos_referencia"].sum())
            sueldos_real, gastos_real = float(uo_ad["SUELDOS"].sum()), float(uo_ad["gastos_operativos"].sum())
            utilidad_real = float(uo_ad["utilidad"].sum())
            ingresos_ppto, costos_ppto = float(ppto_df["ingresos"].sum()), float(ppto_df["costos"].sum())
            sueldos_ppto, gastos_ppto = float(ppto_df["sueldos"].sum()), float(ppto_df["gastos"].sum())
            utilidad_ppto = float(ppto_df["utilidad"].sum())
            pa_ingresos, pa_costos = float(ppto_anual_df["ingresos"].sum()), float(ppto_anual_df["costos"].sum())
            pa_sueldos, pa_gastos = float(ppto_anual_df["sueldos"].sum()), float(ppto_anual_df["gastos"].sum())
            pa_utilidad = float(ppto_anual_df["utilidad"].sum())
            cg_real = costos_real + sueldos_real + gastos_real
            cg_ppto = costos_ppto + sueldos_ppto + gastos_ppto
            cg_ppto_anual = pa_costos + pa_sueldos + pa_gastos

            def _fila_er(nombre, real, ppto, ppto_anual, indent=False):
                nom = ("&nbsp;&nbsp;&nbsp;&nbsp;" if indent else "") + nombre
                return (f'<tr><td>{nom}</td><td class="num">{fmt(real)}</td>'
                       f'<td class="num">{fmt(ppto)}</td>'
                       f'<td class="num">{pct(real / ppto - 1) if ppto else "—"}</td>'
                       f'<td class="num">{fmt(ppto_anual)}</td></tr>')

            filas_er = "".join([
                _fila_er("Ingresos", ingresos_real, ingresos_ppto, pa_ingresos),
                _fila_er("Costos y gastos", cg_real, cg_ppto, cg_ppto_anual),
                _fila_er("Costos de Referencia", costos_real, costos_ppto, pa_costos, indent=True),
                _fila_er("Sueldos", sueldos_real, sueldos_ppto, pa_sueldos, indent=True),
                _fila_er("Gastos Operativos", gastos_real, gastos_ppto, pa_gastos, indent=True),
                _fila_er("Utilidad Operativa", utilidad_real, utilidad_ppto, pa_utilidad),
            ])
            er_body = f"""
              <div class="card">
                <div class="card-header">{_dash_icono('doc')}<p class="card-title">Estado de Resultados</p></div>
                <p class="card-subtitle">Mes: {mes_ppto_txt}. Vista acumulada y comparativo contra 2025
                  pendientes de histórico (ver Resumen Ejecutivo). "Ppto. anual" es una proyección simple
                  (acumulado a la fecha ÷ meses transcurridos × 12), no una meta anual real.</p>
                <table class="dtable">
                  <tr><th>Concepto</th><th class="num">Real</th><th class="num">Presupuesto</th>
                      <th class="num">Var. %</th><th class="num">Ppto. anual (proyectado)</th></tr>
                  {filas_er}
                </table>
              </div>"""
            render_pagina_html([[(1, er_body, 330)]])

# Gastos/Ingresos ya se resuelven arriba, dentro de "if pagina in
# PAGINAS_REDISENADAS:" (componente React unico) -- pagina_analisis_gastos()/
# pagina_analisis_ingresos() quedan sin usar, pendientes de limpieza.

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
    st.subheader("Desglose por Aduana, Personal y Percepción")
    st.caption("Haz clic en un cuadro para profundizar: aduana → empleados de esa aduana → "
               "componentes de su percepción. La miga de pan de arriba regresa a un nivel anterior.")

    if "nom_path" not in st.session_state:
        st.session_state.nom_path = []

    arbol_nomina = _construir_arbol_nomina(sueldos_f)
    aduanas_arbol = {n["name"]: n for n in arbol_nomina}

    # Si cambiaron los filtros globales y la aduana/empleado seleccionado ya no
    # existe en los datos filtrados, se descarta la selección en vez de tronar.
    nom_path = list(st.session_state.nom_path)
    if nom_path and nom_path[0] not in aduanas_arbol:
        nom_path = []
    elif len(nom_path) >= 2:
        empleados_arbol = {e["name"] for e in aduanas_arbol[nom_path[0]]["empleados"]}
        if nom_path[1] not in empleados_arbol:
            nom_path = nom_path[:1]
    if nom_path != st.session_state.nom_path:
        st.session_state.nom_path = nom_path

    clic_path = treemap_nomina_recharts(arbol_nomina, height=460, path=st.session_state.nom_path,
                                        key="tm_nomina")
    if clic_path is not None and clic_path != st.session_state.nom_path:
        st.session_state.nom_path = clic_path
        st.rerun()

    nom_path = st.session_state.nom_path
    aduana_nom_sel = nom_path[0] if len(nom_path) >= 1 else None
    empleado_sel = nom_path[1] if len(nom_path) >= 2 else None

    if aduana_nom_sel is None:
        st.info("Haz clic en una aduana para ver sus empleados.")
    else:
        tabla_nom = sueldos_f[sueldos_f["aduana"] == aduana_nom_sel].copy()
        tot_aduana_nom = tabla_nom["total"].sum()
        tabla_nom["pct"] = tabla_nom["total"] / tot_aduana_nom if tot_aduana_nom else 0.0
        tabla_nom = tabla_nom.sort_values("total", ascending=False)
        st.caption(f"Total nómina {aduana_nom_sel}: {fmt_m(tot_aduana_nom)} · {len(tabla_nom)} empleado(s)")

        st.dataframe(pd.DataFrame({
            "Empleado": tabla_nom["nombre"],
            "% Conformación Aduana": tabla_nom["pct"].map(lambda v: f"{v*100:.1f}%"),
            "Sueldo Total": tabla_nom["total"].map(fmt),
        }), use_container_width=True, hide_index=True)

        if empleado_sel is None:
            st.info("Haz clic en un empleado para ver el desglose de su percepción.")
        else:
            st.divider()
            st.subheader("Ficha Técnica del Trabajador")
            fila_emp = sueldos_f[sueldos_f["nombre"] == empleado_sel]
            if len(fila_emp) > 1:
                st.caption(f"{len(fila_emp)} registros para este empleado en el periodo — se suman.")
            fila_emp_sum = fila_emp[[c for c, _ in COLS_NOMINA] + ["total"]].sum()

            c1, c2 = st.columns(2)
            c1.metric("Aduana", fila_emp["aduana"].iloc[0] if len(fila_emp) else "—")
            c2.metric("Costo Total Neto", fmt_m(fila_emp_sum["total"]))

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

