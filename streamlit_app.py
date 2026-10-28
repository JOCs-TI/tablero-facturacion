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

# ── Paleta de marca JOCS (tablero propio, NO reusa la identidad de Ancora) ─────
JOCS_AZUL = "#244072"
JOCS_VERDE = "#507058"
C_MARGEN = "#c98a2e"
C_NEG = "#c0392b"

_dashboard_component = components.declare_component(
    "facturacion_dashboard", path=str(BASE / "components" / "facturacion_dashboard"))

components.html(
    "<script>try{var d=window.parent.document.documentElement;"
    "d.setAttribute('translate','no');d.classList.add('notranslate');}catch(e){}</script>",
    height=0,
)

NAV_OPCIONES = ["Utilidad Operativa", "Gastos", "Ingresos", "Nómina", "Tarjeta de Crédito"]
if "nav" not in st.session_state:
    st.session_state["nav"] = "Utilidad Operativa"
pagina = st.session_state["nav"]

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

# 2026-10-28: a pedido del usuario se borró toda la fuente de datos vieja
# (Google Sheet único + carpeta de Drive con un Sheet por categoría) y todas
# las reglas de negocio que se habían ido acumulando sobre ella (SOLO_PROPIOS,
# alias de aduanas, "Ingreso = Emitidas + Honorarios AA", exclusión de IMILE
# AIFA, etc.) -- se van a reconstruir de a poco sobre una fuente más
# organizada, todavía sin definir. El componente React de abajo
# (components/facturacion_dashboard/) NO cambió: sigue siendo el diseño de
# las 5 páginas, listo para recibir un payload con datos reales en cuanto
# haya una fuente nueva. Mientras tanto se le manda None en las 5 -- el
# propio componente ya sabe mostrar el estado "sin datos" en cada página
# (ver EmptyCard en index.html).
_dash_payload = {
    "pagina_activa": pagina,
    "nav_opciones": NAV_OPCIONES,
    "utilidad_operativa": None,
    "gastos": None,
    "ingresos": None,
    "nomina": None,
    "tarjeta_credito": None,
    "colores": {"azul": JOCS_AZUL, "verde": JOCS_VERDE, "neg": C_NEG, "margen": C_MARGEN, "ppto": "#9DBFE8"},
}
_dash_resultado = _dashboard_component(datos=_dash_payload, key="fac_dashboard", default=None)
if _dash_resultado and _dash_resultado.get("nav_to"):
    st.session_state["nav"] = _dash_resultado["nav_to"]
    st.rerun()
