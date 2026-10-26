"""
Tablero ejecutivo ISSSTE · Streamlit
Gráfico combinado Trabajadores + Plazas · Mapa coroplético de Tasa.
"""
from pathlib import Path
import json
import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# ═════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ═════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="ISSSTE · Tablero Interactivo",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

# ═════════════════════════════════════════════════════════════════════════
# ESCALA AUTOMÁTICA SEGÚN EL ANCHO DE LA PANTALLA
# K = 1.0 → el tablero se ve exactamente como lo diseñaste en ANCHO_REF.
# En un monitor más grande K sube y TODO crece en proporción (letras, tarjetas,
# panel de filtros, gráficos); en uno más chico baja. Así la composición no cambia.
# ═════════════════════════════════════════════════════════════════════════
import streamlit.components.v1 as components

ANCHO_REF = 1536          # ancho (px CSS) donde el tablero se ve "como lo diseñé"
K_MIN, K_MAX = 0.80, 2.50 # límites para que no se deforme en extremos
_viewport = components.declare_component(
    "viewport", path=str(Path(__file__).parent / "componentes" / "viewport"))

_vw = _viewport(key="vp", default=0)          # ancho real de la ventana del navegador
if _vw:
    st.session_state["_vw"] = int(_vw)
VW = st.session_state.get("_vw", ANCHO_REF)   # 1ª carga: sin dato → K = 1
K = max(K_MIN, min(K_MAX, round(VW / ANCHO_REF / 0.05) * 0.05))

def S(v):
    """Escala un tamaño en px (letra, margen, alto) por el factor K."""
    return round(v * K, 1)

def _escalar_px(txt: str) -> str:
    """Multiplica por K todo valor 'NNpx' dentro de un bloque CSS / HTML inline."""
    return re.sub(r"(\d+(?:\.\d+)?)px", lambda m: f"{float(m.group(1)) * K:.1f}px", txt)

BASE_DIR     = Path(__file__).parent
F            = BASE_DIR / "data" / "RESUMEN_ISSSTE.xlsx"
GEOJSON_PATH = BASE_DIR / "data" / "mexico_states.geojson"

SHEETS = [
    "TABLAS_DIM", "ISSSTE_TRAB_TOT2015_2025", "ISSSTE_TRAB_NOMBRA", "ISSSTE_TRAB_SM",
    "ISSSTE_TRAB_NPLAZAS", "ISSSTE_PLAZ_NOMBRA_PROM", "ISSSTE_PENS_REGPEN",
    "ISSSTE_PENS_PP_PROM", "ISSSTE_PENS_SERV_PROM", "ISSSTE_DEUD_TOT",
    "ISSSTE_DEUD_TBNDEUDO", "ISSSTE_DEUD_TBNDEUDO_PROM", "ISSSTE_DEUD_TBNDEUDO_RET",
    "ISSSTE_FAMILIAR_PARENT",
]

# ═════════════════════════════════════════════════════════════════════════
# CARGA DE DATOS
# ═════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Cargando información del ISSSTE…")
def load(path: str, modified_ns: int, file_size: int):
    file_stat = Path(path).stat()
    if (file_stat.st_mtime_ns, file_stat.st_size) != (modified_ns, file_size):
        raise RuntimeError("El libro de Excel cambió durante la carga; vuelve a ejecutar.")
    return pd.read_excel(path, sheet_name=SHEETS, engine="openpyxl")

if not F.exists():
    st.error(f"No se encontró el archivo de datos:\n\n`{F}`")
    st.stop()

try:
    _file_stat = F.stat()
    # Metadata in the arguments invalidates the cache when the workbook changes.
    _raw = load(str(F), _file_stat.st_mtime_ns, _file_stat.st_size)
    # pandas treats the literal "NA" as missing by default; preserve it as an entity code.
    _ent_trab = pd.read_excel(
        F,
        sheet_name="ISSSTE_TRAB_TOT2015_2025",
        usecols=["ENT_CVE"],
        dtype={"ENT_CVE": "string"},
        keep_default_na=False,
        engine="openpyxl",
    )
    _raw["ISSSTE_TRAB_TOT2015_2025"]["ENT_CVE"] = (
        _ent_trab["ENT_CVE"].replace("", pd.NA)
    )
except Exception as e:
    st.error(f"No se pudo leer `{F}`: {e}")
    st.stop()

def _norm(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    if "Total" in d.columns and "TOTAL" not in d.columns:
        d = d.rename(columns={"Total": "TOTAL"})
    return d

D = {k: _norm(v) for k, v in _raw.items()}

# ═════════════════════════════════════════════════════════════════════════
# PALETA EJECUTIVA
# ═════════════════════════════════════════════════════════════════════════
GUINDA, GUINDA_DARK = "#74192D", "#4B0F1D"
NAVY, BLUE, SKY     = "#0B1F4B", "#1F6FB2", "#4FA8E0"
TEAL, AMBER, CORAL  = "#217A60", "#C9A227", "#B04A5A"
GRAY, BG, GRID      = "#6B7280", "#F4F5F7", "#EDEFF3"
C = [GUINDA, NAVY, BLUE, TEAL, AMBER, CORAL, SKY, GRAY]
SEX_COLORS = {"H": NAVY, "M": CORAL}

# ═════════════════════════════════════════════════════════════════════════
# ESTILOS CSS
# ═════════════════════════════════════════════════════════════════════════
_CSS = f"""
<style>
.stApp {{ background: {BG}; }}
/*.block-container {{ padding-top: 0.8rem; padding-bottom: 2rem; max-width: 1500px; }}*/

.stApp {{ background: {BG}; }}

/* ── Eliminar TODO el espacio superior ── */
[data-testid="stHeader"] {{
    height: 0 !important;
    min-height: 0 !important;
    background: transparent !important;
}}

/* Ocultar el botón "Deploy" y la barra de herramientas */
[data-testid="stToolbar"] {{
    display: none !important;
}}
[data-testid="stDecoration"] {{
    display: none !important;
}}

[data-testid="stAppViewContainer"] > .main {{
    padding-top: 0 !important;
}}

[data-testid="stMainBlockContainer"],
[data-testid="stMain"] > div:first-child {{
    padding-top: 0 !important;
    margin-top: 0 !important;
}}

.block-container {{
    padding-top: 0 !important;
    margin-top: 0 !important;
    padding-left: 4rem !important;
    padding-right: 1.4rem !important;
    padding-bottom: 2rem;
    max-width: 100% !important;
}}

/* Hero pegado al top */
.hero {{
    padding: 10px 18px !important;
    margin-top: 0 !important;
    margin-bottom: 8px !important;
}}

/* Sidebar pegado al top también */
[data-testid="stSidebar"] > div:first-child {{
    width: 300px !important;
    padding-right: 1.2rem !important;   /* ← aire interno a la derecha */
    padding-left: 0.8rem !important;
    padding-top: 0.6rem !important;
}}

/* KPIs en una sola línea */
[data-testid="stHorizontalBlock"] {{
    flex-wrap: nowrap !important;
    gap: 1.2rem !important;
}}
.kpi {{ padding: 10px 14px !important; }}
.kpi .val {{ font-size: 21px !important; line-height: 1.1 !important; }}
.kpi .lbl {{ font-size: 10px !important; }}

/* ── Tarjetas Plotly ── */
.stPlotlyChart {{
    padding: 6px 8px !important;
    margin: 0 !important;
    border-radius: 12px;
}}

/* ── Sidebar fondo ── */
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {GUINDA} 0%, {GUINDA_DARK} 100%);
    border-right: 1px solid rgba(0,0,0,0.06) !important;
}}

/* Regla 2: ancho */
[data-testid="stSidebar"] {{
    min-width: 300px !important;
    max-width: 280px !important;
}}
[data-testid="stSidebar"] > div:first-child {{
    width: 280px !important;
    padding: 0.15rem 1.2rem 0.6rem 1.2rem !important;
}}

/* ── Texto general del sidebar (sin tocar widgets) ── */
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4,
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] .stMarkdown * {{ color: #FFFFFF !important; }}

/* ── Etiquetas de widgets ── */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {{
    color: #F3D7DE !important;
    font-weight: 600 !important;
    font-size: 12.5px !important;
}}

/* ── Selectbox: fondo blanco, texto oscuro ── */
[data-testid="stSidebar"] div[data-baseweb="select"] > div:first-child {{
    background: #FFFFFF !important;
    border-color: rgba(255,255,255,0.55) !important;
}}
[data-testid="stSidebar"] div[data-baseweb="select"] span,
[data-testid="stSidebar"] div[data-baseweb="select"] input,
[data-testid="stSidebar"] div[data-baseweb="select"] div[title] {{
    color: {NAVY} !important;
    -webkit-text-fill-color: {NAVY} !important;
    opacity: 1 !important;
}}
[data-testid="stSidebar"] div[data-baseweb="select"] svg {{
    fill: {GUINDA} !important; color: {GUINDA} !important;
}}

/* ── Chips multiselect ── */
[data-testid="stSidebar"] span[data-baseweb="tag"] {{ background: {GUINDA} !important; }}
[data-testid="stSidebar"] span[data-baseweb="tag"] * {{
    color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important;
}}

/* ── Popup de multiselect / selectbox (se renderiza fuera del sidebar) ── */
ul[data-baseweb="menu"],
div[data-baseweb="popover"] ul,
div[data-baseweb="popover"] [role="listbox"] {{
    background: #FFFFFF !important;
    border: 1px solid rgba(0,0,0,.08) !important;
    border-radius: 10px !important;
    box-shadow: 0 8px 24px rgba(0,0,0,.18) !important;
    padding: 4px !important;
}}

div[data-baseweb="popover"] li,
div[data-baseweb="popover"] [role="option"] {{
    background: #FFFFFF !important;
    color: {NAVY} !important;
    -webkit-text-fill-color: {NAVY} !important;
    border-radius: 6px !important;
    padding: 6px 10px !important;
    font-size: 13px !important;
}}
div[data-baseweb="popover"] li *,
div[data-baseweb="popover"] [role="option"] * {{
    color: {NAVY} !important;
    -webkit-text-fill-color: {NAVY} !important;
}}

div[data-baseweb="popover"] li:hover,
div[data-baseweb="popover"] [role="option"]:hover,
div[data-baseweb="popover"] [aria-selected="true"] {{
    background: #F3D7DE !important;
}}
div[data-baseweb="popover"] li:hover *,
div[data-baseweb="popover"] [aria-selected="true"] * {{
    color: {GUINDA} !important;
    -webkit-text-fill-color: {GUINDA} !important;
}}

/* Casillas / checkboxes del multiselect */
div[data-baseweb="popover"] svg {{
    fill: {GUINDA} !important;
    color: {GUINDA} !important;
}}
div[data-baseweb="popover"] svg[fill="none"],
div[data-baseweb="popover"] svg path {{
    stroke: {GUINDA} !important;
}}

/* ── Radio de navegación ── */
[data-testid="stSidebar"] .stRadio > div {{ gap: 6px; }}
[data-testid="stSidebar"] .stRadio label {{
    background: rgba(255,255,255,0.06);
    border-radius: 8px; padding: 7px 10px;
    color: #FFFFFF !important;
    font-weight: 500;
}}
[data-testid="stSidebar"] .stRadio label:hover {{ background: rgba(255,255,255,0.18); }}

/* ── Hero ── */
.hero {{
    background: linear-gradient(90deg, {GUINDA} 0%, #A1243F 55%, #D0466B 100%);
    color: #fff; padding: 14px 22px; border-radius: 12px;
    box-shadow: 0 8px 22px rgba(116,25,45,.22);
    margin-bottom: 12px;
}}
.hero h2 {{ margin: 0; font-weight: 700; letter-spacing: .3px; font-size: 22px; color: #fff; }}
.hero .sub {{ opacity: .85; font-size: 12.5px; margin-top: 2px; color: #fff; }}

/* ── KPI ── */
.kpi {{
    background: #fff; border-left: 5px solid {BLUE};
    border-radius: 10px; padding: 12px 16px;
    box-shadow: 0 2px 10px rgba(0,0,0,.05);
    height: 100%;
}}
.kpi .lbl {{ color: {GUINDA_DARK}; font-size: 11px; font-weight: 700; letter-spacing: .6px; text-transform: uppercase; }}
.kpi .val {{ color: {NAVY}; font-size: 25px; font-weight: 700; line-height: 1.15; margin-top: 2px; }}
.kpi .dlt {{ font-size: 12px; font-weight: 600; margin-top: 2px; }}
.kpi .up {{ color: {TEAL}; }}
.kpi .dn {{ color: {TEAL}; }}
.kpi .note {{ font-size: 11px; color: {GUINDA_DARK}; margin-top: 3px; }}

.stPlotlyChart {{ background: #fff; border-radius: 12px; padding: 6px;
                  box-shadow: 0 2px 10px rgba(0,0,0,.04); }}
div[data-testid="stDataFrame"] {{ background: #fff; border-radius: 10px; }}
</style>
"""
_CSS_ESCALA = f"""
<style>
html {{ font-size: {16 * K:.2f}px !important; }}   /* los widgets de Streamlit usan rem */
/* oculta el iframe invisible que mide el ancho de la ventana */
.element-container:has(iframe[title*="viewport"]) {{
    position: absolute !important; height: 0 !important; width: 0 !important;
    overflow: hidden !important; margin: 0 !important; padding: 0 !important;
}}
</style>
"""
st.markdown(_escalar_px(_CSS) + _CSS_ESCALA, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════
# CATÁLOGO GEOGRÁFICO
# ═════════════════════════════════════════════════════════════════════════
dim = D["TABLAS_DIM"]
geo = (dim[["cve_issste", "Entidad", "Latitud", "Longitud"]]
       .dropna(subset=["cve_issste"])
       .drop_duplicates("cve_issste"))
names = geo.set_index("cve_issste").Entidad.to_dict()

# ═════════════════════════════════════════════════════════════════════════
# GEOJSON — carga + autodetección de clave de entidad
# ═════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _load_geojson_meta():
    if not GEOJSON_PATH.exists():
        return None, None, 0
    try:
        gj = json.loads(GEOJSON_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None, None, 0

    feats = gj.get("features", [])
    if not feats:
        return gj, None, 0

    sample = feats[0]

    top_id = sample.get("id")
    if top_id is not None:
        v = str(top_id).strip()
        if v.isdigit():
            pad = len(v) if v.startswith("0") else 0
            return gj, "id", pad

    props = sample.get("properties", {}) or {}
    candidates = [
        "CVEGEO", "cve_geo", "CVE_GEO", "cvegeo",
        "cve_ent", "CVE_ENT", "CV_ENT", "CVENT",
        "cve_issste", "CVE_ISSSTE", "cve",
        "id", "ID", "code", "CODIGO", "codigo",
        "state_code", "STATE_CODE", "ent",
    ]
    for key in candidates:
        if key in props:
            v = str(props[key]).strip()
            if v.isdigit():
                pad = len(v) if v.startswith("0") else 0
                return gj, f"properties.{key}", pad

    return gj, None, 0

def _normalizar_codigo(v, pad):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if not s.isdigit():
        return s
    if pad > 0:
        return s.zfill(pad)
    return str(int(s))

# ═════════════════════════════════════════════════════════════════════════
# UTILIDADES DE GRÁFICOS
# ═════════════════════════════════════════════════════════════════════════
PLOTLY_CFG = {
    "displaylogo": False,
    "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"],
}

_CLAVES_FUENTE = {"font", "tickfont", "titlefont", "title_font", "textfont",
                  "insidetextfont", "outsidetextfont"}

def _escalar_fuentes(obj):
    """Recorre el JSON de la figura y multiplica por K cada font.size ya definido."""
    if isinstance(obj, dict):
        for k_, v in obj.items():
            if k_ in _CLAVES_FUENTE and isinstance(v, dict) and isinstance(v.get("size"), (int, float)):
                v["size"] = round(v["size"] * K, 1)
            _escalar_fuentes(v)
    elif isinstance(obj, list):
        for it in obj:
            _escalar_fuentes(it)

def theme(fig, h=340, title=None):
    # 1) lo que cada gráfico ya traía (anotaciones, colorbar, textos de barras…)
    d = fig.to_plotly_json()
    _escalar_fuentes(d)
    fig = go.Figure(d)
    leyenda = dict(
        orientation="h", yanchor="bottom", y=-0.22,
        xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)",
        font=dict(size=S(11)),
    )
    leyenda.update(fig.layout.legend.to_plotly_json())
    margen_superior = fig.layout.margin.t
    if margen_superior is None:
        margen_superior = S(48)
    # 2) estilo base, todo multiplicado por K
    fig.update_layout(
        height=S(h), margin=dict(l=S(12), r=S(12),
                     t=margen_superior, b=S(8)),
        colorway=C, paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Segoe UI, Arial", color=NAVY, size=S(12)),
        title=dict(text=title if title is not None else fig.layout.title.text,
                   font=dict(size=S(14), color=NAVY), x=0.02, xanchor="left"),
        legend=leyenda,
        xaxis=dict(showgrid=False, linecolor=GRID, ticks="outside",
                   tickfont=dict(size=S(11)), zeroline=False),
        yaxis=dict(gridcolor=GRID, linecolor="rgba(0,0,0,0)",
                   tickfont=dict(size=S(11)), zeroline=False),
        hoverlabel=dict(bgcolor="white", font_size=S(12),
                        font_family="Segoe UI", bordercolor=GUINDA),
    )
    return fig

def show(fig, h=340, title=None):
    st.plotly_chart(theme(fig, h, title), use_container_width=True, config=PLOTLY_CFG)

def kpi(col, label, value, delta=None, note=None):
    dlt_html = ""
    if delta is not None:
        cls, arrow = ("up", "▲") if delta >= 0 else ("dn", "▼")
        dlt_html = f'<div class="dlt {cls}">{arrow} {abs(delta):.2f}% vs año previo</div>'
    note_html = f'<div class="note">{note}</div>' if note else ""
    html = (
        '<div class="kpi">'
        f'<div class="lbl">{label}</div>'
        f'<div class="val">{value}</div>'
        f'{dlt_html}'
        f'{note_html}'
        '</div>'
    )
    col.markdown(html, unsafe_allow_html=True)

def head(title, subtitle=""):
    sub = f'<div class="sub">{subtitle}</div>' if subtitle else ""
    st.markdown(f'<div class="hero"><h2>{title}</h2>{sub}</div>', unsafe_allow_html=True)

def fmt(n):
    try: return f"{float(n):,.0f}"
    except Exception: return "—"

def money(n):
    try: return f"${float(n):,.0f}"
    except Exception: return "—"

# ═════════════════════════════════════════════════════════════════════════
# FILTROS LATERALES
# ═════════════════════════════════════════════════════════════════════════
def ctl(b: pd.DataFrame, age: bool = True):
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        _escalar_px("<div style='font-weight:700;font-size:13px;letter-spacing:.4px;'>🔎 Filtros</div>"),
        unsafe_allow_html=True)

    y = None
    if "ANIO" in b.columns:
        y = st.sidebar.selectbox("Año", sorted(b.ANIO.dropna().unique(), reverse=True))

    m = None
    if y is not None and "MES" in b.columns:
        ms = sorted(b.loc[b.ANIO.eq(y), "MES"].dropna().unique())
        if ms:
            m = st.sidebar.selectbox("Mes de corte", ms, index=len(ms) - 1,
                                     format_func=lambda x: f"{int(x):02d}")

        # Entidad: selección única con opción "Todas"
    e = []
    if "ENT_CVE" in b.columns:
        cs = sorted(b.ENT_CVE.dropna().unique())
        opts = ["__TODAS__"] + cs
        sel = st.sidebar.selectbox(
        "Entidad", opts,
        format_func=lambda x: "Todas las entidades" if x == "__TODAS__" else names.get(x, x),
        )
        e = cs if sel == "__TODAS__" else [sel]


    s = st.sidebar.pills(
    "Sexo", ["H", "M"], selection_mode="multi", default=["H", "M"],
    format_func=lambda x: {"H": "Hombres", "M": "Mujeres"}[x],
    ) or ["H", "M"]

    a = []
    if age and "GRUPO_EDAD" in b.columns:
        aa = sorted(b.GRUPO_EDAD.dropna().unique())
        opts = ["__TODOS__"] + list(aa)
        sel = st.sidebar.selectbox(
            "Grupo de edad",
            opts,
            format_func=lambda x: "Todos los grupos"
                                  if x == "__TODOS__" else str(x),
            key="filtro_grupo_edad",
        )
        a = aa if sel == "__TODOS__" else [sel]

    return y, m, e, s, a

def _cve_str(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(2) if s.isdigit() else s


def _tiene_datos(x: pd.DataFrame, col: str) -> bool:
    """True si la columna existe y tiene al menos un valor no nulo/no vacío."""
    if col not in x.columns:
        return False
    serie = x[col]
    if serie.dropna().empty:
        return False
    # También descarta cadenas vacías
    if serie.dtype == object:
        validos = serie.dropna().astype(str).str.strip().replace("", pd.NA).dropna()
        return not validos.empty
    return True


def fil(d, y, m, e, s, a=None):
    x = d.copy()

    # Año: solo si la columna existe y tiene datos
    if y is not None and _tiene_datos(x, "ANIO"):
        x = x[x.ANIO == y]

    # Mes: solo si la columna existe y tiene datos
    if m is not None and _tiene_datos(x, "MES"):
        x = x[x.MES == m]

    # Entidad: SOLO si la columna tiene datos utilizables
    if e and _tiene_datos(x, "ENT_CVE"):
        e_norm = {_cve_str(v) for v in e}
        x = x[x.ENT_CVE.apply(_cve_str).isin(e_norm)]

    # Sexo: solo si la columna tiene datos
    if s and _tiene_datos(x, "SEXO"):
        x = x[x.SEXO.isin(s)]

    # Grupo de edad: solo si hay intersección real
    if a and _tiene_datos(x, "GRUPO_EDAD"):
        disp = set(x.GRUPO_EDAD.dropna().unique())
        efec = [g for g in a if g in disp]
        if efec:
            x = x[x.GRUPO_EDAD.isin(efec)]

    return x

def tot_o_ultimo(d, y, m, e, s, a=None):
    x = fil(d, y, m, e, s, a)
    if not x.empty and "TOTAL" in x.columns:
        return x.TOTAL.sum(), y

    # Fallback: probar el último año disponible de la hoja
    if _tiene_datos(d, "ANIO") and not d.empty:
        ly = int(d.ANIO.max())
        x2 = fil(d, ly, None, e, s, a)
        if not x2.empty and "TOTAL" in x2.columns:
            return x2.TOTAL.sum(), ly

    # Último recurso: sumar todo lo que haya, sin filtros temporales
    if "TOTAL" in d.columns:
        return d.TOTAL.sum(), int(d.ANIO.max()) if _tiene_datos(d, "ANIO") else y

    return 0, y

def tr(d, e, s, a=None):
    x = d.copy()
    if e and "ENT_CVE" in x.columns:    x = x[x.ENT_CVE.isin(e)]
    if s and "SEXO" in x.columns:       x = x[x.SEXO.isin(s)]
    if a and "GRUPO_EDAD" in x.columns: x = x[x.GRUPO_EDAD.isin(a)]
    if "ANIO" not in x.columns or "MES" not in x.columns:
        return pd.DataFrame(columns=["Periodo", "TOTAL"])
    if "TOTAL" not in x.columns:
        return pd.DataFrame(columns=["Periodo", "TOTAL"])
    x["Periodo"] = pd.to_datetime(dict(year=x.ANIO, month=x.MES, day=1))
    return x.groupby("Periodo", as_index=False).TOTAL.sum()

# ─── Columna canónica de plazas ─────────────────────────────────────────
PLAZAS_CANDIDATAS = [
    "NPLAZAS", "N_PLAZAS", "NPLAZA", "NUM_PLAZAS", "NUMPLAZAS",
    "PLAZAS", "PLAZA", "TOT_PLAZAS", "TOTAL_PLAZAS",
    "NPLAZAS_TOT", "PLAZAS_TOT", "PROM_PLAZAS","TU_NOMBRE_AQUI", 
]

def _mult_plazas(rango) -> int:
    """Extrae el multiplicador de plazas desde RANGO_PLAZAS.
    '1 plaza' → 1 · '2 plaza' → 2 · 'Más de 3 plazas' → 4."""
    s = str(rango).strip().lower()
    m = re.search(r"(\d+)", s)
    n = int(m.group(1)) if m else 1
    if "más de" in s or "mas de" in s or ">" in s:
        return n + 1
    return n


def _col_plazas(x: pd.DataFrame):
    """Devuelve una tupla (columna, factor):
    - Si existe una columna tipo NPLAZAS → (nombre, 1)
    - Si existe RANGO_PLAZAS + TOTAL → ('TOTAL', None) y hay que multiplicar
    - Si nada → (None, None)"""
    for c in PLAZAS_CANDIDATAS:
        if c in x.columns and pd.api.types.is_numeric_dtype(x[c]):
            return c, 1

    if "RANGO_PLAZAS" in x.columns and "TOTAL" in x.columns:
        return "TOTAL", None   # señal: multiplicar por _mult_plazas

    return None, None


def _plazas_de(x: pd.DataFrame, col, factor):
    """Serie de plazas agregada por columna de agrupación o total."""
    if col is None:
        return None
    if factor is None:      # hay que multiplicar por el multiplicador del rango
        x = x.copy()
        x["_n"] = x["RANGO_PLAZAS"].apply(_mult_plazas)
        return x["TOTAL"] * x["_n"]
    return x[col]

def tr_plazas(df, e, s):
    """Serie temporal del número de plazas usando ISSSTE_PLAZ_NOMBRA_PROM."""
    x = df.copy()
    if e and "ENT_CVE" in x.columns:  x = x[x.ENT_CVE.isin(e)]
    if s and "SEXO" in x.columns:     x = x[x.SEXO.isin(s)]
    if "ANIO" not in x.columns or "MES" not in x.columns:
        return pd.DataFrame(columns=["Periodo", "TOTAL"])

    x["Periodo"] = pd.to_datetime(dict(year=x.ANIO, month=x.MES, day=1))
    col, factor = _col_plazas(x)

    if col is None:
        return pd.DataFrame(columns=["Periodo", "TOTAL"])

    x["_plazas"] = _plazas_de(x, col, factor)
    return (x.groupby("Periodo", as_index=False)["_plazas"].sum()
              .rename(columns={"_plazas": "TOTAL"}))

def avg(x, g, v):
    columnas_grupo = [g] if isinstance(g, str) else g
    if x.empty or "TOTAL" not in x.columns or v not in x.columns:
        return pd.DataFrame(columns=columnas_grupo + ["Promedio"])
    z = x.copy()
    z["w"] = z.TOTAL * z[v]
    return (z.groupby(g)
              .agg(w=("w", "sum"), TOTAL=("TOTAL", "sum"))
              .assign(Promedio=lambda q: q.w / q.TOTAL)
              .reset_index())

def wavg(x, v):
    if x.empty or v not in x.columns or "TOTAL" not in x.columns: return None
    t = x.TOTAL.sum()
    return (x.TOTAL * x[v]).sum() / t if t else None

def delta_aoa(b, y, m, e, s, a=None):
    if y is None or "ANIO" not in b.columns or "TOTAL" not in b.columns:
        return None
    cur = fil(b, y, m, e, s, a).TOTAL.sum()
    prev = fil(b, y - 1, m, e, s, a).TOTAL.sum()
    return ((cur - prev) / prev * 100) if prev else None

def exportar(x, nombre):
    with st.expander("📋 Ver y descargar datos filtrados"):
        st.dataframe(x, use_container_width=True, height=int(S(300)))
        st.download_button("⬇️ Descargar CSV",
                           x.to_csv(index=False).encode("utf-8"),
                           file_name=f"{nombre}.csv", mime="text/csv")

# ═════════════════════════════════════════════════════════════════════════
# MAPA COROPLÉTICO DE TASA (SM)
# ═════════════════════════════════════════════════════════════════════════
def _mapa_tasa_sm(y, m, e, s, selected_rango=None):
    sm = D["ISSSTE_TRAB_SM"]
    x = fil(sm, y, m, e, s)

    if x.empty:
        return px.scatter(title="Sin datos de salario mínimo para el periodo seleccionado")
    if "Tasa" not in x.columns or "TOTAL" not in x.columns:
        return px.scatter(title="Faltan columnas 'Tasa' o 'TOTAL' en ISSSTE_TRAB_SM")

    g = (x.assign(w=x.TOTAL * x.Tasa)
           .groupby("ENT_CVE", as_index=False)
           .agg(w=("w", "sum"), TOTAL=("TOTAL", "sum"))
           .assign(Tasa=lambda q: q.w / q.TOTAL))

    gj, feat_key, pad = _load_geojson_meta()

    AZUL_ROJO = [
        [0.00, "#0B3D91"],
        [0.25, "#1F6FB2"],
        [0.50, "#9AA6B2"],
        [0.75, "#C75B4A"],
        [1.00, "#8B1A1A"],
    ]

    if gj is None or feat_key is None:
        g["ENT_CVE"] = g["ENT_CVE"].apply(lambda v: _normalizar_codigo(v, 2))
        q = (g.merge(geo, left_on="ENT_CVE", right_on="cve_issste", how="left")
               .dropna(subset=["Latitud"]))
        f = px.scatter_geo(
            q, lat="Latitud", lon="Longitud", size="TOTAL",
            color="Tasa", hover_name="Entidad",
            color_continuous_scale=AZUL_ROJO,
            title="Tasa de distribución por entidad federativa",
        )
        f.update_geos(lataxis_range=[14, 33], lonaxis_range=[-118, -86],
                      showland=True, landcolor="#F1F3F6",
                      showocean=True, oceancolor="#DCEAF5",
                      showcountries=True, countrycolor="#C7CCD3",
                      showcoastlines=True, coastlinecolor="#B8BEC5")
        f.update_layout(coloraxis_colorbar=dict(
            title="Tasa", thickness=12, len=0.7,
            tickfont=dict(size=10), title_font=dict(size=11)))
        return f

    g["ENT_CVE"] = g["ENT_CVE"].apply(lambda v: _normalizar_codigo(v, pad))
    g["Entidad"] = g["ENT_CVE"].map(
        lambda c: names.get(int(c) if c and c.isdigit() else c, names.get(c, c)))

    f = px.choropleth(
        g,
        geojson=gj,
        locations="ENT_CVE",
        featureidkey=feat_key,
        color="Tasa",
        color_continuous_scale=AZUL_ROJO,
        hover_name="Entidad",
        hover_data={"Tasa": ":.3f", "TOTAL": ":,.0f",
                    "ENT_CVE": False, "Entidad": False},
        title="Tasa por entidad federativa",
        labels={"Tasa": "Tasa"},
    )
    f.update_traces(marker_line_color="white", marker_line_width=0.6)
    f.update_geos(fitbounds="locations", visible=False, bgcolor="white")
    f.update_layout(
        coloraxis_colorbar=dict(
            title="Tasa", thickness=12, len=0.7,
            tickfont=dict(size=10), title_font=dict(size=11),
        ),
        margin=dict(l=0, r=0, t=48, b=0),
    )
    return f
#Función para graficar la pirámide poblacional
# ── Mapeo: grupo original -> grupo nuevo ──────────────────────────

def _agrupa_edades(df, col_edad="Grupos_Edad"):
    """
    Recibe el df con la columna original 'Grupos_Edad' y devuelve
    una copia con la columna 'Edad_Agrup' ya recodificada.
    """
    d = df.copy()
    d["Edad_Agrup"] = d[col_edad].map(_MAPA_EDAD)

    # Los grupos que no estén en el mapa se mantienen igual (por seguridad)
    d["Edad_Agrup"] = d["Edad_Agrup"].fillna(d[col_edad])

    return d

def _piramide_issste(df, col_edad="Grupos_Edad", col_sexo="Sexo",
                     col_valor="TOTAL", titulo="Pirámide poblacional ISSSTE"):
    if df is None or df.empty:
        return None

    # 1) Recodificar edades
    d = _agrupa_edades(df, col_edad=col_edad)

    # 2) Normalizar sexo a H / M
    mapa_sexo = {
        "H": "H", "Hombre": "H", "Hombres": "H", "MASCULINO": "H",
        "M": "M", "Mujer": "M",  "Mujeres": "M", "FEMENINO": "M",
    }
    d["_sexo"] = d[col_sexo].map(mapa_sexo).fillna(d[col_sexo])

    # 3) Agrupar por el NUEVO grupo de edad
    piv = (d.groupby(["Edad_Agrup", "_sexo"], as_index=False)[col_valor]
             .sum()
             .pivot(index="Edad_Agrup", columns="_sexo", values=col_valor)
             .fillna(0))

    # 4) Respetar el orden deseado
    piv = piv.reindex([g for g in _ORDEN_EDAD if g in piv.index])

    hombres = -piv.get("H", 0)
    mujeres =  piv.get("M", 0)

    AZUL = "#3b6fb6"
    ROSA = "#e37b8a"

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=piv.index, x=hombres, orientation="h",
        name="Hombres",
        marker=dict(color=AZUL, line=dict(width=0)),
        customdata=piv.get("H", 0),
        hovertemplate="<b>Hombres</b><br>%{y}<br>%{customdata:,.0f}"
                      "<extra></extra>",
    ))

    fig.add_trace(go.Bar(
        y=piv.index, x=mujeres, orientation="h",
        name="Mujeres",
        marker=dict(color=ROSA, line=dict(width=0)),
        hovertemplate="<b>Mujeres</b><br>%{y}<br>%{x:,.0f}<extra></extra>",
    ))

    # 5) Eje X simétrico
    xmax = max(abs(hombres).max(), abs(mujeres).max(), 1)
    paso = 3_000_000
    limite_ticks = int((xmax + paso - 1) // paso) * paso
    ticks = list(range(0, limite_ticks + 1, paso))
    etiquetas_ticks = [
        f"{valor / 1_000_000:g}M" if valor else "0"
        for valor in ticks
    ]

    fig.update_layout(
        title=titulo,
        barmode="overlay",
        bargap=0.08,
        height=460,
        margin=dict(l=10, r=10, t=60, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(
            tickvals=[-t for t in ticks[::-1]] + ticks[1:],
            ticktext=etiquetas_ticks[::-1] + etiquetas_ticks[1:],
            zeroline=True, zerolinecolor="#888", zerolinewidth=1,
        ),
        yaxis=dict(tickfont=dict(size=11)),
        plot_bgcolor="white",
    )

    fig.add_annotation(
        text="", xref="paper", yref="paper",
        x=0.5, y=1.10, showarrow=False, font=dict(size=12, color="#444"),
    )
    fig.add_annotation(
        text="♂", xref="paper", yref="paper", x=0.18, y=1.10,
        showarrow=False, font=dict(size=22, color=AZUL),
    )
    fig.add_annotation(
        text="♀", xref="paper", yref="paper", x=0.82, y=1.10,
        showarrow=False, font=dict(size=22, color=ROSA),
    )

    return fig

_MAPA_EDAD = {
    "0-4": "0-19", "5-9": "0-19", "10-14": "0-19", "15-19": "0-19",
    "20-24": "20-29", "25-29": "20-29",
    "30-34": "30-39", "35-39": "30-39",
    "40-44": "40-49", "45-49": "40-49",
    "50-54": "50-59", "55-59": "50-59",
    "60-64": "60-69", "65-69": "60-69",
    "70-74": "70-79", "75-79": "70-79",
    "80-84": "80 y más", "85 y más": "80 y más",
}

_ORDEN_EDAD = ["0-19", "20-29", "30-39", "40-49",
               "50-59", "60-69", "70-79", "80 y más"]
# ═════════════════════════════════════════════════════════════════════════
# PÁGINAS
# ═════════════════════════════════════════════════════════════════════════
def page_issste():
    b = D["ISSSTE_TRAB_TOT2015_2025"]
    y, m, e, s, a = ctl(b)
    head("ISSSTE · Resumen",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    # ── KPIs ─────────────────────────────────────────────────────────────
    tot_trab, y_trab = tot_o_ultimo(b, y, m, e, s, a)

    # Plazas: usar ISSSTE_PLAZ_NOMBRA_PROM multiplicando por el rango
    _serie_plaz = tr_plazas(D["ISSSTE_PLAZ_NOMBRA_PROM"], e, s)
    if not _serie_plaz.empty and y is not None:
        _p = _serie_plaz[_serie_plaz.Periodo.dt.year == y]
        if not _p.empty:
            tot_plaz = _p.TOTAL.sum()
            y_plaz   = y
        else:
            tot_plaz = _serie_plaz.TOTAL.iloc[-1]
            y_plaz   = int(_serie_plaz.Periodo.dt.year.max())
    else:
        tot_plaz, y_plaz = 0, y

    tot_pen, y_pen = tot_o_ultimo(D["ISSSTE_PENS_REGPEN"],  y, m, e, s, a)
    tot_deu, y_deu = tot_o_ultimo(D["ISSSTE_DEUD_TOT"],      y, m, e, s, a)
    tot_fam, y_fam = tot_o_ultimo(D["ISSSTE_FAMILIAR_PARENT"], y, m, e, s, a)

    c1, c2, c3, c4, c5 = st.columns(5,gap="small")
    kpi(c1, "Trabajadores", fmt(tot_trab), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Plazas",       fmt(tot_plaz),
        note=(f"año {y_plaz}" if y_plaz != y else None))
    kpi(c3, "Pensionados",  fmt(tot_pen),
        note=(f"año {y_pen}" if y_pen != y else None))
    kpi(c4, "Deudos",       fmt(tot_deu),
        note=(f"año {y_deu}" if y_deu != y else None))
    kpi(c5, "Familiares",   fmt(tot_fam),
        note=(f"año {y_fam}" if y_fam != y else None))

    st.markdown("")

    # ── Gráfico combinado + mapa ─────────────────────────────────────────
    l, r = st.columns([3, 2], gap="small")
    with l:
        t_trab = tr(b, e, s, a)                                # 2015–2025
        t_plaz = tr_plazas(D["ISSSTE_PLAZ_NOMBRA_PROM"], e, s) # 2021–2025

        fig = go.Figure()

        if not t_trab.empty:
            fig.add_trace(go.Scatter(
                x=t_trab["Periodo"], y=t_trab["TOTAL"],
                mode="lines+markers",
                connectgaps=False,
                name="Trabajadores",
                line=dict(color=GUINDA, width=2.6),
                marker=dict(size=6, color=GUINDA),
                hovertemplate="<b>Trabajadores</b><br>%{x|%b %Y}<br>"
                              "%{y:,.0f}<extra></extra>",
            ))

        if not t_plaz.empty:
            fig.add_trace(go.Scatter(
                x=t_plaz["Periodo"], y=t_plaz["TOTAL"],
                mode="lines+markers",
                connectgaps=False,
                name="Plazas",
                line=dict(color=NAVY, width=2.6, dash="dot"),
                marker=dict(size=6, color=NAVY,symbol="diamond"),
                hovertemplate="<b>Plazas</b><br>%{x|%b %Y}<br>"
                              "%{y:,.0f}<extra></extra>",
            ))

        if not fig.data:
            st.info("Sin datos de evolución para el filtro seleccionado.")
        else:
            periodos = [
                serie["Periodo"]
                for serie in (t_trab, t_plaz)
                if not serie.empty
            ]
            primer_corte = min(periodo.min() for periodo in periodos)
            ultimo_corte = max(periodo.max() for periodo in periodos)
            fig.update_layout(
                title="Evolución de trabajadores y plazas (2015–2025)",
                xaxis_title=None,
                xaxis_range=[
                    primer_corte - pd.DateOffset(months=1),
                    ultimo_corte,
                ],
                xaxis_tick0=primer_corte,
                xaxis_dtick="M12",
                xaxis_tickformat="%Y",
                yaxis_title="Personas / Plazas (Millones)",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=-0.25,
                            xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)"),
            )
            show(fig, 420)
    with r:
        fig = _mapa_tasa_sm(y, m, e, s)
        show(fig, 420)

        # ── Pirámide poblacional + Dona rango salarial + Barras nombramiento ──
    c1, c2, c3 = st.columns([1.6, 1, 1.2], gap="small")

    # ─── Pirámide poblacional ────────────────────────────────────────────
    with c1:
        df_pir = D.get("ISSSTE_TRAB_TOT2015_2025")
        if df_pir is not None and not df_pir.empty:
            df_pir_f = df_pir.copy()
            if "Año" in df_pir_f.columns and y is not None:
                df_pir_f = df_pir_f[df_pir_f["Año"] == y]
            if "Entidad" in df_pir_f.columns and e is not None:
                df_pir_f = df_pir_f[df_pir_f["Entidad"] == e]

            fig_pir = _piramide_issste(
                df_pir_f,
                col_edad="GRUPO_EDAD",     # <── nombre real de tu columna
                col_sexo="SEXO",            # <── ajusta
                col_valor="TOTAL",          # <── ajusta
                titulo=f"Pirámide poblacional de trabajadores ({y})",
            )
            if fig_pir is not None:
                show(fig_pir, 460)
            else:
                st.info("Sin datos para la pirámide.")
        else:
            st.info("Sin datos ISSSTE_TRAB_TOT2015_2025.")

    # ─── Dona de rango salarial ──────────────────────────────────────────
    with c2:
        q = (fil(D["ISSSTE_TRAB_SM"], y, m, e, s)
             .groupby("RANGO_SM", as_index=False).TOTAL.sum())
        if not q.empty:
            fig = px.pie(q, names="RANGO_SM", values="TOTAL", hole=.55,
                         title="Trabajadores por rango salarial")
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 460)      # ← iguala altura con la pirámide

    # ─── Barras de plazas por nombramiento ───────────────────────────────
    with c3:
        _x = fil(D["ISSSTE_PLAZ_NOMBRA_PROM"], y, m, e, s)
        col, factor = _col_plazas(_x)
        if col and "NOMBRAMIENO_DES" in _x.columns:
            _x = _x.copy()
            _x["_plazas"] = _plazas_de(_x, col, factor)
            q = (_x.groupby("NOMBRAMIENO_DES", as_index=False)["_plazas"].sum()
                    .rename(columns={"_plazas": "TOTAL"})
                    .sort_values("TOTAL").tail(10))
        else:
            q = pd.DataFrame()
        if not q.empty:
            paso_eje = 3_000_000
            limite_ticks = max(
                paso_eje,
                ((int(q["TOTAL"].max()) + paso_eje - 1) // paso_eje) * paso_eje,
            )
            ticks_eje = list(range(0, limite_ticks + 1, paso_eje))
            fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", orientation="h",
                         title="Top 10 plazas por nombramiento")
            
            umbral = q["TOTAL"].max() * 0.10
            posiciones = [
                        "outside" if total <= umbral else "inside"
                        for total in q["TOTAL"]
                        ]
            fig.update_traces(marker_color=BLUE, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                        NAVY if posicion == "outside" else "white"
                                        for posicion in posiciones
                                        ],
                            cliponaxis=False,)            
  
            fig.update_layout(
                yaxis_title="Tipos de Plazas",
                xaxis_title="Trabajadores",
                xaxis=dict(
                    range=[0, limite_ticks * 1.15],
                    tickmode="array",
                    tickvals=ticks_eje,
                    ticktext=[
                        f"{tick / 1_000_000:g}M" if tick else "0"
                        for tick in ticks_eje
                    ],
                ),
            )
            show(fig, 460)      # ← misma altura

def page_trabajadores():
    b = D["ISSSTE_TRAB_TOT2015_2025"]
    y, m, e, s, a = ctl(b)
    x = fil(b, y, m, e, s, a)
    head("Trabajadores afiliados al ISSSTE",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Total trabajadores", fmt(x.TOTAL.sum()), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Entidades",          fmt(len(e)))
    kpi(c3, "Hombres", fmt(x.loc[x.SEXO == "H", "TOTAL"].sum()) if "SEXO" in x else "—")
    kpi(c4, "Mujeres", fmt(x.loc[x.SEXO == "M", "TOTAL"].sum()) if "SEXO" in x else "—")

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s, a)
        if not t.empty:
            t["TOTAL_MILLONES"] = t["TOTAL"] / 1_000_000
            primer_corte = t["Periodo"].min()
            ultimo_corte = t["Periodo"].max()
            fig = px.line(t, x="Periodo", y="TOTAL", text="TOTAL_MILLONES")
            fig.update_traces(
                mode="lines+markers+text",
                texttemplate="%{text:.2f}",
                textposition="top center",
                line=dict(width=2.5, color=GUINDA),
                hovertemplate="<b>Trabajadores</b><br>%{x|%b %Y}<br>"
                                              "%{y:,.0f}<extra></extra>"
            )
            fig.update_layout(
                            title="Evolución de trabajadores (2015–2025)",
                            xaxis_title="Periodo",
                            xaxis_range=[
                                primer_corte - pd.DateOffset(months=1),
                                ultimo_corte,
                            ],
                            xaxis_tick0=primer_corte,
                            xaxis_dtick="M12",
                            xaxis_tickformat="%Y",
                            yaxis_title="Trabajadores (Millones)",
                            hovermode="x unified",
                            legend=dict(orientation="h", yanchor="bottom", y=-0.25,
                                        xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)"),
                        )
            show(fig, 380)  

    with r:
        fig = _mapa_tasa_sm(y, m, e, s)
        show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            edades = _agrupa_edades(x, col_edad="GRUPO_EDAD")
            edades["SEXO"] = edades["SEXO"].replace(
                {"H": "Hombre", "M": "Mujer"}
            )
            ag = (edades.groupby(["Edad_Agrup", "SEXO"], as_index=False)
                         .TOTAL.sum())
            ag = ag[ag["Edad_Agrup"].isin(_ORDEN_EDAD)]
            fig = px.bar(ag, x="Edad_Agrup", y="TOTAL", 
                         color="SEXO",
                         barmode="group",
                         title="Trabajadores por edad y sexo",
                         category_orders={
                             "Edad_Agrup": _ORDEN_EDAD,
                             "SEXO": ["Hombre", "Mujer"],
                         },
                         labels={"Edad_Agrup": "Grupo de edad",
                                 "TOTAL": "Trabajadores"},
                         color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            
            umbral = ag["TOTAL"].max() * 0.05
            for traza in fig.data:
                posiciones = [
                    "outside" if valor <= umbral else "inside"
                    for valor in traza.y
                ]

                traza.texttemplate = "%{y:,.0f}"
                traza.textposition = posiciones
                traza.insidetextanchor = "end"
                traza.textfont = dict(
                    color=[
                        NAVY if posicion == "outside" else "white"
                        for posicion in posiciones
                    ]
                )
                traza.cliponaxis = False            
            
            fig.update_layout(
                xaxis_title="Grupo de edad",
                yaxis_title="Trabajadores (Miles)",
                margin=dict(t=S(90)),
                legend=dict(title=dict(text="Sexo"),
                            orientation="h", 
                            yanchor="bottom", y=1.02,
                            xanchor="center", x=0.5,
                            bgcolor="rgba(0,0,0,0)"),
            )
            show(fig, 460)
    with r:
        nom = fil(D["ISSSTE_TRAB_NOMBRA"], y, m, e, s, a)
        q = (nom.groupby("NOMBRAMIENO_DES", as_index=False).TOTAL.sum()
                .sort_values("TOTAL").tail(10))
        if not q.empty:
            fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", orientation="h",
                         title="Top 10 nombramientos")
                        
            umbral = q["TOTAL"].max() * 0.05
            posiciones = [
                "outside" if total <= umbral else "inside"
                for total in q["TOTAL"]
            ]
            fig.update_traces(marker_color=BLUE, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                NAVY if posicion == "outside" else "white"
                                for posicion in posiciones
                            ],
                            cliponaxis=False,
                              )
            fig.update_layout(
                            xaxis_title="Trabajadores",
                            yaxis_title="Tipos de plazas",
                        )
            show(fig, 460)

    l, r = st.columns(2)
    with l:
        z = avg(nom, "NOMBRAMIENO_DES", "PROM_REMUNARACN").sort_values("Promedio").tail(10)
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="NOMBRAMIENO_DES", orientation="h",
                         title="Remuneración promedio por nombramiento")
            fig.update_traces(marker_color=TEAL, 
                              texttemplate="$%{x:,.0f}",
                              textposition="inside")
            fig.update_layout(
                            xaxis_title="Promedio salario (MXN)",
                            yaxis_title="Tipos de plazas",
                                    )
            show(fig, 460)
    with r:
        sm = (fil(D["ISSSTE_TRAB_SM"], y, m, e, s)
              .groupby("RANGO_SM", as_index=False).TOTAL.sum())
        if not sm.empty:
            fig = px.bar(sm, x="RANGO_SM", y="TOTAL",
                         title="Distribución por rango salarial")
            fig.update_traces(marker_color=GUINDA, texttemplate="%{y:,.0f}",
                              textposition="outside")
            fig.update_layout(
                            yaxis_title="Rango salarial (SM)",
                            xaxis_title="Trabajadores",
                                                )
            show(fig, 460)

    exportar(x, f"trabajadores_{y}_{m}")


def page_plazas():
    b = D["ISSSTE_PLAZ_NOMBRA_PROM"]
    y, m, e, s, _ = ctl(b, age=False)
    x = fil(b, y, m, e, s)
    col, factor = _col_plazas(x)

    head("Plazas de trabajadores afiliados al ISSSTE",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    if col is None:
        st.warning("La hoja ISSSTE_PLAZ_NOMBRA_PROM no contiene columnas "
                   "para calcular plazas (RANGO_PLAZAS + TOTAL).")
        return

    x = x.copy()
    x["_plazas"] = _plazas_de(x, col, factor)
    total_plazas = x["_plazas"].sum()

    c1, c2, c3 = st.columns(3)
    kpi(c1, "Total plazas", fmt(total_plazas))
    kpi(c2, "Nombramientos",
        fmt(x.NOMBRAMIENO_DES.nunique()) if "NOMBRAMIENO_DES" in x else "—")
    kpi(c3, "Entidades", fmt(len(e)))

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr_plazas(b, e, s)
        if not t.empty:
            t["TOTAL_MILLONES"] = t["TOTAL"] / 1_000_000
            fig = px.line(t, x="Periodo", y="TOTAL", text="TOTAL_MILLONES",
                          title="Evolución del número de plazas")
            fig.update_traces(
                mode="lines+markers+text",
                texttemplate="%{text:.2f}",
                textposition="top center",                
                line=dict(width=2.5, color=NAVY),
                hovertemplate="<b>Plazas</b><br>%{x|%b %Y}<br>"
                                            "%{y:,.0f}<extra></extra>")
            fig.update_layout(
                            xaxis_title="Periodo",
                            yaxis_title="Plazas (Millones)",
                                                )
            show(fig, 380)
    with r:
        if "ENT_CVE" in x.columns:
            q = (x.groupby("ENT_CVE", as_index=False)["TOTAL"].sum()
                   .dropna(subset=["ENT_CVE"]))
            gj, feature_key, pad = _load_geojson_meta()
            q["ENT_CVE"] = q["ENT_CVE"].apply(
                lambda code: _normalizar_codigo(code, pad)
            )
            q["Entidad"] = q["ENT_CVE"].map(
                lambda code: names.get(
                    int(code) if code and code.isdigit() else code,
                    names.get(code, code),
                )
            )

            if gj is not None and feature_key is not None:
                fig = px.choropleth(
                    q, geojson=gj, locations="ENT_CVE",
                    featureidkey=feature_key, color="TOTAL",
                    color_continuous_scale=["#D9E6F5", NAVY],
                    hover_name="Entidad",
                    hover_data={"TOTAL": ":,.0f", "ENT_CVE": False,
                                "Entidad": False},
                    labels={"TOTAL": "Plazas"},
                    title="Distribución de plazas por entidad federativa",
                )
                fig.update_geos(fitbounds="locations", visible=False,
                                bgcolor="white")
                fig.update_layout(
                    coloraxis_colorbar=dict(title="Plazas", thickness=12,
                                            len=0.7),
                    margin=dict(t=S(20)),
                )
            else:
                coordenadas = geo[["cve_issste", "Latitud", "Longitud"]]
                q = q.merge(coordenadas, left_on="ENT_CVE",
                            right_on="cve_issste", how="left")
                q = q.dropna(subset=["Latitud"])
                fig = px.scatter_geo(
                    q, lat="Latitud", lon="Longitud", size="TOTAL",
                    hover_name="Entidad", color="TOTAL",
                    color_continuous_scale=["#D9E6F5", NAVY],
                    title="Distribución de plazas por entidad federativa",
                )
                fig.update_geos(lataxis_range=[14, 33], lonaxis_range=[-118, -86],
                                showland=True, landcolor="#F1F3F6",
                                showocean=True, oceancolor="#DCEAF5",
                                showcountries=True, countrycolor="#C7CCD3",
                                showcoastlines=True, coastlinecolor="#B8BEC5")
                fig.update_layout(coloraxis_showscale=False,
                                  margin=dict(t=S(20)))
            show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "RANGO_PLAZAS" in x.columns:
            q = (x.groupby("RANGO_PLAZAS", as_index=False)["_plazas"].sum()
                   .rename(columns={"_plazas": "TOTAL"}))
            if not q.empty:
                fig = px.pie(q, names="RANGO_PLAZAS", values="TOTAL", hole=0.4,
                             title="Proporción según número de plazas")
                fig.update_traces(textposition="outside", textinfo="percent")
                show(fig, 420)
    with r:
        if "NOMBRAMIENO_DES" in x.columns:
            q = (x.groupby("NOMBRAMIENO_DES", as_index=False)["_plazas"].sum()
                   .rename(columns={"_plazas": "TOTAL"})
                   .sort_values("TOTAL").tail(10))
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", 
                             orientation="h",
                             title="Top 10 plazas por nombramiento")
                
                umbral = q["TOTAL"].max() * 0.10
                posiciones = [
                            "outside" if total <= umbral else "inside"
                            for total in q["TOTAL"]
                            ]
                            
                fig.update_traces(marker_color=BLUE, 
                                texttemplate="%{x:,.0f}",
                                textposition=posiciones,
                                textfont_color=[
                                NAVY if posicion == "outside" else "white"
                                for posicion in posiciones
                                ],
                                cliponaxis=False,)
                fig.update_layout(
                                yaxis_title="Tipos de plazas",
                                xaxis_title="Plazas",
                                                    )
                show(fig, 420)

    exportar(x, f"plazas_{y}_{m}")

_ORDEN_EDAD_PEN = ["0-49", "50-59", "60-69", "70-79", "80 y más"]

def page_pensionados():
    b = D["ISSSTE_PENS_REGPEN"]
    y, m, e, s, a = ctl(b)
    x = fil(b, y, m, e, s, a)
    head("Padrón de pensionados y jubilados directos",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Pensionados", fmt(x.TOTAL.sum()), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Hombres", fmt(x.loc[x.SEXO == "H", "TOTAL"].sum()) if "SEXO" in x else "—")
    kpi(c3, "Mujeres", fmt(x.loc[x.SEXO == "M", "TOTAL"].sum()) if "SEXO" in x else "—")
    p = wavg(x, "PROM_IMPORTE")
    kpi(c4, "Pensión promedio", money(p) if p else "—")

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s, a)
        if not t.empty:
            t["TOTAL_MILES"] = t["TOTAL"] / 1000
            primer_corte = t["Periodo"].min()
            ultimo_corte = t["Periodo"].max()
            paso = 50_000
            limite = ((int(t["TOTAL"].max()) + paso - 1) // paso) * paso
            tickvals = list(range(0, limite + paso, paso))
           
            fig = px.line(t, x="Periodo", y="TOTAL", text="TOTAL_MILES",
                          title="Evolución de pensionados y jubilados (2015–2025)")
            fig.update_traces(
                            mode="lines+markers+text",
                            line=dict(width=2.5, color=BLUE, dash="dot"),
                            marker=dict(symbol="diamond"),
                            texttemplate="%{text:,.0f}",
                            textposition="top center",
                             hovertemplate="<b>Pensionados</b><br>%{x|%b %Y}<br>"
                                            "%{y:,.0f}<extra></extra>")
            fig.update_layout(
                        title="Evolución de trabajadores (2015–2025)",
                        
                        xaxis_title="Periodo",
                        xaxis_range=[
                        primer_corte - pd.DateOffset(months=1),
                        ultimo_corte,],
                        xaxis_tick0=primer_corte,
                        xaxis_dtick="M12",
                        xaxis_tickformat="%Y",
                        
                        yaxis_title="Pensionados y jubilados(miles)",
                        yaxis_tickmode="array",
                        yaxis_tickvals=tickvals,
                        yaxis_ticktext=[f"{v // 1000:,}k" for v in tickvals],
            )

            show(fig, 380)
            
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()
        if not q.empty:
            q["SEXO"] = q["SEXO"].replace({"H": "Hombre", "M": "Mujer"})
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=0.4,
                         title="Distribución por sexo", color="SEXO",
                         color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            fig.update_traces(textposition="inside", texttemplate="%{percent}<br>%{value:,.0f}",
            )
            fig.update_layout(
                margin=dict(t=S(90)),
                legend=dict(
                    title=dict(text="Sexo"),
                    orientation="h", yanchor="bottom", y=1.02,
                    xanchor="center", x=0.5,
                    bgcolor="rgba(0,0,0,0)"))
            show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "RECAT_PEN_ACT" in x.columns and "TIP_REG_PENS" in x.columns:
            q = x.groupby(["RECAT_PEN_ACT", "TIP_REG_PENS"], as_index=False).TOTAL.sum()
            q["TIP_REG_PENS"] = q["TIP_REG_PENS"].replace({"DECIMO TRANSITORIO": "Transitorio", "CUENTA INDIVIDUAL": "Individual"})
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="RECAT_PEN_ACT", color="TIP_REG_PENS",
                             orientation="h", barmode="group",
                             title="Régimen pensionario y tipo de pensión recibida")
                
                umbral = q["TOTAL"].max() * 0.1

                for traza in fig.data:
                    valores = traza.x
                    posiciones = [
                        "outside" if valor is not None and valor <= umbral else "inside"
                        for valor in valores
                    ]

                    traza.texttemplate = "%{x:,.0f}"
                    traza.textposition = posiciones
                    traza.textfont = dict(
                        color=[
                            NAVY if posicion == "outside" else "white"
                            for posicion in posiciones
                        ]
                    )
                    traza.cliponaxis = False
                
                fig.update_layout(
                                xaxis_title="Pensionados",
                                yaxis_title="Tipo de pensión",
                                margin=dict(t=S(90)),
                                legend=dict(title=dict(text="Régimen"),
                                            orientation="h", yanchor="bottom", y=1.02,
                                            xanchor="center", x=0.5,
                                            bgcolor="rgba(0,0,0,0)"),
                            )
                show(fig, 440)
    with r:
        sv = fil(D["ISSSTE_PENS_SERV_PROM"], y, m, e, s)
        z = avg(sv, "RANGO_TSERV", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RANGO_TSERV", orientation="h",
                         title="Pensión promedio por años de servicio")
            
            umbral = z["Promedio"].max() * 0.10
            posiciones = [
                        "outside" if total <= umbral else "inside"
                        for total in z["Promedio"]
                        ]
                                        
            fig.update_traces(marker_color=BLUE, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                            NAVY if posicion == "outside" else "white"
                                            for posicion in posiciones
                                            ],
                            cliponaxis=False,
            )
                                            
            fig.update_layout(
                            yaxis_title="Tiempo de servicio (años)",
                            xaxis_title="Pensión promedio (MXN)",
                                                    )
            show(fig, 440)

    l, r = st.columns(2)
    with l:
        z = avg(x, ["RECAT_PEN_ACT", "TIP_REG_PENS"], "PROM_IMPORTE") \
            if "PROM_IMPORTE" in x.columns else pd.DataFrame()
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RECAT_PEN_ACT", color="TIP_REG_PENS",
                         orientation="h", barmode="group",
                         title="Pensión promedio por régimen y tipo")
            
            umbral = z["Promedio"].max() * 0.1
            for traza in fig.data:
                    valores = traza.x
                    posiciones = [
                        "outside" if valor is not None and valor <= umbral else "inside"
                        for valor in valores
                    ]

                    traza.texttemplate = "%{x:,.0f}"
                    traza.textposition = posiciones
                    traza.textfont = dict(
                        color=[
                            NAVY if posicion == "outside" else "white"
                            for posicion in posiciones
                        ]
                    )
            traza.cliponaxis = False
            
            fig.update_layout(
                                            xaxis_title="Promedio salario (MXN)",
                                            yaxis_title="Tipo de pensión",
                                            margin=dict(t=S(90)),
                                            legend=dict(title=dict(text="Régimen"),
                                                        orientation="h", yanchor="bottom", y=1.02,
                                                        xanchor="center", x=0.5,
                                                        bgcolor="rgba(0,0,0,0)"),
                                        )
            show(fig, 440)
    with r:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            edades = _agrupa_edades(x, col_edad="GRUPO_EDAD")
            edades["Edad_Agrup"] = edades["Edad_Agrup"].replace(
                {"0-19": "0-49", "20-29": "0-49",
                 "30-39": "0-49", "40-49": "0-49"}
            )
            edades["SEXO"] = edades["SEXO"].replace({"H": "Hombre", "M": "Mujer"})
            ag = (edades.groupby(["Edad_Agrup", "SEXO"], as_index=False)
                                             .TOTAL.sum())
            ag = ag[ag["Edad_Agrup"].isin(_ORDEN_EDAD_PEN)]
            fig = px.bar(ag, x="TOTAL", y="Edad_Agrup", color="SEXO",
                         orientation="h", barmode="group",
                         title="Pensionados por edad y sexo",
                         category_orders={
                                    "Edad_Agrup": _ORDEN_EDAD_PEN,
                                    "SEXO": ["Hombre", "Mujer"]},
                        labels={"Edad_Agrup": "Grupo de edad","TOTAL": "Pensionados"},
                        color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            
            umbral = ag["TOTAL"].max() * 0.1

            for traza in fig.data:
                    valores = traza.x
                    posiciones = [
                        "outside" if valor is not None and valor <= umbral else "inside"
                        for valor in valores
                    ]

                    traza.texttemplate = "%{x:,.0f}"
                    traza.textposition = posiciones
                    traza.textfont = dict(
                        color=[
                            NAVY if posicion == "outside" else "white"
                            for posicion in posiciones
                        ]
                    )
            traza.cliponaxis = False
                    
            fig.update_layout(
                                    xaxis_title="Pensionados",
                                    yaxis_title="Grupo de edad",
                                    margin=dict(t=S(90)),
                                    legend=dict(title=dict(text="Sexo"),
                                                orientation="h", yanchor="bottom", y=1.02,
                                                xanchor="center", x=0.5,
                                                bgcolor="rgba(0,0,0,0)"),
                                )    
            show(fig, 440)
            
    exportar(x, f"pensionados_{y}_{m}")


def page_deudos():
    b = D["ISSSTE_DEUD_TOT"]
    y, m, e, s, a = ctl(b)
    x = fil(b, y, m, e, s, a)
    head("Padrón de pensionados deudos",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    c1, c2, c3 = st.columns(3)
    kpi(c1, "Total deudos", fmt(x.TOTAL.sum()), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Hombres", fmt(x.loc[x.SEXO == "H", "TOTAL"].sum()) if "SEXO" in x else "—")
    kpi(c3, "Mujeres", fmt(x.loc[x.SEXO == "M", "TOTAL"].sum()) if "SEXO" in x else "—")

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s, a)
        if not t.empty:
            t["TOTAL_MILES"] = t["TOTAL"] / 1000
            primer_corte = t["Periodo"].min()
            ultimo_corte = t["Periodo"].max()
            paso = 10_000
            limite = ((int(t["TOTAL"].max()) + paso - 1) // paso) * paso
            tickvals = list(range(0, limite + paso, paso))
            
            fig = px.line(t, x="Periodo", y="TOTAL", text="TOTAL_MILES",
                          markers=True,
                          title="Evolución de pensionados deudos")
            fig.update_traces(
                mode="lines+markers+text",
                line=dict(width=4, color=GUINDA, dash="dot"),
                hovertemplate="<b>Deudos</b><br>%{x|%b %Y}<br>"
                                "%{y:,.0f}<extra></extra>",
                marker=dict(symbol="diamond"),
                texttemplate="%{text:,.0f}",
                textposition="top center",
            )
            fig.update_layout(
                            yaxis_title="Pnesionados deudos (miles)",
                            xaxis_title="Periodo",
                            xaxis_range=[
                                primer_corte - pd.DateOffset(months=1),
                                ultimo_corte + pd.DateOffset(months=1),
                            ],
                            xaxis_tick0=primer_corte,
                            xaxis_dtick="M12",
                            xaxis_tickformat="%Y",
                            yaxis_tickmode="array",
                            yaxis_tickvals=tickvals,
                            yaxis_ticktext=[f"{v // 1000:,}k" for v in tickvals],
                                )     
            show(fig, 380)
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()        
        if not q.empty:
            q["SEXO"] = q["SEXO"].replace({"H": "Hombre", "M": "Mujer"})
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=0.4,
                         title="Distribución por sexo", color="SEXO",
                         color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            fig.update_traces(textposition="inside", texttemplate="%{percent}<br>%{value:,.0f}",)
            fig.update_layout(
                            margin=dict(t=S(90)),
                            legend=dict(
                                title=dict(text="Sexo"),
                                orientation="h", yanchor="bottom", y=1.02,
                                xanchor="center", x=0.5,
                                bgcolor="rgba(0,0,0,0)"))
            show(fig, 380)

    d = fil(D["ISSSTE_DEUD_TBNDEUDO_PROM"], y, m, e, s)
    l, r = st.columns(2)
    with l:
        if "GRUPO_EDAD" in x.columns:
            #q = x.groupby("GRUPO_EDAD", as_index=False).TOTAL.sum().sort_values("TOTAL")
            edades = _agrupa_edades(x, col_edad="GRUPO_EDAD")
            ag = (edades.groupby(["Edad_Agrup"], as_index=False).TOTAL.sum())
            ag = ag[ag["Edad_Agrup"].isin(_ORDEN_EDAD)]
            
            if not ag.empty:
                paso = 10_000
                limite = ((int(ag["TOTAL"].max()) + paso - 1) // paso) * paso
                tickvals = list(range(0, limite + paso, paso))
                fig = px.bar(ag, x="TOTAL", y="Edad_Agrup", 
                             orientation="h",barmode="stack",
                             title="Deudos por grupo de edad")
                   
                umbral = ag["TOTAL"].max() * 0.10
                posiciones = [
                            "outside" if total <= umbral else "inside"
                            for total in ag["TOTAL"]
                        ]
                                                        
                fig.update_traces(marker_color=BLUE, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                             NAVY if posicion == "outside" else "white"
                                            for posicion in posiciones
                                            ],
                            cliponaxis=False,
                            )
                
                fig.update_layout(
                        yaxis_title="Grupo de edad",
                        xaxis_title="Deudos (miles)",
                        xaxis_tickmode="array",
                        xaxis_tickvals=tickvals,
                        xaxis_ticktext=[str(v // 1000) for v in tickvals],
                    )       
                show(fig, 440)
                
 
    with r:
        if "RECAT_TBN" in d.columns and "RECAT_DEUDO" in d.columns:
            q = d.groupby(["RECAT_TBN", "RECAT_DEUDO"], as_index=False).TOTAL.sum()
            if not q.empty:
                paso = 20_000
                limite = ((int(q["TOTAL"].max()) + paso - 1) // paso) * paso
                tickvals = list(range(0, limite + paso, paso))
                fig = px.bar(q, x="TOTAL", y="RECAT_TBN", color="RECAT_DEUDO",
                             orientation="h", barmode="group",
                             title="Deudos por tipo de pensión")
                fig.update_layout(
                            xaxis_title="Deudos (miles)",
                            yaxis_title="Tipos de pensiones",
                            margin=dict(t=S(90)),
                            legend=dict(
                                title=dict(text="Tipos de deudos"),
                                orientation="h", yanchor="bottom", y=1.02,
                                xanchor="center", x=0.5,
                                bgcolor="rgba(0,0,0,0)"),
                            xaxis_tickmode="array",
                            xaxis_tickvals=tickvals,
                            xaxis_ticktext=[str(v // 1000) for v in tickvals],
                            )
                show(fig, 440)

    l, r = st.columns(2)
    with l:
        z = avg(d, "RECAT_DEUDO", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            tickvals = list(range(0, ((int(z["Promedio"].max()) + 1999) // 2000) * 2000 + 1,2000,))
            fig = px.bar(z, x="Promedio", y="RECAT_DEUDO", orientation="h",
                         title="Monto promedio por tipo de deudo")
            
            umbral = z["Promedio"].max() * 0.10
            posiciones = [
                        "outside" if total <= umbral else "inside"
                        for total in z["Promedio"]
                         ]
                                                                    
            fig.update_traces(marker_color=TEAL, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                            NAVY if posicion == "outside" else "white"
                                            for posicion in posiciones
                                            ],
                            cliponaxis=False,
                            )
                            
            fig.update_layout(
                            yaxis_title="Tipos de deudos",
                            xaxis_title="Monto promedio (miles MXN)",
                            xaxis_tickmode="array",
                            xaxis_tickvals=tickvals,
                            xaxis_ticktext=[str(v // 1000) for v in tickvals],
                                )
            show(fig, 440)
    with r:
        z = avg(d, "RECAT_TBN", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            tickvals = list(range(0, ((int(z["Promedio"].max()) + 1999) // 2000) * 2000 + 1, 2000,))
            fig = px.bar(z, x="Promedio", y="RECAT_TBN", orientation="h",
                         title="Monto promedio por tipo de pensión")
            
            umbral = z["Promedio"].max() * 0.10
            posiciones = [
                        "outside" if total <= umbral else "inside"
                        for total in z["Promedio"]
                        ]
                                                                    
            fig.update_traces(marker_color=AMBER, 
                            texttemplate="%{x:,.0f}",
                            textposition=posiciones,
                            textfont_color=[
                                            NAVY if posicion == "outside" else "white"
                                            for posicion in posiciones
                                            ],
                            cliponaxis=False,
                            )
                            
            fig.update_layout(
                            yaxis_title="Tipos de pensiones",
                            xaxis_title="Monto promedio (miles MXN)",
                            xaxis_tickmode="array",
                            xaxis_tickvals=tickvals,
                            xaxis_ticktext=[str(v // 1000) for v in tickvals],) 
            show(fig, 440)

    exportar(x, f"deudos_{y}_{m}")


def page_familiares():
    b = D["ISSSTE_FAMILIAR_PARENT"]
    y, m, e, s, a = ctl(b)
    x = fil(b, y, m, e, s, a)
    head("Padrón de familiares derechohabientes",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    c1, c2, c3 = st.columns(3)
    kpi(c1, "Familiares", fmt(x.TOTAL.sum()), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Hombres", fmt(x.loc[x.SEXO == "H", "TOTAL"].sum()) if "SEXO" in x else "—")
    kpi(c3, "Mujeres", fmt(x.loc[x.SEXO == "M", "TOTAL"].sum()) if "SEXO" in x else "—")

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s, a)
        if not t.empty:
            t["TOTAL_MILLONES"] = t["TOTAL"] / 1_000_000
            fig = px.line(t, x="Periodo", y="TOTAL", text="TOTAL_MILLONES",
                          title="Evolución de familiares derechohabientes",)
            fig.update_traces(
                mode="lines+markers+text",  
                texttemplate="%{text:.2f}",
                textposition="top center",                
                line=dict(width=2.5, color=GUINDA),
                hovertemplate="<b>Familiares</b><br>%{x|%b %Y}<br>"
                                            "%{y:,.0f}<extra></extra>")
            fig.update_xaxes(title_text="Periodo")
            fig.update_yaxes(title_text="Familiares (Millones)")
            show(fig, 380)
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()
        if not q.empty:
            q["SEXO"] = q["SEXO"].replace({"H": "Hombre", "M": "Mujer"})
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=.50,
                         title="Distribución de familiares derechohabientes por sexo", 
                         color="SEXO",
                         color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            fig.update_traces(textposition="inside", 
                              texttemplate="%{percent}<br>%{value:,.0f}",)
            fig.update_layout(
                            margin=dict(t=S(90)),
                            legend=dict(
                                        title=dict(text="Sexo"),
                                        orientation="h", yanchor="bottom", y=1.02,
                                        xanchor="center", x=0.5,
                                        bgcolor="rgba(0,0,0,0)"
                                        )
                            )
            show(fig, 380)
                    
    l, r = st.columns(2)
    with l:
        if "PARENTESCO_DES" in x.columns and "SEXO" in x.columns:
            q = x.groupby(["PARENTESCO_DES", "SEXO"], as_index=False).TOTAL.sum()
            if not q.empty:
                q["SEXO"] = q["SEXO"].replace({"H": "Hombre", "M": "Mujer"})
                fig = px.bar(q, x="TOTAL", y="PARENTESCO_DES", color="SEXO",
                             orientation="h", barmode="group",
                             title="Total de familiares derechohabientes por parentesco y sexo",
                             color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
                
                umbral = q["TOTAL"].max() * 0.1
                
                for traza in fig.data:
                                    valores = traza.x
                                    posiciones = [
                                        "outside" if valor is not None and valor <= umbral else "inside"
                                        for valor in valores
                                    ]
                
                                    traza.texttemplate = "%{x:,.0f}"
                                    traza.textposition = posiciones
                                    traza.textfont = dict(
                                        color=[
                                            NAVY if posicion == "outside" else "white"
                                            for posicion in posiciones
                                        ]
                                    )
                traza.cliponaxis = False
                            
                fig.update_layout(
                                yaxis_title="Parentesco",
                                xaxis_title="Familiares derechohabientes",
                                legend=dict(
                                            title=dict(text="Sexo"),
                                            orientation="h", yanchor="bottom", y=1.02,
                                            xanchor="center", x=0.5,
                                            bgcolor="rgba(0,0,0,0)"),
                                )
                show(fig, 480)
                
    with r:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            #q = x.groupby(["GRUPO_EDAD", "SEXO"], as_index=False).TOTAL.sum()
            edades = _agrupa_edades(x, col_edad="GRUPO_EDAD")
            edades["SEXO"] = edades["SEXO"].replace(
                {"H": "Hombre", "M": "Mujer"}
            )
            q = (edades.groupby(["Edad_Agrup","SEXO"], as_index=False).TOTAL.sum())
            q = q[q["Edad_Agrup"].isin(_ORDEN_EDAD)]

            fig = px.bar(q, x="TOTAL", y="Edad_Agrup", color="SEXO",
                         orientation="h", barmode="group",
                         title="Familiares derechohabientes por edad y sexo",
                         category_orders={
                                        "Edad_Agrup": _ORDEN_EDAD,
                                        "SEXO": ["Hombre", "Mujer"],
                                        },
                         labels={"Edad_Agrup": "Grupo de edad",
                                "TOTAL": "Familiares derechohabientes"},
                         color_discrete_map={"Hombre": NAVY, "Mujer": CORAL})
            fig.update_layout(
                                yaxis_title="Grupo de edad",
                                xaxis_title="Familiares derechohabientes",
                                margin=dict(t=S(90)),
                                legend=dict(
                                title=dict(text="Sexo"),
                                orientation="h", yanchor="bottom", y=1.02,
                                xanchor="center", x=0.5,
                                bgcolor="rgba(0,0,0,0)"),
                            )
            
            show(fig, 480)

    exportar(x, f"familiares_{y}_{m}")

# ═════════════════════════════════════════════════════════════════════════
# NAVEGACIÓN
# ═════════════════════════════════════════════════════════════════════════
NAV_ITEMS = [
    ("🏛️", "ISSSTE",       page_issste),
    ("👥", "Trabajadores", page_trabajadores),
    ("🧾", "Plazas",       page_plazas),
    ("💰", "Pensionados",  page_pensionados),
    ("🕊️", "Deudos",       page_deudos),
    ("👨‍👩‍👧", "Familiares",   page_familiares),
]

with st.sidebar:
    st.image(
        str(Path(__file__).parent / "Imagen" / "ISSSTE_ICONO.jpg"),
        width=300,
    )
    st.markdown(_escalar_px(
        '<div style="text-align:left;padding:0 0 14px;font-size:16px;'
        'opacity:.8;color:#fff;">Tablero ejecutivo</div>'
    ), unsafe_allow_html=True)

    opts = [f"{ico}  {lbl}" for ico, lbl, _ in NAV_ITEMS]
    sel = st.radio("Navegación", opts, label_visibility="collapsed")
    idx = opts.index(sel)

NAV_ITEMS[idx][2]()

st.caption("Fuente: RESUMEN_ISSSTE.xlsx · Promedios ponderados por TOTAL · "
           "TABLAS_DIM como catálogo de filtros y geografía.")