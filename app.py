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
def load(path: str):
    return pd.read_excel(path, sheet_name=SHEETS, engine="openpyxl")

if not F.exists():
    st.error(f"No se encontró el archivo de datos:\n\n`{F}`")
    st.stop()

try:
    _raw = load(str(F))
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
st.markdown(f"""
<style>
.stApp {{ background: {BG}; }}
.block-container {{ padding-top: 0.8rem; padding-bottom: 2rem; max-width: 1500px; }}

/* ── Sidebar fondo ── */
[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {GUINDA} 0%, {GUINDA_DARK} 100%);
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
    background: #fff; border-left: 5px solid {GUINDA};
    border-radius: 10px; padding: 12px 16px;
    box-shadow: 0 2px 10px rgba(0,0,0,.05);
    height: 100%;
}}
.kpi .lbl {{ color: {GRAY}; font-size: 11px; font-weight: 700; letter-spacing: .6px; text-transform: uppercase; }}
.kpi .val {{ color: {NAVY}; font-size: 25px; font-weight: 700; line-height: 1.15; margin-top: 2px; }}
.kpi .dlt {{ font-size: 12px; font-weight: 600; margin-top: 2px; }}
.kpi .up {{ color: {TEAL}; }}
.kpi .dn {{ color: {CORAL}; }}
.kpi .note {{ font-size: 11px; color: {GRAY}; margin-top: 3px; }}

.stPlotlyChart {{ background: #fff; border-radius: 12px; padding: 6px;
                  box-shadow: 0 2px 10px rgba(0,0,0,.04); }}
div[data-testid="stDataFrame"] {{ background: #fff; border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

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

def theme(fig, h=340, title=None):
    fig.update_layout(
        height=h, margin=dict(l=12, r=12, t=48, b=8),
        colorway=C, paper_bgcolor="white", plot_bgcolor="white",
        font=dict(family="Segoe UI, Arial", color=NAVY, size=12),
        title=dict(text=title if title is not None else fig.layout.title.text,
                   font=dict(size=14, color=NAVY), x=0.02, xanchor="left"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.22,
                    xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)",
                    font=dict(size=11)),
        xaxis=dict(showgrid=False, linecolor=GRID, ticks="outside",
                   tickfont=dict(size=11), zeroline=False),
        yaxis=dict(gridcolor=GRID, linecolor="rgba(0,0,0,0)",
                   tickfont=dict(size=11), zeroline=False),
        hoverlabel=dict(bgcolor="white", font_size=12,
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
    col.markdown(f"""
        <div class="kpi">
            <div class="lbl">{label}</div>
            <div class="val">{value}</div>
            {dlt_html}
            {note_html}
        </div>
    """, unsafe_allow_html=True)

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
        "<div style='font-weight:700;font-size:13px;letter-spacing:.4px;'>🔎 Filtros</div>",
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

def fil(d, y, m, e, s, a=None):
    x = d.copy()
    if y is not None and "ANIO" in x.columns: x = x[x.ANIO == y]
    if m is not None and "MES" in x.columns:  x = x[x.MES == m]
    if e and "ENT_CVE" in x.columns:          x = x[x.ENT_CVE.isin(e)]
    if s and "SEXO" in x.columns:             x = x[x.SEXO.isin(s)]
    if a and "GRUPO_EDAD" in x.columns:       x = x[x.GRUPO_EDAD.isin(a)]
    return x

def tot_o_ultimo(d, y, m, e, s, a=None):
    x = fil(d, y, m, e, s, a)
    if not x.empty and "TOTAL" in x.columns:
        return x.TOTAL.sum(), y
    if "ANIO" in d.columns and not d.empty:
        ly = int(d.ANIO.max())
        x2 = fil(d, ly, None, e, s, a)
        if not x2.empty and "TOTAL" in x2.columns:
            return x2.TOTAL.sum(), ly
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
    if x.empty or "TOTAL" not in x.columns or v not in x.columns:
        return pd.DataFrame(columns=g + ["Promedio"])
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
        st.dataframe(x, use_container_width=True, height=300)
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
            title="Tasa por entidad federativa",
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

# ═════════════════════════════════════════════════════════════════════════
# PÁGINAS
# ═════════════════════════════════════════════════════════════════════════
def page_issste():
    b = D["ISSSTE_TRAB_TOT2015_2025"]
    y, m, e, s, a = ctl(b)
    head("ISSSTE · Resumen ejecutivo",
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

    c1, c2, c3, c4, c5 = st.columns(5)
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
    l, r = st.columns([3, 2])
    with l:
        t_trab = tr(b, e, s, a)                                # 2015–2025
        t_plaz = tr_plazas(D["ISSSTE_PLAZ_NOMBRA_PROM"], e, s) # 2021–2025

        fig = go.Figure()

        if not t_trab.empty:
            fig.add_trace(go.Scatter(
                x=t_trab["Periodo"], y=t_trab["TOTAL"],
                mode="lines+markers",
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
                name="Plazas",
                line=dict(color=NAVY, width=2.6, dash="dot"),
                marker=dict(size=6, color=NAVY),
                hovertemplate="<b>Plazas</b><br>%{x|%b %Y}<br>"
                              "%{y:,.0f}<extra></extra>",
            ))

        if not fig.data:
            st.info("Sin datos de evolución para el filtro seleccionado.")
        else:
            fig.update_layout(
                title="Evolución de trabajadores y plazas (2015–2025)",
                xaxis_title=None,
                yaxis_title="Personas / Plazas",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=-0.25,
                            xanchor="center", x=0.5, bgcolor="rgba(0,0,0,0)"),
            )
            show(fig, 420)
    with r:
        fig = _mapa_tasa_sm(y, m, e, s)
        show(fig, 420)

    # ── Dona de rango salarial + barra de plazas por nombramiento ────────
    l, r = st.columns(2)
    with l:
        q = (fil(D["ISSSTE_TRAB_SM"], y, m, e, s)
             .groupby("RANGO_SM", as_index=False).TOTAL.sum())
        if not q.empty:
            fig = px.pie(q, names="RANGO_SM", values="TOTAL", hole=.55,
                         title="Trabajadores por rango salarial")
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 400)
    with r:
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
            fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", orientation="h",
                         title="Top 10 plazas por nombramiento")
            fig.update_traces(marker_color=GUINDA, texttemplate="%{x:,.0f}",
                              textposition="outside")
            show(fig, 400)


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
            fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                          title="Evolución 2015–2025")
            fig.update_traces(line=dict(width=2.5, color=GUINDA))
            show(fig, 380)
    with r:
        fig = _mapa_tasa_sm(y, m, e, s)
        show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            ag = x.groupby(["GRUPO_EDAD", "SEXO"], as_index=False).TOTAL.sum()
            fig = px.bar(ag, x="TOTAL", y="GRUPO_EDAD", color="SEXO",
                         orientation="h", barmode="group",
                         title="Trabajadores por edad y sexo",
                         color_discrete_map=SEX_COLORS)
            show(fig, 460)
    with r:
        nom = fil(D["ISSSTE_TRAB_NOMBRA"], y, m, e, s, a)
        q = (nom.groupby("NOMBRAMIENO_DES", as_index=False).TOTAL.sum()
                .sort_values("TOTAL").tail(10))
        if not q.empty:
            fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", orientation="h",
                         title="Top 10 nombramientos")
            fig.update_traces(marker_color=BLUE, texttemplate="%{x:,.0f}",
                              textposition="outside")
            show(fig, 460)

    l, r = st.columns(2)
    with l:
        z = avg(nom, "NOMBRAMIENO_DES", "PROM_REMUNERACN").sort_values("Promedio").tail(10)
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="NOMBRAMIENO_DES", orientation="h",
                         title="Remuneración promedio por nombramiento")
            fig.update_traces(marker_color=TEAL, texttemplate="$%{x:,.0f}",
                              textposition="outside")
            show(fig, 460)
    with r:
        sm = (fil(D["ISSSTE_TRAB_SM"], y, m, e, s)
              .groupby("RANGO_SM", as_index=False).TOTAL.sum())
        if not sm.empty:
            fig = px.bar(sm, x="RANGO_SM", y="TOTAL",
                         title="Distribución por rango salarial")
            fig.update_traces(marker_color=GUINDA, texttemplate="%{y:,.0f}",
                              textposition="outside")
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
            fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                          title="Evolución del número de plazas")
            fig.update_traces(line=dict(width=2.5, color=NAVY))
            show(fig, 380)
    with r:
        if "ENT_CVE" in x.columns:
            q = (x.groupby("ENT_CVE", as_index=False)["_plazas"].sum()
                   .rename(columns={"_plazas": "TOTAL"}))
            q = q.merge(geo, left_on="ENT_CVE", right_on="cve_issste",
                        how="left").dropna(subset=["Latitud"])
            fig = px.scatter_geo(q, lat="Latitud", lon="Longitud", size="TOTAL",
                                 hover_name="Entidad",
                                 title="Distribución territorial de plazas",
                                 color="TOTAL",
                                 color_continuous_scale=["#D9E6F5", NAVY])
            fig.update_geos(lataxis_range=[14, 33], lonaxis_range=[-118, -86],
                            showland=True, landcolor="#F1F3F6",
                            showocean=True, oceancolor="#DCEAF5",
                            showcountries=True, countrycolor="#C7CCD3",
                            showcoastlines=True, coastlinecolor="#B8BEC5")
            fig.update_layout(coloraxis_showscale=False)
            show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "RANGO_PLAZAS" in x.columns:
            q = (x.groupby("RANGO_PLAZAS", as_index=False)["_plazas"].sum()
                   .rename(columns={"_plazas": "TOTAL"}))
            if not q.empty:
                fig = px.pie(q, names="RANGO_PLAZAS", values="TOTAL", hole=.55,
                             title="Proporción según número de plazas")
                fig.update_traces(textposition="outside", textinfo="percent+label")
                show(fig, 420)
    with r:
        if "NOMBRAMIENO_DES" in x.columns:
            q = (x.groupby("NOMBRAMIENO_DES", as_index=False)["_plazas"].sum()
                   .rename(columns={"_plazas": "TOTAL"})
                   .sort_values("TOTAL").tail(10))
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="NOMBRAMIENO_DES", orientation="h",
                             title="Top 10 plazas por nombramiento")
                fig.update_traces(marker_color=GUINDA, texttemplate="%{x:,.0f}",
                                  textposition="outside")
                show(fig, 420)

    exportar(x, f"plazas_{y}_{m}")

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
            fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                          title="Evolución de pensionados")
            fig.update_traces(line=dict(width=2.5, color=GUINDA))
            show(fig, 380)
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()
        if not q.empty:
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=.55,
                         title="Distribución por sexo", color="SEXO",
                         color_discrete_map=SEX_COLORS)
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "RECAT_PEN_ACT" in x.columns and "TIP_REG_PENS" in x.columns:
            q = x.groupby(["RECAT_PEN_ACT", "TIP_REG_PENS"], as_index=False).TOTAL.sum()
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="RECAT_PEN_ACT", color="TIP_REG_PENS",
                             orientation="h", barmode="stack",
                             title="Régimen y tipo de pensión")
                show(fig, 440)
    with r:
        sv = fil(D["ISSSTE_PENS_SERV_PROM"], y, m, e, s)
        z = avg(sv, "RANGO_TSERV", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RANGO_TSERV", orientation="h",
                         title="Pensión promedio por años de servicio")
            fig.update_traces(marker_color=TEAL, texttemplate="$%{x:,.0f}",
                              textposition="outside")
            show(fig, 440)

    l, r = st.columns(2)
    with l:
        z = avg(x, ["RECAT_PEN_ACT", "TIP_REG_PENS"], "PROM_IMPORTE") \
            if "PROM_IMPORTE" in x.columns else pd.DataFrame()
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RECAT_PEN_ACT", color="TIP_REG_PENS",
                         orientation="h", barmode="group",
                         title="Pensión promedio por régimen y tipo")
            fig.update_traces(texttemplate="$%{x:,.0f}", textposition="outside")
            show(fig, 440)
    with r:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            ag = x.groupby(["GRUPO_EDAD", "SEXO"], as_index=False).TOTAL.sum()
            fig = px.bar(ag, x="TOTAL", y="GRUPO_EDAD", color="SEXO",
                         orientation="h", barmode="group",
                         title="Pensionados por edad y sexo",
                         color_discrete_map=SEX_COLORS)
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
            fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                          title="Evolución de pensionados deudos")
            fig.update_traces(line=dict(width=2.5, color=GUINDA))
            show(fig, 380)
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()
        if not q.empty:
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=.55,
                         title="Distribución por sexo", color="SEXO",
                         color_discrete_map=SEX_COLORS)
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 380)

    d = fil(D["ISSSTE_DEUD_TBNDEUDO_PROM"], y, m, e, s)
    l, r = st.columns(2)
    with l:
        if "GRUPO_EDAD" in x.columns:
            q = x.groupby("GRUPO_EDAD", as_index=False).TOTAL.sum().sort_values("TOTAL")
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="GRUPO_EDAD", orientation="h",
                             title="Deudos por grupo de edad")
                fig.update_traces(marker_color=BLUE, texttemplate="%{x:,.0f}",
                                  textposition="outside")
                show(fig, 440)
    with r:
        if "RECAT_TBN" in d.columns and "RECAT_DEUDO" in d.columns:
            q = d.groupby(["RECAT_TBN", "RECAT_DEUDO"], as_index=False).TOTAL.sum()
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="RECAT_TBN", color="RECAT_DEUDO",
                             orientation="h", barmode="stack",
                             title="Deudos por tipo de pensión")
                show(fig, 440)

    l, r = st.columns(2)
    with l:
        z = avg(d, "RECAT_DEUDO", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RECAT_DEUDO", orientation="h",
                         title="Monto promedio por tipo de deudo")
            fig.update_traces(marker_color=TEAL, texttemplate="$%{x:,.0f}",
                              textposition="outside")
            show(fig, 440)
    with r:
        z = avg(d, "RECAT_TBN", "PROM_IMPORTE").sort_values("Promedio")
        if not z.empty:
            fig = px.bar(z, x="Promedio", y="RECAT_TBN", orientation="h",
                         title="Monto promedio por tipo de pensión")
            fig.update_traces(marker_color=AMBER, texttemplate="$%{x:,.0f}",
                              textposition="outside")
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
            fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                          title="Evolución de familiares derechohabientes")
            fig.update_traces(line=dict(width=2.5, color=GUINDA))
            show(fig, 380)
    with r:
        q = x.groupby("SEXO", as_index=False).TOTAL.sum()
        if not q.empty:
            fig = px.pie(q, names="SEXO", values="TOTAL", hole=.55,
                         title="Distribución por sexo", color="SEXO",
                         color_discrete_map=SEX_COLORS)
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 380)

    l, r = st.columns(2)
    with l:
        if "PARENTESCO_DES" in x.columns and "SEXO" in x.columns:
            q = x.groupby(["PARENTESCO_DES", "SEXO"], as_index=False).TOTAL.sum()
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="PARENTESCO_DES", color="SEXO",
                             orientation="h", barmode="group",
                             title="Familiares por parentesco y sexo",
                             color_discrete_map=SEX_COLORS)
                fig.update_traces(texttemplate="%{x:,.0f}", textposition="outside")
                show(fig, 480)
    with r:
        if "GRUPO_EDAD" in x.columns and "SEXO" in x.columns:
            q = x.groupby(["GRUPO_EDAD", "SEXO"], as_index=False).TOTAL.sum()
            fig = px.bar(q, x="TOTAL", y="GRUPO_EDAD", color="SEXO",
                         orientation="h", barmode="stack",
                         title="Familiares por edad y sexo",
                         color_discrete_map=SEX_COLORS)
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
    st.markdown(f"""
        <div style="display:flex;align-items:center;gap:12px;padding:4px 4px 14px;">
            <div style="width:44px;height:44px;border-radius:10px;background:#fff;
                        display:flex;align-items:center;justify-content:center;
                        font-weight:800;color:{GUINDA};font-size:16px;">IS</div>
            <div>
                <div style="font-weight:800;font-size:17px;letter-spacing:.5px;color:#fff;">ISSSTE</div>
                <div style="font-size:11px;opacity:.75;color:#fff;">Tablero ejecutivo</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    opts = [f"{ico}  {lbl}" for ico, lbl, _ in NAV_ITEMS]
    sel = st.radio("Navegación", opts, label_visibility="collapsed")
    idx = opts.index(sel)

NAV_ITEMS[idx][2]()

st.caption("Fuente: RESUMEN_ISSSTE.xlsx · Promedios ponderados por TOTAL · "
           "TABLAS_DIM como catálogo de filtros y geografía.")