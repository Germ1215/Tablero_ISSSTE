"""
Tablero ejecutivo ISSSTE · Streamlit
Diseño ejecutivo, interactivo y con navegación optimizada.
Compatible con Python 3.12+ (probado en 3.14).
"""
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px

# ───────────────────────────── Configuración ─────────────────────────────
st.set_page_config(
    page_title="ISSSTE · Tablero Ejecutivo",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).parent
F = BASE_DIR / "data" / "RESUMEN_ISSSTE.xlsx"

SHEETS = [
    "TABLAS_DIM", "ISSSTE_TRAB_TOT2015_2025", "ISSSTE_TRAB_NOMBRA", "ISSSTE_TRAB_SM",
    "ISSSTE_TRAB_NPLAZAS", "ISSSTE_PLAZ_NOMBRA_PROM", "ISSSTE_PENS_REGPEN",
    "ISSSTE_PENS_PP_PROM", "ISSSTE_PENS_SERV_PROM", "ISSSTE_DEUD_TOT",
    "ISSSTE_DEUD_TBNDEUDO", "ISSSTE_DEUD_TBNDEUDO_PROM", "ISSSTE_DEUD_TBNDEUDO_RET",
    "ISSSTE_FAMILIAR_PARENT",
]

# ───────────────────────────── Carga de datos ─────────────────────────────
@st.cache_data(show_spinner="Cargando información del ISSSTE…")
def load(path: str):
    return pd.read_excel(path, sheet_name=SHEETS, engine="openpyxl")

if not F.exists():
    st.error(f"No se encontró el archivo de datos:\n\n`{F}`\n\n"
             "Coloque `RESUMEN_ISSSTE.xlsx` en la carpeta `data/`.")
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

# ───────────────────────────── Paleta ejecutiva ─────────────────────────────
GUINDA, GUINDA_DARK = "#74192D", "#4B0F1D"
NAVY, BLUE, SKY     = "#0B1F4B", "#1F6FB2", "#4FA8E0"
TEAL, AMBER, CORAL  = "#217A60", "#C9A227", "#B04A5A"
GRAY, BG, GRID      = "#6B7280", "#F4F5F7", "#EDEFF3"
C = [GUINDA, NAVY, BLUE, TEAL, AMBER, CORAL, SKY, GRAY]
SEX_COLORS = {"H": NAVY, "M": CORAL}

# ───────────────────────────── Estilos CSS ─────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background: {BG}; }}
.block-container {{ padding-top: 0.8rem; padding-bottom: 2rem; max-width: 1500px; }}

[data-testid="stSidebar"] {{ background: linear-gradient(180deg, {GUINDA} 0%, {GUINDA_DARK} 100%); }}
[data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
[data-testid="stSidebar"] label {{ color: #F3D7DE !important; font-weight: 600 !important; font-size: 12.5px !important; }}

/* ── Selectbox / Multiselect: fondo blanco, TEXTO OSCURO ── */
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] [data-baseweb="input"] > div {{
    background: #FFFFFF !important;
    border-color: rgba(255,255,255,0.55) !important;
    color: {NAVY} !important;
}}
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-baseweb="input"] *,
[data-testid="stSidebar"] [data-baseweb="select"] input,
[data-testid="stSidebar"] [data-baseweb="input"] input {{
    color: {NAVY} !important;
    -webkit-text-fill-color: {NAVY} !important;
    opacity: 1 !important;
}}
[data-testid="stSidebar"] [data-baseweb="select"] svg,
[data-testid="stSidebar"] [data-baseweb="input"] svg {{
    fill: {GUINDA} !important;
    color: {GUINDA} !important;
}}
/* Chips del multiselect */
[data-testid="stSidebar"] [data-baseweb="tag"] {{
    background: {GUINDA} !important;
    color: #FFFFFF !important;
}}
[data-testid="stSidebar"] [data-baseweb="tag"] * {{ color: #FFFFFF !important; }}

/* ── Dropdown desplegable (se renderiza en un portal fuera del sidebar) ── */
[data-baseweb="popover"] [role="listbox"],
[data-baseweb="popover"] ul {{
    background: #FFFFFF !important;
}}
[data-baseweb="popover"] [role="option"],
[data-baseweb="popover"] li {{
    color: {NAVY} !important;
    background: #FFFFFF !important;
}}
[data-baseweb="popover"] [role="option"]:hover,
[data-baseweb="popover"] li:hover,
[data-baseweb="popover"] [aria-selected="true"] {{
    background: #F3D7DE !important;
    color: {GUINDA} !important;
}}

/* ── Radio de navegación ── */
[data-testid="stSidebar"] .stRadio > div {{ gap: 6px; }}
[data-testid="stSidebar"] .stRadio label {{
    background: rgba(255,255,255,0.06);
    border-radius: 8px; padding: 7px 10px;
    transition: background .15s ease;
    font-weight: 500;
}}
[data-testid="stSidebar"] .stRadio label:hover {{ background: rgba(255,255,255,0.18); }}

.hero {{
    background: linear-gradient(90deg, {GUINDA} 0%, #A1243F 55%, #D0466B 100%);
    color: #fff; padding: 14px 22px; border-radius: 12px;
    box-shadow: 0 8px 22px rgba(116,25,45,.22);
    margin-bottom: 12px;
}}
.hero h2 {{ margin: 0; font-weight: 700; letter-spacing: .3px; font-size: 22px; }}
.hero .sub {{ opacity: .85; font-size: 12.5px; margin-top: 2px; }}

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

.stPlotlyChart {{ background: #fff; border-radius: 12px; padding: 6px;
                  box-shadow: 0 2px 10px rgba(0,0,0,.04); }}
div[data-testid="stDataFrame"] {{ background: #fff; border-radius: 10px; }}
</style>
""", unsafe_allow_html=True)

# ───────────────────────────── Catálogo geográfico ─────────────────────────────
dim = D["TABLAS_DIM"]
geo = (dim[["cve_issste", "Entidad", "Latitud", "Longitud"]]
       .dropna(subset=["cve_issste"])
       .drop_duplicates("cve_issste"))
names = geo.set_index("cve_issste").Entidad.to_dict()

# ───────────────────────────── Utilidades de gráficos ─────────────────────────────
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

def kpi(col, label, value, delta=None):
    dlt_html = ""
    if delta is not None:
        cls, arrow = ("up", "▲") if delta >= 0 else ("dn", "▼")
        dlt_html = f'<div class="dlt {cls}">{arrow} {abs(delta):.2f}% vs año previo</div>'
    col.markdown(f"""
        <div class="kpi">
            <div class="lbl">{label}</div>
            <div class="val">{value}</div>
            {dlt_html}
        </div>
    """, unsafe_allow_html=True)

def head(title, subtitle=""):
    sub = f'<div class="sub">{subtitle}</div>' if subtitle else ""
    st.markdown(f'<div class="hero"><h2>{title}</h2>{sub}</div>', unsafe_allow_html=True)

def fmt(n):
    try:
        return f"{float(n):,.0f}"
    except Exception:
        return "—"

def money(n):
    try:
        return f"${float(n):,.0f}"
    except Exception:
        return "—"

# ───────────────────────────── Filtros ─────────────────────────────
def ctl(b: pd.DataFrame, age: bool = True):
    """Barra lateral de filtros adaptativos."""
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

    e = []
    if "ENT_CVE" in b.columns:
        cs = sorted(b.ENT_CVE.dropna().unique())
        e = st.sidebar.multiselect("Entidad", cs, default=cs,
                                   format_func=lambda x: names.get(x, x)) or cs

    s = ["H", "M"]
    if "SEXO" in b.columns:
        s = st.sidebar.multiselect(
            "Sexo", ["H", "M"], default=["H", "M"],
            format_func=lambda x: {"H": "Hombres", "M": "Mujeres"}[x]) or ["H", "M"]

    a = []
    if age and "GRUPO_EDAD" in b.columns:
        aa = sorted(b.GRUPO_EDAD.dropna().unique())
        a = st.sidebar.multiselect("Grupo de edad", aa, default=aa) or aa

    return y, m, e, s, a

def fil(d, y, m, e, s, a=None):
    x = d.copy()
    if y is not None and "ANIO" in x.columns: x = x[x.ANIO == y]
    if m is not None and "MES" in x.columns:  x = x[x.MES == m]
    if e and "ENT_CVE" in x.columns:          x = x[x.ENT_CVE.isin(e)]
    if s and "SEXO" in x.columns:             x = x[x.SEXO.isin(s)]
    if a and "GRUPO_EDAD" in x.columns:       x = x[x.GRUPO_EDAD.isin(a)]
    return x

def tr(d, e, s, a=None):
    x = d.copy()
    if e and "ENT_CVE" in x.columns:    x = x[x.ENT_CVE.isin(e)]
    if s and "SEXO" in x.columns:       x = x[x.SEXO.isin(s)]
    if a and "GRUPO_EDAD" in x.columns: x = x[x.GRUPO_EDAD.isin(a)]
    x["Periodo"] = pd.to_datetime(dict(year=x.ANIO, month=x.MES, day=1))
    return x.groupby("Periodo", as_index=False).TOTAL.sum()

def avg(x, g, v):
    z = x.copy()
    z["w"] = z.TOTAL * z[v]
    return (z.groupby(g)
              .agg(w=("w", "sum"), TOTAL=("TOTAL", "sum"))
              .assign(Promedio=lambda q: q.w / q.TOTAL)
              .reset_index())

def wavg(x, v):
    """Promedio ponderado global."""
    if x.empty or v not in x.columns:
        return None
    t = x.TOTAL.sum()
    return (x.TOTAL * x[v]).sum() / t if t else None

def delta_aoa(b, y, m, e, s, a=None):
    if y is None or "ANIO" not in b.columns:
        return None
    cur = fil(b, y, m, e, s, a).TOTAL.sum()
    prev = fil(b, y - 1, m, e, s, a).TOTAL.sum()
    return ((cur - prev) / prev * 100) if prev else None

def _mapa(x, title, palette):
    q = (x.groupby("ENT_CVE", as_index=False).TOTAL.sum()
           .merge(geo, left_on="ENT_CVE", right_on="cve_issste")
           .dropna(subset=["Latitud"]))
    if q.empty:
        return px.scatter(title=title)
    f = px.scatter_geo(q, lat="Latitud", lon="Longitud", size="TOTAL",
                       hover_name="Entidad", title=title,
                       color="TOTAL", color_continuous_scale=palette)
    f.update_geos(lataxis_range=[14, 33], lonaxis_range=[-118, -86],
                  showland=True, landcolor="#F1F3F6",
                  showocean=True, oceancolor="#DCEAF5",
                  showcountries=True, countrycolor="#C7CCD3",
                  showcoastlines=True, coastlinecolor="#B8BEC5")
    f.update_layout(coloraxis_showscale=False)
    return f

def exportar(x, nombre):
    with st.expander("📋 Ver y descargar datos filtrados"):
        st.dataframe(x, use_container_width=True, height=300)
        st.download_button("⬇️ Descargar CSV",
                           x.to_csv(index=False).encode("utf-8"),
                           file_name=f"{nombre}.csv", mime="text/csv")

# ═════════════════════════════════════════════════════════════════════════
#  PÁGINAS
# ═════════════════════════════════════════════════════════════════════════
def page_issste():
    b = D["ISSSTE_TRAB_TOT2015_2025"]
    y, m, e, s, a = ctl(b)
    head("ISSSTE · Resumen ejecutivo",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    tot_trab = fil(b, y, m, e, s, a).TOTAL.sum()
    tot_plaz = fil(D["ISSSTE_TRAB_NPLAZAS"], y, m, e, s).TOTAL.sum()
    tot_pen  = fil(D["ISSSTE_PENS_REGPEN"], y, m, e, s, a).TOTAL.sum()
    tot_deu  = fil(D["ISSSTE_DEUD_TOT"], y, m, e, s, a).TOTAL.sum()
    tot_fam  = fil(D["ISSSTE_FAMILIAR_PARENT"], y, m, e, s, a).TOTAL.sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    kpi(c1, "Trabajadores", fmt(tot_trab), delta_aoa(b, y, m, e, s, a))
    kpi(c2, "Plazas",       fmt(tot_plaz))
    kpi(c3, "Pensionados",  fmt(tot_pen))
    kpi(c4, "Deudos",       fmt(tot_deu))
    kpi(c5, "Familiares",   fmt(tot_fam))

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s, a)
        fig = px.area(t, x="Periodo", y="TOTAL", markers=True,
                      title="Evolución de trabajadores afiliados")
        fig.update_traces(line=dict(width=2.5, color=GUINDA),
                          fillcolor="rgba(116,25,45,0.10)")
        show(fig, 380)
    with r:
        x = fil(b, y, m, e, s, a)
        show(_mapa(x, "Distribución territorial", ["#F3D7DE", GUINDA]), 380)

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
        q = (fil(D["ISSSTE_TRAB_NPLAZAS"], y, m, e, s)
             .groupby("RANGO_PLAZAS", as_index=False).TOTAL.sum())
        if not q.empty:
            fig = px.pie(q, names="RANGO_PLAZAS", values="TOTAL", hole=.55,
                         title="Plazas por rango")
            fig.update_traces(textposition="outside", textinfo="percent+label")
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
        fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                      title="Evolución 2015–2025")
        fig.update_traces(line=dict(width=2.5, color=GUINDA))
        show(fig, 380)
    with r:
        show(_mapa(x, "Distribución territorial", ["#F3D7DE", GUINDA]), 380)

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
    b = D["ISSSTE_TRAB_NPLAZAS"]
    y, m, e, s, _ = ctl(b, age=False)
    x = fil(b, y, m, e, s)
    head("Plazas de trabajadores afiliados al ISSSTE",
         f"Corte {int(m):02d}/{y}" if m else f"Año {y}")

    c1, c2, c3 = st.columns(3)
    kpi(c1, "Total plazas", fmt(x.TOTAL.sum()), delta_aoa(b, y, m, e, s))
    kpi(c2, "Rangos de plazas",
        fmt(x.RANGO_PLAZAS.nunique()) if "RANGO_PLAZAS" in x else "—")
    kpi(c3, "Nombramientos",
        fmt(x.NOMBRAMIENO_DES.nunique()) if "NOMBRAMIENO_DES" in x else "—")

    st.markdown("")
    l, r = st.columns([3, 2])
    with l:
        t = tr(b, e, s)
        fig = px.line(t, x="Periodo", y="TOTAL", markers=True,
                      title="Evolución del número de plazas")
        fig.update_traces(line=dict(width=2.5, color=NAVY))
        show(fig, 380)
    with r:
        show(_mapa(x, "Distribución territorial de plazas", ["#D9E6F5", NAVY]), 380)

    l, r = st.columns(2)
    with l:
        q = x.groupby("RANGO_PLAZAS", as_index=False).TOTAL.sum()
        if not q.empty:
            fig = px.pie(q, names="RANGO_PLAZAS", values="TOTAL", hole=.55,
                         title="Proporción según número de plazas")
            fig.update_traces(textposition="outside", textinfo="percent+label")
            show(fig, 420)
    with r:
        q = (x.groupby("NOMBRAMIENO_DES", as_index=False).TOTAL.sum()
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
        z = avg(x, ["RECAT_PEN_ACT", "TIP_REG_PENS"], "PROM_IMPORTE")
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
            q = (x.groupby("GRUPO_EDAD", as_index=False).TOTAL.sum()
                   .sort_values("TOTAL"))
            if not q.empty:
                fig = px.bar(q, x="TOTAL", y="GRUPO_EDAD", orientation="h",
                             title="Deudos por grupo de edad")
                fig.update_traces(marker_color=BLUE, texttemplate="%{x:,.0f}",
                                  textposition="outside")
                show(fig, 440)
    with r:
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
#  NAVEGACIÓN
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
                <div style="font-weight:800;font-size:17px;letter-spacing:.5px;">ISSSTE</div>
                <div style="font-size:11px;opacity:.75;">Tablero ejecutivo</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    opts = [f"{ico}  {lbl}" for ico, lbl, _ in NAV_ITEMS]
    sel = st.radio("Navegación", opts, label_visibility="collapsed")
    idx = opts.index(sel)

NAV_ITEMS[idx][2]()

st.caption("Fuente: RESUMEN_ISSSTE.xlsx · Promedios ponderados por TOTAL · "
           "TABLAS_DIM usado como catálogo de filtros y geografía.")