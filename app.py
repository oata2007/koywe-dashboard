"""
app.py
Dashboard OTC de Koywe — Streamlit (fuente: exportación Mongo Excel)
"""
from __future__ import annotations

import os
import subprocess
import tempfile
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

def _slack_token() -> str:
    """Lee el token de Slack desde st.secrets (cloud) o variable de entorno (local)."""
    try:
        return st.secrets["SLACK_TOKEN"]
    except Exception:
        return os.environ.get("SLACK_TOKEN", "")

def _subprocess_env() -> dict:
    """Entorno con el token de Slack para subprocesos."""
    return {**os.environ, "SLACK_TOKEN": _slack_token()}

from data_loader import (
    load_excel_dashboard, client_analysis_excel,
    load_client_names, load_weekly_summary, load_canal_monthly,
    download_drive_file,
    MONTHS_ES, QUARTERS, COUNTRY_MAP_CODE,
)

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Koywe · Dashboard OTC",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Brand Colors ───────────────────────────────────────────────────────────────
DEEP_GREEN  = "#022416"
GREEN_DARK  = "#19502C"
GREEN_MID   = "#1B6B3A"
KOYWE_GREEN = "#0D442E"
LIMA        = "#C8FF1D"
LIMA_3      = "#efffbb"
LAVENDER    = "#ccb6ea"
CHARCOAL    = "#4d4d4d"
WARN        = "#F59E0B"
DANGER      = "#EF4444"

COUNTRY_PALETTE = ["#19502C","#ccb6ea","#1B6B3A","#4d4d4d","#0D442E","#022416","#C8FF1D"]
TIER_COLORS = {"Enterprise":"#022416","Corporate":"#0D442E","Business":"#1B6B3A","Retail":"#4d4d4d"}

# ── Logo SVG ───────────────────────────────────────────────────────────────────
_LOGO_PATH = (
    "M118.148 7.76172C125.247 7.76183 129.977 13.2406 129.977 20.0146L130 20C129.995 20.5221 "
    "129.962 21.0443 129.905 21.5664H111.078C111.67 25.3309 114.381 27.6426 118.177 27.6426C"
    "119.428 27.6853 120.664 27.3959 121.764 26.7979C122.863 26.1997 123.783 25.317 124.423 "
    "24.2441H129.417C128.602 26.5748 127.081 28.5922 125.058 30.0068C123.039 31.4215 120.621 "
    "32.162 118.157 32.124C112.305 32.1238 107.685 28.2168 106.528 22.6865L99.501 26.2803L"
    "97.7383 31.5498H93.8711L87.8906 17.4082L81.9434 31.5498H78.0771L70.3955 8.59766L59.8848 "
    "31.9629C57.2689 37.8208 54.9608 39.8662 50.2314 39.8662H47.876V35.5703H49.8662C52.9322 "
    "35.5703 53.8758 34.6682 55.5723 30.9512L56.1025 29.8838L52.5254 22.2549L45.2363 18.4102C"
    "45.2932 18.9085 45.3271 19.4164 45.3271 19.9385C45.3271 26.9309 40.0188 32.1191 32.9248 "
    "32.1191C25.8311 32.1189 20.5235 26.9212 20.5234 19.9385C20.5234 12.9556 25.8405 7.76195 "
    "32.9248 7.76172C37.1897 7.76172 40.8103 9.63719 43.0186 12.6562L49.6719 16.1689L46 "
    "8.34082H51.1836L58.5576 24.5381L65.4717 8.34082L70.5234 8.3125L70.5137 8.32715H75.2197L"
    "80.4561 23.9785L87.1133 8.57812H88.6494L95.3975 24.0918L100.714 8.32715H105.486L101.32 "
    "20.7881L106.344 18.2197C107.116 12.153 111.931 7.76172 118.148 7.76172ZM4.76758 "
    "0V17.3135H7.69141L13.9512 8.33203H19.6045L11.7471 19.4775L20.4385 31.5547H14.6094L"
    "7.6582 21.6475H4.76758V31.5547H0V0H4.76758ZM32.9297 12.3193C28.5322 12.3195 25.3809 "
    "15.5143 25.3809 19.9385C25.3809 24.3626 28.5418 27.5623 32.9297 27.5625C37.3178 27.5625 "
    "40.4784 24.3864 40.4785 19.9385C40.4785 15.4904 37.3273 12.3193 32.9297 12.3193ZM118.182 "
    "12.0723C114.594 12.0723 111.983 14.2421 111.201 17.6885H125.228C124.37 14.1567 121.769 "
    "12.0724 118.182 12.0723Z"
)

def _logo_svg(fill: str, height: int = 28) -> str:
    w = round(130 * height / 40)
    return (
        f'<svg width="{w}" height="{height}" viewBox="0 0 130 40" fill="none" '
        f'xmlns="http://www.w3.org/2000/svg" style="display:block;flex-shrink:0">'
        f'<path d="{_LOGO_PATH}" fill="{fill}"/></svg>'
    )

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Raleway:wght@300;400;500&display=swap');
  html, body, [class*="css"] {{ font-family: 'Raleway', sans-serif; }}
  h1, h2, h3, h4, .stMetric label {{ font-family: 'DM Sans', sans-serif !important; }}
  [data-testid="stAppViewContainer"] {{ background: #F2F7F2; }}
  [data-testid="stSidebar"] {{ background: {DEEP_GREEN}; }}
  [data-testid="stSidebar"] * {{ color: {LIMA_3} !important; font-family: 'Raleway', sans-serif !important; }}
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stRadio label {{ color: {LIMA_3} !important; font-size: 11px; }}
  [data-testid="stSidebar"] hr {{ border-color: {CHARCOAL} !important; }}
  [data-testid="stSidebar"] [data-baseweb="select"] *,
  [data-testid="stSidebar"] [data-baseweb="input"] *,
  [data-testid="stSidebar"] input {{ color: #111111 !important; }}
  [data-testid="stSidebar"] [data-baseweb="select"] > div {{ background: white !important; }}
  [data-testid="stSidebar"] [data-baseweb="tag"] {{ background: {GREEN_MID} !important; border: none !important; }}
  [data-testid="stSidebar"] [data-baseweb="tag"] span,
  [data-testid="stSidebar"] [data-baseweb="tag"] svg {{ color: {LIMA} !important; fill: {LIMA} !important; }}
  [data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
    background: {GREEN_DARK} !important; border: 1.5px dashed {LIMA} !important; border-radius: 8px !important;
  }}
  [data-testid="stSidebar"] [data-testid="stFileUploader"] section * {{ color: {LIMA_3} !important; }}
  [data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
    background: {LIMA} !important; color: {DEEP_GREEN} !important;
    border: none !important; border-radius: 9999px !important; font-weight: 600 !important;
  }}
  h1 {{ color: {KOYWE_GREEN} !important; font-family: 'DM Sans', sans-serif !important; }}
  h2, h3 {{ color: {GREEN_DARK} !important; font-family: 'DM Sans', sans-serif !important; }}
  .stTabs [data-baseweb="tab-list"] {{ background: {DEEP_GREEN}; border-radius: 8px 8px 0 0; gap: 0; }}
  .stTabs [data-baseweb="tab"] {{ color: {LIMA_3} !important; padding: 10px 20px; font-family: 'DM Sans', sans-serif !important; }}
  .stTabs [aria-selected="true"] {{ color: {LIMA} !important; border-bottom: 3px solid {LIMA} !important; font-weight: 700; }}
  div[data-testid="metric-container"] {{
    background: white; border-radius: 10px; padding: 16px;
    border-top: 4px solid {LIMA}; box-shadow: 0 2px 8px rgba(2,36,22,.10);
  }}
  div[data-testid="metric-container"] label {{ color: {KOYWE_GREEN} !important; font-weight: 500; }}
  div[data-testid="metric-container"] [data-testid="stMetricValue"] {{ color: {DEEP_GREEN} !important; font-family: 'DM Sans', sans-serif !important; font-weight: 700; }}
  [data-testid="stHeader"] {{ display: none !important; }}
  .block-container {{ padding-top: 1.5rem; }}
  .stButton button {{ background: {LIMA}; color: {DEEP_GREEN}; font-family: 'DM Sans', sans-serif; font-weight: 600; border: none; border-radius: 9999px; }}
  .stButton button:hover {{ background: #D4FF4D; color: {DEEP_GREEN}; }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ────────────────────────────────────────────────────────────────────
def fmt_usd(v: float) -> str:
    if pd.isna(v) or v == 0: return "$0"
    if abs(v) >= 1_000_000: return f"${v/1_000_000:.2f}M"
    if abs(v) >= 1_000:     return f"${v/1_000:.2f}K"
    return f"${v:.2f}"

def pct_delta(curr: float, prev: float) -> str | None:
    if not prev or pd.isna(prev) or pd.isna(curr): return None
    return f"{(curr - prev) / prev * 100:+.1f}%"

def chart_defaults(fig, height=260):
    fig.update_layout(
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=0, r=0, t=30, b=0),
        font=dict(family="DM Sans, sans-serif", size=11, color=KOYWE_GREEN),
        legend=dict(font=dict(size=10)),
        title_font=dict(family="DM Sans, sans-serif", color=KOYWE_GREEN, size=13),
    )
    fig.update_xaxes(showgrid=False, linecolor="#D1E8D8")
    fig.update_yaxes(gridcolor="#EBF5EE", linecolor="#D1E8D8")
    for trace in fig.data:
        if trace.type == "bar":
            if getattr(trace, "orientation", None) == "h":
                trace.hovertemplate = "%{y}<br><b>$%{x:,.2f}</b><extra></extra>"
            else:
                trace.hovertemplate = "%{x}<br><b>$%{y:,.2f}</b><extra></extra>"
        elif trace.type in ("scatter", "scattergl"):
            trace.hovertemplate = "%{x|%b %Y}<br><b>$%{y:,.2f}</b><extra></extra>"
    return fig


# ── Auto-descarga desde Google Drive ──────────────────────────────────────────
_DRIVE_METRICS_ID  = st.secrets.get("DRIVE_METRICS_ID", "")
_DRIVE_CHARTS_ID   = st.secrets.get("DRIVE_CHARTS_ID", "")
_DRIVE_METRICS_TMP = "/tmp/koywe_drive_metrics.xlsx"
_DRIVE_CHARTS_TMP  = "/tmp/koywe_drive_charts.xlsx"

if _DRIVE_METRICS_ID and "drive_loaded" not in st.session_state:
    with st.spinner("📡 Cargando datos desde Google Drive…"):
        _ok_m = download_drive_file(_DRIVE_METRICS_ID, _DRIVE_METRICS_TMP)
        _ok_c = (
            download_drive_file(_DRIVE_CHARTS_ID, _DRIVE_CHARTS_TMP)
            if _DRIVE_CHARTS_ID else False
        )
    st.session_state["drive_loaded"]  = _ok_m
    st.session_state["drive_charts"]  = _ok_c
    st.session_state["drive_metrics_id"] = _DRIVE_METRICS_ID
    st.session_state["drive_charts_id"]  = _DRIVE_CHARTS_ID

_drive_metrics_ready = (
    st.session_state.get("drive_loaded", False)
    and os.path.exists(_DRIVE_METRICS_TMP)
)
_drive_charts_ready = (
    st.session_state.get("drive_charts", False)
    and os.path.exists(_DRIVE_CHARTS_TMP)
)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(_logo_svg("#FFFFFF", height=28), unsafe_allow_html=True)
    st.markdown(
        f"<p style='color:{CHARCOAL};font-size:11px;margin-top:4px'>Dashboard OTC · Stablecoins</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    # Estado de Google Drive
    if _DRIVE_METRICS_ID:
        if _drive_metrics_ready:
            st.success("📡 Datos cargados desde Drive", icon="✅")
        else:
            st.warning("⚠️ No se pudo conectar con Drive")

        if st.button("🔄 Actualizar desde Drive", use_container_width=True):
            # Forzar re-descarga borrando caché de sesión y archivos tmp
            for _k in ["drive_loaded", "drive_charts", "drive_metrics_id", "drive_charts_id"]:
                st.session_state.pop(_k, None)
            for _p in [_DRIVE_METRICS_TMP, _DRIVE_CHARTS_TMP]:
                if os.path.exists(_p):
                    os.remove(_p)
            st.cache_data.clear()
            st.rerun()

        st.divider()
        with st.expander("📂 Subir archivo manualmente (override)"):
            st.markdown(
                f"<p style='color:{LIMA_3};font-size:11px;margin-bottom:4px'>OTC Metrics</p>",
                unsafe_allow_html=True,
            )
            uploaded = st.file_uploader(
                "Chart_1/2/4/11 (.xlsx)",
                type=["xlsx"],
                label_visibility="collapsed",
                key="upload_main",
            )
            st.markdown(
                f"<p style='color:{LIMA_3};font-size:11px;margin-bottom:4px;margin-top:8px'>OTC Charts</p>",
                unsafe_allow_html=True,
            )
            uploaded_tx = st.file_uploader(
                "Chart_14 (.xlsx)",
                type=["xlsx"],
                label_visibility="collapsed",
                key="upload_tx",
            )
    else:
        # Sin Drive configurado → uploaders normales
        st.markdown(
            f"<p style='color:{LIMA_3};font-size:11px;margin-bottom:4px'>📂 OTC Metrics</p>",
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Exportación Mongo — Chart_1/2/4/11 (.xlsx)",
            type=["xlsx"],
            label_visibility="collapsed",
            key="upload_main",
        )
        st.markdown(
            f"<p style='color:{LIMA_3};font-size:11px;margin-bottom:4px;margin-top:8px'>📂 OTC Charts</p>",
            unsafe_allow_html=True,
        )
        uploaded_tx = st.file_uploader(
            "Exportación Mongo — Chart_14 (.xlsx)",
            type=["xlsx"],
            label_visibility="collapsed",
            key="upload_tx",
        )
        st.divider()
        if uploaded:
            if st.button("🔄 Recargar datos", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
            st.caption(f"Actualizado: {datetime.now().strftime('%H:%M')}")


# ── Load data ──────────────────────────────────────────────────────────────────
# Resolver fuente de datos: upload manual tiene prioridad sobre Drive
_has_metrics = uploaded is not None or _drive_metrics_ready
_has_charts  = (uploaded_tx is not None
                or _drive_charts_ready
                or os.path.exists("/tmp/koywe_chart14.xlsx"))

if not _has_metrics:
    col_center = st.columns([1, 2, 1])[1]
    with col_center:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(_logo_svg(KOYWE_GREEN, height=40), unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        if _DRIVE_METRICS_ID:
            st.error("❌ No se pudo descargar el archivo de métricas desde Drive. Revisa que el archivo esté compartido con la service account.")
        else:
            st.info("📂 **Sube el archivo de exportación** desde el panel izquierdo para cargar el dashboard.")
    st.stop()

@st.cache_data(show_spinner=False)
def cached_load(file_bytes: bytes) -> dict:
    return load_excel_dashboard(file_bytes)

if uploaded is not None:
    # Upload manual — tiene prioridad
    with st.spinner("Cargando datos…"):
        file_bytes = uploaded.read()
        data = cached_load(file_bytes)
    # Guardar para slack_reporter.py
    _tmp_chart = "/tmp/koywe_current_chart.xlsx"
    with open(_tmp_chart, "wb") as _f:
        _f.write(file_bytes)
else:
    # Datos desde Drive
    _tmp_chart = _DRIVE_METRICS_TMP
    with st.spinner("Procesando datos…"):
        with open(_DRIVE_METRICS_TMP, "rb") as _f:
            file_bytes = _f.read()
        data = cached_load(file_bytes)

# Resolver fuente de Chart_14 (canales)
_tmp_tx = "/tmp/koywe_chart14.xlsx"
if uploaded_tx is not None:
    _tx_bytes = uploaded_tx.read()
    if "tx_file_saved" not in st.session_state or st.session_state.get("tx_name") != uploaded_tx.name:
        with open(_tmp_tx, "wb") as _f:
            _f.write(_tx_bytes)
        st.session_state["tx_file_saved"] = True
        st.session_state["tx_name"] = uploaded_tx.name
    WEEKLY_PATH = _tmp_tx
elif _drive_charts_ready:
    WEEKLY_PATH = _DRIVE_CHARTS_TMP
elif os.path.exists(_tmp_tx):
    WEEKLY_PATH = _tmp_tx
else:
    WEEKLY_PATH = ""

vol_country = data["vol_country"]   # País, Periodo, Fecha, Mes, Año, Volumen_USDT
clients_all = data["clients"]       # Cliente, Periodo, Fecha, Mes, Año, Volumen_USD, Spread, Revenue, Takerate_pct
daily_all   = data["daily"]         # Fecha, País, Volumen_local

# Aplicar directorio de nombres (email → nombre)
_names = load_client_names()
if _names and "Cliente" in clients_all.columns:
    clients_all["Cliente"] = (
        clients_all["Cliente"]
        .str.lower()
        .map(lambda e: _names.get(e, e))
    )

if vol_country.empty:
    st.error("No se pudieron leer los datos del archivo. Verifica que sea la exportación correcta.")
    st.stop()

# ── Leer parámetros de URL (usados por el reporter para replicar la vista) ──────
_qp = st.query_params
_qp_year      = _qp.get("year", "")
_qp_view      = _qp.get("view", "")
_qp_mes       = _qp.get("mes", "")
_qp_trimestre = _qp.get("trimestre", "")
_qp_countries = _qp.get("countries", "")

# ── Period + Country selectors (need data first) ───────────────────────────────
with st.sidebar:
    st.divider()
    avail_years = sorted(vol_country["Año"].dropna().unique().astype(int), reverse=True)
    _year_idx   = avail_years.index(int(_qp_year)) if _qp_year and int(_qp_year) in avail_years else 0
    year = st.selectbox("Año", avail_years, index=_year_idx)

    _view_opts  = ["Mensual", "Trimestral", "YTD"]
    _view_idx   = _view_opts.index(_qp_view) if _qp_view in _view_opts else 0
    view = st.radio("Período", _view_opts, index=_view_idx)

    avail_months_nums = sorted(
        vol_country.loc[vol_country["Año"] == year, "Mes"].dropna().unique().astype(int)
    )
    avail_month_names = [MONTHS_ES[m - 1] for m in avail_months_nums]

    if view == "Mensual":
        _mes_idx    = avail_month_names.index(_qp_mes) if _qp_mes in avail_month_names else len(avail_month_names) - 1
        sel_month   = st.selectbox("Mes", avail_month_names, index=_mes_idx)
        curr_months = [MONTHS_ES.index(sel_month) + 1]
        period_label = f"{sel_month} {year}"
    elif view == "Trimestral":
        q_opts      = [q for q, ms in QUARTERS.items()
                       if any(MONTHS_ES.index(m) + 1 in avail_months_nums for m in ms)]
        _tri_idx    = q_opts.index(_qp_trimestre) if _qp_trimestre in q_opts else len(q_opts) - 1
        sel_q       = st.selectbox("Trimestre", q_opts, index=_tri_idx)
        curr_months = [MONTHS_ES.index(m) + 1 for m in QUARTERS[sel_q]]
        period_label = f"{sel_q} {year}"
    else:
        curr_months  = avail_months_nums
        period_label = f"YTD {year}"
        st.info(f"Mostrando {avail_month_names[0]}–{avail_month_names[-1]} {year}")

    prev_year   = year - 1
    prev_months = curr_months

    st.divider()
    all_countries    = sorted(vol_country["País"].unique())
    _default_countries = (
        [c for c in _qp_countries.split(",") if c in all_countries]
        if _qp_countries else all_countries
    )
    sel_countries = st.multiselect("Países", all_countries, default=_default_countries)


# ── Filter helpers ─────────────────────────────────────────────────────────────
def filt_vol(df, months, yr, countries=None):
    m = (df["Año"] == yr) & (df["Mes"].isin(months))
    if countries:
        m = m & (df["País"].isin(countries))
    return df[m].copy()

def filt_cli(df, months, yr):
    return df[(df["Año"] == yr) & (df["Mes"].isin(months))].copy()


vc_curr = filt_vol(vol_country, curr_months, year, sel_countries)
vc_prev = filt_vol(vol_country, prev_months, prev_year, sel_countries)
cl_curr = filt_cli(clients_all, curr_months, year)
cl_prev = filt_cli(clients_all, prev_months, prev_year)


# ── Global KPIs ────────────────────────────────────────────────────────────────
vol_c  = float(vc_curr["Volumen_USDT"].sum())
vol_p  = float(vc_prev["Volumen_USDT"].sum())
rev_c  = float(cl_curr["Revenue"].sum())
rev_p  = float(cl_prev["Revenue"].sum())
tk_c   = (rev_c / float(cl_curr["Volumen_USD"].sum()) * 100) if cl_curr["Volumen_USD"].sum() else 0
tk_p   = (rev_p / float(cl_prev["Volumen_USD"].sum()) * 100) if cl_prev["Volumen_USD"].sum() else 0
cli_c  = cl_curr["Cliente"].nunique()
cli_p  = cl_prev["Cliente"].nunique()
avg_c  = vol_c / cli_c if cli_c else 0
avg_p  = vol_p / cli_p if cli_p else 0


# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown(
    f"<div style='display:flex;align-items:center;gap:14px;margin-bottom:4px'>"
    f"{_logo_svg(KOYWE_GREEN, height=30)}"
    f"<span style='font-size:13px;color:{CHARCOAL};font-family:Raleway,sans-serif'>"
    f"Dashboard OTC &nbsp;·&nbsp; {period_label} &nbsp;·&nbsp; "
    f"{', '.join(sel_countries) if sel_countries else 'Todos los países'}"
    f"&nbsp;&nbsp;<span style='font-size:11px;color:#888;background:#f0f0f0;"
    f"border-radius:4px;padding:2px 8px'>Fuente: Greenhouse</span>"
    f"</span></div>",
    unsafe_allow_html=True,
)

WEEKLY_PATH = ""  # se sobreescribe si hay archivo subido

# ══════════════════════════════════════════════════════════════════════════════
tab_ov, tab_pais, tab_cli, tab_sem = st.tabs([
    "📊  Overview", "🌎  Por País", "👥  Clientes", "📅  Semanal"
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 · OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab_ov:
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Volumen (USDT)",   fmt_usd(vol_c),  pct_delta(vol_c, vol_p))
    c2.metric("Revenue",          fmt_usd(rev_c),  pct_delta(rev_c, rev_p))
    c3.metric("Takerate",         f"{tk_c:.3f}%",  f"{tk_c - tk_p:+.3f}pp" if tk_p else None)
    c4.metric("Clientes Activos", f"{cli_c:,}",    pct_delta(cli_c, cli_p))
    c5.metric("Vol/Cliente",      fmt_usd(avg_c),  pct_delta(avg_c, avg_p))

    st.markdown("---")

    col_a, col_b = st.columns(2)

    # Volume by country (bar horizontal)
    with col_a:
        if not vc_curr.empty:
            df_byc = vc_curr.groupby("País")["Volumen_USDT"].sum().reset_index().sort_values("Volumen_USDT")
            fig = px.bar(df_byc, x="Volumen_USDT", y="País", orientation="h",
                         title="Volumen por País (USDT)",
                         color="País", color_discrete_sequence=COUNTRY_PALETTE)
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_defaults(fig), use_container_width=True)

    # Revenue donut by country — approximate via vol share
    with col_b:
        if not vc_curr.empty:
            df_byc2 = vc_curr.groupby("País")["Volumen_USDT"].sum().reset_index()
            fig = px.pie(df_byc2, names="País", values="Volumen_USDT",
                         title="Distribución Volumen por País", hole=0.42,
                         color_discrete_sequence=COUNTRY_PALETTE)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(chart_defaults(fig, height=280), use_container_width=True)

    st.markdown("---")

    # Volumen por canal (Chart_14)
    CANAL_COLORS = {"Web OTC": GREEN_DARK, "K3": LAVENDER, "Manual": CHARCOAL}
    df_canal = load_canal_monthly(WEEKLY_PATH)

    if df_canal.empty:
        st.info("Sube el archivo de transacciones (Chart_14) en el panel izquierdo para ver datos por canal.")
    else:
        df_canal_view = df_canal[
            (df_canal["Año"] == year) & (df_canal["Mes"].isin(curr_months))
        ].copy()

        st.markdown("<div id='inicio-canal'></div>", unsafe_allow_html=True)

        st.markdown(
            "##### 🌐 Volumen por Canal — 🇨🇱 Chile",
            help=(
                "El archivo OTC Charts (Chart_14) solo contiene transacciones de Chile. "
                "El volumen aquí puede diferir del total del dashboard (que cubre todos los países). "
                "Además, Chart_14 usa el monto en CLP convertido a USDT via tipo de cambio, "
                "mientras que el KPI principal usa `amountInCrypto` (cripto efectivamente entregada)."
            ),
        )

        # KPIs por canal
        _canal_totals = df_canal_view.groupby("Canal")["Volumen_USDT"].sum()
        _ck1, _ck2, _ck3 = st.columns(3)
        _ck1.metric("🌐 Web OTC", fmt_usd(_canal_totals.get("Web OTC", 0)))
        _ck2.metric("⚙️ K3",     fmt_usd(_canal_totals.get("K3", 0)))
        _ck3.metric("✍️ Manual", fmt_usd(_canal_totals.get("Manual", 0)))

        st.markdown("")
        col_e, col_f = st.columns(2)
        with col_e:
            fig = px.bar(
                df_canal_view, x="Fecha", y="Volumen_USDT", color="Canal",
                title=f"Volumen por Canal (USDT) 🇨🇱 — {period_label}",
                color_discrete_map=CANAL_COLORS,
                barmode="group",
                text_auto=False,
            )
            fig.update_traces(hovertemplate="%{x|%b %Y}<br><b>$%{y:,.0f} USDT</b><extra></extra>")
            fig.update_xaxes(tickformat="%b %Y", dtick="M1")
            fig.update_yaxes(tickprefix="$", tickformat=",.0f")
            st.plotly_chart(chart_defaults(fig, height=280), use_container_width=True)

        with col_f:
            df_canal_tot = df_canal_view.groupby("Canal")["Volumen_USDT"].sum().reset_index()
            fig = px.pie(
                df_canal_tot, names="Canal", values="Volumen_USDT",
                title=f"Distribución por Canal 🇨🇱 — {period_label}",
                color="Canal", color_discrete_map=CANAL_COLORS, hole=0.42,
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(chart_defaults(fig, height=280), use_container_width=True)

        st.markdown("<div id='fin-canal'></div>", unsafe_allow_html=True)

        # ── Enviar a Slack ─────────────────────────────────────────────────────
        st.markdown("---")
        _ck1, _ck2, _ck3 = st.columns([3, 2, 3])
        with _ck2:
            _slack_channels_canal = {
                "#otc-finanzas-testing": "C0AM3K4DR6W",
                "#otc-finanzas":         "C08Q80QS2KK",
                "#comercial-data":       "C0AM3LQ9VDK",
            }
            _sel_ch_canal = st.selectbox("Canal destino", list(_slack_channels_canal.keys()),
                                         label_visibility="collapsed", key="slack_ch_canal")
            if st.button("📤 Enviar a Slack", use_container_width=True, key="btn_slack_canal"):
                _channel_id_canal = _slack_channels_canal[_sel_ch_canal]
                _reporter_canal   = os.path.join(os.path.dirname(__file__), "slack_reporter.py")
                _mes_actual_canal = sel_month if view == "Mensual" else ""
                _tri_actual_canal = sel_q if view == "Trimestral" else ""
                with st.spinner("Generando screenshot y enviando…"):
                    _proc_canal = subprocess.run(
                        ["python3", _reporter_canal,
                         "--file",       _tmp_chart,
                         "--channel",    _channel_id_canal,
                         "--tab",        "overview",
                         "--section",    "canal",
                         "--year",       str(year),
                         "--view",       view,
                         "--mes",        _mes_actual_canal,
                         "--trimestre",  _tri_actual_canal,
                         "--countries",  ",".join(sel_countries),
                        ],
                        capture_output=True, text=True, timeout=120,
                        env=_subprocess_env(),
                    )
                if _proc_canal.returncode == 0:
                    st.success(f"✅ Screenshot enviado a {_sel_ch_canal}")
                else:
                    st.error(f"❌ Error: {_proc_canal.stderr or _proc_canal.stdout}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 · POR PAÍS
# ══════════════════════════════════════════════════════════════════════════════
with tab_pais:
    if vc_curr.empty:
        st.info("No hay datos para el período seleccionado.")
    else:
        country_sel = st.selectbox("País", sel_countries if sel_countries else all_countries, key="pais_sel")

        vc_c = filt_vol(vol_country, curr_months, year, [country_sel])
        vc_p = filt_vol(vol_country, prev_months, prev_year, [country_sel])

        v_c = float(vc_c["Volumen_USDT"].sum())
        v_p = float(vc_p["Volumen_USDT"].sum())

        k1, k2, k3 = st.columns(3)
        k1.metric("Volumen (USDT)", fmt_usd(v_c), pct_delta(v_c, v_p))
        k2.metric("vs año anterior", fmt_usd(v_p) if v_p else "—")
        k3.metric("Variación", pct_delta(v_c, v_p) or "—")

        st.markdown("---")

        # Historical monthly trend for this country
        df_hist = (
            vol_country[vol_country["País"] == country_sel]
            .sort_values("Fecha")
        )
        fig = px.area(df_hist, x="Fecha", y="Volumen_USDT",
                      title=f"Volumen Mensual — {country_sel}",
                      color_discrete_sequence=[GREEN_DARK])
        fig.update_traces(fillcolor="rgba(25,80,44,0.15)", line_color=GREEN_DARK)
        st.plotly_chart(chart_defaults(fig, height=260), use_container_width=True)

        # Daily volume if available
        df_daily_c = daily_all[daily_all["País"] == country_sel].sort_values("Fecha")
        if not df_daily_c.empty:
            fig = px.bar(df_daily_c, x="Fecha", y="Volumen_local",
                         title=f"Volumen Diario — {country_sel} (moneda local)",
                         color_discrete_sequence=[LAVENDER])
            st.plotly_chart(chart_defaults(fig, height=220), use_container_width=True)

        # Monthly table
        df_tbl = df_hist[["Periodo", "Volumen_USDT"]].rename(
            columns={"Periodo": "Mes", "Volumen_USDT": "Volumen (USDT)"}
        ).sort_values("Mes", ascending=False)
        df_tbl["Volumen (USDT)"] = df_tbl["Volumen (USDT)"].apply(fmt_usd)
        st.dataframe(df_tbl, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 · CLIENTES
# ══════════════════════════════════════════════════════════════════════════════
with tab_cli:
    # ── Build retention timeline (se usa tanto en KPIs como en la sección de retención) ──
    monthly_stats = []
    _all_periods_ret = clients_all[["Año", "Mes", "Fecha"]].drop_duplicates().sort_values("Fecha")
    _all_clients_hist: set = set()
    _prev_month_clients: set = set()
    _month_sets: dict = {}   # {(yr, mo): set(clientes)} para reconstruir expanders

    for _, row in _all_periods_ret.iterrows():
        yr_r, mo_r = int(row["Año"]), int(row["Mes"])
        month_clients = set(
            clients_all.loc[(clients_all["Año"] == yr_r) & (clients_all["Mes"] == mo_r), "Cliente"]
        )
        new_m       = month_clients - _all_clients_hist
        recurring_m = month_clients & _prev_month_clients
        lost_m      = _prev_month_clients - month_clients
        react_m     = (month_clients & (_all_clients_hist - _prev_month_clients)) - new_m
        retention   = len(recurring_m) / len(_prev_month_clients) * 100 if _prev_month_clients else None
        monthly_stats.append({
            "Fecha":       row["Fecha"],
            "Periodo":     row.get("Periodo", f"{MONTHS_ES[mo_r-1]} {yr_r}"),
            "Año":         yr_r,
            "Mes":         mo_r,
            "Activos":     len(month_clients),
            "Nuevos":      len(new_m),
            "Recurrentes": len(recurring_m),
            "Reactivados": len(react_m),
            "Perdidos":    len(lost_m),
            "Retención %": round(retention, 1) if retention is not None else None,
            "_new_set":    new_m,
            "_lost_set":   lost_m,
        })
        _month_sets[(yr_r, mo_r)] = month_clients
        _all_clients_hist   |= month_clients
        _prev_month_clients  = month_clients

    df_ret = pd.DataFrame(monthly_stats).sort_values("Fecha")

    # ── Filtros de período de retención ───────────────────────────────────────
    _snap = st.query_params.get("snap", "0") == "1"
    _SNAP_MONTHS = 6

    if _snap:
        df_ret_view = df_ret.tail(_SNAP_MONTHS)
        _period_lbl = f"últ. {_SNAP_MONTHS} meses"
        _desde  = df_ret_view["Periodo"].iloc[0]
        _hasta  = df_ret_view["Periodo"].iloc[-1]
        _mes_tbl = "(Todo el período)"
    else:
        _periodos     = df_ret["Periodo"].tolist()
        _qp_ret_desde = _qp.get("ret_desde", "")
        _qp_ret_hasta = _qp.get("ret_hasta", "")
        _last_fecha   = df_ret["Fecha"].max()
        _desde_fecha  = _last_fecha - pd.DateOffset(months=11)
        _default_desde_idx = next(
            (i for i, f in enumerate(df_ret["Fecha"].tolist()) if f >= _desde_fecha), 0
        )
        if _qp_ret_desde and _qp_ret_desde in _periodos:
            _default_desde_idx = _periodos.index(_qp_ret_desde)

        _f1, _f2, _f3 = st.columns([2, 2, 4])
        with _f1:
            _desde = st.selectbox("Desde", _periodos, index=_default_desde_idx, key="ret_desde")
        with _f2:
            _hasta_opts        = _periodos[_periodos.index(_desde):]
            _default_hasta_idx = (
                _hasta_opts.index(_qp_ret_hasta)
                if _qp_ret_hasta and _qp_ret_hasta in _hasta_opts
                else len(_hasta_opts) - 1
            )
            _hasta = st.selectbox("Hasta", _hasta_opts, index=_default_hasta_idx, key="ret_hasta")

        _idx_desde  = _periodos.index(_desde)
        _idx_hasta  = _periodos.index(_hasta)
        df_ret_view = df_ret.iloc[_idx_desde : _idx_hasta + 1]
        _period_lbl = f"{_desde} – {_hasta}" if _desde != _hasta else _desde
        _mes_tbl    = "(Todo el período)"

    # ── Calcular _view_data: agregar mes a mes desde df_ret_view ─────────────
    _new_agg  = set().union(*df_ret_view["_new_set"].tolist())
    _lost_agg = set().union(*df_ret_view["_lost_set"].tolist()) - _new_agg
    _last_ret_row    = df_ret_view.iloc[-1]
    _first_ret_fecha = df_ret_view["Fecha"].iloc[0]
    _prev_ret_fecha  = _first_ret_fecha - pd.DateOffset(months=1)
    _view_data_full  = client_analysis_excel(
        clients_all,
        [int(_last_ret_row["Fecha"].month)], int(_last_ret_row["Fecha"].year),
        [int(_prev_ret_fecha.month)], int(_prev_ret_fecha.year),
    )
    _view_data = {
        "new":           _new_agg,
        "lost":          _lost_agg,
        "recurring":     _view_data_full["recurring"],
        "reactivated":   _view_data_full["reactivated"],
        "n_new":         len(_new_agg),
        "n_lost":        len(_lost_agg),
        "n_recurring":   int(df_ret_view["Recurrentes"].sum()),
        "n_reactivated": int(df_ret_view["Reactivados"].sum()),
        "client_df":     _view_data_full["client_df"],
    }
    _n_activos = int(_last_ret_row["Activos"])

    # ── KPI bar ───────────────────────────────────────────────────────────────
    cli_data = None
    if not cl_curr.empty:
        cli_data = _view_data   # alias para que el bloque de análisis siga funcionando

    if cli_data is None:
        st.info("No hay datos de clientes para el período seleccionado.")
    else:
        cc1, cc2, cc3, cc4, cc5 = st.columns(5)
        cc1.metric("Clientes Activos",  _n_activos)
        cc2.metric("🟢 Nuevos",         _view_data["n_new"])
        cc3.metric("🔵 Recurrentes",    _view_data["n_recurring"])
        cc4.metric("🟡 Reactivados",    _view_data["n_reactivated"])
        cc5.metric("🔴 Perdidos",       _view_data["n_lost"],
                   delta=f"-{_view_data['n_lost']}" if _view_data["n_lost"] else None,
                   delta_color="inverse")

    st.markdown("---")
    st.subheader("📈 Retención")

    latest       = df_ret_view.iloc[-1] if not df_ret_view.empty else {}
    avg_ret_view = df_ret_view["Retención %"].dropna().mean()

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Retención promedio", f"{avg_ret_view:.1f}%" if avg_ret_view else "—")
    r2.metric("Clientes activos (últ. mes)", f"{int(latest.get('Activos', 0)):,}")
    r3.metric("Churn promedio", f"{100 - avg_ret_view:.1f}%" if avg_ret_view else "—")
    r4.metric("Peak período", f"{int(df_ret_view['Activos'].max()):,}" if not df_ret_view.empty else "—")

    st.markdown("---")

    col_a, col_b = st.columns(2)
    with col_a:
        df_ret_plot = df_ret_view.dropna(subset=["Retención %"])
        fig = px.line(df_ret_plot, x="Fecha", y="Retención %",
                      title=f"Tasa de Retención — {_period_lbl}",
                      color_discrete_sequence=[LIMA])
        fig.update_traces(line_width=2.5,
                          hovertemplate="%{x|%b %Y}<br><b>%{y:.1f}%</b><extra></extra>")
        fig.add_hline(y=avg_ret_view, line_dash="dot", line_color=CHARCOAL,
                      annotation_text=f"Promedio {avg_ret_view:.1f}%")
        fig.update_yaxes(ticksuffix="%", range=[0, 105])
        st.plotly_chart(chart_defaults(fig, height=280), use_container_width=True)

    with col_b:
        df_stack = df_ret_view.copy()
        df_stack["Perdidos_neg"] = -df_stack["Perdidos"]
        fig = go.Figure()
        fig.add_bar(x=df_stack["Fecha"], y=df_stack["Nuevos"],      name="Nuevos",      marker_color=LIMA)
        fig.add_bar(x=df_stack["Fecha"], y=df_stack["Reactivados"],  name="Reactivados", marker_color=LAVENDER)
        fig.add_bar(x=df_stack["Fecha"], y=df_stack["Recurrentes"],  name="Recurrentes", marker_color=GREEN_DARK)
        fig.add_bar(x=df_stack["Fecha"], y=df_stack["Perdidos_neg"], name="Perdidos",    marker_color=DANGER)
        fig.update_layout(barmode="relative", title=f"Movimiento de Clientes — {_period_lbl}")
        st.plotly_chart(chart_defaults(fig, height=280), use_container_width=True)

    st.subheader(_period_lbl)
    df_tbl_ret = df_ret_view.sort_values("Fecha", ascending=False)[
        ["Periodo","Activos","Nuevos","Recurrentes","Reactivados","Perdidos","Retención %"]
    ].copy()
    df_tbl_ret["Retención %"] = df_tbl_ret["Retención %"].apply(
        lambda x: f"{x:.1f}%" if pd.notna(x) else "—"
    )
    st.dataframe(df_tbl_ret.reset_index(drop=True), use_container_width=True, hide_index=True)

    # ── Análisis detallado del período ────────────────────────────────────────
    if cli_data is not None:
        st.markdown("---")
        st.markdown("<div id='seccion-analisis'></div>", unsafe_allow_html=True)
        df_cd = _view_data["client_df"].copy()

        col_i, col_j = st.columns(2)
        with col_i:
            tier_sum = (
                df_cd.groupby("Tier")
                .agg(Clientes=("Cliente", "count"), Volumen=("Volumen_USD", "sum"))
                .reset_index()
            )
            tier_sum["Volumen_fmt"] = tier_sum["Volumen"].apply(fmt_usd)
            fig = px.bar(tier_sum, x="Tier", y="Volumen",
                         title="Volumen por Tier (USD)",
                         color="Tier", color_discrete_map=TIER_COLORS,
                         text="Volumen_fmt")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_defaults(fig, height=360), use_container_width=True)

        with col_j:
            fig = px.pie(tier_sum, names="Tier", values="Clientes",
                         title="Clientes por Tier",
                         color="Tier", color_discrete_map=TIER_COLORS, hole=0.4)
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(chart_defaults(fig, height=360), use_container_width=True)

        # Top 10 bar
        top10 = df_cd.nlargest(10, "Volumen_USD")
        fig = px.bar(top10.sort_values("Volumen_USD"),
                     x="Volumen_USD", y="Cliente", orientation="h",
                     title="Top 10 Clientes por Volumen",
                     color="Tier", color_discrete_map=TIER_COLORS)
        st.plotly_chart(chart_defaults(fig, height=320), use_container_width=True)

        st.markdown("<div id='fin-graficas'></div>", unsafe_allow_html=True)
        st.markdown("---")

        # Filters
        f1, f2, f3 = st.columns([2, 2, 3])
        with f1:
            tier_sel = st.multiselect("Tier", ["Enterprise","Corporate","Business","Retail"],
                                      default=["Enterprise","Corporate","Business","Retail"])
        with f2:
            estado_sel = st.multiselect("Estado", ["Nuevo","Recurrente","Reactivado","Activo"],
                                         default=["Nuevo","Recurrente","Reactivado","Activo"])
        with f3:
            search_cli = st.text_input("Buscar cliente", placeholder="Email o nombre…")

        df_show = df_cd.copy()
        if tier_sel:   df_show = df_show[df_show["Tier"].isin(tier_sel)]
        if estado_sel: df_show = df_show[df_show["Estado"].isin(estado_sel)]
        if search_cli: df_show = df_show[df_show["Cliente"].str.contains(search_cli, case=False, na=False)]

        df_disp = df_show.copy()
        df_disp["Volumen_USD"] = df_disp["Volumen_USD"].apply(fmt_usd)
        df_disp["Revenue"]     = df_disp["Revenue"].apply(fmt_usd)
        df_disp["Takerate_%"]  = df_disp["Takerate_%"].apply(lambda x: f"{x:.4f}%")
        df_disp = df_disp.rename(columns={
            "Volumen_USD": "Volumen (USD)", "Revenue": "Revenue (USD)",
            "Takerate_%": "Takerate %", "Meses": "Meses activo"
        })
        st.dataframe(
            df_disp[["Cliente","Tier","Estado","Volumen (USD)","Revenue (USD)","Takerate %","Meses activo"]],
            use_container_width=True, hide_index=True
        )

        # Lost / new expanders
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            if _view_data["lost"]:
                with st.expander(f"🔴 {_view_data['n_lost']} clientes perdidos"):
                    for c in sorted(_view_data["lost"]): st.markdown(f"- {c}")
        with col_exp2:
            if _view_data["new"]:
                with st.expander(f"🟢 {_view_data['n_new']} clientes nuevos"):
                    for c in sorted(_view_data["new"]): st.markdown(f"- {c}")

        # ── Enviar análisis a Slack ────────────────────────────────────────────
        st.markdown("---")
        _sa1, _sa2, _sa3 = st.columns([3, 2, 3])
        with _sa2:
            _slack_channels2 = {
                "#otc-finanzas-testing": "C0AM3K4DR6W",
                "#otc-finanzas":         "C08Q80QS2KK",
                "#comercial-data":       "C0AM3LQ9VDK",
            }
            _sel_channel2 = st.selectbox("Canal destino", list(_slack_channels2.keys()),
                                         label_visibility="collapsed", key="slack_ch_analisis")
            if st.button("📤 Enviar análisis a Slack", use_container_width=True):
                _channel_id2 = _slack_channels2[_sel_channel2]
                _reporter2   = os.path.join(os.path.dirname(__file__), "slack_reporter.py")
                with st.spinner("Generando screenshot y enviando…"):
                    _proc2 = subprocess.run(
                        ["python3", _reporter2,
                         "--file", _tmp_chart, "--channel", _channel_id2,
                         "--tab", "clientes", "--section", "detalle"],
                        capture_output=True, text=True, timeout=120,
                        env=_subprocess_env(),
                    )
                if _proc2.returncode == 0:
                    st.success(f"✅ Screenshot enviado a {_sel_channel2}")
                else:
                    st.error(f"❌ Error al enviar: {_proc2.stderr or _proc2.stdout}")

    # ── Enviar retención a Slack ───────────────────────────────────────────────
    st.markdown("---")
    _sc1, _sc2, _sc3 = st.columns([3, 2, 3])
    with _sc2:
        _slack_channels = {
            "#otc-finanzas-testing": "C0AM3K4DR6W",
            "#otc-finanzas":         "C08Q80QS2KK",
            "#comercial-data":       "C0AM3LQ9VDK",
        }
        _sel_channel = st.selectbox("Canal destino", list(_slack_channels.keys()),
                                    label_visibility="collapsed", key="slack_ch_retencion")
        if st.button("📤 Enviar retención a Slack", use_container_width=True):
            _channel_id = _slack_channels[_sel_channel]
            _reporter   = os.path.join(os.path.dirname(__file__), "slack_reporter.py")
            _mes_actual = sel_month if view == "Mensual" else ""
            _tri_actual = sel_q if view == "Trimestral" else ""
            with st.spinner("Generando screenshot y enviando…"):
                _proc = subprocess.run(
                    ["python3", _reporter,
                     "--file",       _tmp_chart,
                     "--channel",    _channel_id,
                     "--tab",        "clientes",
                     "--section",    "retencion",
                     "--year",       str(year),
                     "--view",       view,
                     "--mes",        _mes_actual,
                     "--trimestre",  _tri_actual,
                     "--countries",  ",".join(sel_countries),
                     "--ret-desde",  _desde,
                     "--ret-hasta",  _hasta,
                    ],
                    capture_output=True, text=True, timeout=120,
                    env=_subprocess_env(),
                )
            if _proc.returncode == 0:
                st.success(f"✅ Screenshot enviado a {_sel_channel}")
            else:
                st.error(f"❌ Error al enviar: {_proc.stderr or _proc.stdout}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 · SEMANAL
# ══════════════════════════════════════════════════════════════════════════════
with tab_sem:
    st.subheader(f"Resumen Semanal · {year}")

    df_w = load_weekly_summary(WEEKLY_PATH, year)

    if df_w.empty:
        st.info("Sube el archivo de transacciones (Chart_14) en el panel izquierdo para ver el resumen semanal.")
    else:
        # ── KPIs rápidos ──────────────────────────────────────────────────────
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Volumen total", fmt_usd(df_w["Volumen_USDT"].sum()))
        k2.metric("Operaciones",   f"{int(df_w['Operaciones'].sum()):,}")
        k3.metric("Revenue total", fmt_usd(df_w["Revenue_USD"].sum()))
        k4.metric("Semanas activas", len(df_w))

        st.divider()

        # ── Gráfico volumen semanal ───────────────────────────────────────────
        fig_vol = px.bar(
            df_w, x="Semana", y="Volumen_USDT",
            title="Volumen USDT por Semana",
            color_discrete_sequence=[GREEN_DARK],
        )
        for t in fig_vol.data:
            t.hovertemplate = "%{x}<br><b>$%{y:,.0f} USDT</b><extra></extra>"
        st.plotly_chart(chart_defaults(fig_vol, height=280), use_container_width=True)

        col_a, col_b = st.columns(2)

        # ── Gráfico revenue semanal ───────────────────────────────────────────
        with col_a:
            fig_rev = px.bar(
                df_w, x="Semana", y="Revenue_USD",
                title="Revenue USD por Semana",
                color_discrete_sequence=[GREEN_MID],
            )
            for t in fig_rev.data:
                t.hovertemplate = "%{x}<br><b>$%{y:,.2f}</b><extra></extra>"
            st.plotly_chart(chart_defaults(fig_rev, height=260), use_container_width=True)

        # ── Gráfico operaciones y clientes ────────────────────────────────────
        with col_b:
            fig_ops = go.Figure()
            fig_ops.add_bar(x=df_w["Semana"], y=df_w["Operaciones"],
                            name="Operaciones", marker_color=GREEN_DARK)
            fig_ops.add_scatter(x=df_w["Semana"], y=df_w["Clientes"],
                                name="Clientes únicos", mode="lines+markers",
                                line=dict(color=LIMA, width=2),
                                yaxis="y2")
            fig_ops.update_layout(
                title="Operaciones y Clientes por Semana",
                yaxis2=dict(overlaying="y", side="right", showgrid=False),
                legend=dict(orientation="h", y=1.1),
            )
            st.plotly_chart(chart_defaults(fig_ops, height=260), use_container_width=True)

        st.divider()

        # ── Tabla detalle ─────────────────────────────────────────────────────
        st.subheader("Detalle por Semana")
        df_disp = df_w[["Semana", "Volumen_USDT", "Operaciones",
                         "Revenue_USD", "Clientes", "Top Canal", "Top País"]].copy()
        df_disp["Volumen_USDT"] = df_disp["Volumen_USDT"].apply(lambda v: f"${v:,.0f}")
        df_disp["Revenue_USD"]  = df_disp["Revenue_USD"].apply(lambda v: f"${v:,.2f}")
        df_disp["Operaciones"]  = df_disp["Operaciones"].astype(int)
        df_disp.columns         = ["Semana", "Volumen USDT", "Ops",
                                    "Revenue USD", "Clientes", "Canal líder", "País líder"]
        st.dataframe(df_disp.reset_index(drop=True), use_container_width=True, hide_index=True)
