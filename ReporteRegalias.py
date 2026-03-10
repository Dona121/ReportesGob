import streamlit as st
import polars as pl
import pandas as pd
import io
from datetime import date

st.set_page_config(
    page_title="Seguimiento Regalías",
    layout="wide",
    initial_sidebar_state="expanded",
)

C = {
    "azul_oscuro":  "#003d6c",
    "azul_medio":   "#1754ab",
    "verde_oscuro": "#005931",
    "verde_medio":  "#17743d",
    "cian":         "#47b1d5",
    "naranja":      "#d88c16",
    "naranja_osc":  "#cf7000",
    "cafe":         "#9b5b1e",
    "salmon":       "#e68878",
    "bg":           "#e8edf5",
    "white":        "#ffffff",
    "text":         "#1a2332",
    "muted":        "#6b7280",
    "border":       "#e2e8f0",
}

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'Montserrat', sans-serif;
    color: {C['text']};
}}

/* ── Fondo del área principal con gradiente institucional ── */
.stApp {{
    background:
        radial-gradient(ellipse at 0% 0%, rgba(0,61,108,0.10) 0%, transparent 55%),
        radial-gradient(ellipse at 100% 100%, rgba(0,89,49,0.08) 0%, transparent 55%),
        linear-gradient(160deg, #dde5f0 0%, #e8edf5 40%, #eaf0f0 100%);
    min-height: 100vh;
}}

/* Controlar padding del contenedor principal de Streamlit */
.block-container {{
    padding-top: 3.5rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    padding-bottom: 3rem !important;
    max-width: 100% !important;
}}

/* ── Sidebar oscuro ── */
section[data-testid="stSidebar"] > div {{
    background: #001f3f;
    padding-top: 1.5rem;
}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span:not([data-baseweb="tag"] span),
section[data-testid="stSidebar"] div {{
    color: rgba(255,255,255,0.95) !important;
}}
/* File uploader — caja blanca con texto oscuro */
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
    background: rgba(255,255,255,0.08) !important;
    border: 1.5px dashed rgba(255,255,255,0.25) !important;
    border-radius: 8px !important;
}}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section *{{
    color: rgba(255,255,255,0.85) !important;
}}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {{
    background: rgba(255,255,255,0.12) !important;
    color: white !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    border-radius: 6px !important;
}}
/* Tags del multiselect */
span[data-baseweb="tag"] {{
    background: {C['azul_medio']} !important;
    color: white !important;
    max-width: 160px !important;
    border-radius: 4px !important;
}}
span[data-baseweb="tag"] span {{
    color: white !important;
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    overflow: hidden !important;
    white-space: nowrap !important;
    text-overflow: ellipsis !important;
    max-width: 130px !important;
    display: block !important;
}}
/* Contenedor multiselect en sidebar */
section[data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {{
    flex-wrap: wrap !important;
    gap: 4px !important;
    padding: 6px 8px !important;
    background: rgba(255,255,255,0.07) !important;
    border-color: rgba(255,255,255,0.18) !important;
    border-radius: 6px !important;
}}
/* Selectbox valor visible */
section[data-testid="stSidebar"] [data-baseweb="select"] [class*="singleValue"],
section[data-testid="stSidebar"] [data-baseweb="select"] [class*="placeholder"] {{
    color: white !important;
    font-weight: 500 !important;
}}
/* Labels de filtros */
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] .stSelectbox label {{
    color: rgba(255,255,255,0.6) !important;
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
    margin-bottom: 0.2rem !important;
}}
/* Sección título sidebar */
.sidebar-section {{
    font-family: 'Montserrat', sans-serif;
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: {C['cian']} !important;
    margin: 1.5rem 0 0.7rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid rgba(71,177,213,0.3);
}}

/* ── Header ── */
.page-header {{
    background: linear-gradient(120deg, {C['azul_oscuro']} 0%, {C['verde_oscuro']} 100%);
    border-radius: 12px;
    margin: 0 0 0.8rem 0;
    padding: 1.8rem 2rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 4px 24px rgba(0,61,108,0.22);
    position: relative;
    z-index: 1;
}}
.page-header h1 {{
    font-family: 'Montserrat', sans-serif;
    font-size: 1.6rem;
    font-weight: 800;
    color: white;
    margin: 0 0 0.15rem 0;
}}
.page-header p {{
    color: rgba(255,255,255,0.6);
    margin: 0;
    font-size: 0.82rem;
    font-weight: 400;
}}

/* ── Filtros horizontales ── */
.filter-bar {{
    background: {C['white']};
    border-radius: 10px;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    display: flex;
    align-items: center;
    gap: 1rem;
    flex-wrap: wrap;
}}
.filter-label {{
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: {C['muted']};
    white-space: nowrap;
}}

/* ── KPI principal ── */
.kpi-main {{
    background: {C['azul_oscuro']};
    border-radius: 12px;
    padding: 1.4rem 1.6rem;
    color: white;
    height: 100%;
    box-shadow: 0 4px 20px rgba(0,61,108,0.22), 0 1px 4px rgba(0,0,0,0.1);
}}
.kpi-main .label {{
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: rgba(255,255,255,0.6);
    margin-bottom: 0.4rem;
}}
.kpi-main .value {{
    font-family: 'Montserrat', sans-serif;
    font-size: 2.8rem;
    font-weight: 800;
    line-height: 1;
    color: white;
}}
.kpi-main .sub {{
    font-size: 0.78rem;
    color: rgba(255,255,255,0.5);
    margin-top: 0.2rem;
}}

/* ── KPI secundario ── */
.kpi-sec {{
    background: #ffffff;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    border-left: 4px solid;
    box-shadow: 0 3px 16px rgba(0,40,90,0.13), 0 1px 3px rgba(0,0,0,0.07);
    height: 100%;
}}
.kpi-sec .label {{
    font-size: 0.63rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: {C['muted']};
    margin-bottom: 0.3rem;
}}
.kpi-sec .value {{
    font-family: 'Montserrat', sans-serif;
    font-size: 1.6rem;
    font-weight: 800;
    line-height: 1;
}}
.kpi-sec .sub {{
    font-size: 0.7rem;
    color: {C['muted']};
    margin-top: 0.15rem;
}}

/* ── Section heading ── */
.section-heading {{
    font-family: 'Montserrat', sans-serif;
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: {C['azul_oscuro']};
    border-bottom: 2px solid {C['cian']};
    padding-bottom: 0.45rem;
    margin: 0 0 1rem 0;
}}

/* ── Tabs ── */
div[data-testid="stTabs"] [role="tablist"] {{
    background: rgba(255,255,255,0.70);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    border-radius: 10px 10px 0 0;
    padding: 0 1rem;
    border-bottom: 2px solid {C['border']};
    gap: 0;
}}
div[data-testid="stTabs"] [role="tab"] {{
    font-family: 'Montserrat', sans-serif !important;
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    color: {C['muted']} !important;
    padding: 0.75rem 1.2rem !important;
    border-bottom: 2px solid transparent !important;
    margin-bottom: -2px !important;
}}
div[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: {C['azul_oscuro']} !important;
    border-bottom-color: {C['cian']} !important;
}}
div[data-testid="stTabs"] [data-testid="stTabsContent"] {{
    background: #ffffff;
    border-radius: 0 0 12px 12px;
    padding: 1.4rem;
    box-shadow: 0 4px 24px rgba(0,40,90,0.10), 0 1px 4px rgba(0,0,0,0.06);
}}

/* ── Summary table ── */
.summary-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.83rem;
    background: #ffffff;
    border-radius: 8px;
    overflow: hidden;
}}
.summary-table thead tr {{
    background: {C['azul_oscuro']};
    color: white;
}}
.summary-table th {{
    padding: 0.7rem 0.9rem;
    font-family: 'Montserrat', sans-serif;
    font-size: 0.64rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    text-align: left;
    white-space: nowrap;
    line-height: 1.4;
}}
.summary-table td {{
    padding: 0.65rem 0.9rem;
    border-bottom: 1px solid {C['border']};
    vertical-align: middle;
    background: #ffffff;
}}
.summary-table tbody tr:nth-child(even) td {{ background: #f7fafd; }}
.summary-table tbody tr:last-child td {{ border-bottom: none; }}
.summary-table tbody tr:hover td {{ background: #e8f3ff !important; transition: background 0.15s; }}
/* Columna Total destacada */
.summary-table td:last-child,
.summary-table th:last-child {{
    background: rgba(0,61,108,0.05);
    font-weight: 800;
    color: {C['azul_oscuro']};
    border-left: 2px solid {C['border']};
    text-align: center;
}}
.summary-table tbody tr:hover td:last-child {{ background: #d8ecff !important; }}
.entidad-name {{
    font-weight: 600;
    font-size: 0.83rem;
    color: {C['azul_oscuro']};
}}
.dias-val {{
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    font-weight: 500;
}}
.null-cell {{ color: {C['border']}; }}

/* ── Badges ── */
.badge {{
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 3px 9px;
    border-radius: 20px;
    font-size: 0.68rem;
    font-weight: 700;
    margin-left: 5px;
    vertical-align: middle;
    cursor: help;
    position: relative;
}}
.badge-green  {{ background: #d1fae5; color: #065f46; }}
.badge-yellow {{ background: #fef3c7; color: #92400e; }}
.badge-orange {{ background: #ffedd5; color: #9a3412; }}
.badge-red    {{ background: #fee2e2; color: #991b1b; }}
/* Punto de color semáforo */
.badge-dot {{
    width: 8px; height: 8px;
    border-radius: 50%;
    flex-shrink: 0;
    display: inline-block;
}}
.badge-green  .badge-dot {{ background: #059669; }}
.badge-yellow .badge-dot {{ background: #d97706; }}
.badge-orange .badge-dot {{ background: #ea580c; }}
.badge-red    .badge-dot {{ background: #dc2626; }}
/* Tooltip del semáforo */
.badge-tooltip {{
    visibility: hidden;
    opacity: 0;
    position: absolute;
    bottom: calc(100% + 8px);
    left: 50%;
    transform: translateX(-50%);
    background: {C['text']};
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 0.72rem;
    font-weight: 400;
    line-height: 1.5;
    padding: 0.55rem 0.8rem;
    border-radius: 8px;
    width: 240px;
    text-align: left;
    text-transform: none;
    letter-spacing: 0;
    z-index: 9999;
    box-shadow: 0 4px 16px rgba(0,0,0,0.25);
    pointer-events: none;
    transition: opacity 0.15s ease;
    white-space: normal;
}}
.badge-tooltip::after {{
    content: '';
    position: absolute;
    top: 100%;
    left: 50%;
    transform: translateX(-50%);
    border: 5px solid transparent;
    border-top-color: {C['text']};
}}
.badge:hover .badge-tooltip {{ visibility: visible; opacity: 1; }}

/* ── Evaluación — calificación card ── */
.eval-card {{
    background: #ffffff;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    box-shadow: 0 2px 12px rgba(0,40,90,0.09);
    margin-bottom: 0.6rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}}
.eval-bar-wrap {{
    flex: 1;
}}
.eval-label {{
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: {C['muted']};
    margin-bottom: 0.3rem;
}}
.eval-bar-bg {{
    background: {C['border']};
    border-radius: 6px;
    height: 10px;
    overflow: hidden;
    margin-bottom: 0.2rem;
}}
.eval-bar-fill {{
    height: 100%;
    border-radius: 6px;
    transition: width 0.4s ease;
}}
.eval-score {{
    font-family: 'DM Mono', monospace;
    font-size: 1.1rem;
    font-weight: 700;
    min-width: 3rem;
    text-align: right;
    color: {C['azul_oscuro']};
}}

/* ── Detail table ── */
.detail-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.81rem;
    margin-top: 0.4rem;
}}
.detail-table th {{
    background: #f1f5f9;
    color: {C['azul_oscuro']};
    font-family: 'Montserrat', sans-serif;
    font-size: 0.63rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 0.5rem 0.85rem;
    text-align: left;
    border-bottom: 2px solid {C['border']};
}}
.detail-table td {{
    padding: 0.55rem 0.85rem;
    border-bottom: 1px solid {C['border']};
    vertical-align: middle;
}}
.detail-table tbody tr:last-child td {{ border-bottom: none; }}
.detail-table tbody tr:hover td {{ background: #f8fbff; }}
.bpin-tag {{
    font-family: 'DM Mono', monospace;
    font-size: 0.7rem;
    background: #f1f5f9;
    color: {C['muted']};
    padding: 2px 6px;
    border-radius: 4px;
}}
.estado-tag {{
    font-size: 0.7rem;
    background: #eff6ff;
    color: {C['azul_medio']};
    padding: 2px 7px;
    border-radius: 4px;
    font-weight: 500;
    white-space: nowrap;
}}

/* ── Expander ── */
div[data-testid="stExpander"] > details {{
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 2px 10px rgba(0,40,90,0.08) !important;
    margin-bottom: 0.5rem !important;
    border: none !important;
}}
div[data-testid="stExpander"] > details > summary {{
    background: #ffffff !important;
    border: none !important;
    border-left: 4px solid {C['azul_oscuro']} !important;
    border-radius: 10px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.84rem !important;
    color: {C['azul_oscuro']} !important;
    padding: 0.75rem 1.1rem !important;
    transition: background 0.15s, border-color 0.15s !important;
}}
div[data-testid="stExpander"] > details > summary:hover {{
    background: #f0f7ff !important;
    border-left-color: {C['cian']} !important;
}}
div[data-testid="stExpander"] > details[open] > summary {{
    border-radius: 10px 10px 0 0 !important;
    border-left-color: {C['cian']} !important;
    background: #f8fbff !important;
}}
div[data-testid="stExpander"] > details > div {{
    border: none !important;
    border-left: 4px solid {C['cian']} !important;
    border-radius: 0 0 10px 10px !important;
    background: #ffffff !important;
    padding: 0.2rem 1rem 0.8rem 1rem !important;
}}

/* ── Tooltip ── */
.th-wrap {{
    display: inline-flex;
    align-items: center;
    gap: 4px;
    cursor: default;
    position: relative;
}}
.th-icon {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 14px;
    height: 14px;
    border-radius: 50%;
    background: rgba(255,255,255,0.2);
    color: white;
    font-size: 0.6rem;
    font-weight: 700;
    cursor: help;
    line-height: 1;
    flex-shrink: 0;
}}
.th-tooltip {{
    visibility: hidden;
    opacity: 0;
    position: absolute;
    top: calc(100% + 8px);
    left: 50%;
    transform: translateX(-50%);
    background: {C['text']};
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 0.73rem;
    font-weight: 400;
    line-height: 1.55;
    padding: 0.6rem 0.85rem;
    border-radius: 8px;
    white-space: normal;
    width: 230px;
    text-align: left;
    text-transform: none;
    letter-spacing: 0;
    z-index: 9999;
    box-shadow: 0 4px 18px rgba(0,0,0,0.28);
    pointer-events: none;
    transition: opacity 0.15s ease;
}}
.th-tooltip strong {{
    display: block;
    margin-bottom: 4px;
    font-size: 0.75rem;
    color: {C['cian']};
}}
.th-tooltip::before {{
    content: '';
    position: absolute;
    top: -6px;
    left: 50%;
    transform: translateX(-50%);
    border: 6px solid transparent;
    border-top: none;
    border-bottom-color: {C['text']};
}}
.th-wrap:hover .th-tooltip {{ visibility: visible; opacity: 1; }}

/* ── Download button ── */
.stDownloadButton > button {{
    background: {C['verde_oscuro']} !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'Montserrat', sans-serif !important;
    font-size: 0.83rem !important;
    padding: 0.5rem 1.4rem !important;
}}
.stDownloadButton > button:hover {{ background: {C['verde_medio']} !important; }}

/* ── Multiselect / selectbox ── */
span[data-baseweb="tag"] {{
    background: {C['azul_medio']} !important;
    color: white !important;
    border-radius: 4px !important;
}}
span[data-baseweb="tag"] span {{ color: white !important; font-size: 0.75rem !important; }}
/* ── Tarjeta de error ── */
.error-card {{
    background: #fff5f5;
    border: 1.5px solid #fca5a5;
    border-left: 5px solid #dc2626;
    border-radius: 10px;
    padding: 1.2rem 1.5rem;
    margin: 1rem 0;
}}
.error-card .error-title {{
    font-family: 'Montserrat', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    color: #991b1b;
    margin-bottom: 0.4rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}}
.error-card .error-body {{
    font-size: 0.83rem;
    color: #7f1d1d;
    line-height: 1.6;
    margin-bottom: 0.8rem;
}}
.error-card .error-fix {{
    background: #fef2f2;
    border-radius: 6px;
    padding: 0.6rem 0.9rem;
    font-size: 0.8rem;
    color: #991b1b;
    border: 1px solid #fca5a5;
}}
.error-card .error-fix strong {{
    display: block;
    margin-bottom: 0.3rem;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #dc2626;
}}
.error-cols {{
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin-top: 0.5rem;
}}
.col-missing {{
    background: #fee2e2;
    color: #991b1b;
    font-family: 'Montserrat', monospace;
    font-size: 0.73rem;
    font-weight: 700;
    padding: 3px 10px;
    border-radius: 4px;
    border: 1px solid #fca5a5;
}}
.col-wrong-type {{
    background: #fff7ed;
    color: #9a3412;
    font-family: 'Montserrat', monospace;
    font-size: 0.73rem;
    font-weight: 700;
    padding: 3px 10px;
    border-radius: 4px;
    border: 1px solid #fed7aa;
}}
/* Tabla de referencia de columnas */
.ref-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.78rem;
    margin-top: 0.8rem;
    border-radius: 8px;
    overflow: hidden;
}}
.ref-table th {{
    background: #f1f5f9;
    color: {C['azul_oscuro']};
    font-weight: 700;
    font-size: 0.68rem;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 0.45rem 0.75rem;
    text-align: left;
    border-bottom: 2px solid {C['border']};
}}
.ref-table td {{
    padding: 0.4rem 0.75rem;
    border-bottom: 1px solid {C['border']};
    color: {C['text']};
}}
.ref-table tr:last-child td {{ border-bottom: none; }}
.ref-table code {{
    background: #f1f5f9;
    padding: 1px 6px;
    border-radius: 3px;
    font-size: 0.72rem;
    color: {C['azul_medio']};
}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES DE VALIDACIÓN — deben ir ANTES de procesar()
# ─────────────────────────────────────────────────────────────────────────────
TABLA_ESPERADA = "MatrizSeguimientoEvaluacion"

COLUMNAS_ESPERADAS = {
    "ENTIDAD O SECRETARIA":                 ("texto",  [pl.Utf8, pl.String]),
    "BPIN":                                 ("texto",  [pl.Utf8, pl.String]),
    "NOMBRE PROYECTO":                      ("texto",  [pl.Utf8, pl.String]),
    "ESTADO PROYECTO":                      ("texto",  [pl.Utf8, pl.String]),
    "ESTADO CONTRATO":                      ("texto",  [pl.Utf8, pl.String]),
    "CPI":                                  ("número", [pl.Float32, pl.Float64, pl.Int32, pl.Int64]),
    "SPI":                                  ("número", [pl.Float32, pl.Float64, pl.Int32, pl.Int64]),
    "FECHA APROBACIÓN PROYECTO":            ("fecha",  [pl.Date, pl.Datetime]),
    "FECHA DE APERTURA DEL PRIMER PROCESO": ("fecha",  [pl.Date, pl.Datetime]),
    "FECHA SUSCRIPCION":                    ("fecha",  [pl.Date, pl.Datetime]),
    "FECHA ACTA INICIO":                    ("fecha",  [pl.Date, pl.Datetime]),
    "HORIZONTE DEL PROYECTO":               ("fecha",  [pl.Date, pl.Datetime]),
    "FECHA DE FINALIZACIÓN":                ("fecha",  [pl.Date, pl.Datetime]),
    "FECHA DE CORTE GESPROY":               ("fecha",  [pl.Date, pl.Datetime]),
}

TIPO_LABEL = {
    "texto":  "Texto",
    "número": "Número decimal",
    "fecha":  "Fecha",
}

TIPO_EJEMPLO = {
    "texto":  "Ej: «Infraestructura», «SIN CONTRATAR»",
    "número": "Ej: 0, 1.5, 0.87  (sin letras ni símbolos)",
    "fecha":  "Ej: 15/03/2024  (formato fecha de Excel)",
}

# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA
# ─────────────────────────────────────────────────────────────────────────────
INTERVALOS = {
    "hito_1_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
    "hito_2_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
    "hito_3_val": [("0-30",  0,  30), ("31-45",   31,  45), ("46-60",   46,  60),  (">60",  61, None)],
    "hito_5_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
}

# Semáforos: colores y mensajes por hito e intervalo
SEMAFOROS = {
    "hito_1_val": {
        "0-100":   ("green",  "Verde",   "Proyecto dentro de los tiempos para su primera apertura del proceso de contratación."),
        "101-150": ("yellow", "Naranja", "Proyecto en alerta: más de 100 días sin apertura del primer proceso precontractual."),
        "151-180": ("orange", "Rojo",    "Proyecto en alerta roja: más de 150 días sin apertura del primer proceso precontractual."),
        ">180":    ("red",    "Negro",   "Proyecto en alerta negra: más de 180 días sin apertura del primer proceso precontractual."),
    },
    "hito_2_val": {
        "0-100":   ("green",  "Verde",   "Proyecto dentro de los tiempos para la firma del primer contrato."),
        "101-150": ("yellow", "Naranja", "Proyecto en alerta: más de 100 días sin firma del primer contrato."),
        "151-180": ("orange", "Rojo",    "Proyecto en alerta roja: más de 150 días sin firma del primer contrato."),
        ">180":    ("red",    "Negro",   "Proyecto en alerta negra: más de 180 días sin firma del primer contrato."),
    },
    "hito_3_val": {
        "0-30":  ("green",  "Verde",   "Proyecto dentro de los tiempos para la firma del acta de inicio."),
        "31-45": ("yellow", "Naranja", "Proyecto en alerta: más de 30 días sin firma del acta de inicio."),
        "46-60": ("orange", "Rojo",    "Proyecto en alerta roja: más de 45 días sin firma del acta de inicio."),
        ">60":   ("red",    "Negro",   "Proyecto en alerta negra: más de 60 días sin firma del acta de inicio."),
    },
    "hito_4_val": {
        "0-1":   ("green",  "Verde",   "Proyecto presenta horizonte vigente."),
        "1.1-3": ("yellow", "Naranja", "Proyecto con horizonte vencido entre 1 y 3 meses."),
        "3.1-6": ("orange", "Rojo",    "Proyecto con horizonte vencido mayor a 3 meses."),
        ">6":    ("red",    "Negro",   "Proyecto con horizonte vencido mayor a 6 meses."),
    },
    "hito_5_val": {
        "0-100":   ("green",  "Verde",   "Proyecto dentro de los tiempos para pasar a estado 'Para cierre'."),
        "101-150": ("yellow", "Naranja", "Proyecto en alerta: más de 100 días desde su terminación sin pasar a 'Para cierre'."),
        "151-180": ("orange", "Rojo",    "Proyecto en alerta roja: más de 150 días desde su terminación."),
        ">180":    ("red",    "Negro",   "Proyecto en alerta negra: más de 180 días desde su terminación."),
    },
}

TABLA_DESCENTRALIZADAS = "OtrosEjecutoresDescentralizadas"
COLS_EVAL = [
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
    "CALIFICACIÓN INFORMACIÓN A TIEMPO",
    "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
    "CALIFICACIÓN CALIDAD INFORMACIÓN",
]
COLS_EVAL_LABELS = [
    "Desempeño en contratación",
    "Información a tiempo",
    "Ejecución del proyecto",
    "Calidad de la información",
]

def clasificar_hito4_meses(col):
    """Hito 4 se clasifica en meses (días / 30), no en días directos."""
    meses = pl.col(col) / 30.0
    return (
        pl.when(pl.col(col).is_null()).then(None)
        .when(meses <= 1).then(pl.lit("0-1"))
        .when(meses <= 3).then(pl.lit("1.1-3"))
        .when(meses <= 6).then(pl.lit("3.1-6"))
        .otherwise(pl.lit(">6"))
    )

def clasificar(col, intervalos):
    expr = pl.when(pl.col(col).is_null()).then(None)
    for label, lo, hi in intervalos:
        cond = (pl.col(col) >= lo) & (pl.col(col) <= hi) if hi is not None else (pl.col(col) >= lo)
        expr = expr.when(cond).then(pl.lit(label))
    return expr.otherwise(None)

@st.cache_data
def procesar(file_bytes):
    df = pl.read_excel(io.BytesIO(file_bytes), table_name=TABLA_ESPERADA)
    df = (
        df.select(
            "ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO",
            "ESTADO PROYECTO", "ESTADO CONTRATO",
            "CPI", "SPI",
            "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
            "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
            "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY",
        )
        .with_columns(
            pl.col(
                "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
                "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
                "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY",
            ).cast(pl.Date, strict=False)
        )
        .with_columns(
            # Hito 1
            pl.when(
                ((pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")) &
                (~pl.col("FECHA APROBACIÓN PROYECTO").is_null()) & (~pl.col("FECHA DE CORTE GESPROY").is_null())
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA APROBACIÓN PROYECTO")).dt.total_days()).otherwise(None).alias("hito_1_val"),
            # Hito 2
            pl.when(
                ((pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")) &
                (~pl.col("FECHA DE APERTURA DEL PRIMER PROCESO").is_null())
            ).then((pl.col("FECHA ACTA INICIO") - pl.col("FECHA DE APERTURA DEL PRIMER PROCESO")).dt.total_days()).otherwise(None).alias("hito_2_val"),
            # Hito 3
            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO SIN ACTA DE INICIO") &
                (~pl.col("FECHA SUSCRIPCION").is_null())
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA SUSCRIPCION")).dt.total_days()).otherwise(None).alias("hito_3_val"),
            # Hito 4
            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO EN EJECUCIÓN") &
                (pl.col("CPI") == 0) & (pl.col("SPI") == 0) &
                (pl.col("HORIZONTE DEL PROYECTO") <= pl.col("FECHA DE CORTE GESPROY"))
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("HORIZONTE DEL PROYECTO")).dt.total_days()).otherwise(None).alias("hito_4_val"),
            # Hito 5
            pl.when(~pl.col("FECHA DE FINALIZACIÓN").is_null()).then(
                (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA DE FINALIZACIÓN")).dt.total_days()
            ).otherwise(None).alias("hito_5_val"),
            # Suspendidos — basado en ESTADO CONTRATO
            pl.when(
                pl.col("ESTADO CONTRATO").str.strip_chars().str.to_uppercase() == "SUSPENDIDO"
            ).then(pl.lit(1)).otherwise(None).alias("Suspendidos"),
            # Para cierre
            pl.when(pl.col("ESTADO PROYECTO") == "PARA CIERRE").then(pl.lit(1)).otherwise(None).alias("Para cierre"),
        )
        .with_columns(
            clasificar("hito_1_val", INTERVALOS["hito_1_val"]).alias("clasi_1"),
            clasificar("hito_2_val", INTERVALOS["hito_2_val"]).alias("clasi_2"),
            clasificar("hito_3_val", INTERVALOS["hito_3_val"]).alias("clasi_3"),
            clasificar_hito4_meses("hito_4_val").alias("clasi_4"),
            clasificar("hito_5_val", INTERVALOS["hito_5_val"]).alias("clasi_5"),
        )
    )
    return df

@st.cache_data
def _validar_cols_eval(df, cols, col_agrup):
    """
    Intenta castear las columnas de calificación a Float64.
    Retorna (df_casteado, errores) donde errores es lista de dicts con info de cada columna problemática.
    """
    errores = []
    cols_ok  = []
    for c in cols:
        if c not in df.columns:
            continue
        dtype_actual = str(df[c].dtype)
        # Verificar si ya es numérico
        if df[c].dtype in (pl.Float32, pl.Float64, pl.Int32, pl.Int64, pl.Int16, pl.UInt32):
            cols_ok.append(c)
            continue
        # Intentar cast — contar cuántos valores se perderían
        casteada = df[c].cast(pl.Float64, strict=False)
        nulos_antes  = df[c].is_null().sum()
        nulos_despues = casteada.is_null().sum()
        perdidos = int(nulos_despues - nulos_antes)
        if perdidos > 0:
            # Mostrar hasta 3 ejemplos de valores problemáticos
            ejemplos = (
                df.filter(df[c].is_not_null() & casteada.is_null())[c]
                .head(3).to_list()
            )
            errores.append({
                "col":        c,
                "tipo":       dtype_actual,
                "perdidos":   perdidos,
                "total":      df.height,
                "ejemplos":   ejemplos,
            })
        else:
            cols_ok.append(c)

    if cols_ok:
        df = df.with_columns([
            pl.col(c).cast(pl.Float64, strict=False) for c in cols_ok
        ])
    return df, cols_ok, errores


def _render_eval_errors(errores, contexto=""):
    """Muestra tarjetas de error amigables para columnas de calificación no numéricas."""
    tipo_amigable = {
        "string": "texto", "utf8": "texto", "str": "texto",
        "bool": "verdadero/falso", "date": "fecha", "datetime": "fecha",
    }
    st.warning(
        f"⚠️ Algunas calificaciones{' de ' + contexto if contexto else ''} "
        f"tienen datos que no pudieron leerse como números. "
        f"Esas columnas se excluyeron del cálculo.",
        icon=None,
    )
    for e in errores:
        pct = round(e["perdidos"] / e["total"] * 100, 1) if e["total"] else 0
        # Convertir tipo técnico a lenguaje amigable
        tipo_raw = e["tipo"].lower()
        tipo_legible = next((v for k, v in tipo_amigable.items() if k in tipo_raw), "no numérico")
        # Formatear ejemplos
        ejemplos_str = " · ".join([f'<code>{v}</code>' for v in e["ejemplos"]])
        st.markdown(f"""
<div class="error-card">
  <div class="error-title">&#9888; Calificación con valores incorrectos</div>
  <div class="error-body">
    La columna <strong>{e['col']}</strong> contiene valores de tipo <strong>{tipo_legible}</strong>
    en lugar de números.<br>
    Se encontraron <strong>{e['perdidos']} registro(s) con problemas</strong>
    de un total de {e['total']} ({pct}%).<br>
    Ejemplos de valores problemáticos encontrados: {ejemplos_str}
  </div>
  <div class="error-fix">
    <strong>Cómo corregirlo en Excel</strong>
    Abre el archivo, busca la columna <strong>{e['col']}</strong> y revisa
    que cada celda contenga únicamente un número (por ejemplo: <code>3.5</code>, <code>4</code> o <code>2.75</code>).
    Reemplaza cualquier texto, guion, «N/A» o celda vacía con el valor numérico correspondiente.
    Guarda el archivo y vuelve a cargarlo aquí.
  </div>
</div>""", unsafe_allow_html=True)


@st.cache_data
def procesar_descentralizadas(file_bytes):
    """Lee la tabla de descentralizadas y calcula promedios de calificación por EJECUTOR."""
    try:
        df = pl.read_excel(io.BytesIO(file_bytes), table_name=TABLA_DESCENTRALIZADAS)
        cols_disponibles = [c for c in COLS_EVAL if c in df.columns]
        if not cols_disponibles or "EJECUTOR" not in df.columns:
            return None, [], []
        df, cols_ok, errores = _validar_cols_eval(df, cols_disponibles, "EJECUTOR")
        if not cols_ok:
            return None, [], errores
        agg_exprs = [pl.col(c).mean().round(2).alias(c) for c in cols_ok]
        resultado = df.group_by("EJECUTOR").agg(agg_exprs).sort("EJECUTOR")
        return resultado, cols_ok, errores
    except Exception:
        return None, [], []


@st.cache_data
def procesar_eval_sucre(file_bytes):
    """Calcula promedios de calificación por ENTIDAD O SECRETARIA (tabla Sucre)."""
    try:
        df = pl.read_excel(io.BytesIO(file_bytes), table_name=TABLA_ESPERADA)
        cols_disponibles = [c for c in COLS_EVAL if c in df.columns]
        if not cols_disponibles or "ENTIDAD O SECRETARIA" not in df.columns:
            return None, [], []
        df, cols_ok, errores = _validar_cols_eval(df, cols_disponibles, "ENTIDAD O SECRETARIA")
        if not cols_ok:
            return None, [], errores
        agg_exprs = [pl.col(c).mean().round(2).alias(c) for c in cols_ok]
        resultado = df.group_by("ENTIDAD O SECRETARIA").agg(agg_exprs).sort("ENTIDAD O SECRETARIA")
        return resultado, cols_ok, errores
    except Exception:
        return None, [], []


def badge_html(val, hito_key=None):
    """Genera badge con punto de color semáforo y tooltip con mensaje."""
    if val is None: return ""
    val_str = str(val)
    cls_map = {
        "0-100": "badge-green",  "0-30":  "badge-green",  "0-1":   "badge-green",
        "101-150": "badge-yellow","31-45": "badge-yellow", "1.1-3": "badge-yellow",
        "151-180": "badge-orange","46-60": "badge-orange", "3.1-6": "badge-orange",
        ">180": "badge-red",     ">60":   "badge-red",    ">6":    "badge-red",
    }
    cls = cls_map.get(val_str, "badge-yellow")

    tooltip_html = ""
    if hito_key and hito_key in SEMAFOROS and val_str in SEMAFOROS[hito_key]:
        _, color_nombre, mensaje = SEMAFOROS[hito_key][val_str]
        tooltip_html = f'<span class="badge-tooltip"><strong style="color:#47b1d5;display:block;margin-bottom:3px">● {color_nombre}</strong>{mensaje}</span>'

    return (
        f'<span class="badge {cls}">'
        f'<span class="badge-dot"></span>'
        f'{val_str}'
        f'{tooltip_html}'
        f'</span>'
    )

def generar_excel(df_det, df_agr):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_agr.to_pandas().to_excel(writer, sheet_name="Resumen por Entidad", index=False)
        df_det.to_pandas().to_excel(writer, sheet_name="Detalle Proyectos", index=False)
    return buf.getvalue()

def th(label, titulo, desc):
    return f"""<th><div class="th-wrap">{label}<span class="th-icon">?</span>
    <div class="th-tooltip"><strong>{titulo}</strong>{desc}</div></div></th>"""

# ─────────────────────────────────────────────────────────────────────────────
# VALIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def error_card(titulo, cuerpo, solucion):
    return f"""
    <div class="error-card">
        <div class="error-title">&#9888; {titulo}</div>
        <div class="error-body">{cuerpo}</div>
        <div class="error-fix"><strong>Cómo solucionarlo</strong>{solucion}</div>
    </div>"""

def _tipo_amigable(dtype_str):
    """Convierte tipo técnico de Polars a nombre comprensible."""
    d = dtype_str.lower()
    if "utf" in d or "str" in d or "cat" in d: return "Texto"
    if "float" in d or "int" in d:             return "Número"
    if "date" in d or "time" in d:             return "Fecha"
    if "bool" in d:                            return "Verdadero/Falso"
    return dtype_str

def validar_archivo(file_bytes):
    """Retorna (df, errores_html). Si hay errores, df es None."""
    errores = []

    # 1. Verificar que la tabla existe
    try:
        df_raw = pl.read_excel(io.BytesIO(file_bytes), table_name=TABLA_ESPERADA)
    except Exception as e:
        msg = str(e)
        if "table" in msg.lower() or "not found" in msg.lower() or "name" in msg.lower():
            errores.append(error_card(
                "Tabla no encontrada",
                f"No se encontró una tabla con el nombre <b>{TABLA_ESPERADA}</b> en el archivo. "
                f"Es posible que el nombre haya sido cambiado o que los datos no estén definidos como tabla de Excel.",
                f"En Excel, selecciona el rango de datos → <b>Insertar → Tabla</b>, y asegúrate de que el nombre "
                f"sea exactamente <code>{TABLA_ESPERADA}</code> (sin espacios adicionales, respetando mayúsculas)."
            ))
        else:
            errores.append(error_card(
                "Error al leer el archivo",
                f"El archivo no pudo ser leído correctamente.",
                "Verifica que el archivo no esté dañado, que tenga extensión <b>.xlsx</b> y que no esté abierto en Excel al momento de cargarlo."
            ))
        return None, errores

    cols_actuales = set(df_raw.columns)

    # 2. Columnas faltantes
    faltantes = [c for c in COLUMNAS_ESPERADAS if c not in cols_actuales]
    if faltantes:
        chips = "".join(f"<span class='col-missing'>{c}</span>" for c in faltantes)
        errores.append(error_card(
            f"{'Columna faltante' if len(faltantes) == 1 else f'{len(faltantes)} columnas faltantes'}",
            f"Las siguientes columnas no fueron encontradas en la tabla:<div class='error-cols'>{chips}</div>"
            f"<div style='margin-top:0.5rem;font-size:0.79rem;color:#7f1d1d'>Puede que el nombre haya sido modificado por error. "
            f"Compara con la lista de columnas esperadas al final de esta pantalla.</div>",
            "Abre el archivo en Excel, ve a la tabla <b>MatrizSeguimientoEvaluacion</b> y verifica que los encabezados "
            "coincidan exactamente (respeta mayúsculas, tildes y espacios). No renombres las columnas originales."
        ))

    # 3. Tipo de datos incorrecto — solo columnas que sí existen
    tipo_incorrecto = []
    for col, (tipo_label, tipos_validos) in COLUMNAS_ESPERADAS.items():
        if col not in cols_actuales:
            continue
        dtype = df_raw[col].dtype
        if tipo_label == "fecha":
            continue  # el cast strict=False lo maneja en procesar
        if dtype not in tipos_validos:
            tipo_incorrecto.append((col, tipo_label, str(dtype)))

    if tipo_incorrecto:
        chips = "".join(
            f"<span class='col-wrong-type'>{col}</span>"
            for col, _, _ in tipo_incorrecto
        )
        detalles = "".join(
            f"<li style='margin-bottom:4px'><b>{col}</b>: "
            f"el sistema encontró <b>{_tipo_amigable(dtype_actual)}</b>, "
            f"pero esperaba <b>{TIPO_LABEL[tipo_label]}</b> — {TIPO_EJEMPLO[tipo_label]}</li>"
            for col, tipo_label, dtype_actual in tipo_incorrecto
        )
        errores.append(error_card(
            f"{'Tipo de dato incorrecto' if len(tipo_incorrecto) == 1 else f'Tipo de dato incorrecto en {len(tipo_incorrecto)} columnas'}",
            f"Las siguientes columnas tienen un tipo de dato que no corresponde:<div class='error-cols'>{chips}</div>"
            f"<ul style='margin-top:0.6rem;font-size:0.8rem;padding-left:1.2rem'>{detalles}</ul>",
            "En Excel, selecciona la columna señalada y revisa el formato de las celdas (menú <b>Inicio → Número</b>). "
            "Las columnas <b>CPI</b> y <b>SPI</b> deben tener formato <b>Número</b>. "
            "Si los valores están alineados a la izquierda, es probable que estén guardados como texto: "
            "selecciona la columna → <b>Datos → Texto en columnas</b> → finalizar."
        ))

    if errores:
        return None, errores

    return df_raw, []

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — carga y filtros
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div class='sidebar-section'>📁 Datos</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx"], label_visibility="collapsed")
    if uploaded:
        st.success("Archivo cargado correctamente")

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="page-header">
  <div>
    <h1>Seguimiento y Evaluación · Regalías</h1>
    <p>Matriz de seguimiento de proyectos por hito de gestión</p>
  </div>
</div>
""", unsafe_allow_html=True)

if uploaded is None:
    st.info("Carga el archivo Excel de la Matriz de Seguimiento en el panel izquierdo para comenzar.")
    st.stop()

file_bytes = uploaded.read()
df_raw, errores = validar_archivo(file_bytes)

if errores:
    muted = C["muted"]
    ref_rows = "".join(
        f"<tr><td><code>{col}</code></td>"
        f"<td><b>{TIPO_LABEL[tipo]}</b></td>"
        f"<td style='color:{muted}'>{TIPO_EJEMPLO[tipo]}</td></tr>"
        for col, (tipo, _) in COLUMNAS_ESPERADAS.items()
    )
    ref_table = f"""
    <table class="ref-table">
        <thead><tr>
            <th>Nombre exacto de la columna</th>
            <th>Tipo de dato esperado</th>
            <th>Ejemplo</th>
        </tr></thead>
        <tbody>{ref_rows}</tbody>
    </table>"""

    st.markdown(f"""
    <div style="background:{C['white']};border-radius:12px;padding:1.5rem 1.8rem;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-top:0.5rem">
        <div style="font-family:'Montserrat',sans-serif;font-size:1rem;font-weight:700;
                    color:{C['azul_oscuro']};margin-bottom:0.3rem">
            No se pudo cargar el archivo
        </div>
        <div style="font-size:0.83rem;color:{C['muted']};margin-bottom:1rem">
            Se encontraron los siguientes problemas. Corrígelos en Excel y vuelve a cargar el archivo.
        </div>
        {"".join(errores)}
        <details style="margin-top:1.2rem">
            <summary style="font-size:0.8rem;font-weight:600;color:{C['azul_medio']};
                            cursor:pointer;user-select:none;list-style:none;display:flex;align-items:center;gap:6px">
                &#9432; Ver referencia completa: tabla y columnas esperadas
            </summary>
            <div style="margin-top:0.8rem">
                <div style="font-size:0.78rem;font-weight:700;text-transform:uppercase;
                            letter-spacing:0.8px;color:{C['muted']};margin-bottom:0.3rem">
                    Nombre de la tabla en Excel
                </div>
                <div style="font-size:0.85rem;margin-bottom:1rem">
                    La tabla debe llamarse exactamente:
                    <code style="background:#f1f5f9;padding:2px 8px;border-radius:4px;
                    color:{C['azul_medio']};font-size:0.85rem">{TABLA_ESPERADA}</code>
                    <span style="font-size:0.78rem;color:{C['muted']}">
                    — puedes verificarlo en Excel seleccionando cualquier celda de la tabla
                    y mirando el cuadro de nombre en la esquina superior izquierda.</span>
                </div>
                <div style="font-size:0.78rem;font-weight:700;text-transform:uppercase;
                            letter-spacing:0.8px;color:{C['muted']};margin-bottom:0.3rem">
                    Columnas esperadas
                </div>
                <div style="font-size:0.78rem;color:{C['muted']};margin-bottom:0.6rem">
                    Los nombres deben ser exactamente iguales, incluyendo tildes, espacios y mayúsculas.
                </div>
                {ref_table}
            </div>
        </details>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

try:
    df = procesar(file_bytes)
except Exception as e:
    msg = str(e)
    # Detectar columna faltante en procesar (por si pasó la validación pero falla el select)
    col_hint = ""
    for col in COLUMNAS_ESPERADAS:
        if col.lower() in msg.lower():
            col_hint = f"<br>Columna relacionada: <code>{col}</code>"
            break
    st.markdown(f"""
    <div style="background:{C['white']};border-radius:12px;padding:1.5rem 1.8rem;
                box-shadow:0 1px 6px rgba(0,0,0,0.07);margin-top:0.5rem">
        <div style="font-family:'Montserrat',sans-serif;font-size:1rem;font-weight:700;
                    color:{C['azul_oscuro']};margin-bottom:0.8rem">
            Error al procesar el archivo
        </div>
        <div class="error-card">
            <div class="error-title">&#9888; Problema inesperado al leer los datos</div>
            <div class="error-body">
                El archivo fue cargado pero ocurrió un error al procesar su contenido.{col_hint}
            </div>
            <div class="error-fix">
                <strong>Cómo solucionarlo</strong>
                Verifica que ninguna columna haya sido renombrada o eliminada en la tabla
                <code>{TABLA_ESPERADA}</code>. Si el problema persiste, intenta exportar el archivo
                de nuevo desde el sistema origen y cargarlo sin modificaciones.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# FILTROS EN SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
HITOS = {
    "H1 · Sin contratar sin apertura":    ("hito_1_val", "clasi_1"),
    "H2 · Sin contratar con apertura":    ("hito_2_val", "clasi_2"),
    "H3 · Contratado sin acta de inicio": ("hito_3_val", "clasi_3"),
    "H4 · En ejecución rezagado":         ("hito_4_val", "clasi_4"),
    "H5 · Proyectos terminados":          ("hito_5_val", "clasi_5"),
}

with st.sidebar:
    st.markdown("<div class='sidebar-section'>Filtros</div>", unsafe_allow_html=True)

    entidades   = sorted(df["ENTIDAD O SECRETARIA"].drop_nulls().unique().to_list())
    sel_entidades = st.multiselect("Entidad / Secretaría", entidades, default=entidades)

    estados_raw  = sorted(df["ESTADO PROYECTO"].drop_nulls().unique().to_list())
    tiene_sin    = df["ESTADO PROYECTO"].is_null().any() or (df["ESTADO PROYECTO"] == "").any()
    opciones_est = estados_raw + (["(Sin estado)"] if tiene_sin else [])
    sel_estados  = st.multiselect("Estado del proyecto", opciones_est, default=opciones_est)

    st.markdown("<div class='sidebar-section'>Vista de detalle</div>", unsafe_allow_html=True)
    sel_hito_label = st.selectbox("Hito a detallar", list(HITOS.keys()), label_visibility="collapsed")

sel_hito_col, sel_clasi_col = HITOS[sel_hito_label]

# ─────────────────────────────────────────────────────────────────────────────
# FILTRAR
# ─────────────────────────────────────────────────────────────────────────────
filter_expr = pl.col("ENTIDAD O SECRETARIA").is_in(sel_entidades)
if sel_estados:
    incluir_sin   = "(Sin estado)" in sel_estados
    estados_reales = [e for e in sel_estados if e != "(Sin estado)"]
    if estados_reales and incluir_sin:
        estado_expr = pl.col("ESTADO PROYECTO").is_in(estados_reales) | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")
    elif estados_reales:
        estado_expr = pl.col("ESTADO PROYECTO").is_in(estados_reales)
    elif incluir_sin:
        estado_expr = pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")
    else:
        estado_expr = pl.lit(False)
    filter_expr = filter_expr & estado_expr
df_f = df.filter(filter_expr)

# ─────────────────────────────────────────────────────────────────────────────
# KPIs — jerarquía visual
# ─────────────────────────────────────────────────────────────────────────────
total_proy      = df_f.height
total_entidades = df_f["ENTIDAD O SECRETARIA"].n_unique()
suspendidos     = int(df_f["Suspendidos"].drop_nulls().sum()) if df_f["Suspendidos"].drop_nulls().len() > 0 else 0
para_cierre     = int(df_f["Para cierre"].drop_nulls().sum()) if df_f["Para cierre"].drop_nulls().len() > 0 else 0
prom_h1         = df_f["hito_1_val"].mean()
prom_h1_str     = f"{prom_h1:.1f}" if prom_h1 is not None else "N/A"

st.markdown("<div style='height:0.2rem'></div>", unsafe_allow_html=True)
ka, kb, kc, kd, ke = st.columns([1.4, 1.4, 1, 1, 1])

with ka:
    st.markdown(f"""
    <div class="kpi-main">
        <div class="label">Total proyectos</div>
        <div class="value">{total_proy}</div>
        <div class="sub">en los filtros activos</div>
    </div>""", unsafe_allow_html=True)

with kb:
    st.markdown(f"""
    <div class="kpi-main" style="background:{C['verde_oscuro']}">
        <div class="label">Entidades</div>
        <div class="value">{total_entidades}</div>
        <div class="sub">secretarías / dependencias</div>
    </div>""", unsafe_allow_html=True)

with kc:
    st.markdown(f"""
    <div class="kpi-sec" style="border-left-color:{C['naranja_osc']}">
        <div class="label">Suspendidos</div>
        <div class="value" style="color:{C['naranja_osc']}">{suspendidos}</div>
        <div class="sub">proyectos</div>
    </div>""", unsafe_allow_html=True)

with kd:
    st.markdown(f"""
    <div class="kpi-sec" style="border-left-color:{C['cafe']}">
        <div class="label">Para cierre</div>
        <div class="value" style="color:{C['cafe']}">{para_cierre}</div>
        <div class="sub">proyectos</div>
    </div>""", unsafe_allow_html=True)

with ke:
    st.markdown(f"""
    <div class="kpi-sec" style="border-left-color:{C['cian']}">
        <div class="label">Prom. H1</div>
        <div class="value" style="color:{C['cian']}">{prom_h1_str}</div>
        <div class="sub">días promedio</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<div style='height:1.5rem'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# AGRUPACIÓN
# ─────────────────────────────────────────────────────────────────────────────
agrupacion = (
    df_f.group_by("ENTIDAD O SECRETARIA")
    .agg(
        pl.col("hito_1_val").mean().round(1).alias("Hito 1 (días)"),
        pl.col("hito_2_val").mean().round(1).alias("Hito 2 (días)"),
        pl.col("hito_3_val").mean().round(1).alias("Hito 3 (días)"),
        pl.col("hito_4_val").mean().round(1).alias("Hito 4 (días)"),
        pl.col("hito_5_val").mean().round(1).alias("Hito 5 (días)"),
        pl.col("Suspendidos").sum().alias("Suspendidos"),
        pl.col("Para cierre").sum().alias("Para cierre"),
        pl.len().alias("Total"),
    )
    .sort("ENTIDAD O SECRETARIA")
)

clasi_por_entidad = {}
for entidad in agrupacion["ENTIDAD O SECRETARIA"].to_list():
    sub = df_f.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)
    clasi_por_entidad[entidad] = {}
    for cc in ["clasi_1", "clasi_2", "clasi_3", "clasi_4", "clasi_5"]:
        vals = sub[cc].drop_nulls()
        clasi_por_entidad[entidad][cc] = vals.value_counts().sort("count", descending=True)[0, 0] if len(vals) > 0 else None

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab_resumen, tab_detalle, tab_evaluacion, tab_exportar = st.tabs([
    "Resumen por entidad",
    f"Detalle · {sel_hito_label}",
    "Evaluación del modelo",
    "Exportar",
])

# HITO key map para badge_html
HITO_KEY_MAP = {
    "clasi_1": "hito_1_val",
    "clasi_2": "hito_2_val",
    "clasi_3": "hito_3_val",
    "clasi_4": "hito_4_val",
    "clasi_5": "hito_5_val",
}

# ── TAB 1: Tabla resumen ──────────────────────────────────────────────────────
with tab_resumen:
    def hito_cell(dias_val, entidad, clasi_key):
        if dias_val is None or (isinstance(dias_val, float) and dias_val != dias_val):
            return "<td class='null-cell'>—</td>"
        clasi = clasi_por_entidad[entidad].get(clasi_key)
        hito_k = HITO_KEY_MAP.get(clasi_key)
        return f"<td><span class='dias-val'>{dias_val:.1f} d</span>{badge_html(clasi, hito_k)}</td>"

    rows_html = ""
    for row in agrupacion.to_dicts():
        e    = row["ENTIDAD O SECRETARIA"]
        susp = int(row["Suspendidos"]) if row["Suspendidos"] else 0
        pc   = int(row["Para cierre"]) if row["Para cierre"] else 0
        rows_html += f"""<tr>
            <td class="entidad-name">{e}</td>
            {hito_cell(row['Hito 1 (días)'], e, 'clasi_1')}
            {hito_cell(row['Hito 2 (días)'], e, 'clasi_2')}
            {hito_cell(row['Hito 3 (días)'], e, 'clasi_3')}
            {hito_cell(row['Hito 4 (días)'], e, 'clasi_4')}
            {hito_cell(row['Hito 5 (días)'], e, 'clasi_5')}
            <td style="text-align:center;font-weight:500">{susp}</td>
            <td style="text-align:center;font-weight:500">{pc}</td>
            <td style="text-align:center;font-weight:700;color:{C['azul_oscuro']}">{int(row['Total'])}</td>
        </tr>"""

    st.markdown(f"""
    <table class="summary-table">
    <thead><tr>
        <th>Entidad / Secretaría</th>
        {th("Sin contratar<br>sin apertura", "Hito 1 · Sin contratar sin apertura",
            "Promedio de días entre la <b>Fecha de aprobación</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Estado = SIN CONTRATAR y sin fecha de apertura.")}
        {th("Sin contratar<br>con apertura", "Hito 2 · Sin contratar con apertura",
            "Promedio de días entre la <b>Fecha de apertura del primer proceso</b> y la <b>Fecha de acta de inicio</b>.<br><br>Condición: Estado = SIN CONTRATAR con fecha de apertura registrada.")}
        {th("Contratado<br>sin acta de inicio", "Hito 3 · Contratado sin acta de inicio",
            "Promedio de días entre la <b>Fecha de suscripción</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Estado = CONTRATADO SIN ACTA DE INICIO.")}
        {th("En ejecución<br>rezagado", "Hito 4 · En ejecución rezagado",
            "Meses entre el <b>Horizonte del proyecto</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Estado = CONTRATADO EN EJECUCIÓN, CPI = 0, SPI = 0 y horizonte vencido.")}
        {th("Proyectos<br>terminados", "Hito 5 · Proyectos terminados",
            "Promedio de días entre la <b>Fecha de finalización</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Fecha de finalización registrada.")}
        {th("Suspendidos", "Proyectos suspendidos", "Conteo de proyectos cuyo <b>Estado contrato</b> = SUSPENDIDO.")}
        {th("Para cierre", "Proyectos para cierre", "Conteo de proyectos con Estado = PARA CIERRE.")}
        <th>Total</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
    </table>
    """, unsafe_allow_html=True)

# ── TAB 2: Detalle ────────────────────────────────────────────────────────────
with tab_detalle:
    # Identificar hito_key para badge_html según columna de clasificación seleccionada
    hito_key_detalle = HITO_KEY_MAP.get(sel_clasi_col, None)

    df_det = (
        df_f
        .filter(~pl.col(sel_hito_col).is_null())
        .select("ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO", "ESTADO PROYECTO", sel_hito_col, sel_clasi_col)
        .sort(["ENTIDAD O SECRETARIA", sel_hito_col], descending=[False, True])
    )

    if df_det.height == 0:
        st.info("No hay proyectos con valor en este hito para los filtros seleccionados.")
    else:
        for entidad in df_det["ENTIDAD O SECRETARIA"].unique().sort().to_list():
            sub  = df_det.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)
            prom = sub[sel_hito_col].mean()
            n    = sub.height
            prom_str = f"{prom:.1f} días" if prom is not None else "—"

            with st.expander(f"{entidad}   ·   {n} proyecto(s)   ·   Promedio: {prom_str}", expanded=False):
                det_rows = ""
                for r in sub.to_dicts():
                    dias_v   = r[sel_hito_col]
                    clasi_v  = r[sel_clasi_col]
                    dias_str = f"{dias_v:.1f} d" if dias_v is not None else "—"
                    det_rows += f"""<tr>
                        <td><span class="bpin-tag">{r['BPIN'] or '—'}</span></td>
                        <td style="font-size:0.81rem">{r['NOMBRE PROYECTO'] or '—'}</td>
                        <td><span class="estado-tag">{r['ESTADO PROYECTO'] or '(Sin estado)'}</span></td>
                        <td style="font-family:'DM Mono',monospace;font-weight:600;font-size:0.8rem">{dias_str}</td>
                        <td>{badge_html(clasi_v, hito_key_detalle)}</td>
                    </tr>"""
                st.markdown(f"""
                <table class="detail-table">
                <thead><tr>
                    <th>BPIN</th><th>Nombre del proyecto</th><th>Estado</th><th>Días</th><th>Clasificación</th>
                </tr></thead>
                <tbody>{det_rows}</tbody>
                </table>""", unsafe_allow_html=True)

# ── TAB 3: Evaluación del modelo ──────────────────────────────────────────────
with tab_evaluacion:
    st.markdown("<div class='section-heading'>Evaluación del modelo ejecutor</div>", unsafe_allow_html=True)

    modelo_sel = st.radio(
        "Ejecutor",
        ["Departamento de Sucre", "Descentralizadas"],
        horizontal=True,
        label_visibility="collapsed",
    )
    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)

    if modelo_sel == "Departamento de Sucre":
        df_eval, cols_eval_ok, eval_errores = procesar_eval_sucre(file_bytes)
        col_entidad    = "ENTIDAD O SECRETARIA"
        label_entidad  = "Entidad / Secretaría"
        contexto_error = "Departamento de Sucre"
    else:
        df_eval, cols_eval_ok, eval_errores = procesar_descentralizadas(file_bytes)
        col_entidad    = "EJECUTOR"
        label_entidad  = "Ejecutor"
        contexto_error = "Descentralizadas"

    if eval_errores:
        _render_eval_errors(eval_errores, contexto_error)

    if df_eval is None or df_eval.height == 0:
        st.info(f"No se encontraron datos de evaluación para «{modelo_sel}».")
    else:
        # ── Colores de calificación (escala 0-5) ──
        def eval_color(score, max_score=5.0):
            ratio = score / max_score if max_score > 0 else 0
            if ratio >= 0.8:   return C["verde_medio"],  "Sobresaliente"
            elif ratio >= 0.6: return C["cian"],         "Satisfactorio"
            elif ratio >= 0.4: return C["naranja"],      "Aceptable"
            else:              return C["salmon"],        "Por mejorar"

        max_score = 5.0

        # ── CSS extra para tabla de evaluación ──
        st.markdown(f"""
        <style>
        .eval-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 0.83rem;
            background: #ffffff;
            border-radius: 8px;
            overflow: hidden;
            margin-bottom: 0.5rem;
        }}
        .eval-table thead tr {{ background: {C['azul_oscuro']}; color: white; }}
        .eval-table th {{
            padding: 0.65rem 1rem;
            font-size: 0.63rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            text-align: left;
        }}
        .eval-table td {{
            padding: 0.6rem 1rem;
            border-bottom: 1px solid {C['border']};
            vertical-align: middle;
            background: #ffffff;
        }}
        .eval-table tbody tr:nth-child(even) td {{ background: #f7fafd; }}
        .eval-table tbody tr:last-child td {{ border-bottom: none; }}
        .eval-table tbody tr:hover td {{ background: #e8f3ff !important; }}
        .eval-mini-bar {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .eval-mini-bar-bg {{
            flex: 1;
            background: {C['border']};
            border-radius: 4px;
            height: 8px;
            overflow: hidden;
            min-width: 80px;
        }}
        .eval-mini-bar-fill {{
            height: 100%;
            border-radius: 4px;
        }}
        .eval-mini-score {{
            font-family: 'DM Mono', monospace;
            font-weight: 700;
            font-size: 0.82rem;
            min-width: 2.5rem;
            text-align: right;
        }}
        .eval-nivel {{
            font-size: 0.68rem;
            font-weight: 600;
            padding: 2px 7px;
            border-radius: 20px;
            white-space: nowrap;
        }}
        </style>
        """, unsafe_allow_html=True)

        # ── 4 sub-pestañas, una por calificación ──
        tabs_eval = st.tabs([
            "Desempeño en contratación",
            "Información a tiempo",
            "Ejecución del proyecto",
            "Calidad de la información",
        ])

        for i, (col_cal, label_cal) in enumerate(zip(COLS_EVAL, COLS_EVAL_LABELS)):
            with tabs_eval[i]:
                if col_cal not in cols_eval_ok:
                    st.info(f"No hay datos disponibles para «{label_cal}» debido a errores en el archivo.")
                    continue

                filas = []
                for row in df_eval.sort(col_cal, descending=True).to_dicts():
                    nombre = row.get(col_entidad) or "Sin nombre"
                    score  = row.get(col_cal)
                    if score is None:
                        continue
                    color_bar, nivel = eval_color(score, max_score)
                    pct = min(100, round((score / max_score) * 100, 1))

                    # Color de fondo suave para el nivel
                    nivel_bg = {
                        "Sobresaliente": "#d1fae5", "Satisfactorio": "#e0f7fa",
                        "Aceptable":     "#fff7ed", "Por mejorar":   "#fee2e2",
                    }.get(nivel, "#f1f5f9")
                    nivel_color = {
                        "Sobresaliente": "#065f46", "Satisfactorio": "#005931",
                        "Aceptable":     "#9a3412", "Por mejorar":   "#991b1b",
                    }.get(nivel, C["muted"])

                    filas.append(f"""<tr>
                        <td class="entidad-name">{nombre}</td>
                        <td>
                            <div class="eval-mini-bar">
                                <div class="eval-mini-bar-bg">
                                    <div class="eval-mini-bar-fill" style="width:{pct}%;background:{color_bar}"></div>
                                </div>
                                <span class="eval-mini-score" style="color:{color_bar}">{score:.2f}</span>
                            </div>
                        </td>
                        <td style="text-align:center">
                            <span class="eval-nivel" style="background:{nivel_bg};color:{nivel_color}">{nivel}</span>
                        </td>
                    </tr>""")

                if not filas:
                    st.info("No hay registros con calificación para este criterio.")
                else:
                    st.markdown(f"""
                    <table class="eval-table">
                    <thead><tr>
                        <th>{label_entidad}</th>
                        <th>Calificación promedio &nbsp;(escala 0 – {max_score:.0f})</th>
                        <th style="text-align:center">Nivel</th>
                    </tr></thead>
                    <tbody>{"".join(filas)}</tbody>
                    </table>
                    """, unsafe_allow_html=True)


# ── TAB 4: Exportar ───────────────────────────────────────────────────────────
with tab_exportar:
    st.markdown("<div class='section-heading'>Descargar reporte</div>", unsafe_allow_html=True)
    st.markdown("El archivo incluye dos hojas: **Resumen por entidad** con los promedios por hito, y **Detalle proyectos** con el registro completo filtrado.", unsafe_allow_html=False)
    st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

    export_df = df_f.select(
        "ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO", "ESTADO PROYECTO",
        "hito_1_val", "clasi_1", "hito_2_val", "clasi_2",
        "hito_3_val", "clasi_3", "hito_4_val", "clasi_4",
        "hito_5_val", "clasi_5", "Suspendidos", "Para cierre",
    ).rename({
        "hito_1_val": "Hito 1 (días)", "clasi_1": "Clasificación hito 1",
        "hito_2_val": "Hito 2 (días)", "clasi_2": "Clasificación hito 2",
        "hito_3_val": "Hito 3 (días)", "clasi_3": "Clasificación hito 3",
        "hito_4_val": "Hito 4 (días)", "clasi_4": "Clasificación hito 4",
        "hito_5_val": "Hito 5 (días)", "clasi_5": "Clasificación hito 5",
    })

    st.download_button(
        label="Descargar reporte Excel",
        data=generar_excel(export_df, agrupacion),
        file_name=f"regalias_seguimiento_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
