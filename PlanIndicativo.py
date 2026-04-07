"""
Dashboard de Reporte de Avance PDD
Basado en el analisis del notebook ReporteAvance.ipynb
"""

import streamlit as st
import polars as pl
import polars.selectors as cs
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
import traceback
from typing import Optional

# -------------------------------------------------
# PALETA CORPORATIVA
# -------------------------------------------------
COLORS = {
    "verde_oscuro":   "#17743d",
    "verde_medio":    "#005931",
    "cyan":           "#47b1d5",
    "azul_medio":     "#1754ab",
    "azul_oscuro":    "#003d6c",
    "naranja_claro":  "#d88c16",
    "naranja":        "#cf7000",
    "naranja_quemado":"#d37e00",
    "cafe":           "#9b5b1e",
    "salmon":         "#e68878",
    "blanco":         "#ffffff",
    "gris_fondo":     "#f4f6f9",
    "gris_texto":     "#2d3142",
}

CAT_COLORS = {
    "Superior":      COLORS["verde_oscuro"],
    "Alto":          COLORS["cyan"],
    "Medio":         COLORS["naranja_claro"],
    "Bajo":          COLORS["naranja"],
    "Critico":       COLORS["salmon"],
    "Sin Programar": COLORS["cafe"],
}

VIGENCIAS = ["2024", "2025", "2026"]

# -------------------------------------------------
# CONFIGURACION DE PAGINA
# -------------------------------------------------
st.set_page_config(
    page_title="Dashboard PDD - Reporte de Avance",
    page_icon="assets/favicon.ico" if False else None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# -------------------------------------------------
# CSS PERSONALIZADO
# -------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=DM+Sans:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif;
    color: {COLORS['gris_texto']};
}}

.main-header {{
    background: linear-gradient(135deg, {COLORS['azul_oscuro']} 0%, {COLORS['azul_medio']} 60%, {COLORS['cyan']} 100%);
    padding: 2.5rem 3rem 2rem;
    border-radius: 0 0 2rem 2rem;
    margin: -1rem -1rem 2rem -1rem;
    color: white;
}}
.main-header h1 {{
    font-family: 'Sora', sans-serif;
    font-weight: 700;
    font-size: 2.1rem;
    margin: 0;
    letter-spacing: -0.5px;
}}
.main-header p {{
    margin: 0.4rem 0 0;
    font-size: 0.95rem;
    opacity: 0.82;
}}

.kpi-card {{
    background: white;
    border-radius: 1rem;
    padding: 1.4rem 1.6rem;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    border-left: 5px solid {COLORS['azul_medio']};
    margin-bottom: 1rem;
}}
.kpi-card.verde   {{ border-left-color: {COLORS['verde_oscuro']}; }}
.kpi-card.cyan    {{ border-left-color: {COLORS['cyan']}; }}
.kpi-card.naranja {{ border-left-color: {COLORS['naranja_claro']}; }}
.kpi-card.cafe    {{ border-left-color: {COLORS['cafe']}; }}
.kpi-value {{
    font-family: 'Sora', sans-serif;
    font-size: 2.2rem;
    font-weight: 700;
    line-height: 1.1;
}}
.kpi-label {{
    font-size: 0.82rem;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #6b7280;
    margin-top: 0.3rem;
}}
.kpi-tooltip {{
    font-size: 0.75rem;
    color: #9ca3af;
    margin-top: 0.5rem;
    border-top: 1px solid #f3f4f6;
    padding-top: 0.5rem;
}}

.section-title {{
    font-family: 'Sora', sans-serif;
    font-size: 1.15rem;
    font-weight: 600;
    color: {COLORS['azul_oscuro']};
    border-bottom: 2px solid {COLORS['cyan']};
    padding-bottom: 0.4rem;
    margin: 2rem 0 1rem;
}}

section[data-testid="stSidebar"] {{
    background: {COLORS['azul_oscuro']};
}}
section[data-testid="stSidebar"] * {{
    color: white !important;
}}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] .stRadio label {{
    color: #cbd5e1 !important;
    font-size: 0.85rem;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    color: {COLORS['cyan']} !important;
    font-family: 'Sora', sans-serif;
}}

.error-box {{
    background: #fff7f0;
    border: 1.5px solid {COLORS['salmon']};
    border-radius: 0.8rem;
    padding: 1.2rem 1.5rem;
    margin: 1rem 0;
}}
.error-box h4 {{
    color: #c0392b;
    margin: 0 0 0.5rem;
    font-family: 'Sora', sans-serif;
}}
.schema-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
    margin-top: 0.8rem;
}}
.schema-table th {{
    background: {COLORS['azul_oscuro']};
    color: white;
    padding: 0.5rem 0.8rem;
    text-align: left;
}}
.schema-table td {{
    padding: 0.4rem 0.8rem;
    border-bottom: 1px solid #e5e7eb;
}}
.schema-table tr:nth-child(even) td {{
    background: #f9fafb;
}}

.upload-zone {{
    background: #f8fafc;
    border: 2px dashed {COLORS['cyan']};
    border-radius: 1rem;
    padding: 1.5rem;
    margin: 0.5rem 0 1rem;
    text-align: center;
    color: #6b7280;
}}

hr.custom {{ border: none; border-top: 1px solid #e5e7eb; margin: 1.5rem 0; }}

.footer {{
    text-align: center;
    font-size: 0.78rem;
    color: #9ca3af;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid #e5e7eb;
}}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# FUNCIONES AUXILIARES
# -------------------------------------------------
def fmt_miles(v: float) -> str:
    """Formatea numero como millones con simbolo."""
    if v is None or (isinstance(v, float) and v != v):
        return "N/A"
    if abs(v) >= 1_000_000_000:
        return f"${v/1_000_000_000:.1f}B"
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.1f}M"
    if abs(v) >= 1_000:
        return f"${v/1_000:.1f}K"
    return f"${v:,.0f}"


def fmt_pct(v: float) -> str:
    if v is None:
        return "N/A"
    return f"{v*100:.1f}%"


def color_pct(v: float) -> str:
    if v is None:
        return COLORS["cafe"]
    if v >= 0.9:
        return COLORS["verde_oscuro"]
    if v >= 0.6:
        return COLORS["cyan"]
    if v >= 0.3:
        return COLORS["naranja_claro"]
    return COLORS["salmon"]


def kpi_card(label: str, value: str, color_class: str = "", tooltip: str = ""):
    tip_html = f'<div class="kpi-tooltip">Calculo: {tooltip}</div>' if tooltip else ""
    st.markdown(f"""
    <div class="kpi-card {color_class}">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
        {tip_html}
    </div>
    """, unsafe_allow_html=True)


def section_title(text: str):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)


def show_schema_error(file_name: str, schema: list):
    rows = "".join(
        f"<tr><td><code>{r['col']}</code></td><td>{r['tipo']}</td><td>{r['ejemplo']}</td></tr>"
        for r in schema
    )
    st.markdown(f"""
    <div class="error-box">
        <h4>Error al leer {file_name}</h4>
        <p>No se pudo procesar el archivo. Verifica que contenga las siguientes columnas y datos:</p>
        <table class="schema-table">
            <thead><tr><th>Columna esperada</th><th>Tipo</th><th>Ejemplo de dato</th></tr></thead>
            <tbody>{rows}</tbody>
        </table>
        <p style="margin-top:0.8rem;font-size:0.82rem;color:#6b7280;">
            <strong>Tip:</strong> Asegurate de que el archivo Excel tenga la tabla nombrada correctamente
            y que las columnas no tengan espacios extras o caracteres especiales.
        </p>
    </div>
    """, unsafe_allow_html=True)


# -------------------------------------------------
# ESQUEMAS DE ERROR AMIGABLES
# -------------------------------------------------
SCHEMAS = {
    "Plan Indicativo": [
        {"col": "Codigo Meta",                        "tipo": "Texto",   "ejemplo": "MT-001"},
        {"col": "Linea Estrategica",                  "tipo": "Texto",   "ejemplo": "Linea 1: Educacion de calidad"},
        {"col": "Sector PDD",                         "tipo": "Texto",   "ejemplo": "Educacion"},
        {"col": "Numero Programa PDD",                "tipo": "Numero",  "ejemplo": "1"},
        {"col": "Programa PDD",                       "tipo": "Texto",   "ejemplo": "Programa 1.1"},
        {"col": "Meta de cuatrenio",                  "tipo": "Numero",  "ejemplo": "5000"},
        {"col": "Tipo de Acumulacion",                "tipo": "Texto",   "ejemplo": "Acumulado"},
        {"col": "Responsable",                        "tipo": "Texto",   "ejemplo": "Secretaria de Educacion"},
        {"col": "Meta Fisica Esperada 2024",          "tipo": "Numero",  "ejemplo": "1200"},
        {"col": "EJECUCION 2024",                     "tipo": "Numero",  "ejemplo": "1100"},
        {"col": "PORCENTAJE DE EJECUCION 2024",       "tipo": "Decimal", "ejemplo": "0.92"},
        {"col": "CATEGORIA DE EJECUCION FISICA 2024", "tipo": "Texto",   "ejemplo": "Superior / Alto / Medio / Bajo / Critico"},
        {"col": "Programacion recursos propios icld24","tipo": "Numero",  "ejemplo": "500000000"},
        {"col": "Programacion regalias24",            "tipo": "Numero",  "ejemplo": "200000000"},
    ],
    "Hacienda (XLSX)": [
        {"col": "RP",                             "tipo": "Numero",  "ejemplo": "150000000"},
        {"col": "CODIGO META",                    "tipo": "Texto",   "ejemplo": "MT-001"},
        {"col": "CLASIFICACION RECURSOS",         "tipo": "Texto",   "ejemplo": "ICLD / SGP EDUCACION / REGALIAS"},
        {"col": "PROYECTO ARCHIVADO",             "tipo": "Texto",   "ejemplo": "(vacio si activo)"},
        {"col": "SE VA A CARGAR EN PI",           "tipo": "Texto",   "ejemplo": "(vacio si aplica)"},
        {"col": "DISTRIBUIR DE FORMA EQUITATIVA", "tipo": "Texto",   "ejemplo": "SI / NO"},
    ],
    "Regalias (XLSX)": [
        {"col": "PAGO EJECUTADO VALOR / COMPROMISOS", "tipo": "Numero", "ejemplo": "80000000"},
        {"col": "CODIGO META",                        "tipo": "Texto",  "ejemplo": "MT-001"},
        {"col": "CLASIFICACION RECURSOS",             "tipo": "Texto",  "ejemplo": "REGALIAS"},
        {"col": "ULTIMA FECHA PAGO",                  "tipo": "Fecha",  "ejemplo": "2026-03-15"},
    ],
}

# -------------------------------------------------
# CARGA DE ARCHIVOS
# -------------------------------------------------
def read_excel_safe(source, table_name: str, columns: list = None) -> Optional[pl.DataFrame]:
    """Carga un Excel con manejo robusto de errores."""
    try:
        kwargs = {"table_name": table_name}
        if columns:
            kwargs["columns"] = columns
        if isinstance(source, (str,)):
            return pl.read_excel(source, **kwargs)
        else:
            data = source.read() if hasattr(source, "read") else source
            return pl.read_excel(io.BytesIO(data), **kwargs)
    except Exception:
        return None


def fetch_github_file(url: str) -> Optional[bytes]:
    """Descarga un archivo desde GitHub (raw URL)."""
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        return r.content
    except Exception:
        return None


# -------------------------------------------------
# PROCESAMIENTO DE DATOS
# -------------------------------------------------
def programacion_financiera_expr(df: pl.DataFrame, suffix: str) -> pl.Expr:
    cols = [
        f"programacion recursos propios icld{suffix}",
        f"programacion recursos propios icde{suffix}",
        f"programacion sgp educacion{suffix}",
        f"programacion sgp salud{suffix}",
        f"programacion sgp apsb{suffix}",
        f"programacion cofinanciacion municipio{suffix}",
        f"programacion cofinanciacion nacion{suffix}",
        f"programacion credito{suffix}",
        f"programacion regalias{suffix}",
        f"programacion otras fuentes{suffix}",
    ]
    # Normalize actual column names to ASCII-lower for matching
    actual_lower = {c.lower()
                    .replace("á","a").replace("é","e").replace("í","i")
                    .replace("ó","o").replace("ú","u").replace("ñ","n")
                    .replace("ü","u"): c
                    for c in df.columns}
    existing = [actual_lower[c] for c in cols if c in actual_lower]
    if not existing:
        return pl.lit(0.0)
    expr = pl.col(existing[0]).cast(pl.Float64)
    for c in existing[1:]:
        expr = expr + pl.col(c).cast(pl.Float64)
    return expr


def process_regalias(source, year: str, table_name: str) -> Optional[pl.DataFrame]:
    df = read_excel_safe(source, table_name)
    if df is None:
        return None
    try:
        df = df.select(pl.all().name.map(lambda x: x.strip().upper().replace("_X0009_", "")))
        if year == "2024":
            df = (df
                  .select(["CODIGO META", "COMPROMISOS", "CLASIFICACION RECURSOS"])
                  .with_columns(pl.col("CODIGO META").fill_null(""))
                  .filter(pl.col("CODIGO META") != "", pl.col("CODIGO META").str.starts_with("MT"))
                  .rename({"COMPROMISOS": "RP"}))
        elif year == "2025":
            df = (df
                  .select(["PAGOS REGALIAS", "CODIGO META", "CLASIFICACION RECURSOS"])
                  .rename({"PAGOS REGALIAS": "RP"})
                  .with_columns(pl.col("CODIGO META").fill_null(""))
                  .filter(pl.col("CODIGO META") != ""))
        elif year == "2026":
            df = (df
                  .filter(
                      (pl.col("ULTIMA FECHA PAGO") >= pl.date(2026, 1, 1)) &
                      (pl.col("ULTIMA FECHA PAGO") <= pl.date(2026, 12, 31))
                  )
                  .select(["PAGO EJECUTADO VALOR", "CODIGO META", "CLASIFICACION RECURSOS"])
                  .rename({"PAGO EJECUTADO VALOR": "RP"})
                  .with_columns(pl.col("CODIGO META").fill_null(""))
                  .filter(pl.col("CODIGO META") != ""))
        return df
    except Exception:
        return None


def process_hacienda(source, table_name: str, year: str) -> Optional[pl.DataFrame]:
    df = read_excel_safe(source, table_name)
    if df is None:
        return None
    try:
        if year == "2024":
            df = (df
                  .select(["RP", "CODIGO META", "CLASIFICACION RECURSOS"])
                  .with_columns(pl.col("CODIGO META", "CLASIFICACION RECURSOS").fill_null(""))
                  .filter(pl.col("CODIGO META") != "", pl.col("CLASIFICACION RECURSOS") != ""))
        else:
            df = (df
                  .with_columns(
                      pl.col("PROYECTO ARCHIVADO", "CODIGO META", "CLASIFICACION RECURSOS", "SE VA A CARGAR EN PI").fill_null(""),
                      pl.when(pl.col("DISTRIBUIR DE FORMA EQUITATIVA") == "SI")
                        .then(pl.col("RP") / 2)
                        .otherwise(pl.col("RP"))
                  )
                  .filter(
                      pl.col("PROYECTO ARCHIVADO") == "",
                      pl.col("CODIGO META") != "",
                      pl.col("CLASIFICACION RECURSOS") != "",
                      pl.col("SE VA A CARGAR EN PI") == "",
                  )
                  .select(["CODIGO META", "CLASIFICACION RECURSOS", "RP"]))
        return df
    except Exception:
        return None


def merge_ejecucion(reg: Optional[pl.DataFrame], hac: Optional[pl.DataFrame], col_name: str) -> pl.DataFrame:
    frames = [f for f in [reg, hac] if f is not None]
    if not frames:
        return pl.DataFrame({"CODIGO META": pl.Series([], dtype=pl.Utf8),
                             col_name: pl.Series([], dtype=pl.Float64)})
    return (
        pl.concat(frames, how="diagonal")
        .group_by("CODIGO META")
        .agg(pl.col("RP").sum().alias(col_name))
    )


ORDEN_FUENTES = pl.DataFrame({
    "Clasificacion Recursos": ["COFINANCIACION MUNICIPIO","ICDE","OTRAS FUENTES","SGP APSB","SGP SALUD",
                                "SGP EDUCACION","REGALIAS","COFINANCIACION NACION","ICLD","CREDITO"],
    "Orden": [1, 5, 3, 7, 9, 8, 10, 2, 6, 4],
    "Tipo Fuente": ["Otras Fuentes","Recursos Propios","Otras Fuentes",
                    "Sistema General de Participaciones (SGP)",
                    "Sistema General de Participaciones (SGP)",
                    "Sistema General de Participaciones (SGP)",
                    "Sistema General de Regalias",
                    "Otras Fuentes","Recursos Propios","Recursos del Credito"],
})


@st.cache_data(show_spinner=False)
def load_and_process(
    pi_bytes: bytes,
    h24_bytes, r24_bytes,
    h25_bytes, r25_bytes,
    h26_bytes, r26_bytes,
):
    errors = []

    # Plan Indicativo
    pi = read_excel_safe(pi_bytes, "tblPlanIndicativo_2")
    if pi is None:
        errors.append(("Plan Indicativo", "tblPlanIndicativo_2"))
        return None, errors

    orden_lineas    = read_excel_safe(pi_bytes, "orden_lineas")
    orden_sectores  = read_excel_safe(pi_bytes, "orden_sectores")
    orden_programas = read_excel_safe(pi_bytes, "orden_programas")
    homologacion    = read_excel_safe(pi_bytes, "HomologacionSecretarias")

    # Columnas ejecucion fisica
    cols_fisicas = [
        "Codigo Meta", "Linea Estrategica", "Sector PDD", "Numero Programa PDD", "Programa PDD",
        "Meta de cuatrenio", "Tipo de Acumulacion", "Responsable",
        "Meta Fisica Esperada 2024", "Meta Fisica Esperada 2025",
        "Meta Fisica Esperada 2026", "Meta Fisica Esperada 2027",
        "PROYECTOS 2024", "PROYECTOS 2025", "PROYECTOS/GESTIONES PROGRAMADAS 2026",
        "PROYECTOS 2026", "PROYECTOS 2027",
        "EJECUCION 2024", "PORCENTAJE DE EJECUCION 2024", "CATEGORIA DE EJECUCION FISICA 2024",
        "EJECUCION 2025", "PORCENTAJE DE EJECUCION 2025", "CATEGORIA DE EJECUCION FISICA 2025",
        "EJECUCION 2026", "PORCENTAJE DE EJECUCION 2026", "CATEGORIA DE EJECUCION FISICA 2026",
        "EJECUCION ACUMULADA", "PORCENTAJE DE EJECUCION ACUMULADA", "CATEGORIA DE EJECUCION ACUMULADA",
        # Also try with accents as they appear in the real file
        "Línea Estratégica", "Meta Física Esperada 2024", "Meta Física Esperada 2025",
        "Meta Física Esperada 2026", "Meta Física Esperada 2027",
        "EJECUCIÓN 2024", "PORCENTAJE DE EJECUCIÓN 2024", "CATEGORÍA DE EJECUCIÓN FÍSICA 2024",
        "EJECUCIÓN 2025", "PORCENTAJE DE EJECUCIÓN 2025", "CATEGORÍA DE EJECUCIÓN FÍSICA 2025",
        "EJECUCIÓN 2026", "PORCENTAJE DE EJECUCIÓN 2026", "CATEGORÍA DE EJECUCIÓN FÍSICA 2026",
        "EJECUCIÓN ACUMULADA", "PORCENTAJE DE EJECUCIÓN ACUMULADA", "CATEGORÍA DE EJECUCIÓN ACUMULADA",
        "Tipo de Acumulación",
    ]
    available_cols = [c for c in dict.fromkeys(cols_fisicas) if c in pi.columns]
    columnas_fisicas = pi.select(available_cols)

    # Programacion financiera
    pi_lower = pi.select(pl.all().name.map(lambda x: x.strip().lower()))
    prog_cols_names = ["codigo meta"] + [c for c in pi_lower.columns if c.startswith("programaci")]
    prog_cols = pi_lower.select(prog_cols_names) if all(c in pi_lower.columns for c in ["codigo meta"]) else pi_lower

    suffix_map = {"24": "2024", "25": "2025", "26": "2026", "27": "2027"}
    select_exprs = [pl.col("codigo meta")]
    for suf, yr in suffix_map.items():
        expr = programacion_financiera_expr(prog_cols, suf)
        select_exprs.append(expr.alias(f"Programacion Financiera {yr}"))

    prog_financ = prog_cols.select(select_exprs)

    # Ejecuciones financieras
    reg24 = process_regalias(r24_bytes, "2024", "EjecucionRegalias")       if r24_bytes else None
    hac24 = process_hacienda(h24_bytes, "EjecucionHaciendaDiciembre", "2024") if h24_bytes else None
    reg25 = process_regalias(r25_bytes, "2025", "Pagos_Regalias_2025")     if r25_bytes else None
    hac25 = process_hacienda(h25_bytes, "EjecucionHaciendaDiciembre2025", "2025") if h25_bytes else None
    reg26 = process_regalias(r26_bytes, "2026", "Pagos_Regalias_2026")     if r26_bytes else None
    hac26 = process_hacienda(h26_bytes, "EjecucionHacienda2026", "2026")   if h26_bytes else None

    ef24 = merge_ejecucion(reg24, hac24, "Ejecucion Financiera 2024")
    ef25 = merge_ejecucion(reg25, hac25, "Ejecucion Financiera 2025")
    ef26 = merge_ejecucion(reg26, hac26, "Ejecucion Financiera 2026")

    prog_financ = (prog_financ
        .join(ef24, left_on="codigo meta", right_on="CODIGO META", how="left")
        .join(ef25, left_on="codigo meta", right_on="CODIGO META", how="left")
        .join(ef26, left_on="codigo meta", right_on="CODIGO META", how="left")
        .with_columns(
            pl.col("Ejecucion Financiera 2024",
                   "Ejecucion Financiera 2025",
                   "Ejecucion Financiera 2026").fill_null(0)
        )
    )

    # DataFrame principal
    pff = (
        columnas_fisicas
        .join(prog_financ, left_on="Codigo Meta", right_on="codigo meta", how="left")
    )
    # Fill numeric cols
    meta_cols = [c for c in pff.columns if "Meta" in c and "Esperada" in c]
    if meta_cols:
        pff = pff.with_columns([pl.col(c).fill_null(0) for c in meta_cols])

    return {
        "pff": pff,
        "prog_financ": prog_financ,
        "orden_lineas": orden_lineas,
        "orden_sectores": orden_sectores,
        "orden_programas": orden_programas,
        "homologacion": homologacion,
    }, errors


# -------------------------------------------------
# VISUALIZACIONES
# -------------------------------------------------
def gauge_chart(value: float, title: str, color: str):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value * 100,
        number={"suffix": "%", "font": {"size": 32, "color": color, "family": "Sora"}},
        title={"text": title, "font": {"size": 13, "color": "#6b7280"}},
        gauge={
            "axis": {"range": [0, 100], "tickfont": {"size": 10}},
            "bar": {"color": color, "thickness": 0.3},
            "bgcolor": "#f3f4f6",
            "borderwidth": 0,
            "steps": [
                {"range": [0,  30],  "color": "#fee2e2"},
                {"range": [30, 60],  "color": "#fef3c7"},
                {"range": [60, 90],  "color": "#d1fae5"},
                {"range": [90, 100], "color": "#a7f3d0"},
            ],
        }
    ))
    fig.update_layout(height=200, margin=dict(t=40, b=10, l=20, r=20), paper_bgcolor="white")
    return fig


def bar_h_chart(df: pd.DataFrame, x_col: str, y_col: str, pct_col: str,
                title: str, prog_col: str = None, ejec_col: str = None):
    colors = [color_pct(v) for v in df[pct_col].tolist()]
    hover = "<b>%{y}</b><br>Avance: %{text}<extra></extra>"
    customdata = None
    if prog_col and ejec_col and prog_col in df.columns and ejec_col in df.columns:
        customdata = df[[prog_col, ejec_col]].values
        hover = (
            "<b>%{y}</b><br>"
            "Programacion: $%{customdata[0]:,.0f}<br>"
            "Ejecucion: $%{customdata[1]:,.0f}<br>"
            "Avance: %{text}<extra></extra>"
        )
    fig = go.Figure(go.Bar(
        x=df[x_col],
        y=df[y_col],
        orientation="h",
        marker_color=colors,
        text=[f"{v*100:.1f}%" for v in df[pct_col]],
        textposition="outside",
        hovertemplate=hover,
        customdata=customdata,
    ))
    fig.update_layout(
        title=title,
        xaxis_title="Ejecucion ($)",
        height=max(300, len(df) * 42),
        margin=dict(l=20, r=90, t=50, b=20),
        paper_bgcolor="white",
        plot_bgcolor="#fafafa",
        font={"family": "DM Sans"},
    )
    return fig


def pie_chart(labels, values, title, colors_list):
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker_colors=colors_list,
        hole=0.45,
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>%{value}<br>%{percent}<extra></extra>",
    ))
    fig.update_layout(
        title=title, height=370,
        margin=dict(t=50, b=10, l=10, r=10),
        paper_bgcolor="white",
        font={"family": "DM Sans"},
    )
    return fig


# -------------------------------------------------
# SIDEBAR
# -------------------------------------------------
with st.sidebar:
    st.markdown("## Dashboard PDD")
    st.markdown("#### Reporte de Avance 2024-2027")
    st.markdown("---")

    st.markdown("### Fuente de datos")
    modo_carga = st.radio(
        "Selecciona como cargar los archivos:",
        ["GitHub (datos en vivo)", "Subir archivos manualmente"],
        index=0,
    )

    st.markdown("---")
    st.markdown("### Filtros")

    filtro_vigencia = st.selectbox(
        "Vigencia a analizar:",
        VIGENCIAS,
        index=2,
    )

    filtro_linea = st.multiselect(
        "Linea Estrategica:",
        options=[],
        placeholder="Todas las lineas",
        key="filtro_linea",
    )

    filtro_sector = st.multiselect(
        "Sector PDD:",
        options=[],
        placeholder="Todos los sectores",
        key="filtro_sector",
    )

    filtro_responsable = st.multiselect(
        "Dependencia Responsable:",
        options=[],
        placeholder="Todas",
        key="filtro_responsable",
    )

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.75rem; color:#94a3b8; line-height:1.6'>
    Los archivos de 2024 y 2025 son vigencias cerradas.<br>
    Solo el archivo 2026 se actualiza periodicamente.
    </div>
    """, unsafe_allow_html=True)

# -------------------------------------------------
# ENCABEZADO
# -------------------------------------------------
st.markdown(f"""
<div class="main-header">
    <h1>Reporte de Avance del Plan de Desarrollo</h1>
    <p>Ejecucion Fisica y Financiera &middot; Vigencia <strong>{filtro_vigencia}</strong> &middot; Cuatrienio 2024-2027</p>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------
# CARGA DE ARCHIVOS
# -------------------------------------------------
pi_bytes = h24_bytes = r24_bytes = h25_bytes = r25_bytes = h26_bytes = r26_bytes = None

if modo_carga == "GitHub (datos en vivo)":
    with st.expander("Configurar URLs de GitHub", expanded=False):
        st.info("Ingresa las URLs crudas (raw) de los archivos en tu repositorio de GitHub.")
        col1, col2 = st.columns(2)
        with col1:
            url_pi  = st.text_input("Plan Indicativo (.xlsx)",    key="url_pi",  placeholder="https://raw.githubusercontent.com/...")
            url_h24 = st.text_input("Hacienda 2024 (.xlsx)",      key="url_h24", placeholder="https://raw.githubusercontent.com/...")
            url_r24 = st.text_input("Regalias 2024 (.xlsx)",      key="url_r24", placeholder="https://raw.githubusercontent.com/...")
            url_h25 = st.text_input("Hacienda 2025 (.xlsx)",      key="url_h25", placeholder="https://raw.githubusercontent.com/...")
        with col2:
            url_r25 = st.text_input("Regalias 2025 (.xlsx)",      key="url_r25", placeholder="https://raw.githubusercontent.com/...")
            url_h26 = st.text_input("Hacienda 2026 (.xlsx)",      key="url_h26", placeholder="https://raw.githubusercontent.com/...")
            url_r26 = st.text_input("Regalias 2026 (.xlsx)",      key="url_r26", placeholder="https://raw.githubusercontent.com/...")

    if st.session_state.get("url_pi"):
        with st.spinner("Descargando archivos desde GitHub..."):
            pi_bytes  = fetch_github_file(st.session_state["url_pi"])
            h24_bytes = fetch_github_file(st.session_state["url_h24"]) if st.session_state.get("url_h24") else None
            r24_bytes = fetch_github_file(st.session_state["url_r24"]) if st.session_state.get("url_r24") else None
            h25_bytes = fetch_github_file(st.session_state["url_h25"]) if st.session_state.get("url_h25") else None
            r25_bytes = fetch_github_file(st.session_state["url_r25"]) if st.session_state.get("url_r25") else None
            h26_bytes = fetch_github_file(st.session_state["url_h26"]) if st.session_state.get("url_h26") else None
            r26_bytes = fetch_github_file(st.session_state["url_r26"]) if st.session_state.get("url_r26") else None

        if pi_bytes is None:
            st.error("No se pudo descargar el Plan Indicativo. Verifica la URL.")
    else:
        st.markdown("""
        <div class="upload-zone">
            Ingresa las URLs de tus archivos en GitHub para comenzar.<br>
            <small>Usa el panel <strong>Configurar URLs</strong> de arriba.</small>
        </div>
        """, unsafe_allow_html=True)

else:
    st.markdown("### Carga de Archivos")
    with st.expander("Archivos de vigencias cerradas (2024-2025)", expanded=True):
        st.caption("Estos archivos corresponden a vigencias ya cerradas y no se modifican.")
        col1, col2 = st.columns(2)
        with col1:
            pi_file  = st.file_uploader("Plan Indicativo 2024-2027", type=["xlsx"], key="pi")
            h24_file = st.file_uploader("Ejecucion Hacienda 2024",   type=["xlsx"], key="h24")
            r24_file = st.file_uploader("Regalias 2024",              type=["xlsx"], key="r24")
        with col2:
            h25_file = st.file_uploader("Ejecucion Hacienda 2025",   type=["xlsx"], key="h25")
            r25_file = st.file_uploader("Regalias 2025",              type=["xlsx"], key="r25")

    with st.expander("Archivo de vigencia actual (2026)", expanded=True):
        st.caption("Estos archivos se actualizan periodicamente con nueva informacion.")
        col1, col2 = st.columns(2)
        with col1:
            h26_file = st.file_uploader("Ejecucion Hacienda 2026",   type=["xlsx"], key="h26")
        with col2:
            r26_file = st.file_uploader("Regalias 2026",              type=["xlsx"], key="r26")

    pi_bytes  = pi_file.read()  if pi_file  else None
    h24_bytes = h24_file.read() if h24_file else None
    r24_bytes = r24_file.read() if r24_file else None
    h25_bytes = h25_file.read() if h25_file else None
    r25_bytes = r25_file.read() if r25_file else None
    h26_bytes = h26_file.read() if h26_file else None
    r26_bytes = r26_file.read() if r26_file else None

    if not pi_bytes:
        st.markdown("""
        <div class="upload-zone">
            Carga el <strong>Plan Indicativo</strong> para comenzar a visualizar el dashboard.<br>
            <small>Los demas archivos son opcionales. El dashboard mostrara los datos disponibles.</small>
        </div>
        """, unsafe_allow_html=True)

# -------------------------------------------------
# PROCESAMIENTO Y VISUALIZACION
# -------------------------------------------------
if pi_bytes:
    with st.spinner("Procesando datos..."):
        result, errors = load_and_process(
            pi_bytes,
            h24_bytes, r24_bytes,
            h25_bytes, r25_bytes,
            h26_bytes, r26_bytes,
        )

    if result is None:
        st.error("No se pudo cargar el Plan Indicativo. Verifica el archivo.")
        show_schema_error("Plan Indicativo", SCHEMAS["Plan Indicativo"])
        st.stop()

    pff             = result["pff"]
    orden_lineas    = result["orden_lineas"]
    orden_sectores  = result["orden_sectores"]
    orden_programas = result["orden_programas"]
    homologacion    = result["homologacion"]

    # Detectar nombres de columnas con o sin tildes
    def col(base: str) -> Optional[str]:
        """Retorna el nombre real de la columna buscando con y sin tildes."""
        if base in pff.columns:
            return base
        # Try accented variants
        variants = {
            "Linea Estrategica":              "Línea Estratégica",
            "Meta Fisica Esperada":           "Meta Física Esperada",
            "PORCENTAJE DE EJECUCION":        "PORCENTAJE DE EJECUCIÓN",
            "CATEGORIA DE EJECUCION FISICA":  "CATEGORÍA DE EJECUCIÓN FÍSICA",
            "EJECUCION":                      "EJECUCIÓN",
            "PORCENTAJE DE EJECUCION ACUMULADA": "PORCENTAJE DE EJECUCIÓN ACUMULADA",
        }
        for plain, accented in variants.items():
            if plain in base:
                candidate = base.replace(plain, accented)
                if candidate in pff.columns:
                    return candidate
        return None

    # Resolve column names dynamically
    col_linea      = col("Linea Estrategica")    or "Línea Estratégica"
    col_meta_esp   = col(f"Meta Fisica Esperada {filtro_vigencia}") or f"Meta Física Esperada {filtro_vigencia}"
    col_pct_fisica = col(f"PORCENTAJE DE EJECUCION {filtro_vigencia}") or f"PORCENTAJE DE EJECUCIÓN {filtro_vigencia}"
    col_cat        = col(f"CATEGORIA DE EJECUCION FISICA {filtro_vigencia}") or f"CATEGORÍA DE EJECUCIÓN FÍSICA {filtro_vigencia}"
    col_pct_acum   = col("PORCENTAJE DE EJECUCION ACUMULADA") or "PORCENTAJE DE EJECUCIÓN ACUMULADA"

    # Filtros de sidebar - actualizar opciones
    if col_linea in pff.columns:
        lineas_opts = sorted(pff[col_linea].drop_nulls().unique().to_list())
        with st.sidebar:
            filtro_linea = st.multiselect("Linea Estrategica:", lineas_opts,
                                           default=filtro_linea, key="filtro_linea_real")
    if "Sector PDD" in pff.columns:
        sect_opts = sorted(pff["Sector PDD"].drop_nulls().unique().to_list())
        with st.sidebar:
            filtro_sector = st.multiselect("Sector PDD:", sect_opts,
                                            default=filtro_sector, key="filtro_sector_real")
    if "Responsable" in pff.columns:
        resp_opts = sorted(pff["Responsable"].drop_nulls().unique().to_list())
        with st.sidebar:
            filtro_responsable = st.multiselect("Dependencia Responsable:", resp_opts,
                                                 default=filtro_responsable, key="filtro_resp_real")

    # Aplicar filtros
    pff_f = pff.clone()
    if filtro_linea and col_linea in pff_f.columns:
        pff_f = pff_f.filter(pl.col(col_linea).is_in(filtro_linea))
    if filtro_sector and "Sector PDD" in pff_f.columns:
        pff_f = pff_f.filter(pl.col("Sector PDD").is_in(filtro_sector))
    if filtro_responsable and "Responsable" in pff_f.columns:
        pff_f = pff_f.filter(pl.col("Responsable").is_in(filtro_responsable))

    # Columnas financieras (pueden ser con o sin acento)
    col_ejec_fin = f"Ejecucion Financiera {filtro_vigencia}"
    col_prog_fin = f"Programacion Financiera {filtro_vigencia}"

    # ------------------------------------------
    # TABS
    # ------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "Resumen General",
        "Ejecucion Financiera",
        "Ejecucion Fisica",
        "Detalle por Dependencia",
    ])

    # ==========================================
    # TAB 1: RESUMEN GENERAL
    # ==========================================
    with tab1:
        section_title(f"Indicadores Clave - Vigencia {filtro_vigencia}")

        n_total      = len(pff_f)
        n_programadas = 0
        avance_vig   = 0.0
        avance_acum  = 0.0
        n_superiores = 0

        if col_meta_esp in pff_f.columns:
            n_programadas = int(pff_f.filter(pl.col(col_meta_esp).fill_null(0) != 0).height)

        if col_pct_fisica in pff_f.columns and col_meta_esp in pff_f.columns:
            tmp = pff_f.filter(pl.col(col_meta_esp).fill_null(0) != 0)
            avance_vig = float(tmp.select(pl.col(col_pct_fisica).fill_null(0).mean()).item() or 0)

        if col_pct_acum in pff_f.columns:
            avance_acum = float(pff_f.select(pl.col(col_pct_acum).fill_null(0).mean()).item() or 0)

        if col_cat in pff_f.columns and col_meta_esp in pff_f.columns:
            n_superiores = int(
                pff_f.filter(pl.col(col_meta_esp).fill_null(0) != 0)
                     .filter(pl.col(col_cat) == "Superior").height
            )

        ejec_fin_val = 0.0
        prog_fin_val = 0.0
        pct_fin      = 0.0
        if col_ejec_fin in pff_f.columns:
            ejec_fin_val = float(pff_f.select(pl.col(col_ejec_fin).sum()).item() or 0)
        if col_prog_fin in pff_f.columns:
            prog_fin_val = float(pff_f.select(pl.col(col_prog_fin).sum()).item() or 0)
        if prog_fin_val > 0:
            pct_fin = ejec_fin_val / prog_fin_val

        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            kpi_card("Metas Totales", str(n_total), "azul",
                     "Numero total de indicadores de producto del PDD en el filtro actual.")
        with k2:
            kpi_card(f"Metas Programadas {filtro_vigencia}", str(n_programadas), "cyan",
                     f"Metas con Meta Fisica Esperada distinta de cero en {filtro_vigencia}.")
        with k3:
            kpi_card(f"Avance Fisico {filtro_vigencia}", fmt_pct(avance_vig), "verde",
                     f"Promedio del porcentaje de ejecucion fisica de las metas programadas en {filtro_vigencia}.")
        with k4:
            kpi_card("Avance Acumulado Cuatrienio", fmt_pct(avance_acum), "naranja",
                     "Promedio del porcentaje de ejecucion acumulada sobre todas las metas, desde 2024.")
        with k5:
            kpi_card(f"Metas al 100% en {filtro_vigencia}", str(n_superiores), "cafe",
                     f"Metas clasificadas como 'Superior' en {filtro_vigencia}, es decir, con ejecucion igual o mayor al 100%.")

        st.markdown('<hr class="custom">', unsafe_allow_html=True)

        cg1, cg2, cg3 = st.columns(3)
        with cg1:
            st.plotly_chart(
                gauge_chart(avance_vig, f"Ejecucion Fisica {filtro_vigencia}", color_pct(avance_vig)),
                use_container_width=True, key="g1",
            )
        with cg2:
            st.plotly_chart(
                gauge_chart(avance_acum, "Ejecucion Acumulada Cuatrienio", color_pct(avance_acum)),
                use_container_width=True, key="g2",
            )
        with cg3:
            st.plotly_chart(
                gauge_chart(pct_fin, f"Ejecucion Financiera {filtro_vigencia}", color_pct(pct_fin)),
                use_container_width=True, key="g3",
            )

        st.caption(
            f"Ejecucion Fisica: promedio del avance de metas con programacion en {filtro_vigencia}. "
            f"Ejecucion Financiera: recursos ejecutados sobre recursos programados. "
            f"Acumulada: avance respecto a la meta cuatrienal 2024-2027."
        )

        # Distribucion por categoria
        if col_cat in pff_f.columns and col_meta_esp in pff_f.columns:
            section_title(f"Distribucion de Metas por Categoria de Ejecucion {filtro_vigencia}")
            cat_df = (
                pff_f
                .filter(pl.col(col_meta_esp).fill_null(0) != 0)
                .group_by(col_cat)
                .agg(pl.col("Codigo Meta").len().alias("n"))
                .drop_nulls()
                .to_pandas()
            )
            if not cat_df.empty:
                cat_labels = cat_df[col_cat].tolist()
                cat_values = cat_df["n"].tolist()
                cat_colors = [CAT_COLORS.get(l, COLORS["gris_texto"]) for l in cat_labels]
                cp1, cp2 = st.columns([1.2, 1])
                with cp1:
                    st.plotly_chart(
                        pie_chart(cat_labels, cat_values, "Categorias de ejecucion", cat_colors),
                        use_container_width=True, key="pie_cat",
                    )
                with cp2:
                    st.markdown("**Que significa cada categoria?**")
                    categorias_info = {
                        "Superior":      "100% o mas",
                        "Alto":          "80% - 99%",
                        "Medio":         "60% - 79%",
                        "Bajo":          "30% - 59%",
                        "Critico":       "Menos del 30%",
                        "Sin Programar": "Sin meta asignada en la vigencia",
                    }
                    for cat, rango in categorias_info.items():
                        color = CAT_COLORS.get(cat, COLORS["gris_texto"])
                        st.markdown(
                            f'<span style="display:inline-block;padding:2px 10px;border-radius:999px;'
                            f'background:{color}22;color:{color};border:1px solid {color};font-size:0.8rem;font-weight:600">'
                            f'{cat}</span> &nbsp; {rango}',
                            unsafe_allow_html=True
                        )
                        st.write("")

    # ==========================================
    # TAB 2: EJECUCION FINANCIERA
    # ==========================================
    with tab2:
        if col_ejec_fin not in pff_f.columns or col_prog_fin not in pff_f.columns:
            st.info("No hay datos de ejecucion financiera disponibles para esta vigencia. Verifica que los archivos de hacienda y regalias esten cargados.")
        else:
            section_title(f"Ejecucion Financiera por Linea Estrategica - {filtro_vigencia}")
            if col_linea in pff_f.columns and orden_lineas is not None:
                lineas_fin = (
                    pff_f
                    .group_by(col_linea).agg(
                        pl.col(col_prog_fin).sum(),
                        pl.col(col_ejec_fin).sum(),
                    )
                    .join(orden_lineas, left_on=col_linea, right_on="Línea Estratégica", how="inner")
                    .with_columns(
                        pl.when(pl.col(col_prog_fin) == 0)
                          .then(0.0)
                          .otherwise(pl.col(col_ejec_fin) / pl.col(col_prog_fin))
                          .alias("Porcentaje de Ejecucion Financiera")
                    )
                    .sort("Orden Linea")
                    .to_pandas()
                )
                if not lineas_fin.empty:
                    st.plotly_chart(
                        bar_h_chart(lineas_fin, col_ejec_fin, col_linea,
                                    "Porcentaje de Ejecucion Financiera",
                                    f"Ejecucion por Linea Estrategica - {filtro_vigencia}",
                                    prog_col=col_prog_fin, ejec_col=col_ejec_fin),
                        use_container_width=True, key="bar_lineas_fin",
                    )
                    st.caption(
                        "Como se calcula: Ejecucion Financiera dividida entre Programacion Financiera de la vigencia. "
                        "Un porcentaje de 0% en una linea sin presupuesto no indica problema, simplemente no tenia recursos asignados."
                    )

            section_title(f"Ejecucion Financiera por Sector PDD - {filtro_vigencia}")
            if "Sector PDD" in pff_f.columns and orden_sectores is not None:
                sect_fin = (
                    pff_f
                    .group_by("Sector PDD").agg(
                        pl.col(col_prog_fin).sum(),
                        pl.col(col_ejec_fin).sum(),
                    )
                    .join(orden_sectores, on="Sector PDD", how="inner")
                    .with_columns(
                        pl.when(pl.col(col_prog_fin) == 0)
                          .then(0.0)
                          .otherwise(pl.col(col_ejec_fin) / pl.col(col_prog_fin))
                          .alias("Porcentaje de Ejecucion Financiera")
                    )
                    .sort("Orden Sector")
                    .to_pandas()
                )
                if not sect_fin.empty:
                    st.plotly_chart(
                        bar_h_chart(sect_fin, col_ejec_fin, "Sector PDD",
                                    "Porcentaje de Ejecucion Financiera",
                                    f"Ejecucion por Sector - {filtro_vigencia}",
                                    prog_col=col_prog_fin, ejec_col=col_ejec_fin),
                        use_container_width=True, key="bar_sect_fin",
                    )

            section_title("Ejecucion Financiera Acumulada 2024-2026")
            cols_ejec = [c for c in pff_f.columns if "Ejecucion Financiera 20" in c]
            cols_prog  = [c for c in pff_f.columns if "Programacion Financiera 20" in c]
            if cols_ejec:
                years     = sorted(set(c.split(" ")[-1] for c in cols_ejec))
                ejec_vals = [float(pff_f.select(pl.col(f"Ejecucion Financiera {y}").sum()).item() or 0) for y in years]
                prog_vals = [float(pff_f.select(pl.col(f"Programacion Financiera {y}").sum()).item() or 0)
                             for y in years if f"Programacion Financiera {y}" in pff_f.columns]
                while len(prog_vals) < len(years):
                    prog_vals.append(0)

                fig_acum = go.Figure()
                fig_acum.add_trace(go.Bar(name="Programacion", x=years, y=prog_vals,
                                           marker_color=COLORS["cyan"], opacity=0.75))
                fig_acum.add_trace(go.Bar(name="Ejecucion",    x=years, y=ejec_vals,
                                           marker_color=COLORS["azul_medio"]))
                fig_acum.update_layout(
                    barmode="group",
                    title="Programacion vs Ejecucion Financiera por Ano",
                    yaxis_title="Valor ($)", height=380,
                    paper_bgcolor="white", plot_bgcolor="#fafafa",
                    font={"family": "DM Sans"},
                    legend=dict(orientation="h", y=1.1),
                )
                st.plotly_chart(fig_acum, use_container_width=True, key="bar_acum")
                st.caption(
                    "Programacion: recursos apropiados en el plan de desarrollo para cada vigencia. "
                    "Ejecucion: compromisos (RP) y pagos de regalias registrados en los informes financieros."
                )

    # ==========================================
    # TAB 3: EJECUCION FISICA
    # ==========================================
    with tab3:
        section_title(f"Ejecucion Fisica por Linea Estrategica - {filtro_vigencia}")

        if col_pct_fisica in pff_f.columns and col_linea in pff_f.columns and col_meta_esp in pff_f.columns:
            lineas_fis = (
                pff_f
                .filter(pl.col(col_meta_esp).fill_null(0) != 0)
                .group_by(col_linea)
                .agg(
                    pl.col(col_pct_fisica).fill_null(0).mean().alias("Promedio Ejecucion"),
                    pl.col("Codigo Meta").len().alias("N Metas"),
                )
                .to_pandas()
                .sort_values("Promedio Ejecucion", ascending=True)
            )
            if not lineas_fis.empty:
                colors_fis = [color_pct(v) for v in lineas_fis["Promedio Ejecucion"]]
                fig_fis = go.Figure(go.Bar(
                    x=lineas_fis["Promedio Ejecucion"] * 100,
                    y=lineas_fis[col_linea],
                    orientation="h",
                    marker_color=colors_fis,
                    text=[f"{v*100:.1f}%" for v in lineas_fis["Promedio Ejecucion"]],
                    textposition="outside",
                    customdata=lineas_fis[["N Metas"]].values,
                    hovertemplate="<b>%{y}</b><br>Avance: %{x:.1f}%<br>Metas programadas: %{customdata[0]}<extra></extra>",
                ))
                fig_fis.update_layout(
                    xaxis_title="% Promedio Ejecucion Fisica",
                    height=max(300, len(lineas_fis) * 48),
                    paper_bgcolor="white", plot_bgcolor="#fafafa",
                    font={"family": "DM Sans"},
                    margin=dict(l=20, r=90, t=30, b=20),
                )
                st.plotly_chart(fig_fis, use_container_width=True, key="bar_fis_lineas")
                st.caption(
                    f"Como se calcula: promedio del porcentaje de avance de cada indicador de producto "
                    f"con programacion en {filtro_vigencia}. Solo se incluyen metas con Meta Fisica Esperada mayor a cero."
                )

        section_title(f"Ejecucion Fisica por Sector PDD - {filtro_vigencia}")
        if col_pct_fisica in pff_f.columns and "Sector PDD" in pff_f.columns and col_meta_esp in pff_f.columns:
            sect_fis = (
                pff_f
                .filter(pl.col(col_meta_esp).fill_null(0) != 0)
                .group_by("Sector PDD")
                .agg(
                    pl.col(col_pct_fisica).fill_null(0).mean().alias("Promedio Ejecucion"),
                    pl.col("Codigo Meta").len().alias("N Metas"),
                )
                .to_pandas()
                .sort_values("Promedio Ejecucion", ascending=True)
            )
            if not sect_fis.empty:
                colors_sec = [color_pct(v) for v in sect_fis["Promedio Ejecucion"]]
                fig_sec = go.Figure(go.Bar(
                    x=sect_fis["Promedio Ejecucion"] * 100,
                    y=sect_fis["Sector PDD"],
                    orientation="h",
                    marker_color=colors_sec,
                    text=[f"{v*100:.1f}%" for v in sect_fis["Promedio Ejecucion"]],
                    textposition="outside",
                    customdata=sect_fis[["N Metas"]].values,
                    hovertemplate="<b>%{y}</b><br>Avance: %{x:.1f}%<br>Metas: %{customdata[0]}<extra></extra>",
                ))
                fig_sec.update_layout(
                    xaxis_title="% Promedio Ejecucion Fisica",
                    height=max(300, len(sect_fis) * 48),
                    paper_bgcolor="white", plot_bgcolor="#fafafa",
                    font={"family": "DM Sans"},
                    margin=dict(l=20, r=90, t=30, b=20),
                )
                st.plotly_chart(fig_sec, use_container_width=True, key="bar_fis_sect")

        section_title("Tabla de Metas PDD")
        display_cols = [c for c in [
            "Codigo Meta", col_linea, "Sector PDD", "Programa PDD", "Responsable",
            col_meta_esp, f"EJECUCION {filtro_vigencia}", f"EJECUCIÓN {filtro_vigencia}",
            col_pct_fisica, col_cat, col_pct_acum,
        ] if c in pff_f.columns]
        display_cols = list(dict.fromkeys(display_cols))  # deduplicate
        tabla = pff_f.select(display_cols).to_pandas()
        if col_pct_fisica in tabla.columns:
            tabla[col_pct_fisica] = (tabla[col_pct_fisica].fillna(0) * 100).round(1).astype(str) + "%"
        if col_pct_acum in tabla.columns:
            tabla[col_pct_acum] = (tabla[col_pct_acum].fillna(0) * 100).round(1).astype(str) + "%"
        st.dataframe(tabla, use_container_width=True, height=440)

    # ==========================================
    # TAB 4: DEPENDENCIAS
    # ==========================================
    with tab4:
        section_title(f"Avance por Dependencia Responsable - {filtro_vigencia}")

        if col_pct_fisica not in pff_f.columns or col_meta_esp not in pff_f.columns:
            st.info("No hay datos suficientes para mostrar el avance por dependencia.")
        else:
            base_dep = (
                pff_f
                .filter(pl.col(col_meta_esp).fill_null(0) != 0)
                .with_columns(
                    pl.when(pl.col(col_cat) == "Superior").then(1).otherwise(0).alias("Metas_100")
                    if col_cat in pff_f.columns else pl.lit(0).alias("Metas_100"),
                    pl.lit(1).alias("Metas_Prog"),
                )
                .group_by(pl.col("Responsable").str.strip_chars())
                .agg(
                    pl.col(col_pct_fisica).fill_null(0).mean().alias(f"Ejec {filtro_vigencia}"),
                    pl.col("Metas_Prog").sum().alias("Metas Programadas"),
                    pl.col("Metas_100").sum().alias("Metas al 100"),
                    pl.col(col_pct_acum).fill_null(0).mean().alias("Ejec Acumulada")
                    if col_pct_acum in pff_f.columns else pl.lit(0.0).alias("Ejec Acumulada"),
                )
            )

            if homologacion is not None:
                hom_cols = homologacion.columns
                resp_col_hom = next((c for c in hom_cols if "Responsable" in c and "PI" in c), None)
                if resp_col_hom:
                    base_dep = base_dep.join(
                        homologacion.rename({resp_col_hom: "Responsable"}),
                        on="Responsable", how="left"
                    )

            dep_pd = base_dep.to_pandas()

            if dep_pd.empty:
                st.info("No se encontraron datos de dependencias con los filtros actuales.")
            else:
                dep_pd[f"Ejec {filtro_vigencia}"] = (dep_pd[f"Ejec {filtro_vigencia}"] * 100).round(1)
                dep_pd["Ejec Acumulada"] = (dep_pd["Ejec Acumulada"] * 100).round(1)

                name_col = "Dependencia Responsable" if "Dependencia Responsable" in dep_pd.columns else "Responsable"
                dep_sorted = dep_pd.sort_values(f"Ejec {filtro_vigencia}", ascending=True)

                colors_dep = [color_pct(v / 100) for v in dep_sorted[f"Ejec {filtro_vigencia}"]]
                fig_dep = go.Figure(go.Bar(
                    x=dep_sorted[f"Ejec {filtro_vigencia}"],
                    y=dep_sorted[name_col],
                    orientation="h",
                    marker_color=colors_dep,
                    text=[f"{v:.1f}%" for v in dep_sorted[f"Ejec {filtro_vigencia}"]],
                    textposition="outside",
                    customdata=dep_sorted[["Metas Programadas", "Metas al 100", "Ejec Acumulada"]].values,
                    hovertemplate=(
                        "<b>%{y}</b><br>"
                        f"Avance {filtro_vigencia}: %{{x:.1f}}%<br>"
                        "Metas programadas: %{customdata[0]}<br>"
                        "Metas al 100%: %{customdata[1]}<br>"
                        "Avance acumulado: %{customdata[2]:.1f}%<extra></extra>"
                    ),
                ))
                fig_dep.update_layout(
                    xaxis_title=f"% Promedio Ejecucion {filtro_vigencia}",
                    height=max(350, len(dep_sorted) * 48),
                    paper_bgcolor="white", plot_bgcolor="#fafafa",
                    font={"family": "DM Sans"},
                    margin=dict(l=20, r=100, t=30, b=20),
                )
                st.plotly_chart(fig_dep, use_container_width=True, key="bar_dep")
                st.caption(
                    f"Como se calcula: promedio del porcentaje de avance de las metas programadas a cargo de cada dependencia en {filtro_vigencia}. "
                    "Metas al 100%: numero de indicadores clasificados como 'Superior'. "
                    "Avance acumulado: promedio de avance en el cuatrienio completo (2024-2027)."
                )

                st.markdown('<hr class="custom">', unsafe_allow_html=True)
                st.dataframe(dep_pd, use_container_width=True, height=380)

    # Errores de archivos opcionales
    missing = []
    if h26_bytes is None:
        missing.append(("Ejecucion Hacienda 2026", SCHEMAS["Hacienda (XLSX)"]))
    if r26_bytes is None:
        missing.append(("Regalias 2026", SCHEMAS["Regalias (XLSX)"]))
    if missing:
        with st.expander("Archivos faltantes o con errores", expanded=False):
            for fname, schema in missing:
                show_schema_error(fname, schema)

# -------------------------------------------------
# FOOTER
# -------------------------------------------------
st.markdown("""
<div class="footer">
    Dashboard de Avance PDD &middot; Desarrollado con Streamlit y Plotly &middot; Datos procesados con Polars
</div>
""", unsafe_allow_html=True)
