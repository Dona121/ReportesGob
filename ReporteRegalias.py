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

# ── Paleta corporativa ────────────────────────────────────────────────────────
C = {
    "azul_oscuro":   "#003d6c",
    "azul_medio":    "#1754ab",
    "verde_oscuro":  "#005931",
    "verde_medio":   "#17743d",
    "cian":          "#47b1d5",
    "naranja":       "#d88c16",
    "naranja_osc":   "#cf7000",
    "cafe":          "#9b5b1e",
    "salmon":        "#e68878",
    "bg":            "#f2f5f8",
    "white":         "#ffffff",
    "text":          "#1a2332",
    "muted":         "#6b7280",
    "border":        "#e2e8f0",
}

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&family=Sora:wght@600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif;
    background: {C['bg']};
    color: {C['text']};
}}

section[data-testid="stSidebar"] > div {{
    background: {C['azul_oscuro']};
    padding-top: 1.5rem;
}}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div {{
    color: rgba(255,255,255,0.85) !important;
}}
.sidebar-section {{
    font-family: 'Sora', sans-serif;
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: {C['cian']} !important;
    margin: 1.5rem 0 0.5rem 0;
    padding-bottom: 0.4rem;
    border-bottom: 1px solid rgba(71,177,213,0.25);
}}
.page-header {{
    background: linear-gradient(120deg, {C['azul_oscuro']} 0%, {C['verde_oscuro']} 100%);
    border-radius: 12px;
    padding: 1.6rem 2.2rem;
    margin-bottom: 2rem;
}}
.page-header h1 {{
    font-family: 'Sora', sans-serif;
    font-size: 1.75rem;
    font-weight: 800;
    color: white;
    margin: 0 0 0.2rem 0;
    letter-spacing: -0.3px;
}}
.page-header p {{
    color: rgba(255,255,255,0.65);
    margin: 0;
    font-size: 0.88rem;
    font-weight: 300;
}}
.kpi-card {{
    background: {C['white']};
    border-radius: 10px;
    padding: 1.1rem 1.3rem;
    border-top: 3px solid;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
}}
.kpi-label {{
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: {C['muted']};
    margin-bottom: 0.35rem;
}}
.kpi-value {{
    font-family: 'Sora', sans-serif;
    font-size: 1.9rem;
    font-weight: 800;
    line-height: 1;
    margin-bottom: 0.15rem;
}}
.kpi-sub {{
    font-size: 0.75rem;
    color: {C['muted']};
}}
.section-heading {{
    font-family: 'Sora', sans-serif;
    font-size: 0.75rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: {C['azul_oscuro']};
    border-bottom: 2px solid {C['cian']};
    padding-bottom: 0.5rem;
    margin: 2rem 0 1.2rem 0;
}}
.summary-table {{
    width: 100%;
    border-collapse: collapse;
    background: {C['white']};
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    font-size: 0.84rem;
}}
.summary-table thead tr {{
    background: {C['azul_oscuro']};
    color: white;
}}
.summary-table th {{
    padding: 0.75rem 1rem;
    font-family: 'Sora', sans-serif;
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    text-align: left;
    white-space: nowrap;
}}
.summary-table td {{
    padding: 0.7rem 1rem;
    border-bottom: 1px solid {C['border']};
    vertical-align: middle;
}}
.summary-table tbody tr:last-child td {{ border-bottom: none; }}
.summary-table tbody tr:hover td {{ background: #f8fbff; }}
.entidad-name {{
    font-weight: 600;
    font-size: 0.85rem;
    color: {C['azul_oscuro']};
}}
.dias-val {{
    font-family: 'DM Mono', monospace;
    font-size: 0.82rem;
    font-weight: 500;
}}
.null-cell {{ color: {C['border']}; font-size: 0.8rem; }}
.badge {{
    display: inline-block;
    padding: 2px 9px;
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.3px;
    margin-left: 6px;
    vertical-align: middle;
}}
.badge-green  {{ background: #d1fae5; color: #065f46; }}
.badge-yellow {{ background: #fef3c7; color: #92400e; }}
.badge-orange {{ background: #ffedd5; color: #9a3412; }}
.badge-red    {{ background: #fee2e2; color: #991b1b; }}
.detail-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    margin-top: 0.5rem;
}}
.detail-table th {{
    background: #f1f5f9;
    color: {C['azul_oscuro']};
    font-family: 'Sora', sans-serif;
    font-size: 0.66rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 0.55rem 0.9rem;
    text-align: left;
    border-bottom: 2px solid {C['border']};
}}
.detail-table td {{
    padding: 0.6rem 0.9rem;
    border-bottom: 1px solid {C['border']};
    vertical-align: middle;
}}
.detail-table tbody tr:last-child td {{ border-bottom: none; }}
.detail-table tbody tr:hover td {{ background: #f8fbff; }}
.bpin-tag {{
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    background: #f1f5f9;
    color: {C['muted']};
    padding: 2px 7px;
    border-radius: 4px;
}}
.estado-tag {{
    font-size: 0.72rem;
    background: #eff6ff;
    color: {C['azul_medio']};
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 500;
    white-space: nowrap;
}}
div[data-testid="stExpander"] > details > summary {{
    background: {C['white']} !important;
    border: 1px solid {C['border']} !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    color: {C['azul_oscuro']} !important;
    padding: 0.6rem 1rem !important;
}}
div[data-testid="stExpander"] > details > summary:hover {{
    background: #f8fbff !important;
}}
div[data-testid="stExpander"] > details[open] > summary {{
    border-radius: 8px 8px 0 0 !important;
    border-bottom-color: {C['cian']} !important;
}}
div[data-testid="stExpander"] > details > div {{
    border: 1px solid {C['border']} !important;
    border-top: none !important;
    border-radius: 0 0 8px 8px !important;
    background: {C['white']} !important;
    padding: 0 1rem 0.5rem 1rem !important;
}}
.stDownloadButton > button {{
    background: {C['verde_oscuro']} !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.85rem !important;
    padding: 0.5rem 1.4rem !important;
}}
.stDownloadButton > button:hover {{
    background: {C['verde_medio']} !important;
}}
span[data-baseweb="tag"] {{
    background: rgba(71,177,213,0.2) !important;
    color: white !important;
}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# LÓGICA
# ─────────────────────────────────────────────────────────────────────────────

INTERVALOS = {
    "hito_1_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
    "hito_2_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
    "hito_3_val": [("0-30",  0,  30), ("31-45",   31,  45), ("46-60",   46,  60),  (">60",  61, None)],
    "hito_4_val": [("0-30",  0,  30), ("31-90",   31,  90), ("91-180",  91, 180),  (">180", 181, None)],
    "hito_5_val": [("0-100", 0, 100), ("101-150", 101, 150), ("151-180", 151, 180), (">180", 181, None)],
}

def clasificar(col, intervalos):
    expr = pl.when(pl.col(col).is_null()).then(None)
    for label, lo, hi in intervalos:
        cond = (pl.col(col) >= lo) & (pl.col(col) <= hi) if hi is not None else (pl.col(col) >= lo)
        expr = expr.when(cond).then(pl.lit(label))
    return expr.otherwise(None)

@st.cache_data
def procesar(file_bytes):
    df = pl.read_excel(io.BytesIO(file_bytes), table_name="MatrizSeguimientoEvaluacion")
    df = (
        df.select(
            "ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO", "ESTADO PROYECTO",
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
            pl.when(
                ((pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")) &
                (~pl.col("FECHA APROBACIÓN PROYECTO").is_null()) &
                (~pl.col("FECHA DE CORTE GESPROY").is_null())
            ).then(
                (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA APROBACIÓN PROYECTO")).dt.total_days()
            ).otherwise(None).alias("hito_1_val"),

            pl.when(
                ((pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")) &
                (~pl.col("FECHA DE APERTURA DEL PRIMER PROCESO").is_null())
            ).then(
                (pl.col("FECHA ACTA INICIO") - pl.col("FECHA DE APERTURA DEL PRIMER PROCESO")).dt.total_days()
            ).otherwise(None).alias("hito_2_val"),

            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO SIN ACTA DE INICIO") &
                (~pl.col("FECHA SUSCRIPCION").is_null())
            ).then(
                (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA SUSCRIPCION")).dt.total_days()
            ).otherwise(None).alias("hito_3_val"),

            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO EN EJECUCIÓN") &
                (pl.col("CPI") == 0) & (pl.col("SPI") == 0) &
                (pl.col("HORIZONTE DEL PROYECTO") <= pl.col("FECHA DE CORTE GESPROY"))
            ).then(
                (pl.col("FECHA DE CORTE GESPROY") - pl.col("HORIZONTE DEL PROYECTO")).dt.total_days()
            ).otherwise(None).alias("hito_4_val"),

            pl.when(~pl.col("FECHA DE FINALIZACIÓN").is_null()).then(
                (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA DE FINALIZACIÓN")).dt.total_days()
            ).otherwise(None).alias("hito_5_val"),

            pl.when(pl.col("ESTADO PROYECTO") == "SUSPENDIDO").then(pl.lit(1)).otherwise(None).alias("Suspendidos"),
            pl.when(pl.col("ESTADO PROYECTO") == "PARA CIERRE").then(pl.lit(1)).otherwise(None).alias("Para cierre"),
        )
        .with_columns(
            clasificar("hito_1_val", INTERVALOS["hito_1_val"]).alias("clasi_1"),
            clasificar("hito_2_val", INTERVALOS["hito_2_val"]).alias("clasi_2"),
            clasificar("hito_3_val", INTERVALOS["hito_3_val"]).alias("clasi_3"),
            clasificar("hito_4_val", INTERVALOS["hito_4_val"]).alias("clasi_4"),
            clasificar("hito_5_val", INTERVALOS["hito_5_val"]).alias("clasi_5"),
        )
    )
    return df

def badge_html(val):
    if val is None:
        return ""
    mapping = {
        "0-100": "badge-green",  "0-30": "badge-green",
        "101-150": "badge-yellow", "31-45": "badge-yellow", "31-90": "badge-yellow",
        "151-180": "badge-orange", "46-60": "badge-orange", "91-180": "badge-orange",
        ">180": "badge-red",    ">60": "badge-red",
    }
    cls = mapping.get(str(val), "badge-yellow")
    return f"<span class='badge {cls}'>{val}</span>"

def generar_excel(df_det: pl.DataFrame, df_agr: pl.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_agr.to_pandas().to_excel(writer, sheet_name="Resumen por Entidad", index=False)
        df_det.to_pandas().to_excel(writer, sheet_name="Detalle Proyectos", index=False)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div class='sidebar-section'>📁 Datos</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx"], label_visibility="collapsed")
    if uploaded:
        st.success("Archivo cargado correctamente")

if uploaded is None:
    st.markdown("""
    <div class="page-header">
        <h1>Seguimiento y Evaluación · Regalías</h1>
        <p>Matriz de seguimiento de proyectos por hito de gestión</p>
    </div>
    """, unsafe_allow_html=True)
    st.info("Carga el archivo Excel de la Matriz de Seguimiento en el panel izquierdo para comenzar.")
    st.stop()

df = procesar(uploaded.read())

HITOS = {
    "Hito 1 — Sin contratar sin apertura":    ("hito_1_val", "clasi_1"),
    "Hito 2 — Sin contratar con apertura":    ("hito_2_val", "clasi_2"),
    "Hito 3 — Contratado sin acta de inicio": ("hito_3_val", "clasi_3"),
    "Hito 4 — En ejecución":                  ("hito_4_val", "clasi_4"),
    "Hito 5 — Proyectos terminados":           ("hito_5_val", "clasi_5"),
}

with st.sidebar:
    st.markdown("<div class='sidebar-section'>Filtros</div>", unsafe_allow_html=True)
    entidades = sorted(df["ENTIDAD O SECRETARIA"].drop_nulls().unique().to_list())
    sel_entidades = st.multiselect("Entidad / Secretaría", entidades, default=entidades)

    estados = sorted(df["ESTADO PROYECTO"].drop_nulls().unique().to_list())
    sel_estados = st.multiselect("Estado del proyecto", estados, default=estados)

    st.markdown("<div class='sidebar-section'>Vista de detalle</div>", unsafe_allow_html=True)
    sel_hito_label = st.selectbox("Hito a detallar", list(HITOS.keys()), label_visibility="collapsed")

sel_hito_col, sel_clasi_col = HITOS[sel_hito_label]

# ─────────────────────────────────────────────────────────────────────────────
# FILTRAR
# ─────────────────────────────────────────────────────────────────────────────
filter_expr = pl.col("ENTIDAD O SECRETARIA").is_in(sel_entidades)
if sel_estados:
    filter_expr = filter_expr & pl.col("ESTADO PROYECTO").is_in(sel_estados)
df_f = df.filter(filter_expr)

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="page-header">
    <h1>Seguimiento y Evaluación · Regalías</h1>
    <p>Matriz de seguimiento de proyectos por hito de gestión</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────────────────────────────────────
total_proy      = df_f.height
total_entidades = df_f["ENTIDAD O SECRETARIA"].n_unique()
suspendidos     = int(df_f["Suspendidos"].drop_nulls().sum()) if df_f["Suspendidos"].drop_nulls().len() > 0 else 0
para_cierre     = int(df_f["Para cierre"].drop_nulls().sum()) if df_f["Para cierre"].drop_nulls().len() > 0 else 0
prom_h1         = df_f["hito_1_val"].mean()
prom_h1_str     = f"{prom_h1:.1f} días" if prom_h1 is not None else "N/A"

kpi_data = [
    ("Total proyectos", total_proy,       "filtrados",   C["azul_oscuro"]),
    ("Entidades",       total_entidades,  "secretarías", C["verde_oscuro"]),
    ("Suspendidos",     suspendidos,      "proyectos",   C["naranja_osc"]),
    ("Para cierre",     para_cierre,      "proyectos",   C["cafe"]),
    ("Prom. hito 1",    prom_h1_str,      "días prom.",  C["cian"]),
]
for col, (label, val, sub, color) in zip(st.columns(5), kpi_data):
    with col:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:{color}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="color:{color}">{val}</div>
            <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

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

# Clasificación predominante por entidad y hito (calculada una sola vez)
clasi_cols = ["clasi_1", "clasi_2", "clasi_3", "clasi_4", "clasi_5"]
clasi_por_entidad = {}
for entidad in agrupacion["ENTIDAD O SECRETARIA"].to_list():
    sub = df_f.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)
    clasi_por_entidad[entidad] = {}
    for cc in clasi_cols:
        vals = sub[cc].drop_nulls()
        if len(vals) > 0:
            clasi_por_entidad[entidad][cc] = vals.value_counts().sort("count", descending=True)[0, 0]
        else:
            clasi_por_entidad[entidad][cc] = None

# ─────────────────────────────────────────────────────────────────────────────
# TABLA RESUMEN
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-heading'>Resumen por entidad</div>", unsafe_allow_html=True)

def hito_cell(dias_val, entidad, clasi_key):
    if dias_val is None or (isinstance(dias_val, float) and dias_val != dias_val):
        return "<td class='null-cell'>—</td>"
    clasi = clasi_por_entidad[entidad].get(clasi_key)
    return f"<td><span class='dias-val'>{dias_val:.1f} d</span>{badge_html(clasi)}</td>"

rows_html = ""
for row in agrupacion.to_dicts():
    e = row["ENTIDAD O SECRETARIA"]
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
    <th>Hito 1</th><th>Hito 2</th><th>Hito 3</th><th>Hito 4</th><th>Hito 5</th>
    <th>Suspendidos</th><th>Para cierre</th><th>Total</th>
</tr></thead>
<tbody>{rows_html}</tbody>
</table>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# DETALLE POR HITO — usa sel_hito_col y sel_clasi_col del sidebar
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"<div class='section-heading'>Detalle · {sel_hito_label}</div>", unsafe_allow_html=True)

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
        sub = df_det.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)
        prom = sub[sel_hito_col].mean()
        n = sub.height
        prom_str = f"{prom:.1f} días" if prom is not None else "—"

        with st.expander(f"{entidad}   ·   {n} proyecto(s)   ·   Promedio: {prom_str}", expanded=False):
            det_rows = ""
            for r in sub.to_dicts():
                dias_v  = r[sel_hito_col]
                clasi_v = r[sel_clasi_col]
                dias_str = f"{dias_v:.1f} d" if dias_v is not None else "—"
                det_rows += f"""<tr>
                    <td><span class="bpin-tag">{r['BPIN'] or '—'}</span></td>
                    <td style="font-size:0.83rem">{r['NOMBRE PROYECTO'] or '—'}</td>
                    <td><span class="estado-tag">{r['ESTADO PROYECTO'] or '—'}</span></td>
                    <td style="font-family:'DM Mono',monospace;font-weight:600;font-size:0.82rem">{dias_str}</td>
                    <td>{badge_html(clasi_v)}</td>
                </tr>"""
            st.markdown(f"""
            <table class="detail-table">
            <thead><tr>
                <th>BPIN</th><th>Nombre del proyecto</th><th>Estado</th><th>Días</th><th>Clasificación</th>
            </tr></thead>
            <tbody>{det_rows}</tbody>
            </table>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# EXPORTAR
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-heading'>Exportar reporte</div>", unsafe_allow_html=True)

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

col_dl, _ = st.columns([1, 4])
with col_dl:
    st.download_button(
        label="Descargar reporte Excel",
        data=generar_excel(export_df, agrupacion),
        file_name=f"regalias_seguimiento_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
