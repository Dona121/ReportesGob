import streamlit as st
import polars as pl
import pandas as pd
import io
from datetime import date

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Seguimiento Regalías",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Paleta corporativa ───────────────────────────────────────────────────────
COLORS = {
    "verde_oscuro":  "#005931",
    "verde_medio":   "#17743d",
    "cian":          "#47b1d5",
    "azul_medio":    "#1754ab",
    "azul_oscuro":   "#003d6c",
    "naranja":       "#d88c16",
    "naranja_oscuro":"#cf7000",
    "rojo_tierra":   "#d37e00",
    "cafe":          "#9b5b1e",
    "salmon":        "#e68878",
}

# ── CSS global ───────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {{
    font-family: 'Barlow', sans-serif;
    background: #f4f6f9;
}}

/* Header institucional */
.header-bar {{
    background: linear-gradient(135deg, {COLORS['azul_oscuro']} 0%, {COLORS['verde_oscuro']} 100%);
    padding: 1.4rem 2rem;
    border-radius: 0 0 16px 16px;
    margin-bottom: 1.8rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}}
.header-bar h1 {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    color: white;
    margin: 0;
    letter-spacing: 0.5px;
}}
.header-bar p {{
    color: rgba(255,255,255,0.75);
    margin: 0;
    font-size: 0.9rem;
    font-weight: 300;
}}

/* Tarjetas de métricas */
.metric-card {{
    background: white;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    border-left: 4px solid;
    box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    margin-bottom: 0.5rem;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}}
.metric-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 6px 18px rgba(0,0,0,0.12);
}}
.metric-card .label {{
    font-size: 0.75rem;
    font-weight: 600;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 0.3rem;
}}
.metric-card .value {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    line-height: 1;
}}
.metric-card .sub {{
    font-size: 0.78rem;
    color: #9ca3af;
    margin-top: 0.2rem;
}}

/* Tabla estilizada */
.styled-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
    background: white;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}}
.styled-table thead tr {{
    background: {COLORS['azul_oscuro']};
    color: white;
    font-family: 'Barlow Condensed', sans-serif;
    font-weight: 600;
    font-size: 0.82rem;
    letter-spacing: 0.5px;
    text-transform: uppercase;
}}
.styled-table th, .styled-table td {{
    padding: 0.7rem 1rem;
    text-align: left;
    border-bottom: 1px solid #f0f0f0;
}}
.styled-table tbody tr:hover {{
    background: #f0f7ff;
}}
.styled-table tbody tr:last-child td {{
    border-bottom: none;
}}

/* Badges de clasificación */
.badge {{
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
}}

/* Sección título */
.section-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.3rem;
    font-weight: 700;
    color: {COLORS['azul_oscuro']};
    border-bottom: 3px solid {COLORS['cian']};
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem 0;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}

/* Sidebar */
section[data-testid="stSidebar"] {{
    background: {COLORS['azul_oscuro']};
}}
section[data-testid="stSidebar"] * {{
    color: white !important;
}}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label {{
    color: rgba(255,255,255,0.8) !important;
    font-size: 0.82rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}}
section[data-testid="stSidebar"] .sidebar-title {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: {COLORS['cian']} !important;
    border-bottom: 1px solid rgba(71,177,213,0.3);
    padding-bottom: 0.5rem;
    margin-bottom: 1rem;
}}

/* Expander */
.streamlit-expanderHeader {{
    background: #f8fafc !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    color: {COLORS['azul_oscuro']} !important;
}}

/* Botón descarga */
.stDownloadButton > button {{
    background: {COLORS['verde_oscuro']} !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.2rem !important;
    font-family: 'Barlow', sans-serif !important;
    transition: background 0.2s !important;
}}
.stDownloadButton > button:hover {{
    background: {COLORS['verde_medio']} !important;
}}
</style>
""", unsafe_allow_html=True)

# ── Lógica de procesamiento ───────────────────────────────────────────────────

def hito_1(df):
    condicion_estado = (pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | (pl.col("ESTADO PROYECTO") == "") | (pl.col("ESTADO PROYECTO").is_null())
    condicion_fechas = (~pl.col("FECHA APROBACIÓN PROYECTO").is_null()) & (~pl.col("FECHA DE CORTE GESPROY").is_null())
    return pl.when(condicion_estado & condicion_fechas).then(
        (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA APROBACIÓN PROYECTO")).dt.total_days()
    ).otherwise(None)

def hito_2(df):
    condicion_estado = (pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | (pl.col("ESTADO PROYECTO") == "") | (pl.col("ESTADO PROYECTO").is_null())
    condicion_fechas = ~pl.col("FECHA DE APERTURA DEL PRIMER PROCESO").is_null()
    return pl.when(condicion_estado & condicion_fechas).then(
        (pl.col("FECHA ACTA INICIO") - pl.col("FECHA DE APERTURA DEL PRIMER PROCESO")).dt.total_days()
    ).otherwise(None)

def hito_3(df):
    condicion_estado = pl.col("ESTADO PROYECTO") == "CONTRATADO SIN ACTA DE INICIO"
    condicion_fecha = ~pl.col("FECHA SUSCRIPCION").is_null()
    return pl.when(condicion_estado & condicion_fecha).then(
        (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA SUSCRIPCION")).dt.total_days()
    ).otherwise(None)

def hito_4(df):
    condicion_estado = pl.col("ESTADO PROYECTO") == "CONTRATADO EN EJECUCIÓN"
    condicion_cpi_spi = (pl.col("CPI") == 0) & (pl.col("SPI") == 0)
    condicion_horizonte = pl.col("HORIZONTE DEL PROYECTO") <= pl.col("FECHA DE CORTE GESPROY")
    return pl.when(condicion_estado & condicion_cpi_spi & condicion_horizonte).then(
        (pl.col("FECHA DE CORTE GESPROY") - pl.col("HORIZONTE DEL PROYECTO")).dt.total_days()
    ).otherwise(None)

def hito_5(df):
    condicion_fecha = ~pl.col("FECHA DE FINALIZACIÓN").is_null()
    return pl.when(condicion_fecha).then(
        (pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA DE FINALIZACIÓN")).dt.total_days()
    ).otherwise(None)

def clasificar(col, intervalos):
    expr = pl.when(pl.col(col).is_null()).then(None)
    for etiqueta, (lo, hi) in intervalos.items():
        if hi is None:
            expr = expr.when(pl.col(col) > lo).then(pl.lit(etiqueta))
        else:
            expr = expr.when((pl.col(col) >= lo) & (pl.col(col) <= hi)).then(pl.lit(etiqueta))
    return expr.otherwise(None)

INTERVALOS = {
    "hito_1": {"0-100": (0, 100), "101-150": (101, 150), "151-180": (151, 180), ">180": (180, None)},
    "hito_2": {"0-100": (0, 100), "101-150": (101, 150), "151-180": (151, 180), ">180": (180, None)},
    "hito_3": {"0-30": (0, 30), "31-45": (31, 45), "46-60": (46, 60), ">60": (60, None)},
    "hito_4": {"0-30": (0, 30), "31-90": (31, 90), "91-180": (91, 180), ">180": (180, None)},
    "hito_5": {"0-100": (0, 100), "101-150": (101, 150), "151-180": (151, 180), ">180": (180, None)},
}

BADGE_COLORS = {
    "0-100": ("#d1fae5", "#065f46"),
    "101-150": ("#fef9c3", "#854d0e"),
    "0-30": ("#d1fae5", "#065f46"),
    "31-45": ("#fef9c3", "#854d0e"),
    "31-90": ("#fef9c3", "#854d0e"),
    "46-60": ("#fed7aa", "#9a3412"),
    "91-180": ("#fed7aa", "#9a3412"),
    "151-180": ("#fed7aa", "#9a3412"),
    ">60":  ("#fee2e2", "#991b1b"),
    ">180": ("#fee2e2", "#991b1b"),
}

HITOS_INFO = [
    ("hito_1_val", "Clasificacion hito 1", "Proyectos sin contratar sin apertura", COLORS["verde_oscuro"]),
    ("hito_2_val", "Clasificacion hito 2", "Proyectos sin contratar con apertura", COLORS["azul_medio"]),
    ("hito_3_val", "Clasificacion hito 3", "Contratados sin acta de inicio", COLORS["naranja"]),
    ("hito_4_val", "Clasificacion hito 4", "Proyectos en ejecución", COLORS["cian"]),
    ("hito_5_val", "Clasificacion hito 5", "Proyectos terminados", COLORS["cafe"]),
]

@st.cache_data
def procesar(file_bytes):
    df = pl.read_excel(io.BytesIO(file_bytes), table_name="MatrizSeguimientoEvaluacion")
    df = (
        df.select(
            "ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO", "ESTADO PROYECTO",
            "CPI", "SPI", "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
            "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
            "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY"
        )
        .with_columns(
            pl.col("FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
                   "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
                   "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY").cast(pl.Date, strict=False)
        )
        .with_columns(
            hito_1(None).alias("hito_1_val"),
            hito_2(None).alias("hito_2_val"),
            hito_3(None).alias("hito_3_val"),
            hito_4(None).alias("hito_4_val"),
            hito_5(None).alias("hito_5_val"),
            pl.when(pl.col("ESTADO PROYECTO") == "SUSPENDIDO").then(pl.lit(1)).otherwise(None).alias("Suspendidos"),
            pl.when(pl.col("ESTADO PROYECTO") == "PARA CIERRE").then(pl.lit(1)).otherwise(None).alias("Para cierre"),
        )
        .with_columns(
            clasificar("hito_1_val", INTERVALOS["hito_1"]).alias("Clasificacion hito 1"),
            clasificar("hito_2_val", INTERVALOS["hito_2"]).alias("Clasificacion hito 2"),
            clasificar("hito_3_val", INTERVALOS["hito_3"]).alias("Clasificacion hito 3"),
            clasificar("hito_4_val", INTERVALOS["hito_4"]).alias("Clasificacion hito 4"),
            clasificar("hito_5_val", INTERVALOS["hito_5"]).alias("Clasificacion hito 5"),
        )
    )
    return df

def badge_html(text):
    if text is None or (isinstance(text, float) and text != text):
        return "<span style='color:#9ca3af;font-size:0.8rem'>—</span>"
    bg, fg = BADGE_COLORS.get(str(text), ("#e5e7eb", "#374151"))
    return f"<span class='badge' style='background:{bg};color:{fg}'>{text}</span>"

def fmt_dias(val):
    if val is None:
        return "—"
    try:
        return f"{float(val):.1f} días"
    except:
        return "—"

def generar_excel(df_det: pl.DataFrame, df_agr: pl.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_agr.to_pandas().to_excel(writer, sheet_name="Resumen por Entidad", index=False)
        df_det.to_pandas().to_excel(writer, sheet_name="Detalle Proyectos", index=False)
    return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────

# Header
st.markdown("""
<div class="header-bar">
  <div>
    <h1>📊 Seguimiento y Evaluación · Regalías</h1>
    <p>Matriz de seguimiento de proyectos por hito de gestión</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Carga de datos ────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("<div class='sidebar-title'>⚙ Configuración</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Cargar Excel", type=["xlsx"], help="Archivo con tabla MatrizSeguimientoEvaluacion")

    if uploaded:
        st.success("✓ Archivo cargado")
        st.markdown("---")
        st.markdown("<div class='sidebar-title'>🔍 Filtros</div>", unsafe_allow_html=True)

if uploaded is None:
    st.info("👈 Carga el archivo Excel de la Matriz de Seguimiento para comenzar.")
    st.stop()

df = procesar(uploaded.read())

# ── Filtros en sidebar ────────────────────────────────────────────────────────
with st.sidebar:
    entidades = sorted(df["ENTIDAD O SECRETARIA"].drop_nulls().unique().to_list())
    sel_entidades = st.multiselect("Entidad / Secretaría", entidades, default=entidades)

    estados = sorted(df["ESTADO PROYECTO"].drop_nulls().unique().to_list())
    sel_estados = st.multiselect("Estado del proyecto", estados, default=estados)

    hitos_disp = {
        "Hito 1 – Sin contratar sin apertura": "hito_1_val",
        "Hito 2 – Sin contratar con apertura": "hito_2_val",
        "Hito 3 – Contratado sin acta":        "hito_3_val",
        "Hito 4 – En ejecución":               "hito_4_val",
        "Hito 5 – Terminados":                 "hito_5_val",
    }
    sel_hito_label = st.selectbox("Ver detalle de hito", list(hitos_disp.keys()))
    sel_hito_col   = hitos_disp[sel_hito_label]

# ── Filtrar ───────────────────────────────────────────────────────────────────
df_f = df.filter(
    pl.col("ENTIDAD O SECRETARIA").is_in(sel_entidades) &
    pl.col("ESTADO PROYECTO").is_in(sel_estados)
)

# ── Agrupación ────────────────────────────────────────────────────────────────
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
        pl.len().alias("Total proyectos"),
    )
    .sort("ENTIDAD O SECRETARIA")
)

# ── KPIs globales ─────────────────────────────────────────────────────────────
total_proy     = df_f.height
total_entidades = df_f["ENTIDAD O SECRETARIA"].n_unique()
suspendidos    = int(df_f["Suspendidos"].sum() or 0)
para_cierre    = int(df_f["Para cierre"].sum() or 0)
promedio_h1    = df_f["hito_1_val"].mean()
promedio_h1_str = f"{promedio_h1:.1f} días" if promedio_h1 else "N/A"

k1, k2, k3, k4, k5 = st.columns(5)
kpis = [
    (k1, "Total proyectos",    total_proy,          "filtrados",        COLORS["azul_oscuro"]),
    (k2, "Entidades",          total_entidades,     "secretarías",      COLORS["verde_oscuro"]),
    (k3, "Suspendidos",        suspendidos,         "proyectos",        COLORS["naranja_oscuro"]),
    (k4, "Para cierre",        para_cierre,         "proyectos",        COLORS["cafe"]),
    (k5, "Prom. Hito 1",       promedio_h1_str,     "días promedio",    COLORS["cian"]),
]
for col, label, val, sub, color in kpis:
    with col:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color:{color}">
            <div class="label">{label}</div>
            <div class="value" style="color:{color}">{val}</div>
            <div class="sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

# ── Tabla resumen por entidad ─────────────────────────────────────────────────
st.markdown("<div class='section-title'>Resumen por Entidad</div>", unsafe_allow_html=True)

agr_pd = agrupacion.to_pandas()

table_rows = ""
for _, row in agr_pd.iterrows():
    def cell(v, col_key):
        clasi_map = {
            "Hito 1 (días)": "Clasificacion hito 1",
            "Hito 2 (días)": "Clasificacion hito 2",
            "Hito 3 (días)": "Clasificacion hito 3",
            "Hito 4 (días)": "Clasificacion hito 4",
            "Hito 5 (días)": "Clasificacion hito 5",
        }
        dias = f"{v:.1f}" if pd.notna(v) else "—"
        # get clasificacion from df_f
        clasi_col = clasi_map.get(col_key)
        entidad = row["ENTIDAD O SECRETARIA"]
        clasi_val = None
        if clasi_col:
            sub = df_f.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)[clasi_col].drop_nulls()
            if len(sub) > 0:
                clasi_val = sub[0]
        badge = badge_html(clasi_val)
        return f"<td>{dias} días &nbsp;{badge}</td>"

    table_rows += f"""<tr>
        <td><strong>{row['ENTIDAD O SECRETARIA']}</strong></td>
        {cell(row['Hito 1 (días)'], 'Hito 1 (días)')}
        {cell(row['Hito 2 (días)'], 'Hito 2 (días)')}
        {cell(row['Hito 3 (días)'], 'Hito 3 (días)')}
        {cell(row['Hito 4 (días)'], 'Hito 4 (días)')}
        {cell(row['Hito 5 (días)'], 'Hito 5 (días)')}
        <td style="text-align:center">{int(row['Suspendidos']) if pd.notna(row['Suspendidos']) else 0}</td>
        <td style="text-align:center">{int(row['Para cierre']) if pd.notna(row['Para cierre']) else 0}</td>
        <td style="text-align:center"><strong>{int(row['Total proyectos'])}</strong></td>
    </tr>"""

st.markdown(f"""
<table class="styled-table">
<thead><tr>
    <th>Entidad / Secretaría</th>
    <th>Hito 1</th><th>Hito 2</th><th>Hito 3</th><th>Hito 4</th><th>Hito 5</th>
    <th>Suspendidos</th><th>Para cierre</th><th>Total</th>
</tr></thead>
<tbody>{table_rows}</tbody>
</table>
""", unsafe_allow_html=True)

# ── Detalle por hito ──────────────────────────────────────────────────────────
st.markdown(f"<div class='section-title'>Detalle · {sel_hito_label}</div>", unsafe_allow_html=True)

clasi_col_map = {
    "hito_1_val": "Clasificacion hito 1",
    "hito_2_val": "Clasificacion hito 2",
    "hito_3_val": "Clasificacion hito 3",
    "hito_4_val": "Clasificacion hito 4",
    "hito_5_val": "Clasificacion hito 5",
}
clasi_col = clasi_col_map[sel_hito_col]

df_det = (
    df_f.filter(~pl.col(sel_hito_col).is_null())
    .select("ENTIDAD O SECRETARIA", "BPIN", "NOMBRE PROYECTO", "ESTADO PROYECTO",
            sel_hito_col, clasi_col)
    .rename({sel_hito_col: "Días", clasi_col: "Clasificación"})
    .sort(["ENTIDAD O SECRETARIA", "Días"], descending=[False, True])
)

# Agrupar por entidad con expander
entidades_det = df_det["ENTIDAD O SECRETARIA"].unique().sort().to_list()

for entidad in entidades_det:
    sub = df_det.filter(pl.col("ENTIDAD O SECRETARIA") == entidad)
    prom = sub["Días"].mean()
    n = sub.height
    with st.expander(f"**{entidad}** — {n} proyecto(s) · Promedio: {prom:.1f} días", expanded=False):
        det_rows = ""
        for r in sub.to_dicts():
            det_rows += f"""<tr>
                <td style="font-size:0.78rem;color:#6b7280">{r['BPIN'] or '—'}</td>
                <td>{r['NOMBRE PROYECTO'] or '—'}</td>
                <td><span style="font-size:0.8rem;background:#f1f5f9;padding:2px 8px;border-radius:6px">{r['ESTADO PROYECTO'] or '—'}</span></td>
                <td style="font-weight:600">{r['Días']:.1f} días</td>
                <td>{badge_html(r['Clasificación'])}</td>
            </tr>"""
        st.markdown(f"""
        <table class="styled-table">
        <thead><tr><th>BPIN</th><th>Nombre del proyecto</th><th>Estado</th><th>Días</th><th>Clasificación</th></tr></thead>
        <tbody>{det_rows}</tbody>
        </table>""", unsafe_allow_html=True)

# ── Exportar Excel ────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Exportar</div>", unsafe_allow_html=True)

col_dl, _ = st.columns([1, 3])
with col_dl:
    excel_bytes = generar_excel(df_f, agrupacion)
    st.download_button(
        label="⬇ Descargar reporte Excel",
        data=excel_bytes,
        file_name=f"reporte_regalias_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
