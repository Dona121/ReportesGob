import streamlit as st
import polars as pl
import pandas as pd
import plotly.graph_objects as go
import io
from datetime import date

st.set_page_config(
    page_title="Seguimiento Regalías",
    layout="wide",
    initial_sidebar_state="collapsed",
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
    "bg":           "#f2f5f8",
    "white":        "#ffffff",
    "text":         "#1a2332",
    "muted":        "#6b7280",
    "border":       "#e2e8f0",
}

HITO_COLORS = [C["azul_medio"], C["cian"], C["naranja"], C["verde_medio"], C["cafe"]]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'Montserrat', sans-serif;
    background: {C['bg']};
    color: {C['text']};
}}

/* ── Header ── */
.page-header {{
    background: linear-gradient(120deg, {C['azul_oscuro']} 0%, {C['verde_oscuro']} 100%);
    border-radius: 12px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
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
    background: {C['white']};
    border-radius: 10px;
    padding: 0.9rem 1.1rem;
    border-left: 3px solid;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
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
    background: {C['white']};
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
    background: {C['white']};
    border-radius: 0 0 10px 10px;
    padding: 1.4rem;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
}}

/* ── Summary table ── */
.summary-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.83rem;
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
}}
.summary-table tbody tr:last-child td {{ border-bottom: none; }}
.summary-table tbody tr:hover td {{ background: #f0f7ff; }}
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
    display: inline-block;
    padding: 2px 8px;
    border-radius: 20px;
    font-size: 0.68rem;
    font-weight: 700;
    margin-left: 5px;
    vertical-align: middle;
}}
.badge-green  {{ background: #d1fae5; color: #065f46; }}
.badge-yellow {{ background: #fef3c7; color: #92400e; }}
.badge-orange {{ background: #ffedd5; color: #9a3412; }}
.badge-red    {{ background: #fee2e2; color: #991b1b; }}

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
div[data-testid="stExpander"] > details > summary {{
    background: #f8fafc !important;
    border: 1px solid {C['border']} !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.85rem !important;
    color: {C['azul_oscuro']} !important;
    padding: 0.55rem 1rem !important;
}}
div[data-testid="stExpander"] > details > summary:hover {{ background: #f0f7ff !important; }}
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
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 4px;
    border: 1px solid #fca5a5;
}}
.col-wrong-type {{
    background: #fff7ed;
    color: #9a3412;
    font-family: 'DM Mono', monospace;
    font-size: 0.75rem;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 4px;
    border: 1px solid #fed7aa;
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
    df = pl.read_excel(io.BytesIO(file_bytes), table_name=TABLA_ESPERADA)
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
                (~pl.col("FECHA APROBACIÓN PROYECTO").is_null()) & (~pl.col("FECHA DE CORTE GESPROY").is_null())
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA APROBACIÓN PROYECTO")).dt.total_days()).otherwise(None).alias("hito_1_val"),

            pl.when(
                ((pl.col("ESTADO PROYECTO") == "SIN CONTRATAR") | pl.col("ESTADO PROYECTO").is_null() | (pl.col("ESTADO PROYECTO") == "")) &
                (~pl.col("FECHA DE APERTURA DEL PRIMER PROCESO").is_null())
            ).then((pl.col("FECHA ACTA INICIO") - pl.col("FECHA DE APERTURA DEL PRIMER PROCESO")).dt.total_days()).otherwise(None).alias("hito_2_val"),

            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO SIN ACTA DE INICIO") &
                (~pl.col("FECHA SUSCRIPCION").is_null())
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("FECHA SUSCRIPCION")).dt.total_days()).otherwise(None).alias("hito_3_val"),

            pl.when(
                (pl.col("ESTADO PROYECTO") == "CONTRATADO EN EJECUCIÓN") &
                (pl.col("CPI") == 0) & (pl.col("SPI") == 0) &
                (pl.col("HORIZONTE DEL PROYECTO") <= pl.col("FECHA DE CORTE GESPROY"))
            ).then((pl.col("FECHA DE CORTE GESPROY") - pl.col("HORIZONTE DEL PROYECTO")).dt.total_days()).otherwise(None).alias("hito_4_val"),

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
    if val is None: return ""
    cls = {
        "0-100": "badge-green", "0-30": "badge-green",
        "101-150": "badge-yellow", "31-45": "badge-yellow", "31-90": "badge-yellow",
        "151-180": "badge-orange", "46-60": "badge-orange", "91-180": "badge-orange",
        ">180": "badge-red", ">60": "badge-red",
    }.get(str(val), "badge-yellow")
    return f"<span class='badge {cls}'>{val}</span>"

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
TABLA_ESPERADA = "MatrizSeguimientoEvaluacion"

COLUMNAS_ESPERADAS = {
    "ENTIDAD O SECRETARIA":                 ("texto",  [pl.Utf8, pl.String]),
    "BPIN":                                 ("texto",  [pl.Utf8, pl.String]),
    "NOMBRE PROYECTO":                      ("texto",  [pl.Utf8, pl.String]),
    "ESTADO PROYECTO":                      ("texto",  [pl.Utf8, pl.String]),
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
    "texto":  "Texto (Utf8)",
    "número": "Número (Float / Int)",
    "fecha":  "Fecha (Date / Datetime)",
}

def error_card(titulo, cuerpo, solucion):
    return f"""
    <div class="error-card">
        <div class="error-title">&#9888; {titulo}</div>
        <div class="error-body">{cuerpo}</div>
        <div class="error-fix"><strong>Cómo solucionarlo</strong>{solucion}</div>
    </div>"""

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
                f"El archivo no pudo ser leído correctamente. Detalle técnico: <code>{msg}</code>",
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
            f"Las siguientes columnas no fueron encontradas en la tabla:<div class='error-cols'>{chips}</div>",
            "Verifica que los encabezados de la tabla coincidan exactamente con los nombres esperados "
            "(respeta mayúsculas, tildes y espacios). No renombres las columnas originales de la matriz."
        ))

    # 3. Tipo de datos incorrecto — solo columnas que sí existen
    tipo_incorrecto = []
    for col, (tipo_label, tipos_validos) in COLUMNAS_ESPERADAS.items():
        if col not in cols_actuales:
            continue
        dtype = df_raw[col].dtype
        # fechas: también aceptar Utf8 que se pueda castear (se parsea luego con strict=False)
        if tipo_label == "fecha":
            continue  # el cast strict=False maneja esto en procesar
        if dtype not in tipos_validos:
            tipo_incorrecto.append((col, tipo_label, str(dtype)))

    if tipo_incorrecto:
        chips = "".join(
            f"<span class='col-wrong-type' title='Tipo actual: {dtype_actual}'>{col}</span>"
            for col, _, dtype_actual in tipo_incorrecto
        )
        detalles = "".join(
            f"<li><b>{col}</b>: se encontró <code>{dtype_actual}</code>, se esperaba <code>{TIPO_LABEL[tipo_label]}</code></li>"
            for col, tipo_label, dtype_actual in tipo_incorrecto
        )
        errores.append(error_card(
            f"{'Tipo de dato incorrecto' if len(tipo_incorrecto) == 1 else f'Tipos de datos incorrectos en {len(tipo_incorrecto)} columnas'}",
            f"Las siguientes columnas tienen un tipo de dato inesperado:<div class='error-cols'>{chips}</div><ul style='margin-top:0.5rem;font-size:0.8rem'>{detalles}</ul>",
            "En Excel, selecciona la columna indicada y verifica su formato de celda. "
            "Las columnas CPI y SPI deben ser <b>numéricas</b>. "
            "Si los valores aparecen como texto (alineados a la izquierda), selecciona la columna → <b>Datos → Texto en columnas</b> y elige formato General o Número."
        ))

    if errores:
        return None, errores

    return df_raw, []

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

# ─────────────────────────────────────────────────────────────────────────────
# CARGA
# ─────────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("Cargar archivo Excel", type=["xlsx"], label_visibility="collapsed")

if uploaded is None:
    st.info("Carga el archivo Excel de la Matriz de Seguimiento para comenzar.")
    st.stop()

file_bytes = uploaded.read()
df_raw, errores = validar_archivo(file_bytes)

if errores:
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
    </div>
    """, unsafe_allow_html=True)
    st.stop()

df = procesar(file_bytes)

# ─────────────────────────────────────────────────────────────────────────────
# FILTROS HORIZONTALES
# ─────────────────────────────────────────────────────────────────────────────
entidades    = sorted(df["ENTIDAD O SECRETARIA"].drop_nulls().unique().to_list())
estados_raw  = sorted(df["ESTADO PROYECTO"].drop_nulls().unique().to_list())
tiene_sin    = df["ESTADO PROYECTO"].is_null().any() or (df["ESTADO PROYECTO"] == "").any()
opciones_est = estados_raw + (["(Sin estado)"] if tiene_sin else [])

HITOS = {
    "H1 · Sin contratar sin apertura":    ("hito_1_val", "clasi_1"),
    "H2 · Sin contratar con apertura":    ("hito_2_val", "clasi_2"),
    "H3 · Contratado sin acta de inicio": ("hito_3_val", "clasi_3"),
    "H4 · En ejecución rezagado":         ("hito_4_val", "clasi_4"),
    "H5 · Proyectos terminados":          ("hito_5_val", "clasi_5"),
}

fc1, fc2, fc3 = st.columns([2, 2, 1.5])
with fc1:
    st.caption("**Entidad / Secretaría**")
    sel_entidades = st.multiselect("Entidad", entidades, default=entidades, label_visibility="collapsed")
with fc2:
    st.caption("**Estado del proyecto**")
    sel_estados = st.multiselect("Estado", opciones_est, default=opciones_est, label_visibility="collapsed")
with fc3:
    st.caption("**Vista de detalle**")
    sel_hito_label = st.selectbox("Hito", list(HITOS.keys()), label_visibility="collapsed")

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
# TABS: Resumen | Gráfico | Detalle | Exportar
# ─────────────────────────────────────────────────────────────────────────────
tab_resumen, tab_grafico, tab_detalle, tab_exportar = st.tabs([
    "Resumen por entidad",
    "Comparativo por hito",
    f"Detalle · {sel_hito_label}",
    "Exportar",
])

# ── TAB 1: Tabla resumen ──────────────────────────────────────────────────────
with tab_resumen:
    def hito_cell(dias_val, entidad, clasi_key):
        if dias_val is None or (isinstance(dias_val, float) and dias_val != dias_val):
            return "<td class='null-cell'>—</td>"
        clasi = clasi_por_entidad[entidad].get(clasi_key)
        return f"<td><span class='dias-val'>{dias_val:.1f} d</span>{badge_html(clasi)}</td>"

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
            "Promedio de días entre el <b>Horizonte del proyecto</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Estado = CONTRATADO EN EJECUCIÓN, CPI = 0, SPI = 0 y horizonte vencido.")}
        {th("Proyectos<br>terminados", "Hito 5 · Proyectos terminados",
            "Promedio de días entre la <b>Fecha de finalización</b> y la <b>Fecha de corte GESPROY</b>.<br><br>Condición: Fecha de finalización registrada.")}
        {th("Suspendidos", "Proyectos suspendidos", "Conteo de proyectos con Estado = SUSPENDIDO.")}
        {th("Para cierre", "Proyectos para cierre", "Conteo de proyectos con Estado = PARA CIERRE.")}
        <th>Total</th>
    </tr></thead>
    <tbody>{rows_html}</tbody>
    </table>
    """, unsafe_allow_html=True)

# ── TAB 2: Gráfico comparativo ────────────────────────────────────────────────
with tab_grafico:
    HITO_LABELS = [
        "Sin contratar sin apertura",
        "Sin contratar con apertura",
        "Contratado sin acta",
        "En ejecución rezagado",
        "Proyectos terminados",
    ]
    HITO_COLS = ["Hito 1 (días)", "Hito 2 (días)", "Hito 3 (días)", "Hito 4 (días)", "Hito 5 (días)"]

    agr_pd    = agrupacion.to_pandas()
    entidades_g = agr_pd["ENTIDAD O SECRETARIA"].tolist()

    g_col1, g_col2 = st.columns([2, 1])

    with g_col1:
        st.markdown("<div class='section-heading'>Días promedio por hito y entidad</div>", unsafe_allow_html=True)
        fig = go.Figure()
        for i, (hcol, hlabel, color) in enumerate(zip(HITO_COLS, HITO_LABELS, HITO_COLORS)):
            vals = agr_pd[hcol].tolist()
            fig.add_trace(go.Bar(
                name=hlabel,
                x=entidades_g,
                y=vals,
                marker_color=color,
                marker_line_width=0,
                text=[f"{v:.0f}d" if pd.notna(v) else "" for v in vals],
                textposition="outside",
                textfont=dict(size=10, family="Montserrat"),
            ))

        fig.update_layout(
            barmode="group",
            plot_bgcolor="white",
            paper_bgcolor="white",
            font=dict(family="Montserrat", size=11, color=C["text"]),
            legend=dict(
                orientation="h",
                yanchor="bottom", y=1.02,
                xanchor="left", x=0,
                font=dict(size=10),
            ),
            xaxis=dict(
                tickfont=dict(size=10),
                gridcolor=C["border"],
                linecolor=C["border"],
            ),
            yaxis=dict(
                title="Días promedio",
                gridcolor=C["border"],
                linecolor=C["border"],
                tickfont=dict(size=10),
            ),
            margin=dict(t=50, b=20, l=10, r=10),
            height=380,
        )
        st.plotly_chart(fig, use_container_width=True)

    with g_col2:
        st.markdown("<div class='section-heading'>Ranking por hito seleccionado</div>", unsafe_allow_html=True)

        hito_sel_col = HITO_COLS[list(HITOS.keys()).index(sel_hito_label)]
        ranking = agr_pd[["ENTIDAD O SECRETARIA", hito_sel_col]].dropna(subset=[hito_sel_col])
        ranking = ranking.sort_values(hito_sel_col, ascending=False)

        if len(ranking) == 0:
            st.info("No hay datos para este hito.")
        else:
            max_val = ranking[hito_sel_col].max()
            rank_rows = ""
            for i, (_, r) in enumerate(ranking.iterrows()):
                pct = (r[hito_sel_col] / max_val * 100) if max_val > 0 else 0
                bar_color = C["salmon"] if pct > 66 else (C["naranja"] if pct > 33 else C["verde_medio"])
                rank_rows += f"""
                <div style="margin-bottom:0.8rem">
                    <div style="display:flex;justify-content:space-between;margin-bottom:3px">
                        <span style="font-size:0.78rem;font-weight:600;color:{C['azul_oscuro']}">{r['ENTIDAD O SECRETARIA']}</span>
                        <span style="font-family:'DM Mono',monospace;font-size:0.78rem;font-weight:600;color:{C['text']}">{r[hito_sel_col]:.1f}d</span>
                    </div>
                    <div style="background:{C['border']};border-radius:4px;height:7px;overflow:hidden">
                        <div style="background:{bar_color};width:{pct:.1f}%;height:100%;border-radius:4px;transition:width 0.3s"></div>
                    </div>
                </div>"""
            st.markdown(f"<div style='padding-top:0.3rem'>{rank_rows}</div>", unsafe_allow_html=True)

# ── TAB 3: Detalle ────────────────────────────────────────────────────────────
with tab_detalle:
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
                        <td>{badge_html(clasi_v)}</td>
                    </tr>"""
                st.markdown(f"""
                <table class="detail-table">
                <thead><tr>
                    <th>BPIN</th><th>Nombre del proyecto</th><th>Estado</th><th>Días</th><th>Clasificación</th>
                </tr></thead>
                <tbody>{det_rows}</tbody>
                </table>""", unsafe_allow_html=True)

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
