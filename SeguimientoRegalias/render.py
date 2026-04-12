"""
render.py
Funciones de renderizado HTML y helpers de presentación visual.
Incluye: badge_html, _pill, _fmt_date, _dias_tooltip, eval_color,
         _clasificar_promedio, _contratos_panel, _calcular_clasi_modal,
         constantes de color de estado.
"""
from constants import C, INTERVALOS, SEMAFOROS
import streamlit as st
import polars as pl
import html
import logging

_log = logging.getLogger(__name__)

def badge_html(val, hito_key=None):
    """Genera badge con punto de color semáforo y tooltip con mensaje."""
    if val is None: return ""
    val_str = str(val)
    cls_map = {
        "0-100":   "badge-green",  "0-30":  "badge-green",  "0-1":   "badge-green",
        "101-150": "badge-yellow", "31-45": "badge-yellow", "1.1-3": "badge-yellow",
        "151-180": "badge-orange", "46-60": "badge-orange", "3.1-6": "badge-orange",
        ">180":    "badge-black",  ">60":   "badge-black",  ">6":    "badge-black",
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

def _calcular_clasi_modal(df: pl.DataFrame, cols: list) -> dict:
    result: dict = {}
    for col in cols:
        modal = (
            df.filter(pl.col(col).is_not_null())
            .group_by(["ENTIDAD O SECRETARIA", col])
            .agg(pl.len().alias("_n"))
            .sort(["ENTIDAD O SECRETARIA", "_n"], descending=[False, True])
            .group_by("ENTIDAD O SECRETARIA")
            .first()
            .select(["ENTIDAD O SECRETARIA", col])
        )
        for row in modal.to_dicts():
            ent = row["ENTIDAD O SECRETARIA"]
            if ent not in result:
                result[ent] = {c: None for c in cols}
            result[ent][col] = row[col]
    return result

# Ajuste 2: usar df (sin filtros) para que el semáforo sea consistente
# independientemente de los filtros del sidebar. df_f puede excluir proyectos
# y cambiar la moda de clasificación por entidad.
clasi_por_entidad = _calcular_clasi_modal(df, _CLASI_COLS)

ESTADO_PROY_COLORS = {
    "SIN CONTRATAR":                 (C["cian"],        "#e0f7fa"),
    "CONTRATADO EN EJECUCIÓN":       (C["verde_medio"], "#d1fae5"),
    "CONTRATADO SIN ACTA DE INICIO": (C["azul_medio"],  "#dbeafe"),
    "TERMINADO":                     (C["muted"],       "#f1f5f9"),
    "PARA CIERRE":                   (C["cafe"],        "#fef3c7"),
}
ESTADO_CONT_COLORS = {
    "EN EJECUCIÓN":  (C["verde_medio"], "#d1fae5"),
    "TERMINADO":     (C["muted"],       "#f1f5f9"),
    "LIQUIDADO":     (C["azul_medio"],  "#dbeafe"),
    "SUSPENDIDO":    (C["naranja_osc"], "#ffedd5"),
    "SIN CONTRATO":  (C["cian"],        "#e0f7fa"),
}
CTTO_ESTADO_COLORS = {
    "EN EJECUCIÓN":  (C["verde_medio"],  "#d1fae5"),
    "EJECUTADO":     (C["verde_oscuro"], "#d1fae5"),
    "TERMINADO":     (C["muted"],        "#f1f5f9"),
    "LIQUIDADO":     (C["azul_medio"],   "#dbeafe"),
    "SUSPENDIDO":    (C["naranja_osc"],  "#ffedd5"),
    "RESCINDIDO":    (C["salmon"],       "#fee2e2"),
    "SUSCRITO":      (C["cian"],         "#e0f7fa"),
}

def _pill(texto, color_map, default_fg=None, default_bg=None):
    if not texto:
        return '<span class="proy-pill proy-pill--empty">—</span>'
    eu = texto.strip().upper()
    fg, bg = color_map.get(eu, (default_fg or C["muted"], default_bg or "#f1f5f9"))
    extra = "font-weight:700;" if eu == "SUSPENDIDO" else ""
    return (f'<span class="proy-pill" '
            f'style="background:{bg};color:{fg};border:1px solid {fg}40;{extra}">'
            f'{html.escape(texto)}</span>')

def _fmt_valor(v):
    """Formatea un valor float como moneda COP."""
    if v is None or (isinstance(v, float) and v != v):
        return "—"
    try:
        return f"$ {v:,.0f}"
    except Exception:
        return str(v)

def _valor_a_gradiente(valor, v_min, v_max):
    """Devuelve un color de fondo suave proporcional al valor del contrato."""
    if valor is None or v_max == v_min:
        return "#ffffff"
    ratio = max(0.0, min(1.0, (valor - v_min) / (v_max - v_min)))
    r = int(255 - ratio * (255 - 219))
    g = int(255 - ratio * (255 - 234))
    b = int(255 - ratio * (255 - 254))
    return f"rgb({r},{g},{b})"

def _contratos_panel(bpin_str, df_cttos):
    """Genera el HTML del panel de contratos para un BPIN dado."""
    if df_cttos is None:
        return '<div class="ctto-panel"><div class="ctto-panel-empty">Archivo de contratos no disponible.</div></div>'

    bpin_norm = (
        str(bpin_str).strip()
        .replace(".", "").replace("-", "").replace(",", "").replace(" ", "")
    )
    cttos = df_cttos.filter(pl.col("BPIN") == bpin_norm)

    if cttos.height == 0:
        return '<div class="ctto-panel"><div class="ctto-panel-empty">Sin contratos registrados para este proyecto.</div></div>'

    valores     = [r.get("CONTRATO VALOR TOTAL") for r in cttos.to_dicts()]
    valores_num = [v for v in valores if v is not None and isinstance(v, float) and v == v]
    v_min = min(valores_num) if valores_num else 0
    v_max = max(valores_num) if valores_num else 0

    n      = cttos.height
    header = f"""
    <div class="ctto-panel-header">
        <span class="ctto-panel-title">Contratos</span>
        <span class="ctto-panel-count">{n}</span>
    </div>"""

    rows_list = []
    for ctto in cttos.to_dicts():
        valor    = ctto.get("CONTRATO VALOR TOTAL")
        bg_grad  = _valor_a_gradiente(valor, v_min, v_max)
        estado_c = (ctto.get("ESTADO CONTRATO") or "").strip().upper()
        fg_e, bg_e = CTTO_ESTADO_COLORS.get(estado_c, (C["muted"], "#f1f5f9"))
        proceso   = html.escape(ctto.get("NO. PROCESO PRECONTRACTUAL") or "—")
        modalidad = html.escape(ctto.get("MODALIDAD CONTRATACION") or "—")
        tipo      = html.escape(ctto.get("TIPO CONTRATO") or "—")
        objeto    = html.escape(ctto.get("CONTRATO OBJETO") or "—")

        bar_px = 0
        if valor and v_max > v_min:
            bar_px = int(max(6, min(60, (valor - v_min) / (v_max - v_min) * 60)))
        elif valor:
            bar_px = 60

        rows_list.append(f"""<tr style="background:{bg_grad}">
            <td class="ctto-col1"><span class="ctto-proceso">{proceso}</span></td>
            <td style="font-size:0.73rem;color:{C['text']}">{modalidad}</td>
            <td style="font-size:0.73rem;color:{C['muted']}">{tipo}</td>
            <td>
                <div class="ctto-valor-wrap">
                    <span class="ctto-valor">{_fmt_valor(valor)}</span>
                    <div class="ctto-valor-bar" style="width:{bar_px}px"></div>
                </div>
            </td>
            <td><span class="ctto-estado-pill" style="background:{bg_e};color:{fg_e};border:1px solid {fg_e}40">{html.escape(ctto.get("ESTADO CONTRATO") or "—")}</span></td>
            <td><div class="ctto-objeto">{objeto}</div></td>
        </tr>""")

    rows = "".join(rows_list)
    tabla = f"""
    <div style="border-radius:10px;overflow:hidden;box-shadow:0 1px 10px rgba(0,40,90,0.10);">
    <table class="ctto-table">
    <thead><tr>
        <th class="ctto-col1">No. proceso</th>
        <th>Modalidad</th><th>Tipo</th>
        <th>Valor total</th><th>Estado</th>
        <th>Objeto del contrato</th>
    </tr></thead>
    <tbody>{rows}</tbody>
    </table></div>"""
    return f'<div class="ctto-panel">{header}{tabla}</div>'

def _fmt_date(val):
    """Formatea una fecha como DD/MM/YYYY o devuelve '—'."""
    if val is None:
        return "—"
    try:
        return val.strftime("%d/%m/%Y")
    except Exception:
        return str(val)

HITO_CALC_META = {
    "hito_1_val": (
        "Fecha aprobación",    "FECHA APROBACIÓN PROYECTO",
        "Fecha corte GESPROY", "FECHA DE CORTE GESPROY",
        "Días desde la aprobación del proyecto hasta el corte, sin proceso de contratación abierto.",
    ),
    "hito_2_val": (
        "Fecha apertura proceso", "FECHA DE APERTURA DEL PRIMER PROCESO",
        "Fecha acta de inicio",   "FECHA ACTA INICIO",
        "Días desde la apertura del primer proceso hasta el acta de inicio.",
    ),
    "hito_3_val": (
        "Fecha suscripción",   "FECHA SUSCRIPCION",
        "Fecha corte GESPROY", "FECHA DE CORTE GESPROY",
        "Días desde la suscripción del contrato hasta el corte, sin acta de inicio.",
    ),
    "hito_4_val": (
        "Horizonte del proyecto", "HORIZONTE DEL PROYECTO",
        "Fecha corte GESPROY",    "FECHA DE CORTE GESPROY",
        "Días de retraso sobre el horizonte (CPI=0, SPI=0). El resultado se muestra en meses.",
    ),
    "hito_5_val": (
        "Fecha finalización",  "FECHA DE FINALIZACIÓN",
        "Fecha corte GESPROY", "FECHA DE CORTE GESPROY",
        "Días entre la fecha de finalización registrada y el corte.",
    ),
}

def _dias_tooltip(r, hito_col):
    """Genera el HTML del tooltip con las fechas y fórmula del cálculo."""
    meta = HITO_CALC_META.get(hito_col)
    if not meta:
        return ""
    lbl_a, col_a, lbl_b, col_b, nota = meta
    fecha_a = _fmt_date(r.get(col_a))
    fecha_b = _fmt_date(r.get(col_b))
    dias_v  = r.get(hito_col)
    dias_display = f"{dias_v:.0f} días" if dias_v is not None else "—"
    es_h4 = hito_col == "hito_4_val"
    resultado_label = f"{dias_v/30:.1f} meses ({dias_display})" if es_h4 and dias_v else dias_display
    return (
        f'<div class="dias-tip-box">'
        f'  <div class="dias-tip-title">Cálculo del hito</div>'
        f'  <div class="dias-tip-row"><span class="dias-tip-lbl">{lbl_b}</span>'
        f'    <span class="dias-tip-val">{fecha_b}</span></div>'
        f'  <div class="dias-tip-op">menos (−)</div>'
        f'  <div class="dias-tip-row"><span class="dias-tip-lbl">{lbl_a}</span>'
        f'    <span class="dias-tip-val">{fecha_a}</span></div>'
        f'  <div class="dias-tip-sep"></div>'
        f'  <div class="dias-tip-result">= &nbsp;{resultado_label}</div>'
        f'  <div class="dias-tip-nota">{nota}</div>'
        f'</div>'
    )

def eval_color(score, max_score=100.0):
    """Devuelve (color_hex, etiqueta) según el score en escala 0-100."""
    ratio = score / max_score if max_score > 0 else 0
    if ratio >= 0.8:   return C["verde_medio"],  "Sobresaliente"
    elif ratio >= 0.6: return C["cian"],         "Satisfactorio"
    elif ratio >= 0.4: return C["naranja"],      "Aceptable"
    else:              return C["salmon"],        "Por mejorar"

HITO_KEY_MAP = {
    "clasi_1": "hito_1_val",
    "clasi_2": "hito_2_val",
    "clasi_3": "hito_3_val",
    "clasi_4": "hito_4_val",
    "clasi_5": "hito_5_val",
}
# Mapa inverso: clasi_key → hito_val_col (para clasificar el promedio)
CLASI_TO_HITO = {
    "clasi_1": "hito_1_val",
    "clasi_2": "hito_2_val",
    "clasi_3": "hito_3_val",
    "clasi_4": "hito_4_val",
    "clasi_5": "hito_5_val",
}

def _clasificar_promedio(dias_val, clasi_key):
    """Clasifica el promedio de días según los intervalos del hito.
    Siempre coherente con el número mostrado, independiente de filtros."""
    if dias_val is None or (isinstance(dias_val, float) and dias_val != dias_val):
        return None
    hito_col = CLASI_TO_HITO.get(clasi_key)
    if hito_col == "hito_4_val":
        # H4 se mide en meses
        meses = dias_val / 30.0
        if   meses <= 1: return "0-1"
        elif meses <= 3: return "1.1-3"
        elif meses <= 6: return "3.1-6"
        else:            return ">6"
    else:
        intervalos = INTERVALOS.get(hito_col, [])
        for label, lo, hi in intervalos:
            if hi is None and dias_val >= lo:            return label
            if hi is not None and lo <= dias_val <= hi:  return label
        return None
