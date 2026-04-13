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

# ─────────────────────────────────────────────────────────────────────────────
# Tooltip contextual de estado de proyecto
# Datos del Excel ToolTipEstadoProyectos
# ─────────────────────────────────────────────────────────────────────────────

ESTADO_INFO = {
    "SIN CONTRATAR": {
        "icono":       "📋",
        "descripcion": "El proyecto fue aprobado y migrado a GESPROY, pero aún no cuenta con contrato suscrito ni proceso precontractual activo.",
        "vino_de":     "Aprobación del proyecto (ingreso a GESPROY)",
        "fecha_entrada": "Fecha de aprobación del proyecto",
        "para_avanzar": "Suscribir el primer contrato del proyecto.",
        "fecha_avance": "Fecha de suscripción del contrato → cambia a <em>Contratado sin acta de inicio</em>",
        "accion":      "Iniciar y registrar el proceso precontractual en GESPROY a la mayor brevedad.",
    },
    "CONTRATADO SIN ACTA DE INICIO": {
        "icono":       "📝",
        "descripcion": "El proyecto tiene al menos un contrato suscrito, pero aún no se ha firmado el acta de inicio que da comienzo formal a la ejecución.",
        "vino_de":     "Suscripción del primer contrato",
        "fecha_entrada": "Fecha de suscripción del primer contrato",
        "para_avanzar": "Firmar el acta de inicio del proyecto y contar con programación inicial.",
        "fecha_avance": "Fecha de acta de inicio → cambia a <em>Contratado en ejecución</em>",
        "accion":      "Registrar el acta de inicio y la programación inicial en GESPROY.",
    },
    "CONTRATADO EN EJECUCIÓN": {
        "icono":       "⚙️",
        "descripcion": "El proyecto se encuentra en ejecución activa. Se cuentan indicadores de avance (CPI/SPI) y el horizonte de ejecución está vigente o puede estar vencido.",
        "vino_de":     "Firma del acta de inicio",
        "fecha_entrada": "Fecha de acta de inicio",
        "para_avanzar": "Cumplir las metas e indicadores del proyecto y contar con las actas finales de los contratos.",
        "fecha_avance": "Cambio manual en GESPROY una vez finalizadas las metas → <em>Terminado</em>",
        "accion":      "Reportar avances periódicamente en GESPROY. Verificar CPI y SPI. Si el horizonte está vencido, gestionar su actualización.",
    },
    "TERMINADO": {
        "icono":       "✅",
        "descripcion": "El proyecto ha cumplido sus metas e indicadores y se declaró finalizado. Aún pueden quedar contratos pendientes de liquidación.",
        "vino_de":     "Finalización de metas e indicadores (actas finales enviadas)",
        "fecha_entrada": "Registro manual en GESPROY",
        "para_avanzar": "Liquidar todos los contratos, completar los pagos y expedir el acto administrativo de cierre.",
        "fecha_avance": "No hay fecha automática → <em>Para cierre</em> (gestión manual)",
        "accion":      "Remitir a la Secretaría Técnica las actas finales de todos los contratos y avanzar en la liquidación.",
    },
    "PARA CIERRE": {
        "icono":       "🔒",
        "descripcion": "El proyecto está liquidado y finalizado. Solo falta el acto administrativo que formaliza oficialmente su cierre.",
        "vino_de":     "Liquidación total de contratos y pagos",
        "fecha_entrada": "No hay fecha automática",
        "para_avanzar": "Expedir el acto administrativo de cierre.",
        "fecha_avance": "Sin estado siguiente — este es el estado final.",
        "accion":      "Gestionar la expedición del acto administrativo de cierre ante la entidad competente.",
    },
    "SUSPENDIDO": {
        "icono":       "⏸️",
        "descripcion": "El proyecto o alguno de sus contratos fue suspendido temporalmente. La ejecución está detenida.",
        "vino_de":     "Decisión administrativa o técnica de suspensión",
        "fecha_entrada": "Fecha de acto de suspensión",
        "para_avanzar": "Levantar la causal de suspensión y reactivar mediante acto administrativo.",
        "fecha_avance": "Acto de reactivación → retoma el estado previo",
        "accion":      "Identificar y resolver la causal de suspensión. Coordinar con la interventoría y la entidad para la reactivación.",
    },
}

def _fmt_date_short(d):
    """Formatea una fecha Python date como DD/MM/AAAA, o retorna '—'."""
    if d is None:
        return "—"
    try:
        return d.strftime("%d/%m/%Y")
    except Exception:
        return str(d)

def _estado_tooltip_html(est_proy, row_data=None):
    """
    Genera un pill de estado con tooltip contextual enriquecido.
    row_data: dict con los campos del proyecto (fechas, CPI, SPI, etc.).
    """
    if not est_proy:
        return '<span class="proy-pill proy-pill--empty">—</span>'

    eu = est_proy.strip().upper()
    fg, bg = ESTADO_PROY_COLORS.get(eu, (C["muted"], "#f1f5f9"))
    info   = ESTADO_INFO.get(eu)

    if not info:
        # Estado sin tooltip definido → pill simple
        extra = "font-weight:700;" if eu == "SUSPENDIDO" else ""
        return (f'<span class="proy-pill" '
                f'style="background:{bg};color:{fg};border:1px solid {fg}40;{extra}">'
                f'{html.escape(est_proy)}</span>')

    # ── Construir cuerpo del tooltip ──────────────────────────────────────────
    # 1. Fechas clave del proyecto (si están disponibles)
    fechas_html = ""
    if row_data:
        campos_fecha = [
            ("Aprobación",      row_data.get("FECHA APROBACIÓN PROYECTO")),
            ("Apertura proceso",row_data.get("FECHA DE APERTURA DEL PRIMER PROCESO")),
            ("Suscripción",     row_data.get("FECHA SUSCRIPCION")),
            ("Acta de inicio",  row_data.get("FECHA ACTA INICIO")),
            ("Horizonte",       row_data.get("HORIZONTE DEL PROYECTO")),
            ("Finalización",    row_data.get("FECHA DE FINALIZACIÓN")),
            ("Corte GESPROY",   row_data.get("FECHA DE CORTE GESPROY")),
        ]
        filas = [
            f'<tr><td class="etip-flabel">{lbl}</td>'
            f'<td class="etip-fval">{_fmt_date_short(v)}</td></tr>'
            for lbl, v in campos_fecha if v is not None
        ]
        if filas:
            fechas_html = f'''
            <div class="etip-section-title">📅 Fechas del proyecto</div>
            <table class="etip-dates">{"".join(filas)}</table>'''

        # 2. Indicadores CPI/SPI si aplica
        indicadores_html = ""
        cpi = row_data.get("CPI")
        spi = row_data.get("SPI")
        if cpi is not None or spi is not None:
            def _ind(v, lbl):
                if v is None: return ""
                try:
                    fv = float(v)
                    color = "#17743d" if fv >= 0.8 else ("#d88c16" if fv >= 0.5 else "#e68878")
                    return f'<span class="etip-ind" style="color:{color}"><strong>{lbl}:</strong> {fv:.2f}</span>'
                except Exception:
                    return ""
            indicadores_html = (
                f'<div class="etip-section-title">📊 Indicadores de ejecución</div>'
                f'<div class="etip-inds">{_ind(cpi,"CPI")} {_ind(spi,"SPI")}</div>'
            )
    else:
        indicadores_html = ""

    tooltip_body = f'''
    <div class="etip-header">
        <span class="etip-icono">{info["icono"]}</span>
        <span class="etip-estado">{html.escape(est_proy)}</span>
    </div>
    <p class="etip-desc">{info["descripcion"]}</p>

    <div class="etip-section-title">🔄 ¿Cómo llegó aquí?</div>
    <div class="etip-row"><span class="etip-label">Vino de:</span> {info["vino_de"]}</div>
    <div class="etip-row"><span class="etip-label">Fecha clave:</span> {info["fecha_entrada"]}</div>

    <div class="etip-section-title">➡️ ¿Qué se necesita para avanzar?</div>
    <div class="etip-row">{info["para_avanzar"]}</div>
    <div class="etip-row etip-small">{info["fecha_avance"]}</div>

    {fechas_html}
    {indicadores_html}

    <div class="etip-accion">
        <span class="etip-accion-label">💡 Acción sugerida</span>
        {info["accion"]}
    </div>'''

    extra_style = "font-weight:700;" if eu == "SUSPENDIDO" else ""
    return f'''<span class="proy-pill etip-trigger"
        style="background:{bg};color:{fg};border:1px solid {fg}40;{extra_style}cursor:pointer">
        {html.escape(est_proy)}&thinsp;<span class="etip-i">ⓘ</span>
        <span class="etip-popup">{tooltip_body}</span>
    </span>'''
