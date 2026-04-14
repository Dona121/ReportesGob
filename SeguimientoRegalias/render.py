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
        "descripcion": (
            "El proyecto fue aprobado y registrado en GESPROY, pero todavía no tiene "
            "contratos suscritos ni procesos precontractuales abiertos. "
            "Mientras permanezca aquí, no hay ejecución de recursos."
        ),
        "vino_de": "Aprobación del proyecto y migración a GESPROY.",
        "fecha_entrada": (
            "Fecha de aprobación del proyecto. "
            "Desde ese día se mide cuánto tiempo lleva sin contratar (Hito 1)."
        ),
        "para_avanzar": "Se debe suscribir el primer contrato del proyecto.",
        "fecha_avance": (
            "Una vez registrada la fecha de suscripción del contrato en GESPROY, "
            "el sistema cambia automáticamente el estado a Contratado sin acta de inicio."
        ),
        "accion": (
            "Verifique si el proceso precontractual ya fue publicado en SECOP. "
            "Si no, coordine con la entidad ejecutora para abrirlo y registrarlo en GESPROY "
            "a la mayor brevedad, pues cada día adicional suma al indicador de alerta."
        ),
        "porque_fechas": (
            "La fecha de aprobación es el punto de partida del cronograma del proyecto. "
            "La fecha de apertura de proceso (si existe) indica que ya se inició la "
            "contratación. La fecha de corte es el momento hasta el cual se calcula el indicador."
        ),
    },
    "CONTRATADO SIN ACTA DE INICIO": {
        "descripcion": (
            "El proyecto tiene al menos un contrato suscrito, pero aún no se ha firmado "
            "el acta de inicio. Sin ese documento, la ejecución física y financiera "
            "del contrato no ha comenzado formalmente."
        ),
        "vino_de": (
            "Suscripción del primer contrato. "
            "El sistema actualizó el estado automáticamente al registrar esa fecha en GESPROY."
        ),
        "fecha_entrada": (
            "Fecha de suscripción del primer contrato. "
            "Desde ese día se mide el tiempo sin acta de inicio (Hito 3)."
        ),
        "para_avanzar": (
            "Se debe firmar el acta de inicio del contrato y contar con la programación "
            "inicial registrada en GESPROY."
        ),
        "fecha_avance": (
            "Al registrar la fecha del acta de inicio, el sistema cambia automáticamente "
            "el estado a Contratado en ejecución."
        ),
        "accion": (
            "Confirme con el contratista y la interventoría la fecha de inicio de actividades "
            "y formalice el acta. Registre también la programación inicial del proyecto en "
            "GESPROY, ya que ambos son requisitos para avanzar al estado de ejecución."
        ),
        "porque_fechas": (
            "La fecha de suscripción marca el inicio del contrato legal. "
            "La fecha de acta de inicio marca el comienzo real de la obra o servicio. "
            "El tiempo entre ambas es el indicador de alerta de este estado (Hito 3)."
        ),
    },
    "CONTRATADO EN EJECUCIÓN": {
        "descripcion": (
            "El proyecto está en ejecución activa. Se están desarrollando las actividades "
            "del contrato y se reportan indicadores de avance. "
            "El horizonte de ejecución puede estar vigente o vencido."
        ),
        "vino_de": (
            "Firma del acta de inicio, que confirma el inicio formal de la ejecución "
            "del contrato en GESPROY."
        ),
        "fecha_entrada": (
            "Fecha de acta de inicio. "
            "A partir de aquí se mide el avance mediante los indicadores CPI y SPI, "
            "y se controla el horizonte de ejecución (Hito 4)."
        ),
        "para_avanzar": (
            "El proyecto debe haber cumplido sus metas e indicadores. Las entidades "
            "sectoriales deben remitir las actas finales como soporte. "
            "El cambio se realiza manualmente en GESPROY."
        ),
        "fecha_avance": (
            "No hay una fecha automática. El estado cambia a Terminado cuando la "
            "Secretaría Técnica registra el cierre de metas en GESPROY."
        ),
        "accion": (
            "Reporte los avances mensualmente en GESPROY. Revise que el CPI y el SPI "
            "estén por encima de 0. Si el horizonte del proyecto ya venció, solicite su "
            "actualización antes de que se active la alerta de rezago (Hito 4)."
        ),
        "porque_fechas": (
            "La fecha de acta de inicio es el punto de partida de la ejecución. "
            "El horizonte indica hasta cuándo debería estar ejecutado. "
            "Si la fecha de corte supera el horizonte con CPI=0 y SPI=0, "
            "el proyecto entra en alerta de rezago."
        ),
    },
    "TERMINADO": {
        "descripcion": (
            "El proyecto ha cumplido sus metas e indicadores y fue declarado finalizado. "
            "Es posible que aún queden contratos pendientes de liquidación o trámites "
            "administrativos por completar."
        ),
        "vino_de": (
            "Cumplimiento de las metas del proyecto. Las entidades sectoriales remitieron "
            "las actas finales y la Secretaría Técnica registró el cierre en GESPROY."
        ),
        "fecha_entrada": (
            "El cambio es manual; no hay una fecha de sistema que lo active automáticamente. "
            "El Hito 5 mide los días desde la finalización hasta la fecha de corte."
        ),
        "para_avanzar": (
            "Se deben liquidar todos los contratos, completar el pago total del proyecto "
            "y expedir el acto administrativo de cierre."
        ),
        "fecha_avance": (
            "No hay fecha automática. El paso a Para cierre requiere gestión manual "
            "por parte de la entidad ejecutora y la Secretaría Técnica."
        ),
        "accion": (
            "Verifique que todos los contratos estén liquidados y los pagos completados. "
            "Coordine con la entidad la elaboración y expedición del acto administrativo "
            "de cierre para avanzar al estado final."
        ),
        "porque_fechas": (
            "La fecha de finalización registrada en GESPROY determina desde cuándo "
            "el proyecto acumula días en el Hito 5. "
            "Cuanto más tiempo pase sin llegar a Para cierre, mayor será la alerta."
        ),
    },
    "PARA CIERRE": {
        "descripcion": (
            "El proyecto está totalmente liquidado y finalizado. Solo falta el acto "
            "administrativo que formaliza oficialmente su cierre ante las entidades "
            "competentes. Este es el último estado antes de cerrar el expediente."
        ),
        "vino_de": (
            "Liquidación total de contratos y pagos. La entidad ejecutora confirmó "
            "que no quedan obligaciones pendientes."
        ),
        "fecha_entrada": (
            "No hay una fecha automática de sistema para este cambio de estado. "
            "Es un registro manual en GESPROY."
        ),
        "para_avanzar": (
            "Solo se requiere expedir el acto administrativo de cierre. "
            "Este es el estado final del ciclo del proyecto."
        ),
        "fecha_avance": (
            "No hay un estado siguiente. Una vez expedido el acto de cierre, "
            "el expediente del proyecto queda formalmente cerrado."
        ),
        "accion": (
            "Gestione con la secretaría o entidad competente la elaboración y expedición "
            "del acto administrativo de cierre. Este trámite es el último paso para "
            "formalizar el cierre del proyecto ante el DNP y el SGR."
        ),
        "porque_fechas": (
            "Las fechas registradas sirven como soporte del historial completo del proyecto. "
            "Confirme que todas estén correctamente registradas en GESPROY antes de expedir "
            "el acto de cierre."
        ),
    },
    "SUSPENDIDO": {
        "descripcion": (
            "El proyecto o alguno de sus contratos fue suspendido temporalmente. "
            "La ejecución de actividades y el desembolso de recursos están detenidos "
            "mientras dure la suspensión."
        ),
        "vino_de": (
            "Decisión administrativa, técnica o judicial que ordenó la suspensión. "
            "Debe estar soportada en un acto administrativo."
        ),
        "fecha_entrada": (
            "Fecha del acto administrativo de suspensión. "
            "Desde ese momento el proyecto no debe reportar avances de ejecución."
        ),
        "para_avanzar": (
            "Se debe resolver la causal que originó la suspensión y expedir un acto "
            "administrativo de reactivación firmado por la entidad competente."
        ),
        "fecha_avance": (
            "Al levantar la suspensión, el proyecto retoma el estado que tenía antes "
            "de ser suspendido. El cambio se hace manualmente en GESPROY."
        ),
        "accion": (
            "Identifique y documente la causal de suspensión. Coordine con la interventoría, "
            "el contratista y la entidad ejecutora las acciones necesarias para levantarla. "
            "Una vez resuelta, gestione el acto administrativo de reactivación y actualice "
            "el estado en GESPROY."
        ),
        "porque_fechas": (
            "Las fechas ayudan a calcular cuánto tiempo lleva suspendido el proyecto "
            "y si los plazos contractuales deben ser ajustados como consecuencia "
            "de la suspensión."
        ),
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

# Etiquetas y justificación de cada fecha para el tooltip
_FECHA_INFO = [
    ("FECHA APROBACIÓN PROYECTO",           "Aprobación",        "Inicio del ciclo del proyecto en GESPROY"),
    ("FECHA DE APERTURA DEL PRIMER PROCESO","Apertura proceso",   "Inicio del proceso precontractual"),
    ("FECHA SUSCRIPCION",                   "Suscripción",        "Firma del primer contrato"),
    ("FECHA ACTA INICIO",                   "Acta de inicio",     "Inicio formal de la ejecución"),
    ("HORIZONTE DEL PROYECTO",              "Horizonte",          "Fecha límite de ejecución prevista"),
    ("FECHA DE FINALIZACIÓN",               "Finalización",       "Cierre de metas del proyecto"),
    ("FECHA DE CORTE GESPROY",              "Corte GESPROY",      "Fecha hasta la que se calculan los indicadores"),
]

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
        extra = "font-weight:700;" if eu == "SUSPENDIDO" else ""
        return (f'<span class="proy-pill" '
                f'style="background:{bg};color:{fg};border:1px solid {fg}40;{extra}">'
                f'{html.escape(est_proy)}</span>')

    # ── Fechas clave del proyecto ─────────────────────────────────────────────
    fechas_html = ""
    indicadores_html = ""
    if row_data:
        filas = []
        for col_key, lbl, razon in _FECHA_INFO:
            v = row_data.get(col_key)
            if v is not None:
                filas.append(
                    f'<tr>'
                    f'<td class="etip-flabel">{lbl}</td>'
                    f'<td class="etip-fval">{_fmt_date_short(v)}</td>'
                    f'</tr>'
                )
        if filas:
            fechas_html = (
                f'<div class="etip-section-title">Fechas registradas</div>'
                f'<table class="etip-dates">{"".join(filas)}</table>'
                f'<div class="etip-small" style="margin-top:0.2rem">'
                f'{html.escape(info["porque_fechas"])}</div>'
            )

        # ── Indicadores CPI/SPI ───────────────────────────────────────────────
        cpi = row_data.get("CPI")
        spi = row_data.get("SPI")
        if cpi is not None or spi is not None:
            def _ind(v, lbl, desc):
                if v is None: return ""
                try:
                    fv = float(v)
                    color = "#17743d" if fv >= 0.8 else ("#d88c16" if fv >= 0.5 else "#e68878")
                    return (f'<span class="etip-ind" style="color:{color}">'
                            f'<strong>{lbl}:</strong> {fv:.2f} <span style="opacity:0.6;font-size:0.62rem">({desc})</span></span>')
                except Exception:
                    return ""
            indicadores_html = (
                f'<div class="etip-section-title">Indicadores de ejecución</div>'
                f'<div class="etip-inds">'
                f'{_ind(cpi, "CPI", "costo")}'
                f'{_ind(spi, "SPI", "tiempo")}'
                f'</div>'
                f'<div class="etip-small" style="margin-top:0.15rem">'
                f'Valores &#8805; 0.8 son satisfactorios. Por debajo de 0.5 indican rezago crítico.</div>'
            )

    parts = [
        f'<span class="etip-estado">{html.escape(est_proy)}</span>',
        f'<p class="etip-desc">{html.escape(info["descripcion"])}</p>',

        f'<div class="etip-section-title">¿Cómo llegó a este estado?</div>',
        f'<div class="etip-row">{html.escape(info["vino_de"])}</div>',
        f'<div class="etip-row etip-small">'
        f'<span class="etip-label">Fecha de entrada:</span>'
        f'{html.escape(info["fecha_entrada"])}</div>',

        f'<div class="etip-section-title">¿Qué se necesita para avanzar?</div>',
        f'<div class="etip-row">{html.escape(info["para_avanzar"])}</div>',
        f'<div class="etip-row etip-small">{html.escape(info["fecha_avance"])}</div>',

        fechas_html,
        indicadores_html,

        f'<div class="etip-accion">'
        f'<span class="etip-accion-label">Acción sugerida</span>'
        f'{html.escape(info["accion"])}'
        f'</div>',
    ]
    tooltip_body = "".join(parts)

    extra_style = "font-weight:700;" if eu == "SUSPENDIDO" else ""
    return (
        f'<span class="proy-pill etip-trigger" '
        f'style="background:{bg};color:{fg};border:1px solid {fg}40;{extra_style}cursor:pointer">'
        f'{html.escape(est_proy)}&thinsp;<span class="etip-i">&#9432;</span>'
        f'<span class="etip-popup">{tooltip_body}</span>'
        f'</span>'
    )
