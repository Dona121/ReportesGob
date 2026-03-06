#!/usr/bin/env python
# coding: utf-8

# In[1]:


import streamlit as st
import polars as pl
import pandas
import xlsxwriter
import os
import io
from datetime import datetime

st.set_page_config(
    page_title="Matriz de Seguimiento y Evaluación",
    layout="wide",
)

st.title("Matriz de Seguimiento y Evaluación — Regalías SGR")
st.markdown("Carga los archivos fuente para generar la matriz mensual.")

# ── Funciones de Ayuda ─────────────────────────────────────────────────────────
def ruta_reciente(archivos: list[str], prefijo: str) -> str | None:
    filtrados = [f for f in archivos if f.startswith(prefijo)]
    if not filtrados:
        return None
    fechas = []
    for f in filtrados:
        fecha = pl.Series([f]).str.extract(r"(\d{8})").item()
        fechas.append(fecha)
    reciente = (
        pl.DataFrame({"Nombre archivo": filtrados, "Fecha": fechas})
        .sort("Fecha", descending=True)
        .row(0, named=True)
    )
    return reciente["Nombre archivo"]


def leer_excel_regalias(contenido_bytes: bytes) -> pl.DataFrame:
    df = pl.read_excel(
        io.BytesIO(contenido_bytes),
        has_header=False,
        infer_schema_length=0,
    )
    encabezados = dict(zip(df.columns, df.row(1)))
    return (
        df
        .rename(encabezados)
        .slice(2)
        .select(pl.all().name.map(lambda x: x.strip()))
    )


# ── Funciones de cálculo ───────────────────────────────────────────────────────
def dias_desde_aprobacion_hasta_primer_proceso(estado_proyecto, fecha_aprobacion, fecha_corte_gesproy):
    condicion_estado = pl.col(estado_proyecto) == "SIN CONTRATAR"
    condicion_fechas = (~pl.col(fecha_aprobacion).is_null()) & (~pl.col(fecha_corte_gesproy).is_null())
    return (
        pl.when(condicion_estado & condicion_fechas)
        .then(pl.col(fecha_corte_gesproy) - pl.col(fecha_aprobacion))
        .otherwise(None)
    )


def dias_desde_apertura_hasta_primer_contrato(estado_proyecto, fecha_acta_inicio, fecha_primer_proceso):
    condicion_estado = pl.col(estado_proyecto) == "SIN CONTRATAR"
    condicion_fechas = ~pl.col(fecha_primer_proceso).is_null()
    return (
        pl.when(condicion_estado & condicion_fechas)
        .then(pl.col(fecha_acta_inicio) - pl.col(fecha_primer_proceso))
        .otherwise(None)
    )


def dias_desde_suscripcion_hasta_fecha_acta_inicio(estado_proyecto, fecha_corte_gesproy, fecha_suscripcion):
    condicion_estado = pl.col(estado_proyecto) == "CONTRATADO SIN ACTA DE INICIO"
    return (
        pl.when(condicion_estado)
        .then(pl.col(fecha_corte_gesproy) - pl.col(fecha_suscripcion))
        .otherwise(None)
    )


def calificacion_desempeño_contratacion(
    estado_proyecto, fecha_aprobacion, fecha_apertura_primer_proceso,
    fecha_corte_gesproy, fecha_suscripcion, fecha_acta_inicio
):
    condicion_estado   = pl.col(estado_proyecto) == "SIN CONTRATAR"
    condicion_estado_2 = pl.col(estado_proyecto) == "CONTRATADO SIN ACTA DE INICIO"
    condicion_escenario_2 = (~pl.col(fecha_aprobacion).is_null()) & (pl.col(fecha_apertura_primer_proceso).is_null())
    condicion_escenario_3 = (pl.col(fecha_suscripcion).is_null()) & (~pl.col(fecha_apertura_primer_proceso).is_null())
    condicion_escenario_4 = ~pl.col(fecha_suscripcion).is_null()

    dias_aprobacion_corte  = (pl.col(fecha_corte_gesproy) - pl.col(fecha_aprobacion)).dt.total_days()
    dias_apertura_corte    = (pl.col(fecha_corte_gesproy) - pl.col(fecha_apertura_primer_proceso)).dt.total_days()
    dias_corte_suscripcion = (pl.col(fecha_corte_gesproy) - pl.col(fecha_suscripcion)).dt.total_days()

    formula_condicion_3 = (180 - dias_apertura_corte) / (180 - 121)
    formula_condicion_4 = (60 - dias_corte_suscripcion) / 29 * 100

    return (
        pl.when(condicion_estado & condicion_escenario_2 & (dias_aprobacion_corte >= 0) & (dias_aprobacion_corte <= 180)).then(pl.lit(100))
        .when(condicion_estado & condicion_escenario_2 & (dias_aprobacion_corte > 180)).then(pl.lit(0))
        .when(condicion_estado & condicion_escenario_3 & (dias_apertura_corte >= 0) & (dias_apertura_corte <= 120)).then(pl.lit(100))
        .when(condicion_estado & condicion_escenario_3 & (dias_apertura_corte > 120) & (dias_apertura_corte <= 180)).then(formula_condicion_3)
        .when(condicion_estado & condicion_escenario_3 & (dias_apertura_corte > 180)).then(pl.lit(0))
        .when(condicion_estado_2 & condicion_escenario_4 & (dias_corte_suscripcion >= 0) & (dias_corte_suscripcion <= 30)).then(pl.lit(100))
        .when(condicion_estado_2 & condicion_escenario_4 & (dias_corte_suscripcion > 30) & (dias_corte_suscripcion <= 60)).then(formula_condicion_4)
        .when(condicion_estado_2 & condicion_escenario_4 & (dias_corte_suscripcion > 60)).then(pl.lit(0))
        .otherwise(pl.lit(None))
    )


# ── Carga de archivos ──────────────────────────────────────────────────────────
st.header("Carga de archivos fuente")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Reportes Gesproy")
    uploads_gesproy = st.file_uploader(
        "Sube los archivos CG-proy, CG-cttos y CG-carga",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="gesproy",
    )

with col2:
    st.subheader("Versión anterior de la Matriz")
    upload_version_anterior = st.file_uploader(
        "Sube el archivo de la versión anterior (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        key="version_anterior",
    )

# ── Botón de generación ────────────────────────────────────────────────────────
st.divider()
st.header("Generar Matriz")

if st.button("Generar Matriz", type="primary", use_container_width=True):

    # Validar que se subieron todos los archivos necesarios
    errores = []
    if not uploads_gesproy:
        errores.append("Faltan los archivos de Gesproy (CG-proy, CG-cttos, CG-carga).")
    else:
        nombres_gesproy = [f.name for f in uploads_gesproy]
        nombre_proy  = ruta_reciente(nombres_gesproy, "CG-proy")
        nombre_cttos = ruta_reciente(nombres_gesproy, "CG-cttos")
        nombre_carga = ruta_reciente(nombres_gesproy, "CG-carga")
        if not nombre_proy:  errores.append("No se encontró archivo CG-proy.")
        if not nombre_cttos: errores.append("No se encontró archivo CG-cttos.")
        if not nombre_carga: errores.append("No se encontró archivo CG-carga.")
    if not upload_version_anterior:
        errores.append("Falta el archivo de la versión anterior de la Matriz.")

    if errores:
        for e in errores:
            st.error(e)
        st.stop()

    # Mapa nombre → bytes de los archivos Gesproy
    gesproy_bytes = {f.name: f.read() for f in uploads_gesproy}

    progress = st.progress(0, text="Iniciando...")

    try:
        # ── Proyectos ──────────────────────────────────────────────────────────
        progress.progress(10, text="Leyendo proyectos...")
        regalias_proyectos = (
            leer_excel_regalias(gesproy_bytes[nombre_proy])
            .select(
                "BPIN", "NOMBRE PROYECTO", "SECTOR", "ESTADO PROYECTO",
                "VALOR SGR", "VALOR NACIÓN", "VALOR OTROS",
                "VALOR OTRAS FUENTES NO SUIFP", "VALOR TOTAL PROYECTO", "VALOR PAGOS",
            )
            .filter(~pl.col("ESTADO PROYECTO").is_in(["CERRADO", "DESAPROBADO"]))
            .with_columns(
                pl.col(
                    "VALOR NACIÓN", "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP",
                    "VALOR TOTAL PROYECTO", "VALOR SGR", "VALOR PAGOS",
                ).cast(pl.Float64)
            )
        )

        # ── Versión anterior ───────────────────────────────────────────────────
        progress.progress(25, text="Leyendo versión anterior...")
        BPINes_version_anterior = (
            pl.read_excel(
                io.BytesIO(upload_version_anterior.read()),
                table_name="MatrizSeguimientoEvaluacion",
            )
            .select(
                "BPIN", "ALCANCE DEL PROYECTO", "FUENTE DE FINANCIACIÓN",
                "ENTIDAD O SECRETARIA", "INDICADOR DE PRODUCTO MGA",
                "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
                "FECHA DE INCORPORACIÓN DE RECURSOS", "HORIZONTE DEL PROYECTO",
                "FECHA DE FINALIZACIÓN",
            )
        )

        # ── Contratos ──────────────────────────────────────────────────────────
        progress.progress(40, text="Procesando contratos...")
        regalias_contratos = (
            leer_excel_regalias(gesproy_bytes[nombre_cttos])
            .select(
                "BPIN", "FECHA ACT ADTIVO APERT", "ESTADO CONTRATO",
                "FECHA INICIAL", "FECHA SUSCRIPCION", "ULTIMA FECHA PAGO",
                "TIPO CONTRATO", "VALOR TOTAL FUENTES SGR",
            )
            .with_columns(
                pl.col("FECHA INICIAL", "FECHA ACT ADTIVO APERT", "FECHA SUSCRIPCION", "ULTIMA FECHA PAGO")
                .str.head(10)
                .str.to_date(strict=False)
            )
            .rename({
                "FECHA ACT ADTIVO APERT": "FECHA DE APERTURA DEL PRIMER PROCESO",
                "FECHA INICIAL": "FECHA ACTA INICIO",
            })
            .sort(by="VALOR TOTAL FUENTES SGR", descending=True)
        )

        contratos = []
        for (bpin,), j in regalias_contratos.group_by("BPIN"):
            condicion_filtro = (
                j.sort("VALOR TOTAL FUENTES SGR", descending=True)
                .with_columns(
                    pl.when(pl.col("TIPO CONTRATO").is_in(["Obra pública", "Consultoría", "Convenios de Cooperación", "Interadministrativos"]))
                    .then(pl.lit("SI")).otherwise(pl.lit("NO")).alias("PRIORIZADO")
                )
                .with_columns(
                    pl.when(
                        (pl.col("PRIORIZADO") == "NO") &
                        pl.col("TIPO CONTRATO").is_in(["Suministro", "Contratos o convenios con entidades sin ánimo de lucro"])
                    )
                    .then(pl.lit("SI")).otherwise(pl.lit("NO")).alias("PRIORIZADO_2")
                )
                .with_columns(
                    ((pl.col("PRIORIZADO") == "SI") | (pl.col("PRIORIZADO_2") == "SI")).alias("NIVEL")
                )
            )
            tiene_nivel = condicion_filtro.select(pl.col("NIVEL").any()).item()
            if tiene_nivel:
                contratos.append(condicion_filtro.filter(pl.col("NIVEL")).head(1).drop("PRIORIZADO", "PRIORIZADO_2", "NIVEL"))
            else:
                contratos.append(condicion_filtro.head(1).drop("PRIORIZADO", "PRIORIZADO_2", "NIVEL"))
        regalias_contratos = pl.concat(contratos)

        # ── Cargue ─────────────────────────────────────────────────────────────
        progress.progress(55, text="Leyendo cargue...")
        regalias_cargue = (
            leer_excel_regalias(gesproy_bytes[nombre_carga])
            .select("BPIN", "FECHA APROBACIÓN PROYECTO", "AVANCE FISICO", "AVANCE FINANCIERO")
            .with_columns(
                pl.col("FECHA APROBACIÓN PROYECTO").str.head(10).str.to_date(strict=False),
                pl.col("AVANCE FISICO", "AVANCE FINANCIERO").cast(pl.Float64, strict=False),
            )
        )

        # ── Consolidación ──────────────────────────────────────────────────────
        progress.progress(65, text="Consolidando datos...")
        df = (
            regalias_proyectos
            .join(BPINes_version_anterior, on="BPIN", how="left")
            .join(regalias_contratos, on="BPIN", how="left")
            .join(regalias_cargue, on="BPIN", how="left")
            .select(
                "BPIN", "ENTIDAD O SECRETARIA", "NOMBRE PROYECTO",
                "ALCANCE DEL PROYECTO", "SECTOR", "INDICADOR DE PRODUCTO MGA",
                "ESTADO PROYECTO", "ESTADO CONTRATO", "TIPO CONTRATO",
                "FUENTE DE FINANCIACIÓN", "VALOR SGR", "VALOR NACIÓN",
                "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP",
                "VALOR TOTAL PROYECTO", "VALOR PAGOS", "ULTIMA FECHA PAGO",
                "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
                "FECHA DE INCORPORACIÓN DE RECURSOS", "AVANCE FISICO", "AVANCE FINANCIERO",
                pl.lit("").alias("CPI"),
                pl.lit("").alias("SPI"),
                "FECHA APROBACIÓN PROYECTO",
                "FECHA DE APERTURA DEL PRIMER PROCESO",
                "FECHA SUSCRIPCION", "FECHA ACTA INICIO",
                "HORIZONTE DEL PROYECTO", "FECHA DE FINALIZACIÓN",
                pl.date(2026, datetime.now().month, 15).alias("FECHA DE CORTE GESPROY"),
                pl.lit("").alias("INFORMACIÓN SOLICITADA"),
                pl.lit("").alias("INFORMACIÓN RECIBIDA"),
                pl.lit("").alias("FECHA DE RECIBO DE INFORMACIÓN"),
                pl.lit("").alias("DESEMPEÑO EN EL CRONOGRAMA"),
                pl.lit("").alias("DESEMPEÑO EN EL COSTO"),
                pl.lit("").alias("COLUMNA APOYO"),
                pl.lit("").alias("BRECHA FISICO - FINANCIERA"),
                pl.lit("").alias("CONTROL EXTERNALIDADES"),
                pl.lit("").alias("CALIFICACIÓN INFORMACIÓN A TIEMPO"),
                pl.lit("").alias("CALIFICACIÓN CALIDAD INFORMACIÓN"),
                pl.lit("").alias("COLUMNA APOYO 2"),
                pl.lit("").alias("CALIFICACIÓN EJECUCIÓN DEL PROYECTO"),
                pl.lit("").alias("COMENTARIOS CALIFICACIÓN"),
            )
        )

        # ── Cálculos ───────────────────────────────────────────────────────────
        progress.progress(75, text="Calculando indicadores...")

        columnas_dias = {
            "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
            "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
            "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
        }

        df = (
            df.with_columns(
                dias_desde_aprobacion_hasta_primer_proceso(
                    "ESTADO PROYECTO", "FECHA APROBACIÓN PROYECTO", "FECHA DE CORTE GESPROY"
                ).alias("DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO"),
                dias_desde_apertura_hasta_primer_contrato(
                    "ESTADO PROYECTO", "FECHA ACTA INICIO", "FECHA DE APERTURA DEL PRIMER PROCESO"
                ).alias("DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO"),
                dias_desde_suscripcion_hasta_fecha_acta_inicio(
                    "ESTADO PROYECTO", "FECHA DE CORTE GESPROY", "FECHA SUSCRIPCION"
                ).alias("DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO"),
                calificacion_desempeño_contratacion(
                    "ESTADO PROYECTO", "FECHA APROBACIÓN PROYECTO",
                    "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA DE CORTE GESPROY",
                    "FECHA SUSCRIPCION", "FECHA ACTA INICIO"
                ).alias("CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN"),
            )
        )

        # ── Definición columnas para exportar ──────────────────────────────────
        columnas_datos_generales = [
            "BPIN", "ENTIDAD O SECRETARIA", "NOMBRE PROYECTO", "ALCANCE DEL PROYECTO",
            "SECTOR", "INDICADOR DE PRODUCTO MGA", "ESTADO PROYECTO", "ESTADO CONTRATO",
            "TIPO CONTRATO", "FUENTE DE FINANCIACIÓN", "VALOR SGR", "VALOR NACIÓN",
            "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP", "VALOR TOTAL PROYECTO",
            "VALOR PAGOS", "ULTIMA FECHA PAGO", "FECHA DE MIGRACIÓN A GESPROY",
            "FECHA DE ASIGNACIÓN DE RECURSOS", "FECHA DE INCORPORACIÓN DE RECURSOS",
        ]
        columnas_datos_calificacion = [
            "AVANCE FISICO", "AVANCE FINANCIERO", "CPI", "SPI",
            "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
            "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
            "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY",
            "INFORMACIÓN SOLICITADA", "INFORMACIÓN RECIBIDA",
            "FECHA DE RECIBO DE INFORMACIÓN",
            "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
            "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
            "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
            "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
            "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA", "CONTROL EXTERNALIDADES",
        ]
        columnas_evaluacion = [
            "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
            "CALIFICACIÓN INFORMACIÓN A TIEMPO",
            "CALIFICACIÓN CALIDAD INFORMACIÓN",
            "COLUMNA APOYO 2",
            "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
            "COMENTARIOS CALIFICACIÓN",
        ]
        todas_las_columnas = columnas_datos_generales + columnas_datos_calificacion + columnas_evaluacion

        # ── Vista previa ───────────────────────────────────────────────────────
        progress.progress(80, text="Preparando vista previa...")
        st.divider()
        st.header("Vista previa de datos")
        st.caption(f"**{len(df)} proyectos** encontrados")
        cols_preview = [
            "BPIN", "ENTIDAD O SECRETARIA", "NOMBRE PROYECTO",
            "ESTADO PROYECTO", "VALOR SGR", "AVANCE FISICO", "AVANCE FINANCIERO",
            "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
        ]
        st.dataframe(
            df.select([c for c in cols_preview if c in df.columns]).to_pandas(),
            use_container_width=True,
            height=350,
        )

        # ── Exportar Excel ─────────────────────────────────────────────────────
        progress.progress(88, text="Generando Excel...")

        AZUL         = "#B7DEE8"
        AZUL_HEADER  = "#31869B"
        NARANJA        = "#F1A983"
        NARANJA_HEADER = "#BE5014"

        color_por_columna = (
            {c: (AZUL, AZUL_HEADER) for c in columnas_datos_generales}
            | {c: (NARANJA, NARANJA_HEADER) for c in columnas_datos_calificacion}
            | {c: (AZUL, AZUL_HEADER) for c in columnas_evaluacion}
        )

        columnas_fecha = {
            "ULTIMA FECHA PAGO", "FECHA APROBACIÓN PROYECTO",
            "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA SUSCRIPCION",
            "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
            "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
            "FECHA DE INCORPORACIÓN DE RECURSOS", 
            "FECHA DE FINALIZACIÓN",
        }
        columnas_numero = {
            "VALOR SGR", "VALOR NACIÓN", "VALOR OTROS",
            "VALOR OTRAS FUENTES NO SUIFP", "VALOR TOTAL PROYECTO",
            "VALOR PAGOS", "AVANCE FISICO", "AVANCE FINANCIERO",
        }
        columnas_con_formula = {
            "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
            "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA",
            "CALIFICACIÓN INFORMACIÓN A TIEMPO",
            "COLUMNA APOYO 2", "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
        }

        def col_letter(i: int) -> str:
            result = ""
            while True:
                result = chr(i % 26 + 65) + result
                i = i // 26 - 1
                if i < 0:
                    break
            return result

        idx = {col: todas_las_columnas.index(col) for col in todas_las_columnas}

        def formulas_para_fila(r: int) -> dict:
            def ref(c): return f"${col_letter(idx[c])}{r}"
            estado         = ref("ESTADO PROYECTO")
            spi            = ref("SPI")
            cpi            = ref("CPI")
            af             = ref("AVANCE FINANCIERO")
            afis           = ref("AVANCE FISICO")
            columna_apoyo  = ref("COLUMNA APOYO")
            columna_apoyo2 = ref("COLUMNA APOYO 2")
            info_sol       = ref("INFORMACIÓN SOLICITADA")
            cron           = ref("DESEMPEÑO EN EL CRONOGRAMA")
            cost           = ref("DESEMPEÑO EN EL COSTO")
            brecha_ref     = ref("BRECHA FISICO - FINANCIERA")
            ext            = ref("CONTROL EXTERNALIDADES")
            fecha_recibo   = ref("FECHA DE RECIBO DE INFORMACIÓN")
            en_ejecucion   = f'{estado}="CONTRATADO EN EJECUCI\u00d3N"'

            def formula_indicador(v):
                p = f"(({v}-0.38)/(0.84-0.38))*100"
                return (
                    f'=IF(AND({v}>1.3,{en_ejecucion}),0,'
                    f'IF(AND({v}>1.25,{v}<=1.3,{en_ejecucion}),30,'
                    f'IF(AND({v}>1.2,{v}<=1.25,{en_ejecucion}),90,'
                    f'IF(AND({v}>=0.84,{v}<1.2,{en_ejecucion}),100,'
                    f'IF(AND({v}>=0.38,{v}<0.84,{en_ejecucion}),{p},'
                    f'IF(AND({v}<0.38,{en_ejecucion}),0,'
                    f'"")'
                    f')))))'
                )

            b = f"({af}-{afis})"
            formula_columna_apoyo = (
                f'=IF(AND({af}<60,{b}<=50),100,IF(AND({af}<60,{b}>50),0,'
                f'IF(AND({af}>=60,{af}<70,{b}<=40),100,IF(AND({af}>=60,{af}<70,{b}>40),0,'
                f'IF(AND({af}>=70,{af}<80,{b}<=30),100,IF(AND({af}>=70,{af}<80,{b}>30),0,'
                f'IF(AND({af}>=80,{af}<90,{b}<=20),100,IF(AND({af}>=80,{af}<90,{b}>20),0,'
                f'IF(AND({af}>=90,{af}<=100,{b}<=10),100,'
                f'IF(AND({af}>=90,{af}<=100,{b}>10),0,""))))))))))'
            )
            formula_brecha = (
                f'=IF(AND({afis}>{af},{en_ejecucion}),100,'
                f'IF(AND({af}>{afis},{b}>50,{en_ejecucion}),0,'
                f'IF(AND({af}>{afis},{b}<=50,{af}<=60),100,'
                f'IF(AND({af}>{afis},{b}<=50,{af}>60),{columna_apoyo},""))))'
            )
            formula_info_tiempo = (
                f'=IF(ISNUMBER({fecha_recibo}),'
                f'IF(DAY({fecha_recibo})=10,100,'
                f'IF(DAY({fecha_recibo})=11,80,'
                f'IF(DAY({fecha_recibo})=12,50,0))),"")'
            )
            formula_apoyo2 = (
                f'=IF({ext}=0,100,IF({ext}=1,90,IF({ext}=2,75,'
                f'IF({ext}=3,60,IF({ext}=4,50,IF({ext}=5,25,'
                f'IF({ext}>=6,0,"")))))))'
            )
            ponderado = f"({cron}*0.4+{cost}*0.2+{brecha_ref}*0.4)"
            calculo   = f'IFERROR(IF(ISTEXT({info_sol}),{ponderado},""),"")'
            formula_ejecucion = (
                f'=IFERROR(IF(AND({en_ejecucion},{ext}>1),'
                f'({calculo})*{columna_apoyo2}/100,{calculo}),0)'
            )
            return {
                "DESEMPEÑO EN EL CRONOGRAMA":          formula_indicador(spi),
                "DESEMPEÑO EN EL COSTO":               formula_indicador(cpi),
                "COLUMNA APOYO":                       formula_columna_apoyo,
                "BRECHA FISICO - FINANCIERA":          formula_brecha,
                "CALIFICACIÓN INFORMACIÓN A TIEMPO":   formula_info_tiempo,
                "COLUMNA APOYO 2":                     formula_apoyo2,
                "CALIFICACIÓN EJECUCIÓN DEL PROYECTO": formula_ejecucion,
            }

        # Preparar datos
        df_export = (
            df.select(todas_las_columnas)
            .with_columns([
                pl.col(c).dt.total_seconds().truediv(86400).cast(pl.Int64).alias(c)
                for c in columnas_dias
            ])
        )
        datos = df_export.to_pandas()

        ROW_HEIGHT = 70
        COL_WIDTH  = 25
        N_COLS     = len(todas_las_columnas)
        N_ROWS     = len(datos)
        FILA_SECCION = 0
        FILA_HEADER  = 1
        FILA_DATOS   = 2

        # Escribir en buffer en memoria
        output_buffer = io.BytesIO()
        workbook = xlsxwriter.Workbook(output_buffer, {"in_memory": True})
        ws = workbook.add_worksheet("MatrizEvaluacionSeguimiento")

        base = dict(font_name="Roboto", font_size=10, border=1, text_wrap=True)
        def fmt(extra): return workbook.add_format({**base, **extra})

        fmt_titulo_azul    = fmt({"bold": True, "font_color": "white", "bg_color": AZUL,           "align": "center", "valign": "vcenter", "font_size": 12})
        fmt_titulo_naranja = fmt({"bold": True, "font_color": "white", "bg_color": NARANJA,         "align": "center", "valign": "vcenter", "font_size": 12})
        fmt_header_azul    = fmt({"bold": True, "font_color": "white", "bg_color": AZUL_HEADER,    "align": "center", "valign": "vcenter"})
        fmt_header_naranja = fmt({"bold": True, "font_color": "white", "bg_color": NARANJA_HEADER, "align": "center", "valign": "vcenter"})
        fmt_celda          = fmt({"valign": "vcenter", "align": "center"})
        fmt_fecha          = fmt({"valign": "vcenter", "align": "center", "num_format": "yyyy-mm-dd"})
        fmt_numero         = fmt({"valign": "vcenter", "align": "center", "num_format": "#,##0.000"})
        fmt_dias           = fmt({"valign": "vcenter", "align": "center"})

        for col_idx in range(N_COLS):
            ws.set_column(col_idx, col_idx, COL_WIDTH)
        ws.set_row(FILA_SECCION, 40)
        ws.set_row(FILA_HEADER, ROW_HEIGHT)
        for row_idx in range(N_ROWS):
            ws.set_row(FILA_DATOS + row_idx, ROW_HEIGHT)

        secciones = [
            ("DATOS GENERALES",         columnas_datos_generales,    fmt_titulo_azul),
            ("DATOS PARA CALIFICACIÓN", columnas_datos_calificacion, fmt_titulo_naranja),
            ("EVALUACIÓN",              columnas_evaluacion,         fmt_titulo_azul),
        ]
        col_offset = 0
        for titulo, columnas, ft in secciones:
            n = len(columnas)
            ws.merge_range(FILA_SECCION, col_offset, FILA_SECCION, col_offset + n - 1, titulo, ft)
            col_offset += n

        ws.add_table(
            FILA_HEADER, 0, FILA_HEADER + N_ROWS, N_COLS - 1,
            {
                "name": "MatrizEvaluacionSeguimiento",
                "style": "Table Style Medium 2",
                "autofilter": True,
                "header_row": True,
                "columns": [{"header": col_name} for col_name in todas_las_columnas],
            },
        )

        for col_idx, col_name in enumerate(todas_las_columnas):
            fmt_hdr = fmt_header_azul if color_por_columna[col_name][0] == AZUL else fmt_header_naranja
            ws.write(FILA_HEADER, col_idx, col_name, fmt_hdr)

        for col_name in ["COLUMNA APOYO", "COLUMNA APOYO 2"]:
            ci = todas_las_columnas.index(col_name)
            ws.set_column(ci, ci, COL_WIDTH, None, {"hidden": True})

        for row_idx in range(N_ROWS):
            excel_row_num = row_idx + FILA_DATOS + 1
            formulas_fila = formulas_para_fila(excel_row_num)
            for col_idx, col_name in enumerate(todas_las_columnas):
                er = row_idx + FILA_DATOS
                ec = col_idx
                if col_name in columnas_con_formula:
                    ws.write_formula(er, ec, formulas_fila[col_name], fmt_celda)
                    continue
                valor = datos.iloc[row_idx, col_idx]
                if col_name in columnas_fecha:
                    if pandas.isnull(valor):
                        ws.write_blank(er, ec, None, fmt_fecha)
                    else:
                        ws.write_datetime(er, ec, valor.to_pydatetime(), fmt_fecha)
                elif col_name in columnas_numero:
                    if pandas.isnull(valor):
                        ws.write_blank(er, ec, None, fmt_numero)
                    else:
                        ws.write_number(er, ec, float(valor), fmt_numero)
                elif col_name in columnas_dias:
                    if pandas.isnull(valor):
                        ws.write_blank(er, ec, None, fmt_dias)
                    else:
                        ws.write_number(er, ec, int(valor), fmt_dias)
                else:
                    ws.write(er, ec, None if pandas.isnull(valor) else valor, fmt_celda)

        workbook.close()
        output_buffer.seek(0)

        progress.progress(100, text="Listo!")

        # ── Descarga ───────────────────────────────────────────────────────────
        nombre_archivo = f"MatrizSeguimientoEvaluacion_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.success(f"Matriz generada con **{N_ROWS} proyectos**.")
        st.divider()
        st.header("Descargar")
        st.download_button(
            label="Descargar Matriz Excel",
            data=output_buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except Exception as e:
        progress.empty()
        st.error(f"Error al procesar los archivos: {e}")
        st.exception(e)

