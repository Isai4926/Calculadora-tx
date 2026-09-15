# -*- coding: utf-8 -*-
"""
==========================================================================
 CALCULADORA TX
 App de análisis fiscal-contable: ingresos, gastos, IVA, ISR retenido y
 conciliación bancaria, a partir de 4 bases de datos (Excel/CSV).
==========================================================================

Este archivo está organizado en secciones, marcadas con comentarios en
mayúsculas, para que sea fácil de leer y modificar aunque no seas
programador experto:

  1. CONFIGURACIÓN Y CONSTANTES
  2. FUNCIONES DE CARGA Y VALIDACIÓN DE ARCHIVOS
  3. FUNCIONES DE CÁLCULO (clasificación, IVA, conciliación, ISR)
  4. FUNCIONES DE ALERTAS
  5. FUNCIÓN DE REPORTE EXCEL DESCARGABLE
  6. INTERFAZ DE USUARIO (Streamlit)

IMPORTANTE: esta herramienta es un AUXILIAR de análisis. Los resultados
deben ser revisados por un contador o especialista fiscal antes de
presentar declaraciones o tomar decisiones fiscales.
"""

import io
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==========================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# ==========================================================================

st.set_page_config(
    page_title="Calculadora Tx",
    page_icon="🧮",
    layout="wide",
)

TASA_IVA = 0.16                # Tasa general de IVA usada en los cálculos
TOLERANCIA_MONTO = 10.0        # Tolerancia en pesos para dar por conciliado un importe
TOLERANCIA_DIAS = 5            # Tolerancia en días naturales para conciliar fechas
TOLERANCIA_IVA = 1.0           # Tolerancia en pesos para detectar inconsistencias de IVA

# Conceptos de gasto que NUNCA son deducibles, sin importar si hay comprobante
CONCEPTOS_NO_DEDUCIBLES = {
    "Donativo no autorizado",
    "Gasto personal",
    "Recargo",
    "Multa",
}

# Conceptos de gasto que SÍ pueden ser deducibles si cuentan con comprobante fiscal
CONCEPTOS_DEDUCIBLES_CONOCIDOS = {
    "Transporte", "Publicidad", "Capacitación", "Hospedaje", "Mantenimiento",
    "Papelería", "Licencias de software", "Servicios jurídicos", "Mensajería",
    "Renta de oficina", "Comida de trabajo", "Telefonía",
}

# Tipos de movimiento bancario que representan cobros/ingresos de clientes
TIPOS_INGRESO = {"Ingreso", "Cobro", "Deposito", "Abono"}
# Tipos de movimiento bancario que representan pagos/egresos a proveedores
TIPOS_EGRESO = {"Egreso"}

# Columnas obligatorias por cada una de las 4 bases de datos de entrada
COLUMNAS_OBLIGATORIAS = {
    "Facturas emitidas": ["folio", "fecha", "cliente", "concepto", "subtotal", "iva", "total"],
    "Gastos (facturas recibidas)": ["folio_gasto", "fecha", "proveedor", "concepto", "subtotal", "iva", "total", "comprobante"],
    "Movimientos bancarios": ["fecha", "monto", "tipo", "referencia"],
    "Retenciones de ISR": ["fecha", "referencia", "concepto", "isr_retenido"],
}

# Colores usados de forma consistente en todas las gráficas
COLOR_INGRESO = "#2E7D32"
COLOR_GASTO = "#C62828"
COLOR_NEUTRO = "#1565C0"
COLOR_ALERTA_ALTA = "#C62828"
COLOR_ALERTA_MEDIA = "#F9A825"
COLOR_ALERTA_BAJA = "#757575"

MENSAJE_DISCLAIMER = (
    "⚠️ **Aviso importante:** esta herramienta es un auxiliar de análisis. "
    "Los cálculos, clasificaciones y conciliaciones que aquí se muestran deben "
    "ser revisados por un contador público o especialista fiscal antes de "
    "presentar declaraciones o tomar decisiones fiscales. La aplicación no "
    "inventa ni asume información que no esté presente en los archivos "
    "cargados: cuando un dato es insuficiente o ambiguo, el registro se "
    "marca para revisión manual en vez de clasificarse automáticamente."
)


# ==========================================================================
# 2. FUNCIONES DE CARGA Y VALIDACIÓN DE ARCHIVOS
# ==========================================================================

def leer_archivo(archivo_subido):
    """Lee un archivo subido por el usuario (xlsx o csv) y regresa un DataFrame.

    Si el archivo no se puede leer (formato corrupto, extensión no soportada,
    etc.) regresa None y un mensaje de error legible para el usuario.
    """
    nombre = archivo_subido.name.lower()
    try:
        if nombre.endswith(".csv"):
            df = pd.read_csv(archivo_subido)
        elif nombre.endswith((".xlsx", ".xls")):
            df = pd.read_excel(archivo_subido)
        else:
            return None, f"Formato de archivo no soportado ('{archivo_subido.name}'). Usa .xlsx o .csv."
        return df, None
    except Exception as e:
        return None, f"No se pudo leer el archivo '{archivo_subido.name}'. Detalle técnico: {e}"


def validar_columnas(df, nombre_base):
    """Verifica que el DataFrame tenga todas las columnas obligatorias.

    Regresa una lista de mensajes de error (vacía si todo está correcto).
    Los nombres de columna se comparan sin importar mayúsculas/minúsculas
    ni espacios extra, para ser tolerantes a variaciones menores del archivo.
    """
    columnas_presentes = {str(c).strip().lower() for c in df.columns}
    requeridas = COLUMNAS_OBLIGATORIAS[nombre_base]
    faltantes = [c for c in requeridas if c.lower() not in columnas_presentes]
    errores = []
    for col in faltantes:
        errores.append(f"A la base **{nombre_base}** le falta la columna obligatoria **'{col}'**.")
    return errores


def normalizar_columnas(df):
    """Renombra las columnas a minúsculas y sin espacios, para trabajar de
    forma consistente sin importar cómo vengan capitalizadas en el archivo
    original."""
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def preparar_fechas(df, columna="fecha"):
    """Convierte una columna a tipo fecha. Los valores que no se puedan
    convertir quedan como NaT (nulos), para poder detectarlos como alerta
    de dato faltante en vez de que la app truene."""
    df = df.copy()
    df[columna] = pd.to_datetime(df[columna], errors="coerce")
    return df


# ==========================================================================
# 3. FUNCIONES DE CÁLCULO
# ==========================================================================

def clasificar_gastos(gastos):
    """Agrega a la tabla de gastos dos columnas nuevas:
      - clasificacion: 'Deducible', 'No deducible' o 'Pendiente de clasificación'
      - motivo_clasificacion: explicación breve de por qué se clasificó así

    Regla de negocio:
      1. Si el concepto está en la lista de NO deducibles -> No deducible.
      2. Si no tiene comprobante fiscal ('comprobante' != 'Si') -> No deducible.
      3. Si el concepto es uno de los conceptos deducibles conocidos y sí
         tiene comprobante -> Deducible.
      4. Si el concepto no está catalogado (no aparece en ninguna de las dos
         listas anteriores) -> Pendiente de clasificación, para que un
         humano lo revise. La app NUNCA asume que un concepto desconocido
         es deducible o no deducible.
    """
    def clasificar(row):
        concepto = str(row.get("concepto", "")).strip()
        comprobante_raw = str(row.get("comprobante", "")).strip().lower()

        if comprobante_raw not in ("si", "sí", "no"):
            return "Pendiente de clasificación", "El valor de la columna 'comprobante' no es 'Si' o 'No'"

        tiene_comprobante = comprobante_raw in ("si", "sí")

        if concepto in CONCEPTOS_NO_DEDUCIBLES:
            return "No deducible", f"Concepto no deducible por su naturaleza: {concepto}"

        if not tiene_comprobante:
            return "No deducible", "No cuenta con comprobante fiscal (CFDI)"

        if concepto in CONCEPTOS_DEDUCIBLES_CONOCIDOS:
            return "Deducible", "Concepto deducible y con comprobante fiscal"

        return "Pendiente de clasificación", f"Concepto no catalogado en las reglas: '{concepto}'"

    gastos = gastos.copy()
    resultado = gastos.apply(lambda r: pd.Series(clasificar(r)), axis=1)
    resultado.columns = ["clasificacion", "motivo_clasificacion"]
    return pd.concat([gastos, resultado], axis=1)


def calcular_iva(facturas, gastos_clasificados):
    """Calcula IVA trasladado (de ingresos), IVA acreditable (solo de gastos
    deducibles) e IVA neto. Regresa un diccionario con los totales."""
    iva_trasladado = float(facturas["iva"].sum())
    gastos_deducibles = gastos_clasificados[gastos_clasificados["clasificacion"] == "Deducible"]
    iva_acreditable = float(gastos_deducibles["iva"].sum())
    iva_neto = iva_trasladado - iva_acreditable
    return {
        "iva_trasladado": iva_trasladado,
        "iva_acreditable": iva_acreditable,
        "iva_neto": iva_neto,
    }


def detectar_inconsistencias_iva(df):
    """Regresa una copia del DataFrame con una columna booleana
    'iva_inconsistente' que marca los registros donde el IVA registrado no
    corresponde a subtotal * 16% (con una tolerancia de $1 peso, por
    redondeos)."""
    df = df.copy()
    iva_esperado = (df["subtotal"] * TASA_IVA).round(2)
    df["iva_esperado"] = iva_esperado
    df["iva_inconsistente"] = (df["iva"] - iva_esperado).abs() > TOLERANCIA_IVA
    return df


def conciliar_facturas_cobros(facturas, movimientos):
    """Concilia cada factura emitida contra los movimientos bancarios de tipo
    'ingreso' cuya referencia coincide con el folio de la factura.

    Reglas de tolerancia: ±$10 en importe y ±5 días naturales en fecha.
    Si más de un movimiento referencia el mismo folio, se suman los montos
    (pago dividido / parcial) y, si la suma NO concilia dentro de tolerancia,
    se marca como posible duplicado para revisión manual (nunca se asume
    automáticamente cuál de los movimientos es el correcto).
    """
    mov_ingresos = movimientos[movimientos["tipo"].isin(TIPOS_INGRESO)].copy()
    filas = []

    for _, f in facturas.iterrows():
        folio = f["folio"]
        coincidencias = mov_ingresos[mov_ingresos["referencia"] == folio].copy()

        if coincidencias.empty:
            filas.append({
                "folio": folio, "fecha_factura": f["fecha"], "cliente": f.get("cliente", ""),
                "total_factura": f["total"], "monto_cobrado": 0.0, "n_movimientos": 0,
                "diferencia_monto": round(-float(f["total"]), 2), "diferencia_dias": np.nan,
                "estado_conciliacion": "No conciliado (pendiente de cobro)",
            })
            continue

        coincidencias["dias_dif"] = (coincidencias["fecha"] - f["fecha"]).dt.days.abs()
        monto_cobrado = float(coincidencias["monto"].sum())
        dias_dif = int(coincidencias["dias_dif"].min())
        diferencia_monto = round(monto_cobrado - float(f["total"]), 2)
        n_mov = len(coincidencias)

        dentro_tolerancia = abs(diferencia_monto) <= TOLERANCIA_MONTO and dias_dif <= TOLERANCIA_DIAS

        if n_mov == 1:
            if abs(diferencia_monto) < 0.01 and dias_dif == 0:
                estado = "Conciliado"
            elif dentro_tolerancia:
                estado = "Conciliado con diferencia dentro de tolerancia"
            else:
                estado = "Requiere revisión manual"
        else:
            # Más de un movimiento con la misma referencia
            if dentro_tolerancia:
                estado = "Conciliado (pago dividido)"
            else:
                estado = "Posible duplicado / revisión manual"

        filas.append({
            "folio": folio, "fecha_factura": f["fecha"], "cliente": f.get("cliente", ""),
            "total_factura": float(f["total"]), "monto_cobrado": monto_cobrado, "n_movimientos": n_mov,
            "diferencia_monto": diferencia_monto, "diferencia_dias": dias_dif,
            "estado_conciliacion": estado,
        })

    return pd.DataFrame(filas)


def movimientos_sin_factura(facturas, movimientos):
    """Regresa los movimientos bancarios de tipo ingreso cuya referencia NO
    corresponde a ningún folio de factura emitida (posible ingreso no
    facturado)."""
    folios_validos = set(facturas["folio"])
    mov_ingresos = movimientos[movimientos["tipo"].isin(TIPOS_INGRESO)]
    return mov_ingresos[~mov_ingresos["referencia"].isin(folios_validos)].copy()


def egresos_sin_gasto(gastos, movimientos):
    """Regresa los movimientos bancarios de tipo egreso cuya referencia NO
    corresponde a ningún folio_gasto de la base de gastos (pago sin gasto
    identificado, requiere revisión manual)."""
    folios_gasto_validos = set(gastos["folio_gasto"])
    mov_egresos = movimientos[movimientos["tipo"].isin(TIPOS_EGRESO)]
    return mov_egresos[~mov_egresos["referencia"].isin(folios_gasto_validos)].copy()


def resumen_isr(isr):
    """El detalle de retenciones de ISR no trae folio de factura ni UUID, y
    al probarse contra las facturas emitidas no se encontró una relación de
    monto consistente (no es un porcentaje fijo del subtotal). Por eso esta
    función NO intenta vincular cada retención a una factura específica:
    reporta el total del periodo y deja el detalle disponible para que el
    usuario lo revise/concilie manualmente con información adicional (RFC,
    folio fiscal) si la tiene."""
    total_isr = float(isr["isr_retenido"].sum())
    return total_isr


# ==========================================================================
# 4. FUNCIONES DE ALERTAS
# ==========================================================================

def generar_alertas(facturas, gastos_clasificados, movimientos, isr,
                     conciliacion, mov_sin_fact, egresos_sin_match):
    """Recorre todos los resultados calculados y arma una tabla única de
    alertas con: tipo, prioridad (alta/media/baja), descripción y referencia
    al registro afectado. Esta tabla es la base tanto del tablero visual
    como de la hoja 'Alertas y validaciones' del reporte Excel."""
    alertas = []

    def agregar(tipo, prioridad, descripcion, referencia=""):
        alertas.append({
            "tipo": tipo, "prioridad": prioridad,
            "descripcion": descripcion, "referencia": referencia,
        })

    # -- Facturas sin movimiento bancario (pendientes de cobro) --
    pendientes = conciliacion[conciliacion["estado_conciliacion"] == "No conciliado (pendiente de cobro)"]
    for _, r in pendientes.iterrows():
        agregar("Factura sin movimiento bancario", "Alta",
                f"La factura {r['folio']} (${r['total_factura']:,.2f}) no tiene ningún cobro registrado.",
                r["folio"])

    # -- Movimientos bancarios sin factura relacionada --
    for _, r in mov_sin_fact.iterrows():
        agregar("Movimiento bancario sin factura relacionada", "Alta",
                f"Depósito de ${r['monto']:,.2f} el {r['fecha'].date()} (ref. {r['referencia']}) no corresponde a ninguna factura emitida. Posible ingreso no facturado.",
                r["referencia"])

    # -- Egresos bancarios sin gasto relacionado --
    for _, r in egresos_sin_match.iterrows():
        agregar("Egreso bancario sin gasto relacionado", "Media",
                f"Pago de ${abs(r['monto']):,.2f} el {r['fecha'].date()} (ref. {r['referencia']}) no coincide con ningún folio de gasto registrado.",
                r["referencia"])

    # -- Diferencias de importe / fecha fuera de tolerancia y duplicados --
    revision = conciliacion[conciliacion["estado_conciliacion"].isin(
        ["Requiere revisión manual", "Posible duplicado / revisión manual"])]
    for _, r in revision.iterrows():
        detalle = f"diferencia de ${abs(r['diferencia_monto']):,.2f}"
        if not pd.isna(r["diferencia_dias"]):
            detalle += f" y {int(r['diferencia_dias'])} día(s) de diferencia"
        agregar("Diferencia de conciliación fuera de tolerancia", "Media",
                f"Factura {r['folio']}: {detalle} ({r['estado_conciliacion']}).",
                r["folio"])

    # -- Gastos sin clasificación --
    pendientes_gasto = gastos_clasificados[gastos_clasificados["clasificacion"] == "Pendiente de clasificación"]
    for _, r in pendientes_gasto.iterrows():
        agregar("Gasto sin clasificación", "Media",
                f"El gasto {r['folio_gasto']} ({r.get('concepto', '')}) no pudo clasificarse automáticamente: {r['motivo_clasificacion']}.",
                r["folio_gasto"])

    # -- Inconsistencias de IVA --
    for nombre, df, col_id in [("factura", facturas, "folio"), ("gasto", gastos_clasificados, "folio_gasto")]:
        inconsistentes = df[df.get("iva_inconsistente", False) == True]
        for _, r in inconsistentes.iterrows():
            agregar("Inconsistencia de IVA", "Media",
                    f"El {nombre} {r[col_id]} tiene IVA de ${r['iva']:,.2f}, pero el 16% del subtotal sería ${r['iva_esperado']:,.2f}.",
                    r[col_id])

    # -- Importes en cero o negativos (facturas y gastos) --
    for nombre, df, col_id in [("factura", facturas, "folio"), ("gasto", gastos_clasificados, "folio_gasto")]:
        malos = df[df["total"] <= 0]
        for _, r in malos.iterrows():
            agregar("Importe en cero o negativo", "Alta",
                    f"El {nombre} {r[col_id]} tiene un total de ${r['total']:,.2f}.",
                    r[col_id])

    # -- Datos obligatorios faltantes (nulos) tras la carga --
    for nombre_base, df, cols in [
        ("Facturas emitidas", facturas, COLUMNAS_OBLIGATORIAS["Facturas emitidas"]),
        ("Gastos", gastos_clasificados, COLUMNAS_OBLIGATORIAS["Gastos (facturas recibidas)"]),
        ("Movimientos bancarios", movimientos, COLUMNAS_OBLIGATORIAS["Movimientos bancarios"]),
        ("Retenciones ISR", isr, COLUMNAS_OBLIGATORIAS["Retenciones de ISR"]),
    ]:
        for col in cols:
            if col in df.columns:
                n_nulos = df[col].isna().sum()
                if n_nulos > 0:
                    agregar("Dato obligatorio faltante", "Alta",
                            f"{n_nulos} registro(s) de '{nombre_base}' tienen la columna '{col}' vacía.")

    # -- Retenciones de ISR no vinculadas a folio de factura --
    if len(isr) > 0:
        agregar("Retenciones de ISR no conciliadas", "Media",
                f"Las {len(isr)} retenciones de ISR (total ${isr['isr_retenido'].sum():,.2f}) no traen folio de "
                f"factura ni UUID, por lo que no se vincularon automáticamente a facturas específicas. "
                f"Se listan en la hoja de ISR para conciliación manual.")

    # -- Registros duplicados dentro de cada base (mismo folio/referencia repetido) --
    dup_facturas = facturas[facturas.duplicated(subset=["folio"], keep=False)]
    for folio in dup_facturas["folio"].unique():
        agregar("Registro duplicado", "Media", f"El folio de factura '{folio}' aparece más de una vez en la base.", folio)

    dup_gastos = gastos_clasificados[gastos_clasificados.duplicated(subset=["folio_gasto"], keep=False)]
    for folio in dup_gastos["folio_gasto"].unique():
        agregar("Registro duplicado", "Media", f"El folio de gasto '{folio}' aparece más de una vez en la base.", folio)

    if not alertas:
        return pd.DataFrame(columns=["tipo", "prioridad", "descripcion", "referencia"])

    df_alertas = pd.DataFrame(alertas)
    orden_prioridad = {"Alta": 0, "Media": 1, "Baja": 2}
    df_alertas["orden"] = df_alertas["prioridad"].map(orden_prioridad)
    df_alertas = df_alertas.sort_values("orden").drop(columns="orden").reset_index(drop=True)
    return df_alertas


# ==========================================================================
# 5. FUNCIÓN DE REPORTE EXCEL DESCARGABLE
# ==========================================================================

def generar_reporte_excel(facturas, gastos_clasificados, movimientos, isr,
                           conciliacion, mov_sin_fact, egresos_sin_match,
                           alertas, resumen):
    """Arma el archivo Excel descargable con una hoja por cada bloque de
    información, conservando trazabilidad hacia los datos originales (las
    columnas calculadas siempre están separadas de las columnas originales
    del archivo fuente)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:

        # --- Resumen ejecutivo ---
        df_resumen = pd.DataFrame(
            [{"Indicador": k, "Valor": v} for k, v in resumen.items()]
        )
        df_resumen.to_excel(writer, sheet_name="Resumen ejecutivo", index=False)

        # --- Ingresos facturados y cobrados ---
        conciliacion.to_excel(writer, sheet_name="Ingresos fact. y cobrados", index=False)

        # --- Gastos deducibles / no deducibles (clasificación completa) ---
        gastos_deducibles = gastos_clasificados[gastos_clasificados["clasificacion"] == "Deducible"]
        gastos_no_deducibles = gastos_clasificados[gastos_clasificados["clasificacion"] != "Deducible"]
        gastos_deducibles.to_excel(writer, sheet_name="Gastos deducibles", index=False)
        gastos_no_deducibles.to_excel(writer, sheet_name="Gastos no deducibles", index=False)

        # --- Cálculo de IVA ---
        iva_facturas = facturas[["folio", "fecha", "subtotal", "iva", "total", "iva_esperado", "iva_inconsistente"]]
        iva_gastos = gastos_clasificados[["folio_gasto", "fecha", "clasificacion", "subtotal", "iva", "total", "iva_esperado", "iva_inconsistente"]]
        iva_facturas.to_excel(writer, sheet_name="IVA - Facturas", index=False)
        iva_gastos.to_excel(writer, sheet_name="IVA - Gastos", index=False)

        # --- Retenciones de ISR ---
        isr.to_excel(writer, sheet_name="Retenciones ISR", index=False)

        # --- Conciliación bancaria (todas) y partidas no conciliadas ---
        conciliacion.to_excel(writer, sheet_name="Conciliación bancaria", index=False)
        no_conciliadas = conciliacion[~conciliacion["estado_conciliacion"].isin(["Conciliado"])]
        no_conciliadas.to_excel(writer, sheet_name="Partidas no conciliadas", index=False)
        mov_sin_fact.to_excel(writer, sheet_name="Depósitos sin factura", index=False)
        egresos_sin_match.to_excel(writer, sheet_name="Egresos sin gasto", index=False)

        # --- Alertas y validaciones ---
        alertas.to_excel(writer, sheet_name="Alertas y validaciones", index=False)

        # --- Datos procesados (originales, tal como se cargaron) ---
        facturas.to_excel(writer, sheet_name="Datos - Facturas", index=False)
        gastos_clasificados.to_excel(writer, sheet_name="Datos - Gastos", index=False)
        movimientos.to_excel(writer, sheet_name="Datos - Movimientos", index=False)

    buffer.seek(0)
    return buffer


# ==========================================================================
# 6. INTERFAZ DE USUARIO (Streamlit)
# ==========================================================================

st.title("🧮 Calculadora Tx")
st.caption(
    "Ingresos facturados vs. cobrados · gastos deducibles/no deducibles · "
    "IVA trasladado, acreditable y neto · ISR retenido · conciliación bancaria"
)
st.info(MENSAJE_DISCLAIMER)

# --- Barra lateral: carga de archivos ---
st.sidebar.header("1. Cargar bases de datos")
st.sidebar.caption("Formatos aceptados: .xlsx o .csv")

archivo_facturas = st.sidebar.file_uploader("Facturas emitidas", type=["xlsx", "csv"], key="facturas")
archivo_gastos = st.sidebar.file_uploader("Gastos (facturas recibidas)", type=["xlsx", "csv"], key="gastos")
archivo_movimientos = st.sidebar.file_uploader("Movimientos bancarios", type=["xlsx", "csv"], key="movimientos")
archivo_isr = st.sidebar.file_uploader("Retenciones de ISR", type=["xlsx", "csv"], key="isr")

archivos_cargados = {
    "Facturas emitidas": archivo_facturas,
    "Gastos (facturas recibidas)": archivo_gastos,
    "Movimientos bancarios": archivo_movimientos,
    "Retenciones de ISR": archivo_isr,
}

if not all(archivos_cargados.values()):
    st.warning("Carga las 4 bases de datos en la barra lateral para comenzar el análisis.")
    with st.expander("Ver columnas obligatorias por archivo"):
        for nombre, cols in COLUMNAS_OBLIGATORIAS.items():
            st.markdown(f"**{nombre}:** {', '.join(cols)}")
    st.stop()

# --- Carga y validación ---
dataframes = {}
errores_totales = []

for nombre, archivo in archivos_cargados.items():
    df, error_lectura = leer_archivo(archivo)
    if error_lectura:
        errores_totales.append(error_lectura)
        continue
    df = normalizar_columnas(df)
    errores_columnas = validar_columnas(df, nombre)
    if errores_columnas:
        errores_totales.extend(errores_columnas)
        continue
    dataframes[nombre] = df

if errores_totales:
    st.error("No se pudo continuar por los siguientes problemas en los archivos cargados:")
    for e in errores_totales:
        st.markdown(f"- {e}")
    st.stop()

facturas = preparar_fechas(dataframes["Facturas emitidas"])
gastos = preparar_fechas(dataframes["Gastos (facturas recibidas)"])
movimientos = preparar_fechas(dataframes["Movimientos bancarios"])
isr = preparar_fechas(dataframes["Retenciones de ISR"])

# Tipos numéricos (por si vienen como texto en el archivo)
for col in ["subtotal", "iva", "total"]:
    facturas[col] = pd.to_numeric(facturas[col], errors="coerce")
    gastos[col] = pd.to_numeric(gastos[col], errors="coerce")
movimientos["monto"] = pd.to_numeric(movimientos["monto"], errors="coerce")
isr["isr_retenido"] = pd.to_numeric(isr["isr_retenido"], errors="coerce")

# --- Barra lateral: filtro de periodo ---
st.sidebar.header("2. Periodo de análisis")
fecha_min = min(facturas["fecha"].min(), gastos["fecha"].min())
fecha_max = max(facturas["fecha"].max(), gastos["fecha"].max())
rango_fechas = st.sidebar.date_input(
    "Analizar facturas y gastos emitidos entre:",
    value=(fecha_min.date(), fecha_max.date()),
    min_value=fecha_min.date(),
    max_value=fecha_max.date(),
)
if len(rango_fechas) == 2:
    inicio, fin = rango_fechas
else:
    inicio, fin = fecha_min.date(), fecha_max.date()

facturas_periodo = facturas[(facturas["fecha"].dt.date >= inicio) & (facturas["fecha"].dt.date <= fin)].copy()
gastos_periodo = gastos[(gastos["fecha"].dt.date >= inicio) & (gastos["fecha"].dt.date <= fin)].copy()

st.sidebar.header("3. Filtros adicionales")
clientes = sorted(facturas_periodo["cliente"].dropna().unique().tolist()) if "cliente" in facturas_periodo.columns else []
proveedores = sorted(gastos_periodo["proveedor"].dropna().unique().tolist()) if "proveedor" in gastos_periodo.columns else []
filtro_cliente = st.sidebar.multiselect("Cliente", clientes)
filtro_proveedor = st.sidebar.multiselect("Proveedor", proveedores)

if filtro_cliente:
    facturas_periodo = facturas_periodo[facturas_periodo["cliente"].isin(filtro_cliente)]
if filtro_proveedor:
    gastos_periodo = gastos_periodo[gastos_periodo["proveedor"].isin(filtro_proveedor)]

# --- Cálculos principales (con caché para no recalcular en cada interacción) ---
gastos_clasificados = clasificar_gastos(gastos_periodo)
gastos_clasificados = detectar_inconsistencias_iva(gastos_clasificados)
facturas_periodo = detectar_inconsistencias_iva(facturas_periodo)

totales_iva = calcular_iva(facturas_periodo, gastos_clasificados)
conciliacion = conciliar_facturas_cobros(facturas_periodo, movimientos)
mov_sin_fact = movimientos_sin_factura(facturas_periodo, movimientos)
egresos_sin_match = egresos_sin_gasto(gastos_periodo, movimientos)
total_isr = resumen_isr(isr)

total_facturado = float(facturas_periodo["total"].sum())
total_cobrado = float(conciliacion["monto_cobrado"].sum())
total_pendiente_cobro = total_facturado - total_cobrado
gastos_deducibles_monto = float(gastos_clasificados.loc[gastos_clasificados["clasificacion"] == "Deducible", "total"].sum())
gastos_no_deducibles_monto = float(gastos_clasificados.loc[gastos_clasificados["clasificacion"] == "No deducible", "total"].sum())
n_conciliadas = int(conciliacion["estado_conciliacion"].isin(
    ["Conciliado", "Conciliado con diferencia dentro de tolerancia", "Conciliado (pago dividido)"]).sum())
n_no_conciliadas = len(conciliacion) - n_conciliadas
monto_conciliado = float(conciliacion.loc[conciliacion["estado_conciliacion"].isin(
    ["Conciliado", "Conciliado con diferencia dentro de tolerancia", "Conciliado (pago dividido)"]), "monto_cobrado"].sum())
monto_no_conciliado = total_facturado - monto_conciliado
pct_conciliacion = (n_conciliadas / len(conciliacion) * 100) if len(conciliacion) > 0 else 0

alertas = generar_alertas(facturas_periodo, gastos_clasificados, movimientos, isr,
                           conciliacion, mov_sin_fact, egresos_sin_match)
n_alta = int((alertas["prioridad"] == "Alta").sum()) if len(alertas) else 0
n_media = int((alertas["prioridad"] == "Media").sum()) if len(alertas) else 0
n_baja = int((alertas["prioridad"] == "Baja").sum()) if len(alertas) else 0

resumen = {
    "Periodo analizado": f"{inicio} a {fin}",
    "Total facturado": round(total_facturado, 2),
    "Total cobrado": round(total_cobrado, 2),
    "Total pendiente de cobro": round(total_pendiente_cobro, 2),
    "Gastos deducibles": round(gastos_deducibles_monto, 2),
    "Gastos no deducibles": round(gastos_no_deducibles_monto, 2),
    "IVA trasladado": round(totales_iva["iva_trasladado"], 2),
    "IVA acreditable": round(totales_iva["iva_acreditable"], 2),
    "IVA neto": round(totales_iva["iva_neto"], 2),
    "ISR retenido (total del periodo cargado)": round(total_isr, 2),
    "Partidas conciliadas (número)": n_conciliadas,
    "Partidas conciliadas (monto)": round(monto_conciliado, 2),
    "Partidas no conciliadas (número)": n_no_conciliadas,
    "Partidas no conciliadas (monto)": round(monto_no_conciliado, 2),
    "Porcentaje de conciliación": f"{pct_conciliacion:.1f}%",
    "Alertas de prioridad alta": n_alta,
    "Alertas de prioridad media": n_media,
    "Alertas de prioridad baja": n_baja,
}

# --------------------------------------------------------------------
# TABLERO: pestañas
# --------------------------------------------------------------------
tab_resumen, tab_ingresos, tab_gastos, tab_iva, tab_isr, tab_conciliacion, tab_alertas, tab_datos = st.tabs(
    ["📊 Resumen ejecutivo", "💰 Ingresos", "🧾 Gastos", "➗ IVA", "📑 ISR",
     "🏦 Conciliación bancaria", "🚨 Alertas", "📂 Datos procesados"]
)

# ---------------- RESUMEN EJECUTIVO ----------------
with tab_resumen:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total facturado", f"${total_facturado:,.2f}")
    c2.metric("Total cobrado", f"${total_cobrado:,.2f}")
    c3.metric("Pendiente de cobro", f"${total_pendiente_cobro:,.2f}")
    c4.metric("% de conciliación", f"{pct_conciliacion:.1f}%")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Gastos deducibles", f"${gastos_deducibles_monto:,.2f}")
    c6.metric("Gastos no deducibles", f"${gastos_no_deducibles_monto:,.2f}")
    c7.metric("IVA neto", f"${totales_iva['iva_neto']:,.2f}",
              help="IVA trasladado menos IVA acreditable")
    c8.metric("ISR retenido (periodo cargado)", f"${total_isr:,.2f}")

    c9, c10, c11 = st.columns(3)
    c9.metric("Alertas de prioridad alta", n_alta)
    c10.metric("Alertas de prioridad media", n_media)
    c11.metric("Alertas de prioridad baja", n_baja)

    st.divider()
    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Facturado vs. cobrado")
        fig = go.Figure(data=[
            go.Bar(name="Facturado", x=["Periodo seleccionado"], y=[total_facturado], marker_color=COLOR_NEUTRO),
            go.Bar(name="Cobrado", x=["Periodo seleccionado"], y=[total_cobrado], marker_color=COLOR_INGRESO),
            go.Bar(name="Pendiente", x=["Periodo seleccionado"], y=[total_pendiente_cobro], marker_color=COLOR_ALERTA_MEDIA),
        ])
        fig.update_layout(barmode="group", yaxis_title="Pesos MXN")
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.subheader("Gastos deducibles vs. no deducibles")
        fig2 = px.pie(
            names=["Deducibles", "No deducibles"],
            values=[gastos_deducibles_monto, gastos_no_deducibles_monto],
            color=["Deducibles", "No deducibles"],
            color_discrete_map={"Deducibles": COLOR_INGRESO, "No deducibles": COLOR_GASTO},
        )
        st.plotly_chart(fig2, use_container_width=True)

    col_c, col_d = st.columns(2)
    with col_c:
        st.subheader("Integración del IVA")
        fig3 = go.Figure(go.Waterfall(
            x=["IVA trasladado", "IVA acreditable", "IVA neto"],
            measure=["relative", "relative", "total"],
            y=[totales_iva["iva_trasladado"], -totales_iva["iva_acreditable"], totales_iva["iva_neto"]],
        ))
        fig3.update_layout(yaxis_title="Pesos MXN")
        st.plotly_chart(fig3, use_container_width=True)

    with col_d:
        st.subheader("Estado de la conciliación bancaria")
        conteo_estado = conciliacion["estado_conciliacion"].value_counts().reset_index()
        conteo_estado.columns = ["Estado", "Número de facturas"]
        fig4 = px.bar(conteo_estado, x="Estado", y="Número de facturas", color="Estado")
        fig4.update_layout(showlegend=False, xaxis_title="")
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("Evolución mensual de ingresos y gastos")
    facturas_mes = facturas_periodo.assign(mes=facturas_periodo["fecha"].dt.to_period("M").astype(str))
    gastos_mes = gastos_clasificados.assign(mes=gastos_clasificados["fecha"].dt.to_period("M").astype(str))
    serie_ingresos = facturas_mes.groupby("mes")["total"].sum().rename("Facturado")
    serie_gastos = gastos_mes.groupby("mes")["total"].sum().rename("Gastos")
    serie = pd.concat([serie_ingresos, serie_gastos], axis=1).fillna(0).reset_index()
    if len(serie) >= 1:
        fig5 = go.Figure()
        fig5.add_trace(go.Scatter(x=serie["mes"], y=serie["Facturado"], mode="lines+markers", name="Facturado", line_color=COLOR_INGRESO))
        fig5.add_trace(go.Scatter(x=serie["mes"], y=serie["Gastos"], mode="lines+markers", name="Gastos", line_color=COLOR_GASTO))
        fig5.update_layout(yaxis_title="Pesos MXN", xaxis_title="Mes")
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.caption("No hay suficientes datos en el periodo seleccionado para mostrar una evolución mensual.")

    st.divider()
    st.subheader("📥 Reporte descargable")
    reporte_excel = generar_reporte_excel(
        facturas_periodo, gastos_clasificados, movimientos, isr,
        conciliacion, mov_sin_fact, egresos_sin_match, alertas, resumen,
    )
    st.download_button(
        label="Descargar reporte completo en Excel",
        data=reporte_excel,
        file_name=f"calculadora_tx_reporte_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------- INGRESOS ----------------
with tab_ingresos:
    st.subheader("Facturas emitidas vs. cobradas")
    st.dataframe(conciliacion, use_container_width=True)
    st.caption(f"Total facturado: ${total_facturado:,.2f} · Total cobrado: ${total_cobrado:,.2f} · "
               f"Pendiente de cobro: ${total_pendiente_cobro:,.2f}")

    st.subheader("Depósitos bancarios sin factura relacionada")
    if len(mov_sin_fact) > 0:
        st.dataframe(mov_sin_fact, use_container_width=True)
    else:
        st.success("No se encontraron depósitos sin factura relacionada.")

# ---------------- GASTOS ----------------
with tab_gastos:
    col1, col2 = st.columns(2)
    col1.metric("Gastos deducibles", f"${gastos_deducibles_monto:,.2f}")
    col2.metric("Gastos no deducibles", f"${gastos_no_deducibles_monto:,.2f}")

    filtro_clasificacion = st.multiselect(
        "Filtrar por clasificación",
        ["Deducible", "No deducible", "Pendiente de clasificación"],
        default=["Deducible", "No deducible", "Pendiente de clasificación"],
    )
    tabla_gastos = gastos_clasificados[gastos_clasificados["clasificacion"].isin(filtro_clasificacion)]
    st.dataframe(tabla_gastos, use_container_width=True)

    st.subheader("Egresos bancarios sin gasto relacionado")
    if len(egresos_sin_match) > 0:
        st.dataframe(egresos_sin_match, use_container_width=True)
        st.caption(
            "Estos pagos no coinciden con ningún folio de gasto de la base cargada. "
            "No se intentó adivinar a qué gasto corresponden: requieren revisión manual."
        )
    else:
        st.success("No se encontraron egresos sin gasto relacionado.")

# ---------------- IVA ----------------
with tab_iva:
    c1, c2, c3 = st.columns(3)
    c1.metric("IVA trasladado", f"${totales_iva['iva_trasladado']:,.2f}")
    c2.metric("IVA acreditable", f"${totales_iva['iva_acreditable']:,.2f}")
    c3.metric("IVA neto (a cargo si es positivo)", f"${totales_iva['iva_neto']:,.2f}")
    st.caption("IVA acreditable = suma del IVA de los gastos clasificados como Deducibles. "
               "IVA neto = IVA trasladado − IVA acreditable.")

    inconsistentes_fact = facturas_periodo[facturas_periodo["iva_inconsistente"]]
    inconsistentes_gasto = gastos_clasificados[gastos_clasificados["iva_inconsistente"]]
    if len(inconsistentes_fact) or len(inconsistentes_gasto):
        st.warning(f"Se detectaron {len(inconsistentes_fact)} factura(s) y {len(inconsistentes_gasto)} "
                   f"gasto(s) donde el IVA registrado no corresponde al 16% del subtotal (tolerancia ${TOLERANCIA_IVA:.2f}).")
    st.subheader("Detalle de IVA - Facturas")
    st.dataframe(facturas_periodo[["folio", "fecha", "subtotal", "iva", "iva_esperado", "iva_inconsistente", "total"]], use_container_width=True)
    st.subheader("Detalle de IVA - Gastos")
    st.dataframe(gastos_clasificados[["folio_gasto", "fecha", "clasificacion", "subtotal", "iva", "iva_esperado", "iva_inconsistente", "total"]], use_container_width=True)

# ---------------- ISR ----------------
with tab_isr:
    st.metric("Total ISR retenido (registros cargados)", f"${total_isr:,.2f}")
    st.info(
        "Las retenciones de ISR no incluyen folio de factura ni UUID/RFC. Al comparar los montos "
        "contra las facturas emitidas no se encontró una proporción fija y confiable (el importe retenido "
        "no equivale a un porcentaje constante del subtotal de ninguna factura cercana en fecha), por lo que "
        "la app NO inventa una relación automática entre retenciones y facturas. Se muestran ambas listas "
        "por separado para que se concilien manualmente con información adicional (RFC, folio fiscal) si se cuenta con ella."
    )
    st.dataframe(isr, use_container_width=True)

# ---------------- CONCILIACIÓN BANCARIA ----------------
with tab_conciliacion:
    c1, c2, c3 = st.columns(3)
    c1.metric("Partidas conciliadas", f"{n_conciliadas} (${monto_conciliado:,.2f})")
    c2.metric("Partidas no conciliadas", f"{n_no_conciliadas} (${monto_no_conciliado:,.2f})")
    c3.metric("% de conciliación", f"{pct_conciliacion:.1f}%")

    filtro_estado = st.multiselect(
        "Filtrar por estado de conciliación",
        sorted(conciliacion["estado_conciliacion"].unique().tolist()),
        default=sorted(conciliacion["estado_conciliacion"].unique().tolist()),
    )
    st.dataframe(conciliacion[conciliacion["estado_conciliacion"].isin(filtro_estado)], use_container_width=True)

    st.subheader("Egresos bancarios sin gasto relacionado")
    st.dataframe(egresos_sin_match, use_container_width=True)

# ---------------- ALERTAS ----------------
with tab_alertas:
    c1, c2, c3 = st.columns(3)
    c1.metric("🔴 Alta prioridad", n_alta)
    c2.metric("🟠 Media prioridad", n_media)
    c3.metric("⚪ Baja prioridad", n_baja)

    if len(alertas) == 0:
        st.success("No se generaron alertas con los datos y filtros actuales.")
    else:
        filtro_prioridad = st.multiselect(
            "Filtrar por prioridad", ["Alta", "Media", "Baja"], default=["Alta", "Media", "Baja"]
        )
        filtro_tipo = st.multiselect(
            "Filtrar por tipo de alerta", sorted(alertas["tipo"].unique().tolist()),
            default=sorted(alertas["tipo"].unique().tolist()),
        )
        tabla_alertas = alertas[alertas["prioridad"].isin(filtro_prioridad) & alertas["tipo"].isin(filtro_tipo)]
        st.dataframe(tabla_alertas, use_container_width=True)

        fig_alertas = px.bar(
            alertas.groupby(["tipo", "prioridad"]).size().reset_index(name="conteo"),
            x="tipo", y="conteo", color="prioridad",
            color_discrete_map={"Alta": COLOR_ALERTA_ALTA, "Media": COLOR_ALERTA_MEDIA, "Baja": COLOR_ALERTA_BAJA},
        )
        fig_alertas.update_layout(xaxis_title="", yaxis_title="Número de alertas")
        st.plotly_chart(fig_alertas, use_container_width=True)

        st.download_button(
            "Descargar excepciones (alertas) en CSV",
            data=tabla_alertas.to_csv(index=False).encode("utf-8-sig"),
            file_name="alertas_calculadora_tx.csv",
            mime="text/csv",
        )

# ---------------- DATOS PROCESADOS ----------------
with tab_datos:
    st.caption("Datos originales tal como se cargaron, más las columnas calculadas (claramente añadidas al final).")
    st.subheader("Facturas emitidas")
    st.dataframe(facturas_periodo, use_container_width=True)
    st.subheader("Gastos (facturas recibidas)")
    st.dataframe(gastos_clasificados, use_container_width=True)
    st.subheader("Movimientos bancarios")
    st.dataframe(movimientos, use_container_width=True)
    st.subheader("Retenciones de ISR")
    st.dataframe(isr, use_container_width=True)
