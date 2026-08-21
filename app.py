"""
App Streamlit - Analisis de Tendencias
Unidades: Destilacion y Polimerizacion (LIPP La Plata)
Fuente de datos: Tendencias.xlsx (hojas DESTILACION y POLIMERIZACION)

Notas:
- Rel_TEA_CDonor se espera ya calculada en Excel.
- La carga de datos valida y mapea columnas por TAG de IP21, no solo por posicion.
- La app detecta cambios en Tendencias.xlsx por fecha de modificacion del archivo.
- Incluye un panel simple para crear variables calculadas: Variable A + operacion + Variable B.
"""

import os
import re
from io import BytesIO
from itertools import combinations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

try:
    from sklearn.ensemble import GradientBoostingRegressor
    from sklearn.impute import SimpleImputer
    from sklearn.pipeline import make_pipeline
    SKLEARN_DISPONIBLE = True
except Exception:
    GradientBoostingRegressor = None
    SimpleImputer = None
    make_pipeline = None
    SKLEARN_DISPONIBLE = False


# ==============================================================================
# CONFIGURACION GENERAL
# ==============================================================================

st.set_page_config(page_title="Analisis de Tendencias", layout="wide")
st.title("Analisis de Tendencias")


def aplicar_layout_leyenda_inferior(fig, titulo=None, altura=650):
    """Aplica layout común para evitar que la leyenda lateral achique la gráfica."""
    layout_kwargs = dict(
        height=altura,
        margin=dict(l=40, r=40, t=70, b=130),
        legend=dict(
            orientation="h",
            y=-0.22,
            x=0.5,
            xanchor="center",
            yanchor="top",
        ),
        hovermode="x unified",
    )
    if titulo is not None:
        layout_kwargs["title"] = titulo
    fig.update_layout(**layout_kwargs)
    return fig

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO = os.path.join(CARPETA, "Tendencias.xlsx")


# ==============================================================================
# CONFIGURACION DE UNIDADES
# Cada unidad tiene:
#   - hoja: nombre exacto de la hoja en el Excel
#   - variables: dict {nombre_python: "Descripcion [TAG]"}
#     El orden es importante: debe coincidir con el orden de columnas del Excel
#   - default_vars: variables preseleccionadas al abrir la app
# ==============================================================================

CONFIG_UNIDADES = {
    "Destilación": {
        "hoja": "DESTILACIÓN",
        "variables": {
            "Caudal_P1006": "Caudal de bombas P-1006 [C-1003.10FIC020]",
            "Caudal_fondo_C1003": "Caudal fondo de C-1003 [P-1002A/B.10FRQ008]",
            "Ingreso_C1001_ctrl": "Ingreso a C-1001 controlador [C-1001.10FIC350]",
            "Presion_C1001": "Presion C-1001 [C-1001.10PIC004]",
            "Caudal_reflujo_C1001": "Caudal reflujo C-1001 [P-1001.10FYC003]",
            "Ingreso_splitter": "Ingreso a splitter [C-1003.10FIC006]",
            "Extraccion_propileno": "Extraccion de propileno [C-1002.10FQC011]",
            "Rel_extraccion_reflujo": "Relacion extraccion/reflujo [C-1002.10REL002]",
            # Las columnas DP_C1002 y DP_C1003 estan sin datos en IP21, por eso se excluyen.
            "Propileno_YPF": "Propileno de YPF [ANALIZ YPF.50QR003]",
            "Propileno_a_splitter": "Propileno a splitter [C-1001.10QR012]",
            "Propano_a_splitter": "Propano a splitter [C-1001.10QR013]",
            "Pesados_a_splitter": "Pesados a splitter [C-1001.10QR014]",
            "Propano_tope_splitter": "Propano en tope splitter [C-1002.10QR041]",
            "Etano_tope_splitter": "Etano en tope splitter [C-1002.10QR040]",
            "Conc_propileno_tope": "Concentracion propileno tope splitter [C-1002.10QR043]",
            "Conc_propileno_fondo": "Concentracion propileno fondo splitter [C-1003.10QR024]",
            "Conc_pesados_fondo": "Concentracion pesados fondo splitter [C-1003.10QR026]",
            "Conc_propano_fondo": "Concentracion propano fondo splitter [C-1003.10QR025]",
            "Temp_tope_C1003": "Temperatura tope C-1003 [C-1003.10TI020]",
            "Temp_fondo_C1003": "Temperatura fondo C-1003 [C-1003.10TI007]",
            "Temp_tope_C1002": "Temperatura tope C-1002 [C-1002.10TI006]",
            "P_succion_K1001": "Presion succion K-1001 [V-1003.10PIC021]",
            "Off_gas_V1001": "Off gas V-1001 [V-1001.10FIC004]",
            "Temp_tope_C1001": "Temp. tope C-1001 [C-1001.10TI003]",
            "Conc_livianos_V1001": "Conc. livianos V-1001 [C-1001.10QC019]",
            "Caudal_vapor_E1001": "Caudal vapor a E-1001 [E-1001.10FIC002]",
            "Nivel_cond_E1001": "Nivel condensado E-1001 [E-1001.10LI002]",
            "Temp_fondo_C1001": "Temp. fondo C-1001 [C-1001.10TI004]",
            "Nivel_fondo_C1001": "Nivel fondo C-1001 [C-1001.10LIC001]",
        },
        "default_vars": ["Extraccion_propileno", "Temp_tope_C1003", "Temp_fondo_C1003"],
    },
    "Polimerización": {
        "hoja": "POLIMERIZACIÓN",
        "variables": {
            "Nivel_R2301": "Nivel reactor R-2301 [R-2301.23LRC004]",
            "Caudal_alim_R2301": "Caudal alimentacion R-2301 [R-2301.23FRC001]",
            "Rendimiento": "Rendimiento [R-2301.23REND01]",
            "Calor_reaccion_URA": "Calor de reaccion URA [E-2301A/B.23URA001]",
            "Gases_tope_R2301": "Gases de tope R-2301 [R-2301.23FR009]",
            "Presion_R2301": "Presion R-2301 [R-2301.23PCZ005]",
            "Caudal_E2301A": "Caudal E-2301 A [E-2301A.23FR007]",
            "Caudal_E2301B": "Caudal E-2301 B [E-2301B.23FR008]",
            "Caudal_E2301C": "Caudal E-2301 C [E-2301C.23FI029]",
            "Caudal_CW": "Caudal de CW [E-2301A/B.23FRA004]",
            "Sumatoria_condensadores": "Sumatoria condensadores [Q E-2301]",
            "MFI_polvo": "MFI Polvo [ENS.MFI_POLVO]",
            "MFI_pellets": "MFI Pellets [ENS.MFI_PELLETS]",
            "XS": "Porcentaje XS [ENS.PORCENTAJE_XS]",
            "Produccion_PP": "Produccion de polvo PP [E-2401.23PROD01]",
            "Caudal_purga": "Caudal de purga [P-2402A/B.24FRC049]",
            "Slurry": "Slurry [R-2301.23LIS005]",
            "Catalizador_A": "Caudal catalizador A [P-2209A.22FQC031]",
            "Catalizador_B": "Caudal catalizador B [P-2209B.22FQC054]",
            "TEA": "Caudal de TEA [P-2205A/B.22FIC131]",
            "C_Donor": "Caudal de C-Donor [P-2213A/B.22FIC051]",
            "Alcohol": "Caudal de alcohol [P-2204A/B.22FIT050A]",
            "H2_a_K2301": "Caudal H2 a K-2301 [K-2301.23FRC021]",
            "H2_a_R2301": "Caudal H2 a R-2301 [R-2301.23FRC005]",
            "Conc_propano": "Concentracion propano [R-2301.23QC007A]",
            "Conc_H2": "Concentracion H2 [R-2301.23QC006A]",
            "Presion_cat_2209A": "Presion Cat. P-2209A [P-2209A.22PIZ071]",
            "Presion_cat_2209B": "Presion Cat. P-2209B [P-2209B.22PIZ061]",
            "Temperatura_R2301": "Temperatura R-2301 [E-2301A/B.23TRC002]",
            "Caudal_fresco": "Caudal fresco [V-2102A/B.21FRQ001]",
            "Salida_fondo_R2301": "Salida fondo R-2301 [R-2301.24FRC006]",
            "Gas_S2402_S2404": "Gas de S-2402 a S-2404 [E-2404.24FRC042]",
            # Esta columna se espera ya calculada en Excel.
            "Rel_TEA_CDonor": "Relacion TEA/C-Donor (calculado en Excel)",
        },
        "default_vars": ["Rendimiento", "Presion_R2301", "MFI_polvo"],
    },
}


# ==============================================================================
# VARIABLES QUE SON CAUDALES
# Se pueden filtrar automaticamente para descartar valores negativos.
# ==============================================================================

CAUDALES = {
    "Caudal_P1006", "Caudal_fondo_C1003", "Ingreso_C1001_ctrl", "Caudal_reflujo_C1001",
    "Ingreso_splitter", "Extraccion_propileno", "Propileno_YPF", "Propileno_a_splitter",
    "Propano_a_splitter", "Pesados_a_splitter", "Off_gas_V1001", "Caudal_vapor_E1001",
    "Caudal_alim_R2301", "Gases_tope_R2301", "Caudal_E2301A", "Caudal_E2301B",
    "Caudal_E2301C", "Caudal_CW", "Produccion_PP", "Caudal_purga", "Catalizador_A",
    "Catalizador_B", "TEA", "C_Donor", "Alcohol", "H2_a_K2301", "H2_a_R2301",
    "Caudal_fresco", "Salida_fondo_R2301", "Gas_S2402_S2404",
    "Caudal_ZN306_activo", "Caudal_ZN389_activo", "Caudal_catalizador_activo",
}


# ==============================================================================
# FUNCIONES AUXILIARES
# ==============================================================================

def nombre_largo(var: str, nombres_legibles: dict) -> str:
    """Devuelve el nombre legible de una variable."""
    return nombres_legibles.get(var, var)


def calcular_variable_personalizada(df: pd.DataFrame, definicion: dict) -> pd.Series:
    """Calcula una variable A operacion B segun una definicion guardada."""
    var_a = definicion["var_a"]
    var_b = definicion["var_b"]
    operacion = definicion["operacion"]

    a = pd.to_numeric(df[var_a], errors="coerce")
    b = pd.to_numeric(df[var_b], errors="coerce")

    with np.errstate(divide="ignore", invalid="ignore"):
        if operacion == "+":
            resultado = a + b
        elif operacion == "-":
            resultado = a - b
        elif operacion == "×":
            resultado = a * b
        elif operacion == "÷":
            resultado = a / b
        elif operacion == "A/B × 100":
            resultado = (a / b) * 100
        elif operacion == "|A - B|":
            resultado = (a - b).abs()
        elif operacion == "Promedio":
            resultado = (a + b) / 2
        else:
            resultado = pd.Series(np.nan, index=df.index)

    resultado = pd.to_numeric(resultado, errors="coerce")
    resultado = resultado.replace([np.inf, -np.inf], np.nan)
    return resultado


def texto_formula(definicion: dict) -> str:
    """Arma una vista previa legible de la variable calculada."""
    a = definicion["var_a"]
    b = definicion["var_b"]
    op = definicion["operacion"]

    if op == "A/B × 100":
        return f"({a} / {b}) × 100"
    if op == "|A - B|":
        return f"|{a} - {b}|"
    if op == "Promedio":
        return f"({a} + {b}) / 2"
    return f"{a} {op} {b}"


def aplicar_variables_calculadas_guardadas(df: pd.DataFrame, nombres_legibles: dict, todas_variables: list, unidad: str) -> None:
    """Aplica todas las variables calculadas guardadas en la sesion."""
    key_defs = f"variables_calculadas_{unidad}"
    if key_defs not in st.session_state:
        st.session_state[key_defs] = []

    for definicion in st.session_state[key_defs]:
        clave = definicion["clave"]
        try:
            df[clave] = calcular_variable_personalizada(df, definicion)
            nombres_legibles[clave] = definicion["nombre"] + f" [{clave}]"
            if clave not in todas_variables:
                todas_variables.append(clave)
        except Exception:
            # Si cambia el Excel o se borra una variable base, se ignora esa calculada.
            pass


def aplicar_agrupacion(df: pd.DataFrame, agrupacion: str) -> pd.DataFrame:
    """
    Aplica agrupacion temporal por promedio para variables numericas.

    Si existe la columna Producto, conserva para cada periodo el producto dominante
    (moda). Esto permite filtrar campañas y calcular targets por producto aun con
    agrupacion diaria/semanal/mensual.
    """
    if agrupacion == "Sin agrupación":
        return df.copy()

    freq_map = {"Hora": "h", "Día": "D", "Semana": "W", "Mes": "ME"}
    freq = freq_map[agrupacion]

    df_idx = df.set_index("Fecha_y_hora")

    df_num = (
        df_idx
        .resample(freq)
        .mean(numeric_only=True)
        .asfreq(freq)
    )

    if "Producto" in df_idx.columns:
        def _producto_dominante(serie):
            s = serie.dropna().astype(str).str.strip()
            s = s[s != ""]
            if len(s) == 0:
                return np.nan
            return s.value_counts().idxmax()

        producto_agrup = df_idx["Producto"].resample(freq).agg(_producto_dominante)
        df_num["Producto"] = producto_agrup

    return df_num.reset_index()


def top_correlaciones(df: pd.DataFrame, variables: list, nombres_legibles: dict, metodo: str, min_puntos: int) -> pd.DataFrame:
    """Calcula ranking de correlaciones entre pares de variables."""
    filas = []
    for var1, var2 in combinations(variables, 2):
        tmp = df[[var1, var2]].dropna()
        n = len(tmp)
        if n < min_puntos:
            continue
        r = tmp[var1].corr(tmp[var2], method=metodo)
        if pd.isna(r):
            continue
        filas.append({
            "Variable 1": nombre_largo(var1, nombres_legibles),
            "Variable 2": nombre_largo(var2, nombres_legibles),
            "r": round(float(r), 3),
            "abs(r)": round(abs(float(r)), 3),
            "Puntos validos": int(n),
        })

    if not filas:
        return pd.DataFrame(columns=["Variable 1", "Variable 2", "r", "abs(r)", "Puntos validos"])

    return pd.DataFrame(filas).sort_values("abs(r)", ascending=False).reset_index(drop=True)


def calcular_lags(df: pd.DataFrame, var_x: str, var_y: str, metodo: str, max_lag: int, min_puntos: int) -> pd.DataFrame:
    """
    Calcula correlacion entre X(t) e Y(t+lag).
    lag > 0: X precede a Y.
    lag < 0: Y precede a X.
    """
    filas = []
    base = df[["Fecha_y_hora", var_x, var_y]].copy()

    for lag in range(-max_lag, max_lag + 1):
        y_desfasada = base[var_y].shift(-lag)
        tmp = pd.concat([base[var_x], y_desfasada], axis=1)
        tmp.columns = [var_x, var_y]
        tmp = tmp.dropna()
        n = len(tmp)
        if n < min_puntos:
            r = np.nan
        else:
            r = tmp[var_x].corr(tmp[var_y], method=metodo)

        filas.append({
            "Lag": lag,
            "Correlacion": r,
            "abs_r": abs(r) if pd.notna(r) else np.nan,
            "Puntos validos": n,
        })

    return pd.DataFrame(filas)


# ==============================================================================
# CARGA DE DATOS
# Lee el Excel, descarta la primera columna, asigna nombres python.
# IMPORTANTE: el orden de columnas del Excel debe coincidir con el orden
# del dict 'variables' en CONFIG_UNIDADES.
# ==============================================================================


def extraer_tag_desde_label(label):
    """
    Extrae el TAG ubicado entre corchetes en el nombre legible.
    Ejemplo:
    "Presion succion K-1001 [V-1003.10PIC021]" -> "V-1003.10PIC021"
    """
    if label is None:
        return None

    texto = str(label)
    match = re.search(r"\[([^\]]+)\]", texto)

    if not match:
        return None

    return match.group(1).strip()


def normalizar_tag(tag):
    """
    Normaliza TAGs para poder comparar el TAG de la app contra el TAG del Excel.
    El Excel suele traer sufijo .PV y la app no.
    """
    if tag is None or pd.isna(tag):
        return ""

    texto = str(tag).strip().upper()
    texto = re.sub(r"\s+", "", texto)

    if texto.endswith(".PV"):
        texto = texto[:-3]

    return texto


def buscar_columna_por_tag(tags_excel, tag_esperado):
    """
    Busca en la fila de TAGs del Excel la columna correspondiente al TAG esperado.
    Compara ignorando espacios, mayusculas/minusculas y sufijo .PV.
    """
    if not tag_esperado:
        return None

    esperado = normalizar_tag(tag_esperado)

    for idx, tag_excel in enumerate(tags_excel):
        actual = normalizar_tag(tag_excel)

        if actual == esperado:
            return idx

    return None


def normalizar_texto(texto):
    """
    Normaliza texto descriptivo para busquedas simples por descripcion.
    """
    if texto is None or pd.isna(texto):
        return ""

    texto = str(texto).strip().lower()
    texto = texto.replace("á", "a").replace("é", "e").replace("í", "i")
    texto = texto.replace("ó", "o").replace("ú", "u").replace("ñ", "n")
    texto = re.sub(r"\s+", " ", texto)

    return texto




def normalizar_producto_operativo(producto):
    """
    Normaliza el producto para análisis de campañas.

    Reglas:
    - Todos los KFM se agrupan como KFM6110.
    - Los sufijos TE de transición no se discriminan.
      Ejemplos:
        RFD6140K TE063  -> RFD6140K
        KED6270 TE-079  -> KED6270
        KYD 6110 TE-088 -> KYD6110
    - También se corrigen espacios internos del grado:
        HYS 6200 -> HYS6200
    - Si hay texto adicional entre paréntesis o sufijos de ensayo, se toma
      el primer código de grado detectado.
    """
    if producto is None or pd.isna(producto):
        return pd.NA

    p = str(producto).strip()
    if p == "" or p.lower() in ["nan", "none", "<na>", "0", "0.0"]:
        return pd.NA

    p_upper = p.upper().strip()

    # Regla específica solicitada: KFM se muestra como KFM6110.
    if "KFM" in p_upper:
        return "KFM6110"

    # Buscar el primer código de grado: letras + 4 dígitos + sufijo opcional.
    # Funciona para RFD6140K, KED6270, HYS 6200, KYD 6110, XMD6279T, etc.
    m = re.search(r"([A-Z]{2,4})\s*[- ]?\s*(\d{4})([A-Z]?)", p_upper)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"

    # Fallback: eliminar sufijos TE si no se detectó código estructurado.
    p_upper = re.sub(r"\s*[-]?\s*TE\s*[-]?\s*\d+.*$", "", p_upper).strip()
    p_upper = re.sub(r"\s+", "", p_upper)

    return p_upper



def preparar_serie_producto_filtrado_para_grafico(df_plot, y_col, agrupacion):
    """
    Para gráficos con filtro por producto/grado activo:
    - Une puntos consecutivos.
    - Corta la línea cuando hay saltos grandes de tiempo entre campañas.
    - Mantiene todos los puntos visibles.

    Regla de continuidad:
    - hora: gap máximo 1.5 h
    - dia: gap máximo 1.5 días
    - semana: gap máximo 10 días
    - mes: gap máximo 45 días
    """
    df_linea = df_plot[["Fecha_y_hora", y_col]].copy()
    df_linea = df_linea.sort_values("Fecha_y_hora").reset_index(drop=True)

    if agrupacion == "hora":
        max_gap = pd.Timedelta(hours=1.5)
    elif agrupacion == "dia":
        max_gap = pd.Timedelta(days=1.5)
    elif agrupacion == "semana":
        max_gap = pd.Timedelta(days=10)
    elif agrupacion == "mes":
        max_gap = pd.Timedelta(days=45)
    else:
        max_gap = pd.Timedelta(days=1.5)

    fechas = pd.to_datetime(df_linea["Fecha_y_hora"], errors="coerce")
    gaps = fechas.diff() > max_gap

    # Insertar NaN en y después de un salto para que Plotly corte la línea.
    y = pd.to_numeric(df_linea[y_col], errors="coerce").copy()
    y.loc[gaps] = np.nan

    return df_linea["Fecha_y_hora"], y


def obtener_cambios_producto_para_marcas(df_plot: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve los cambios de producto/campaña para marcar en la gráfica.

    Importante:
    - Usa la secuencia temporal de Producto.
    - El primer punto visible se marca como inicio de campaña visible.
    - Si hay valores vacíos de Producto, se ignoran.
    """
    if "Producto" not in df_plot.columns or "Fecha_y_hora" not in df_plot.columns:
        return pd.DataFrame(columns=["Fecha_y_hora", "Producto"])

    datos_producto = df_plot[["Fecha_y_hora", "Producto"]].copy()
    datos_producto["Fecha_y_hora"] = pd.to_datetime(
        datos_producto["Fecha_y_hora"],
        errors="coerce",
    )
    datos_producto["Producto"] = datos_producto["Producto"].map(
        normalizar_producto_operativo
    )

    datos_producto = datos_producto.dropna(subset=["Fecha_y_hora", "Producto"])

    if len(datos_producto) == 0:
        return pd.DataFrame(columns=["Fecha_y_hora", "Producto"])

    datos_producto = datos_producto.sort_values("Fecha_y_hora").reset_index(drop=True)

    cambios = datos_producto.loc[
        datos_producto["Producto"].astype(str).ne(
            datos_producto["Producto"].astype(str).shift()
        )
    ].copy()

    return cambios[["Fecha_y_hora", "Producto"]]


def agregar_marcas_producto_a_figura(
    fig,
    df_plot,
    en_subplots=False,
    mostrar_etiquetas=True,
):
    """
    Agrega líneas verticales en todos los cambios de producto/campaña.

    Estilo tipo versión anterior:
    - línea vertical punteada;
    - nombre del grado horizontal arriba de la gráfica;
    - sin recuadros ni etiquetas verticales.
    """
    cambios = obtener_cambios_producto_para_marcas(df_plot)

    if len(cambios) == 0:
        return fig

    usar_texto = bool(mostrar_etiquetas)

    # Margen superior moderado para que entren los nombres de grado.
    try:
        margen_actual = 0
        if getattr(fig.layout, "margin", None) is not None:
            margen_actual = getattr(fig.layout.margin, "t", 0) or 0
        if usar_texto:
            fig.update_layout(margin=dict(t=max(margen_actual, 85)))
    except Exception:
        pass

    for _, fila in cambios.iterrows():
        try:
            x_val = fila["Fecha_y_hora"]
            producto = str(fila["Producto"])

            fig.add_shape(
                type="line",
                x0=x_val,
                x1=x_val,
                y0=0,
                y1=1,
                xref="x",
                yref="paper",
                line=dict(
                    width=1,
                    dash="dot",
                    color="rgba(70,70,70,0.55)",
                ),
                opacity=0.65,
                layer="below",
            )

            if usar_texto:
                fig.add_annotation(
                    x=x_val,
                    y=1.025,
                    xref="x",
                    yref="paper",
                    text=producto,
                    showarrow=False,
                    textangle=0,
                    font=dict(
                        size=9,
                        color="rgba(70,70,70,0.95)",
                    ),
                    align="center",
                    xanchor="center",
                    yanchor="bottom",
                )

        except Exception:
            # Las marcas son auxiliares y nunca deben romper la app.
            pass

    return fig


def buscar_columna_por_descripcion(descripciones_excel, descripcion_buscada):
    """
    Busca una columna por descripcion. Se usa como respaldo para variables calculadas
    en Excel que no tienen TAG IP21.
    """
    buscada = normalizar_texto(descripcion_buscada)

    if not buscada:
        return None

    for idx, descripcion_excel in enumerate(descripciones_excel):
        actual = normalizar_texto(descripcion_excel)

        if buscada in actual or actual in buscada:
            return idx

    return None

def obtener_mtime_archivo() -> float | None:
    """
    Devuelve la fecha de modificacion del Excel local.
    En Streamlit Cloud normalmente se usa carga manual de archivo.
    """
    if not os.path.exists(ARCHIVO):
        return None
    return os.path.getmtime(ARCHIVO)


@st.cache_data(show_spinner="Cargando Excel...")
def cargar_datos(
    hoja: str,
    nombres_cols: list,
    nombres_legibles: dict,
    excel_mtime: float | None,
    excel_bytes: bytes | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    # excel_mtime no se usa directamente, pero fuerza a Streamlit a invalidar la cache
    # cuando cambia la fecha de modificacion de Tendencias.xlsx en modo local.
    _ = excel_mtime

    fuente_excel = None
    nombre_fuente = "archivo cargado manualmente"

    if excel_bytes is not None:
        fuente_excel = BytesIO(excel_bytes)
    else:
        if not os.path.exists(ARCHIVO):
            st.error(f"No se encontro el archivo local: {ARCHIVO}")
            st.info("En Streamlit Cloud, usa la opción 'Cargar Excel manualmente' desde la barra lateral.")
            st.stop()

        fuente_excel = ARCHIVO
        nombre_fuente = "Tendencias.xlsx"

    try:
        df_raw = pd.read_excel(fuente_excel, sheet_name=hoja, header=None)
    except ValueError:
        st.error(f"No existe la hoja '{hoja}' en el {nombre_fuente}.")
        st.stop()
    except PermissionError:
        st.error("No se puede leer Tendencias.xlsx. Cerralo en Excel y volve a correr la app.")
        st.stop()
    except Exception as exc:
        st.error(f"Error leyendo el Excel: {exc}")
        st.stop()

    if df_raw.shape[0] < 6 or df_raw.shape[1] < 3:
        st.error("La hoja no tiene la estructura esperada de IP21.")
        st.stop()

    # Estructura IP21 usada:
    # Fila 3 = descripcion de variables -> indice 2
    # Fila 4 = TAGs -> indice 3
    # Fila 6+ = datos -> indice 5 en adelante
    descripciones_excel = list(df_raw.iloc[2, :])
    tags_excel = list(df_raw.iloc[3, :])

    # La fecha suele estar en columna B. Si en el futuro se mueve, se intenta buscar por texto.
    col_fecha = 1
    for idx, tag in enumerate(tags_excel):
        if tag is not None and str(tag).strip().lower() in ["fecha y hora", "fecha_y_hora", "fecha"]:
            col_fecha = idx
            break

    df_datos = df_raw.iloc[5:, :].copy()

    df_out = pd.DataFrame()
    df_out["Fecha_y_hora"] = pd.to_datetime(df_datos.iloc[:, col_fecha], errors="coerce")

    filas_mapeo = []

    for var in nombres_cols:
        label = nombres_legibles[var]
        tag_esperado = extraer_tag_desde_label(label)
        col_excel = buscar_columna_por_tag(tags_excel, tag_esperado)

        # Caso especial: relacion TEA/C-Donor viene calculada en Excel y no tiene TAG IP21.
        if col_excel is None and var == "Rel_TEA_CDonor":
            col_excel = buscar_columna_por_descripcion(descripciones_excel, "Relación TEA/C-Donor")

        if col_excel is None:
            df_out[var] = np.nan
            filas_mapeo.append({
                "Variable app": var,
                "Nombre": label,
                "TAG esperado": tag_esperado if tag_esperado else "Sin TAG",
                "Columna Excel": "NO ENCONTRADA",
                "TAG Excel": "NO ENCONTRADO",
                "Descripcion Excel": "NO ENCONTRADA",
                "Estado": "ERROR",
            })
            continue

        df_out[var] = pd.to_numeric(df_datos.iloc[:, col_excel], errors="coerce")
        filas_mapeo.append({
            "Variable app": var,
            "Nombre": label,
            "TAG esperado": tag_esperado if tag_esperado else "Sin TAG",
            "Columna Excel": col_excel + 1,
            "TAG Excel": tags_excel[col_excel],
            "Descripcion Excel": descripciones_excel[col_excel],
            "Estado": "OK",
        })

    # Columna opcional de producto/campaña.
    # Se busca PRODUCTO de forma robusta porque no tiene TAG IP21.
    # Importante: en algunos Excel la columna tiene muchas filas iniciales vacías,
    # por eso no se valida solamente con las primeras filas de datos.
    if normalizar_texto(hoja).startswith("polimer"):
        col_producto = None

        # 1) Buscar en descripciones, tags y primeras filas.
        for idx in range(df_raw.shape[1]):
            valores_header = []
            for rr in range(min(6, df_raw.shape[0])):
                valores_header.append(normalizar_texto(df_raw.iloc[rr, idx]))

            desc_norm = normalizar_texto(descripciones_excel[idx]) if idx < len(descripciones_excel) else ""
            tag_norm = normalizar_texto(tags_excel[idx]) if idx < len(tags_excel) else ""

            if "producto" in valores_header or desc_norm == "producto" or tag_norm == "producto":
                col_producto = idx
                break

        # 2) Fallback: si no lo encontró por header, tomar una columna llamada exactamente
        # PRODUCTO si pandas/openpyxl la detectó desplazada en las primeras filas.
        if col_producto is None:
            for idx in range(df_raw.shape[1]):
                serie_cabecera = df_raw.iloc[:10, idx].astype("string").str.strip().str.upper()
                if (serie_cabecera == "PRODUCTO").any():
                    col_producto = idx
                    break

        if col_producto is not None:
            producto = df_datos.iloc[:, col_producto].astype("string").str.strip()

            producto = producto.replace({
                "": pd.NA,
                "nan": pd.NA,
                "None": pd.NA,
                "NONE": pd.NA,
                "NaN": pd.NA,
                "<NA>": pd.NA,
                "0": pd.NA,
                "0.0": pd.NA,
            })

            # Asignar por posición para evitar problemas de índice.
            df_out["Producto"] = producto.map(normalizar_producto_operativo).to_numpy()

            filas_mapeo.append({
                "Variable app": "Producto",
                "Nombre": "Producto / campaña",
                "TAG esperado": "Sin TAG",
                "Columna Excel": col_producto + 1,
                "TAG Excel": tags_excel[col_producto] if col_producto < len(tags_excel) else "",
                "Descripcion Excel": descripciones_excel[col_producto] if col_producto < len(descripciones_excel) else "PRODUCTO",
                "Estado": "OK",
            })
        else:
            df_out["Producto"] = pd.NA
            filas_mapeo.append({
                "Variable app": "Producto",
                "Nombre": "Producto / campaña",
                "TAG esperado": "Sin TAG",
                "Columna Excel": "NO ENCONTRADA",
                "TAG Excel": "NO ENCONTRADO",
                "Descripcion Excel": "NO ENCONTRADA",
                "Estado": "ADVERTENCIA",
            })


    df_out = df_out.dropna(subset=["Fecha_y_hora"])
    df_out = df_out.sort_values("Fecha_y_hora").reset_index(drop=True)

    mapeo = pd.DataFrame(filas_mapeo)

    return df_out, mapeo

# ==============================================================================
# SIDEBAR - SELECTORES
# ==============================================================================

# Contenedores visuales del sidebar.
# El contenido se llena en distintos momentos del script, pero queda ordenado así:
# 1) Controles principales de uso diario
# 2) Modelo de rendimiento estimado
# 3) Datos / validación del Excel
# 4) Herramientas avanzadas
sidebar_principal = st.sidebar.container()
sidebar_modelo = st.sidebar.container()
sidebar_datos = st.sidebar.container()
sidebar_herramientas = st.sidebar.container()

with sidebar_principal:
    st.header("Unidad")
    unidad = st.radio("Seleccionar unidad:", options=list(CONFIG_UNIDADES.keys()))

config = CONFIG_UNIDADES[unidad]
nombres_legibles = dict(config["variables"])
todas_variables = list(nombres_legibles.keys())
default_vars = list(config["default_vars"])

with sidebar_datos:
    st.markdown("---")
    st.subheader("Datos y validación")

    fuente_datos = st.radio(
        "Seleccionar fuente:",
        options=["Cargar Excel manualmente", "Usar Excel de la carpeta"],
        index=0,
    )

    excel_bytes = None
    excel_mtime = None

    if fuente_datos == "Cargar Excel manualmente":
        archivo_cargado = st.file_uploader(
            "Cargar archivo Excel",
            type=["xlsx", "xlsm", "xls"],
        )

        if archivo_cargado is None:
            st.info("Cargá un archivo Excel desde la barra lateral para comenzar el análisis.")
            st.stop()

        excel_bytes = archivo_cargado.getvalue()
        st.success(f"Archivo cargado: {archivo_cargado.name}")
    else:
        excel_mtime = obtener_mtime_archivo()

    if st.button("Actualizar datos", width="stretch"):
        st.cache_data.clear()
        st.rerun()

    df, mapeo_columnas = cargar_datos(
        config["hoja"],
        todas_variables,
        nombres_legibles,
        excel_mtime,
        excel_bytes,
    )

    if df.empty:
        st.warning("La hoja seleccionada no tiene datos validos de fecha/hora.")
        st.stop()

    st.caption(f"Hoja: {config['hoja']} | Registros crudos: {len(df):,}")

    if not mapeo_columnas.empty:
        errores_mapeo = mapeo_columnas[mapeo_columnas["Estado"] != "OK"]

        if len(errores_mapeo) > 0:
            st.error(f"Hay {len(errores_mapeo)} variables sin mapear. Revisar Excel/app.")
        else:
            st.success("Columnas validadas por TAG IP21")

        with st.expander("Validacion columnas Excel", expanded=False):
            st.dataframe(mapeo_columnas, width="stretch", hide_index=True)


# ==============================================================================
# VARIABLES CALCULADAS - FORMATO SIMPLE TIPO CARD
# ==============================================================================

aplicar_variables_calculadas_guardadas(df, nombres_legibles, todas_variables, unidad)


# ==============================================================================
# TARGET OPTIMO POR FAMILIA DE PRODUCTO - POLIMERIZACION
# ==============================================================================
# Objetivo:
# - Comparar el rendimiento real contra el mejor historico comparable.
# - No se usan nombres de producto.
# - La familia se define con variables de proceso/calidad:
#   catalizador activo, MFI, H2 y XS.
#
# Base candidata:
# - 28/04/2024 a 06/10/2024
# - 01/05/2025 a 31/12/2025
# - Exclusion: 01/11/2024 a 30/04/2025 por evento de baja productividad/bajo nivel.
#
# Presion normal para construir la base:
# - 30.0 <= Presion_R2301 <= 31.0 bar
#
# Target:
# - Dentro de cada familia, se usa el TOP 25% de rendimiento.
# - Si la familia tiene pocos puntos, se relaja automaticamente:
#   misma familia -> mismo catalizador -> base completa.
#
# Interpretacion:
# - Productividad_estimada = target esperado para la familia/condicion.
# - Desvio = Rendimiento real - Target.
# - Desvio negativo = estamos por debajo del target.

if unidad.startswith("Polimer") and all(
    col in df.columns
    for col in [
        "Rendimiento",
        "MFI_polvo",
        "Conc_H2",
        "Conc_propano",
        "Catalizador_A",
        "Catalizador_B",
        "Presion_cat_2209A",
        "Presion_cat_2209B",
        "Presion_R2301",
        "XS",
        "Caudal_fresco",
        "Caudal_alim_R2301",
        "TEA",
        "C_Donor",
        "Rel_TEA_CDonor",
    ]
):
    rendimiento = pd.to_numeric(df["Rendimiento"], errors="coerce")
    mfi = pd.to_numeric(df["MFI_polvo"], errors="coerce")
    h2 = pd.to_numeric(df["Conc_H2"], errors="coerce")
    propano = pd.to_numeric(df["Conc_propano"], errors="coerce")
    xs = pd.to_numeric(df["XS"], errors="coerce")
    presion_r = pd.to_numeric(df["Presion_R2301"], errors="coerce")

    cat_a = pd.to_numeric(df["Catalizador_A"], errors="coerce")
    cat_b = pd.to_numeric(df["Catalizador_B"], errors="coerce")
    p_cat_a = pd.to_numeric(df["Presion_cat_2209A"], errors="coerce")
    p_cat_b = pd.to_numeric(df["Presion_cat_2209B"], errors="coerce")

    # Identificacion de catalizador activo por presion de descarga.
    zn306_activo = p_cat_b >= 30
    zn389_activo = p_cat_a >= 30

    # Asignacion actual:
    # P-2209B = ZN-306
    # P-2209A = ZN-389
    df["Caudal_ZN306_activo"] = np.where(zn306_activo, cat_b, 0.0)
    df["Caudal_ZN389_activo"] = np.where(zn389_activo, cat_a, 0.0)

    df["Caudal_catalizador_activo"] = (
        pd.to_numeric(df["Caudal_ZN306_activo"], errors="coerce")
        + pd.to_numeric(df["Caudal_ZN389_activo"], errors="coerce")
    )

    df.loc[df["Caudal_catalizador_activo"] <= 0, "Caudal_catalizador_activo"] = np.nan

    cat_activo = pd.to_numeric(df["Caudal_catalizador_activo"], errors="coerce")

    df["Fraccion_ZN389_activo"] = (
        pd.to_numeric(df["Caudal_ZN389_activo"], errors="coerce") / cat_activo
    ).replace([np.inf, -np.inf], np.nan)

    # Codigo de catalizador:
    # 0 = ZN-306 dominante
    # 1 = mixto/transicion
    # 2 = ZN-389 dominante
    df["Tipo_catalizador_productividad"] = np.select(
        [
            df["Fraccion_ZN389_activo"] <= 0.25,
            df["Fraccion_ZN389_activo"] >= 0.75,
        ],
        [
            0.0,
            2.0,
        ],
        default=1.0,
    )

    df.loc[df["Fraccion_ZN389_activo"].isna(), "Tipo_catalizador_productividad"] = np.nan

    df["Indicador_ZN389_activo"] = np.where(
        df["Fraccion_ZN389_activo"] >= 0.5,
        1.0,
        0.0,
    )

    # Variables derivadas adicionales.
    if "MFI_pellets" in df.columns:
        mfi_pellets = pd.to_numeric(df["MFI_pellets"], errors="coerce")
        df["Delta_MFI_polvo_pellets"] = mfi - mfi_pellets
    else:
        df["Delta_MFI_polvo_pellets"] = np.nan

    if "Conc_propano" in df.columns:
        propano = pd.to_numeric(df["Conc_propano"], errors="coerce")
        df["Ratio_H2_propano"] = (h2 / propano).replace([np.inf, -np.inf], np.nan)
    else:
        df["Ratio_H2_propano"] = np.nan

    if all(c in df.columns for c in ["H2_a_K2301", "H2_a_R2301"]):
        df["H2_total"] = (
            pd.to_numeric(df["H2_a_K2301"], errors="coerce")
            + pd.to_numeric(df["H2_a_R2301"], errors="coerce")
        )
    else:
        df["H2_total"] = np.nan

    if all(c in df.columns for c in ["Caudal_E2301A", "Caudal_E2301B", "Caudal_E2301C"]):
        df["Carga_condensadores"] = (
            pd.to_numeric(df["Caudal_E2301A"], errors="coerce")
            + pd.to_numeric(df["Caudal_E2301B"], errors="coerce")
            + pd.to_numeric(df["Caudal_E2301C"], errors="coerce")
        )
    else:
        df["Carga_condensadores"] = np.nan

    df["Rel_TEA_CDonor_calc"] = (
        pd.to_numeric(df["TEA"], errors="coerce")
        / pd.to_numeric(df["C_Donor"], errors="coerce")
    ).replace([np.inf, -np.inf], np.nan)

    fecha = pd.to_datetime(df["Fecha_y_hora"], errors="coerce")

    base_2024 = (
        (fecha >= pd.Timestamp("2024-04-28"))
        & (fecha <= pd.Timestamp("2024-10-06 23:59:59"))
    )

    base_2025 = (
        (fecha >= pd.Timestamp("2025-05-01"))
        & (fecha <= pd.Timestamp("2025-12-31 23:59:59"))
    )

    periodo_excluido = (
        (fecha >= pd.Timestamp("2024-11-01"))
        & (fecha <= pd.Timestamp("2025-04-30 23:59:59"))
    )

    presion_operacion_ok = (presion_r >= 30.0) & (presion_r <= 31.0)

    df["Presion_operacion_OK"] = np.where(presion_operacion_ok, 1.0, 0.0)
    df["Periodo_base_productividad_2024"] = np.where((base_2024 | base_2025) & (~periodo_excluido), 1.0, 0.0)

    # Se inicializa. El calculo final se hace despues de la agrupacion temporal.
    # Nombres historicos mantenidos para compatibilidad:
    # - Productividad_estimada = Rendimiento estimado.
    # - Desvio_vs_productividad_estimada = Rendimiento real - Rendimiento esperado.
    df["Productividad_estimada"] = np.nan
    df["Rendimiento_esperado_producto"] = np.nan
    df["Target_operativo_producto"] = np.nan
    df["Maximo_historico_validado_producto"] = np.nan
    df["Desvio_vs_productividad_estimada"] = np.nan
    df["Gap_vs_target_operativo"] = np.nan
    df["Gap_vs_maximo_historico"] = np.nan
    df["Distancia_benchmark_productividad"] = np.nan
    df["Confiabilidad_benchmark_productividad"] = np.nan
    df["Indicador_target_optimo_productividad"] = np.nan
    df["Percentil_rendimiento_base"] = np.nan
    df["Familia_productividad_codigo"] = np.nan
    df["MFI_bin_productividad"] = np.nan
    df["H2_bin_productividad"] = np.nan
    df["XS_bin_productividad"] = np.nan
    df["Propano_bin_productividad"] = np.nan

    # Variables derivadas visibles en la interfaz.
    # Las variables auxiliares del modelo de target se calculan internamente,
    # pero no se agregan al selector para evitar saturar la lista de variables.
    variables_productividad = {
        "Productividad_estimada": "Rendimiento estimado",
        "Desvio_vs_productividad_estimada": "Desvio vs rendimiento estimado [Real - Estimado]",
        "Confiabilidad_benchmark_productividad": "Confiabilidad modelo dic-25/feb-26 [%]",
    }

    for var_prod, label_prod in variables_productividad.items():
        nombres_legibles[var_prod] = label_prod

        if var_prod not in todas_variables:
            todas_variables.append(var_prod)

    if "Productividad_estimada" not in default_vars:
        default_vars.append("Productividad_estimada")


with sidebar_herramientas:
    st.markdown("---")
    st.subheader("Herramientas avanzadas")

    key_defs = f"variables_calculadas_{unidad}"
    if key_defs not in st.session_state:
        st.session_state[key_defs] = []

    with st.expander("＋ Agregar variable calculada", expanded=False):
        st.caption("Armá una variable nueva combinando dos variables existentes.")

        nombre_nueva = st.text_input(
            "Nombre",
            value="",
            placeholder="Ejemplo: Relación TEA/C-Donor",
            key=f"calc_nombre_nueva_{unidad}",
        )

        opciones_calc = list(nombres_legibles.keys())

        var_a_nueva = st.selectbox(
            "Variable A",
            options=opciones_calc,
            format_func=lambda x: nombres_legibles[x],
            key=f"calc_var_a_nueva_{unidad}",
        )

        operacion_nueva = st.selectbox(
            "Operación",
            options=["+", "-", "×", "÷", "A/B × 100", "|A - B|", "Promedio"],
            key=f"calc_operacion_nueva_{unidad}",
        )

        var_b_nueva = st.selectbox(
            "Variable B",
            options=opciones_calc,
            index=min(1, len(opciones_calc) - 1),
            format_func=lambda x: nombres_legibles[x],
            key=f"calc_var_b_nueva_{unidad}",
        )

        definicion_preview = {
            "var_a": var_a_nueva,
            "operacion": operacion_nueva,
            "var_b": var_b_nueva,
        }
        st.caption("Vista previa")
        st.code(texto_formula(definicion_preview))

        if st.button("Agregar", width="stretch", key=f"btn_agregar_calc_{unidad}"):
            if not nombre_nueva.strip():
                st.error("Escribí un nombre para la variable calculada.")
            else:
                numero = len(st.session_state[key_defs]) + 1
                clave = f"Calc_{numero}"
                definicion = {
                    "clave": clave,
                    "nombre": nombre_nueva.strip(),
                    "var_a": var_a_nueva,
                    "operacion": operacion_nueva,
                    "var_b": var_b_nueva,
                }
                st.session_state[key_defs].append(definicion)
                st.success(f"Variable agregada: {nombre_nueva.strip()}")
                st.rerun()

    if st.session_state[key_defs]:
        with st.expander("Variables creadas", expanded=False):
            for idx, definicion in enumerate(st.session_state[key_defs]):
                st.markdown(f"**{definicion['nombre']}**")
                st.code(f"{definicion['clave']} = {texto_formula(definicion)}")
                col_ver, col_del = st.columns([1, 1])
                with col_ver:
                    clave = definicion["clave"]
                    if clave in df.columns:
                        validos = int(df[clave].notna().sum())
                        st.caption(f"Puntos válidos: {validos:,}")
                with col_del:
                    if st.button("Eliminar", key=f"delete_calc_{unidad}_{idx}", width="stretch"):
                        st.session_state[key_defs].pop(idx)
                        st.rerun()
                st.markdown("---")


with sidebar_principal:
    st.markdown("---")
    st.header("Controles principales")
    fecha_min = df["Fecha_y_hora"].min().date()
    fecha_max = df["Fecha_y_hora"].max().date()

    # Selector de fechas robusto.
    # Se usan campos de texto para evitar problemas del calendario de Streamlit,
    # y presets antes de crear los text_input para que los botones funcionen bien.
    st.caption(
        f"Rango disponible en Excel: {fecha_min.strftime('%Y/%m/%d')} a {fecha_max.strftime('%Y/%m/%d')}"
    )

    key_desde_text = f"fecha_desde_text_{unidad}"
    key_hasta_text = f"fecha_hasta_text_{unidad}"

    if key_desde_text not in st.session_state:
        st.session_state[key_desde_text] = fecha_min.strftime("%Y/%m/%d")

    if key_hasta_text not in st.session_state:
        st.session_state[key_hasta_text] = fecha_max.strftime("%Y/%m/%d")

    def _parse_fecha_sidebar(valor, fallback):
        if valor is None or str(valor).strip() == "":
            return fallback

        texto = str(valor).strip()

        for formato in ["%Y/%m/%d", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
            try:
                return pd.to_datetime(texto, format=formato).date()
            except Exception:
                pass

        fecha = pd.to_datetime(texto, errors="coerce", dayfirst=False)
        if pd.isna(fecha):
            fecha = pd.to_datetime(texto, errors="coerce", dayfirst=True)

        if pd.isna(fecha):
            return None

        return fecha.date()

    def _nombre_mes_es(fecha_ts):
        meses = {
            1: "Enero",
            2: "Febrero",
            3: "Marzo",
            4: "Abril",
            5: "Mayo",
            6: "Junio",
            7: "Julio",
            8: "Agosto",
            9: "Septiembre",
            10: "Octubre",
            11: "Noviembre",
            12: "Diciembre",
        }
        return f"{meses[int(fecha_ts.month)]} {int(fecha_ts.year)}"

    def _fin_de_mes(fecha_ts):
        fecha_ts = pd.Timestamp(fecha_ts)
        return (fecha_ts + pd.offsets.MonthEnd(0)).date()

    # Presets de período. Importante: van ANTES de los text_input.
    # Si se actualiza st.session_state después de crear el text_input,
    # Streamlit puede no reflejar el cambio correctamente.
    st.markdown("**Selección rápida de período**")

    _meses_disponibles = (
        pd.to_datetime(df["Fecha_y_hora"], errors="coerce")
        .dropna()
        .dt.to_period("M")
        .drop_duplicates()
        .sort_values()
    )

    if len(_meses_disponibles) > 0:
        _meses_ts = [p.to_timestamp() for p in _meses_disponibles]
        _labels_meses = [_nombre_mes_es(m) for m in _meses_ts]

        # Seleccionar por defecto el mes del campo Hasta si existe; si no, el último mes disponible.
        _hasta_actual = _parse_fecha_sidebar(st.session_state.get(key_hasta_text), fecha_max)
        _periodo_hasta_actual = pd.Period(pd.Timestamp(_hasta_actual), freq="M") if _hasta_actual else pd.Period(pd.Timestamp(fecha_max), freq="M")
        _periodos_lista = list(_meses_disponibles)

        if _periodo_hasta_actual in _periodos_lista:
            _idx_mes_default = _periodos_lista.index(_periodo_hasta_actual)
        else:
            _idx_mes_default = len(_meses_ts) - 1

        _mes_seleccionado_label = st.selectbox(
            "Mes para informe",
            options=_labels_meses,
            index=_idx_mes_default,
            key=f"mes_informe_{unidad}",
            help="Aplica el mes completo. Ejemplo: Julio 2026 = 2026/07/01 a 2026/07/31.",
        )

        _idx_mes_sel = _labels_meses.index(_mes_seleccionado_label)
        _mes_sel_ts = pd.Timestamp(_meses_ts[_idx_mes_sel])
        _inicio_mes_sel = _mes_sel_ts.date()
        _fin_mes_sel = _fin_de_mes(_mes_sel_ts)

        c_preset_1, c_preset_2 = st.columns(2)
        with c_preset_1:
            if st.button("Aplicar mes", key=f"btn_aplicar_mes_{unidad}", width="stretch"):
                st.session_state[key_desde_text] = _inicio_mes_sel.strftime("%Y/%m/%d")
                st.session_state[key_hasta_text] = _fin_mes_sel.strftime("%Y/%m/%d")
                st.rerun()

        with c_preset_2:
            if st.button("Usar todo", key=f"btn_fechas_todo_{unidad}", width="stretch"):
                st.session_state[key_desde_text] = fecha_min.strftime("%Y/%m/%d")
                st.session_state[key_hasta_text] = fecha_max.strftime("%Y/%m/%d")
                st.rerun()
    else:
        if st.button("Usar todo", key=f"btn_fechas_todo_{unidad}", width="stretch"):
            st.session_state[key_desde_text] = fecha_min.strftime("%Y/%m/%d")
            st.session_state[key_hasta_text] = fecha_max.strftime("%Y/%m/%d")
            st.rerun()

    st.markdown("**Rango seleccionado**")

    desde_txt = st.text_input(
        "Desde",
        key=key_desde_text,
        help="Escribir como YYYY/MM/DD. Ejemplo: 2026/07/01",
    )

    hasta_txt = st.text_input(
        "Hasta",
        key=key_hasta_text,
        help="Escribir como YYYY/MM/DD. Ejemplo: 2026/07/31",
    )

    desde = _parse_fecha_sidebar(desde_txt, fecha_min)
    hasta = _parse_fecha_sidebar(hasta_txt, fecha_max)

    if desde is None or hasta is None:
        st.error("Formato de fecha inválido. Usá YYYY/MM/DD, por ejemplo 2026/08/18.")
        st.stop()

    if desde > hasta:
        st.error("La fecha 'Desde' no puede ser posterior a 'Hasta'.")
        st.stop()

    if desde < fecha_min or hasta > fecha_max:
        st.warning(
            "El rango seleccionado queda parcialmente fuera del Excel. "
            "Se mostrarán solo los datos disponibles."
        )

    st.markdown("---")
    st.subheader("Variables a graficar")

    key_vars_grafico = f"variables_grafico_{unidad}"
    key_quitar_grafico = f"quitar_grafico_pendiente_{unidad}"

    if key_vars_grafico not in st.session_state:
        st.session_state[key_vars_grafico] = [v for v in default_vars if v in todas_variables]

    # Si el usuario tocó "✕" en la corrida anterior, se quita acá,
    # antes de crear el widget multiselect. Streamlit no permite modificar
    # st.session_state[key_vars_grafico] después de instanciar el widget.
    if key_quitar_grafico in st.session_state:
        _var_a_quitar = st.session_state.pop(key_quitar_grafico)
        st.session_state[key_vars_grafico] = [
            v for v in st.session_state[key_vars_grafico]
            if v != _var_a_quitar
        ]

    # Evita que queden variables de otra unidad o auxiliares ya ocultas.
    st.session_state[key_vars_grafico] = [
        v for v in st.session_state[key_vars_grafico]
        if v in todas_variables
    ]

    variables_sel = st.multiselect(
        "Selecciona variables:",
        options=todas_variables,
        format_func=lambda x: nombres_legibles[x],
        key=key_vars_grafico,
    )

    if variables_sel:
        with st.container(border=True):
            st.caption("Variables seleccionadas")
            for _var_sel in list(variables_sel):
                col_nombre, col_quitar = st.columns([4, 1])
                with col_nombre:
                    st.markdown(f"**{nombres_legibles[_var_sel]}**")
                with col_quitar:
                    if st.button("✕", key=f"quitar_grafico_{unidad}_{_var_sel}", help="Quitar variable"):
                        st.session_state[key_quitar_grafico] = _var_sel
                        st.rerun()


    st.markdown("---")
    st.subheader("Modo de grafico")
    modo_grafico = st.radio("Modo:", options=["Separados", "Combinados"], index=0)

    mostrar_cambios_producto = False
    if unidad.startswith("Polimer"):
        mostrar_cambios_producto = st.checkbox(
            "Mostrar cambios de grado",
            value=False,
            help=(
                "Dibuja una línea vertical en cada cambio de producto/campaña "
                "a lo largo de todo el período visible."
            ),
        )

    # Se elimina la opción rápida "Rendimiento real vs estimado".
    # Con el modo Combinados + selección de variables + eje Y secundario, ya no hace falta.
    grafico_rendimiento_target = False

    # Configuración del eje Y secundario desde el sidebar izquierdo.
    # No se usa panel lateral derecho para no achicar la gráfica.
    escala_y1_manual = None
    escala_y2_manual = None
    variables_y2_sidebar = []

    variables_para_ejes = [v for v in list(variables_sel) if v in todas_variables]

    if modo_grafico == "Combinados" and len(variables_para_ejes) >= 2:
        st.markdown("---")
        st.subheader("Eje Y secundario")
        variables_y2_sidebar = st.multiselect(
            "Variables en eje derecho:",
            options=variables_para_ejes,
            default=variables_para_ejes[1:],
            format_func=lambda x: nombres_legibles[x],
            key=f"variables_y2_sidebar_{unidad}",
            help="Las variables no seleccionadas quedan en el eje izquierdo.",
        )

        # Evitar que todas las variables queden en el eje derecho.
        if len(variables_y2_sidebar) == len(variables_para_ejes):
            variables_y2_sidebar = variables_y2_sidebar[1:]
            st.caption("Se dejó al menos una variable en el eje izquierdo.")

    st.markdown("---")
    st.subheader("Escalas Y")

    usar_escalas_manual = st.checkbox(
        "Ajustar escalas Y",
        value=False,
        help="Permite fijar manualmente los mínimos y máximos visibles de los ejes Y.",
    )

    # Contenedor fijo para los campos de escala. Se completa más adelante,
    # cuando ya existen df_f, variables_y1 y variables_y2.
    escala_y_sidebar_container = st.container()

    st.markdown("---")
    st.subheader("Agrupacion temporal")
    agrupacion = st.selectbox(
        "Frecuencia:",
        options=["Hora", "Día", "Semana", "Mes"],
        index=2,
    )

    st.markdown("---")
    st.subheader("Filtros recomendados")
    filtrar_caudales_negativos = st.checkbox(
        "Filtros recomendados",
        value=True,
        help=(
            "Aplica filtros automaticos recomendados segun la unidad seleccionada. "
            "En Polimerizacion usa nivel, URA, presion, MFI polvo y slurry. "
            "En Destilacion usa presion C-1001, ingreso a splitter y presion de succion K-1001."
        ),
    )

    if filtrar_caudales_negativos:
        if unidad.startswith("Polimer"):
            st.caption(
                "Activos Polimerizacion: 60<Nivel<70; 8000<URA<13000; "
                "30<P R-2301<31; MFI polvo>0; Slurry>50."
            )
        elif unidad.startswith("Destil"):
            st.caption(
                "Activos Destilacion: P C-1001>28; Ingreso splitter>36; "
                "P succion K-1001>10."
            )

    st.markdown("---")
    st.subheader("Filtros por variable")

    filtros_recomendados_vars = []
    if filtrar_caudales_negativos:
        if unidad.startswith("Polimer"):
            filtros_recomendados_vars = [
                v for v in [
                    "Nivel_R2301",
                    "Calor_reaccion_URA",
                    "Presion_R2301",
                    "MFI_polvo",
                    "Slurry",
                ]
                if v in todas_variables
            ]
        elif unidad.startswith("Destil"):
            filtros_recomendados_vars = [
                v for v in [
                    "Presion_C1001",
                    "Ingreso_splitter",
                    "P_succion_K1001",
                ]
                if v in todas_variables
            ]

    # Mantiene seleccionadas las variables recomendadas cuando el tilde está activo.
    if "vars_filtro_estado" not in st.session_state:
        st.session_state["vars_filtro_estado"] = []

    # Evita que queden variables de la otra unidad cuando se cambia Destilación/Polimerización.
    st.session_state["vars_filtro_estado"] = [
        v for v in st.session_state["vars_filtro_estado"]
        if v in todas_variables
    ]

    if filtrar_caudales_negativos:
        for _v_rec in filtros_recomendados_vars:
            if _v_rec not in st.session_state["vars_filtro_estado"]:
                st.session_state["vars_filtro_estado"].append(_v_rec)

    vars_filtro = st.multiselect(
        "Variables a filtrar:",
        options=todas_variables,
        format_func=lambda x: nombres_legibles[x],
        key="vars_filtro_estado",
    )



    # ==============================================================================
    # FILTRO POR PRODUCTO / CAMPAÑA
    # ==============================================================================
    productos_sel = []
    df_producto_marcas = None

    if unidad.startswith("Polimer"):
        st.markdown("---")
        st.subheader("Filtro por producto")

        if "Producto" in df.columns:
            serie_producto = (
                df["Producto"]
                .astype("string")
                .str.strip()
                .replace({
                    "": pd.NA,
                    "nan": pd.NA,
                    "None": pd.NA,
                    "NONE": pd.NA,
                    "NaN": pd.NA,
                    "<NA>": pd.NA,
                    "0": pd.NA,
                    "0.0": pd.NA,
                })
            )

            df["Producto"] = serie_producto.map(normalizar_producto_operativo)

            # Guardamos la línea de tiempo completa de productos ANTES de filtrar.
            # Se usa solo para dibujar cambios de grado/campaña en las gráficas.
            df_producto_marcas = df[["Fecha_y_hora", "Producto"]].copy()

            productos_disponibles = sorted([
                str(p).strip()
                for p in df["Producto"].dropna().unique()
                if str(p).strip()
            ])

            if productos_disponibles:
                st.caption(f"Productos detectados: {len(productos_disponibles)}")

                productos_sel = st.multiselect(
                    "Producto:",
                    options=productos_disponibles,
                    default=[],
                    help=(
                        "Sin selección = todos los productos. Si seleccionás uno o más, "
                        "la app filtra las campañas antes de agrupar y calcular el target."
                    ),
                )

                if productos_sel:
                    df = df[df["Producto"].astype(str).isin(productos_sel)].copy()
                    st.caption(f"Productos seleccionados: {len(productos_sel)}")
                else:
                    st.caption("Sin selección: se incluyen todos los productos.")
            else:
                st.warning(
                    "La columna PRODUCTO fue encontrada, pero no se detectaron valores no vacíos. "
                    "Verificá que el Excel cargado tenga productos escritos en la hoja POLIMERIZACIÓN."
                )
        else:
            st.warning(
                "No se encontró la columna PRODUCTO. Cargá el Excel actualizado con la nueva columna PRODUCTO."
            )



# ==============================================================================
# AGRUPACION Y FILTROS
# ==============================================================================

filtro_producto_activo = unidad.startswith("Polimer") and len(productos_sel) > 0

df_agrup = aplicar_agrupacion(df, agrupacion)


# AJUSTE FINAL POST-AGRUPACION RENDIMIENTO ESTIMADO POR VARIABLES CRITICAS
#
# Criterio conceptual:
# - El rendimiento estimado NO depende del cambio de lecho.
# - El rendimiento estimado NO se corrige contra el rendimiento real de la campaña actual.
# - El estimado se calcula por variables críticas de proceso/catalizador.
# - El desvío Real - Estimado se usa como señal de posible pérdida de actividad,
#   contaminante o necesidad de revisar/cambiar lechos.
#
# Cambio principal de esta versión:
# - Se reemplaza el estimador por vecinos como método principal porque tendía a
#   generar una curva demasiado plana.
# - El método default es Gradient Boosting quantile, que captura respuestas no
#   lineales de MFI, H2, propano, donor, slurry, temperatura, etc.
if unidad.startswith("Polimer") and all(
    col in df_agrup.columns
    for col in [
        "Fecha_y_hora",
        "Rendimiento",
        "Presion_R2301",
        "Productividad_estimada",
    ]
):
    _fecha = pd.to_datetime(df_agrup["Fecha_y_hora"], errors="coerce")
    _rendimiento = pd.to_numeric(df_agrup["Rendimiento"], errors="coerce")
    _presion = pd.to_numeric(df_agrup["Presion_R2301"], errors="coerce")

    with sidebar_modelo:
        st.markdown("---")
        st.subheader("Modelo rendimiento estimado")

        with st.expander("Benchmark por variables críticas", expanded=True):
            st.caption(
                "El estimado se calcula con variables críticas de proceso/catalizador. "
                "No usa cambio de lecho como variable y no se ajusta contra el real de la campaña actual."
            )

            _metodo_estimacion = st.selectbox(
                "Método de estimación",
                options=[
                    "Gradient Boosting quantile",
                    "Gradient Boosting promedio",
                ],
                index=0,
                key="metodo_estimacion_rendimiento",
                help=(
                    "Gradient Boosting captura relaciones no lineales. "
                    "Quantile usa un percentil configurable como benchmark esperado."
                ),
            )

            _percentil_benchmark = st.number_input(
                "Percentil benchmark",
                min_value=0.40,
                max_value=0.90,
                value=0.65,
                step=0.05,
                format="%.2f",
                key="percentil_benchmark_gbr",
                help=(
                    "Percentil usado por el modelo quantile. "
                    "0.50 se aproxima a mediana; 0.65/0.70 es un benchmark algo exigente."
                ),
            )

            c_gbr_1, c_gbr_2 = st.columns(2)
            with c_gbr_1:
                _gbr_n_estimators = st.number_input(
                    "Árboles",
                    min_value=80,
                    max_value=800,
                    value=300,
                    step=20,
                    key="gbr_n_estimators_rendimiento",
                )
            with c_gbr_2:
                _gbr_max_depth = st.number_input(
                    "Profundidad",
                    min_value=2,
                    max_value=5,
                    value=3,
                    step=1,
                    key="gbr_max_depth_rendimiento",
                )

            c_train_1, c_train_2 = st.columns(2)
            with c_train_1:
                _solo_historico_anterior = st.checkbox(
                    "Entrenar solo con datos anteriores al rango visible",
                    value=True,
                    key="gbr_solo_historico_anterior",
                    help=(
                        "Evita que el período que estás evaluando se use para calibrar su propio estimado."
                    ),
                )
            with c_train_2:
                _dias_exclusion_entrenamiento = st.number_input(
                    "Exclusión antes del rango [días]",
                    min_value=0,
                    max_value=90,
                    value=14,
                    step=1,
                    key="gbr_dias_exclusion_entrenamiento",
                )

            st.markdown("**Exclusiones por datos no confiables**")
            _excluir_evento_bajo_nivel = st.checkbox(
                "Excluir evento bajo nivel reactor",
                value=True,
                key="excluir_evento_bajo_nivel_modelo",
                help=(
                    "Excluye el período con baja productividad por bajo nivel no detectado. "
                    "Abril 2025 queda incluido por defecto."
                ),
            )

            c_exc_1, c_exc_2 = st.columns(2)
            with c_exc_1:
                _inicio_evento_bajo_nivel = st.date_input(
                    "Inicio evento",
                    value=pd.Timestamp("2024-11-01").date(),
                    key="inicio_evento_bajo_nivel_modelo",
                )
            with c_exc_2:
                _fin_evento_bajo_nivel = st.date_input(
                    "Fin evento",
                    value=pd.Timestamp("2025-03-31").date(),
                    key="fin_evento_bajo_nivel_modelo",
                )

            _min_puntos_entrenamiento = st.number_input(
                "Mínimo puntos entrenamiento",
                min_value=30,
                max_value=2000,
                value=120,
                step=10,
                key="min_puntos_entrenamiento_gbr",
            )

            if not SKLEARN_DISPONIBLE:
                st.error(
                    "scikit-learn no está disponible. Agregá scikit-learn a requirements.txt "
                    "o usá la versión anterior."
                )

            st.info(
                "Uso diagnóstico: si Real - Estimado es negativo de forma sostenida, "
                "puede indicar pérdida de actividad del catalizador por contaminante o necesidad de revisar lechos."
            )

    _presion_ok = (_presion > 30.0) & (_presion < 31.0)

    # ----------------------------------------------------------------------
    # Variables derivadas para el modelo
    # ----------------------------------------------------------------------
    if "C_Donor" in df_agrup.columns:
        _cdonor_num = pd.to_numeric(df_agrup["C_Donor"], errors="coerce")
        df_agrup["Sin_C_Donor_modelo"] = np.where(
            _cdonor_num.fillna(0.0) <= 0.001,
            1.0,
            0.0,
        )
    else:
        df_agrup["Sin_C_Donor_modelo"] = np.nan

    # Relación TEA/C-Donor. Cuando no hay donor, la relación es indefinida:
    # se deja en NaN y el imputador del modelo usa la mediana del histórico válido.
    _rel_excel = None
    if "Rel_TEA_CDonor" in df_agrup.columns:
        _rel_excel = pd.to_numeric(df_agrup["Rel_TEA_CDonor"], errors="coerce")

    _rel_calc = None
    if all(c in df_agrup.columns for c in ["TEA", "C_Donor"]):
        _tea = pd.to_numeric(df_agrup["TEA"], errors="coerce")
        _cdonor = pd.to_numeric(df_agrup["C_Donor"], errors="coerce")
        _rel_calc = (_tea / _cdonor.where(_cdonor.abs() > 0.001)).replace(
            [np.inf, -np.inf],
            np.nan,
        )

    if _rel_excel is not None and _rel_calc is not None:
        _rel = _rel_excel.fillna(_rel_calc)
    elif _rel_excel is not None:
        _rel = _rel_excel
    elif _rel_calc is not None:
        _rel = _rel_calc
    else:
        _rel = pd.Series(np.nan, index=df_agrup.index)

    df_agrup["Rel_TEA_CDonor_modelo"] = _rel
    if "Sin_C_Donor_modelo" in df_agrup.columns:
        df_agrup.loc[
            pd.to_numeric(df_agrup["Sin_C_Donor_modelo"], errors="coerce").fillna(0.0) > 0.5,
            "Rel_TEA_CDonor_modelo",
        ] = np.nan

    # Variables críticas candidatas.
    # PRODUCTO no entra en el cálculo; el grado queda solo para filtrar/visualizar.
    _features_candidatas = [
        "MFI_polvo",
        "XS",
        "Conc_H2",
        "Conc_propano",
        "Slurry",
        "Rel_TEA_CDonor_modelo",
        "C_Donor",
        "Sin_C_Donor_modelo",
        "Temperatura_R2301",
        "Tipo_catalizador_productividad",
        "Indicador_ZN389_activo",
    ]

    _features = []
    for _f in _features_candidatas:
        if _f in df_agrup.columns:
            _serie = pd.to_numeric(df_agrup[_f], errors="coerce")
            if _serie.notna().sum() >= 20:
                _features.append(_f)

    _estimado = np.full(len(df_agrup), np.nan)
    _conf = np.full(len(df_agrup), np.nan)
    _dist_min = np.full(len(df_agrup), np.nan)
    _n_vecinos_usados = np.full(len(df_agrup), np.nan)
    _indicador_base = np.zeros(len(df_agrup))
    _modelo_status = "Sin modelo"
    _n_train = 0

    if len(_features) >= 3 and SKLEARN_DISPONIBLE:
        # Conjunto histórico válido. Se excluyen datos malos, no cambios de lecho.
        _base_mask = (
            _presion_ok
            & _rendimiento.notna()
            & _fecha.notna()
        )

        if bool(_excluir_evento_bajo_nivel):
            _evento_bajo_nivel = (
                (_fecha >= pd.Timestamp(_inicio_evento_bajo_nivel))
                & (
                    _fecha
                    <= pd.Timestamp(_fin_evento_bajo_nivel)
                    + pd.Timedelta(days=1)
                    - pd.Timedelta(seconds=1)
                )
            )
            _base_mask = _base_mask & (~_evento_bajo_nivel)

        # Evitar auto-calibración del rango visible.
        if bool(_solo_historico_anterior):
            _inicio_visible = pd.Timestamp(desde) - pd.Timedelta(days=int(_dias_exclusion_entrenamiento))
            _base_mask = _base_mask & (_fecha < _inicio_visible)

        # Filtros de operación normal si existen.
        if "Nivel_R2301" in df_agrup.columns:
            _nivel = pd.to_numeric(df_agrup["Nivel_R2301"], errors="coerce")
            _base_mask = _base_mask & (_nivel > 60.0) & (_nivel < 70.0)

        if "Calor_reaccion_URA" in df_agrup.columns:
            _ura = pd.to_numeric(df_agrup["Calor_reaccion_URA"], errors="coerce")
            _base_mask = _base_mask & (_ura > 8000.0) & (_ura < 13000.0)

        if "Slurry" in _features:
            _base_mask = _base_mask & (
                pd.to_numeric(df_agrup["Slurry"], errors="coerce") > 50.0
            )

        if "MFI_polvo" in _features:
            _base_mask = _base_mask & (
                pd.to_numeric(df_agrup["MFI_polvo"], errors="coerce") > 0.0
            )

        if "Conc_propano" in _features:
            _base_mask = _base_mask & (
                pd.to_numeric(df_agrup["Conc_propano"], errors="coerce") > 0.0
            )

        if "Conc_H2" in _features:
            _base_mask = _base_mask & (
                pd.to_numeric(df_agrup["Conc_H2"], errors="coerce") > 0.0
            )

        _x_all = df_agrup[_features].copy()
        for _f in _features:
            _x_all[_f] = pd.to_numeric(_x_all[_f], errors="coerce")

        # El modelo acepta NaN por imputación, pero se exige que cada fila tenga
        # una cantidad mínima de variables críticas observadas.
        _min_features_validas = max(3, int(np.ceil(len(_features) * 0.55)))
        _filas_con_features = _x_all.notna().sum(axis=1) >= _min_features_validas
        _base_final = _base_mask & _filas_con_features

        _n_train = int(_base_final.sum())

        if _n_train >= int(_min_puntos_entrenamiento):
            _x_train = _x_all.loc[_base_final, _features]
            _y_train = _rendimiento.loc[_base_final]

            if _metodo_estimacion.startswith("Gradient Boosting quantile"):
                _reg = GradientBoostingRegressor(
                    loss="quantile",
                    alpha=float(_percentil_benchmark),
                    n_estimators=int(_gbr_n_estimators),
                    max_depth=int(_gbr_max_depth),
                    learning_rate=0.03,
                    min_samples_leaf=5,
                    random_state=42,
                )
                _modelo_status = f"Gradient Boosting quantile P{int(float(_percentil_benchmark) * 100)}"
            else:
                _reg = GradientBoostingRegressor(
                    loss="absolute_error",
                    n_estimators=int(_gbr_n_estimators),
                    max_depth=int(_gbr_max_depth),
                    learning_rate=0.03,
                    min_samples_leaf=5,
                    random_state=42,
                )
                _modelo_status = "Gradient Boosting promedio robusto"

            _modelo = make_pipeline(
                SimpleImputer(strategy="median"),
                _reg,
            )

            try:
                _modelo.fit(_x_train, _y_train)
                _pred = _modelo.predict(_x_all[_features])

                # Evitar extrapolaciones absurdas fuera del rango histórico válido.
                _q01 = float(_y_train.quantile(0.01))
                _q99 = float(_y_train.quantile(0.99))
                _iqr_y = float(_y_train.quantile(0.75) - _y_train.quantile(0.25))
                _margen_y = max(1.50, 0.75 * _iqr_y)

                _pred = np.clip(
                    _pred,
                    _q01 - _margen_y,
                    _q99 + _margen_y,
                )

                _estimado = _pred.astype(float)

                # Diagnóstico simple de confianza.
                _conf_base = min(100.0, 100.0 * _n_train / max(float(_min_puntos_entrenamiento), 1.0))
                _conf[:] = _conf_base
                _n_vecinos_usados[:] = float(_n_train)
                _indicador_base[_base_final.to_numpy(dtype=bool)] = 1.0

            except Exception as _exc_modelo:
                _modelo_status = f"Error modelo: {_exc_modelo}"

    df_agrup["Productividad_estimada"] = pd.Series(_estimado, index=df_agrup.index)
    df_agrup["Rendimiento_esperado_producto"] = pd.Series(_estimado, index=df_agrup.index)

    # Variables auxiliares mantenidas por compatibilidad con el resto de la app.
    df_agrup["Target_operativo_producto"] = np.nan
    df_agrup["Maximo_historico_validado_producto"] = np.nan
    df_agrup["Distancia_benchmark_productividad"] = pd.Series(_dist_min, index=df_agrup.index)
    df_agrup["Confiabilidad_benchmark_productividad"] = pd.Series(_conf, index=df_agrup.index)
    df_agrup["N_vecinos_benchmark_productividad"] = pd.Series(_n_vecinos_usados, index=df_agrup.index)
    df_agrup["Indicador_target_optimo_productividad"] = pd.Series(_indicador_base, index=df_agrup.index)
    df_agrup["Percentil_rendimiento_base"] = np.nan

    df_agrup["Desvio_vs_productividad_estimada"] = (
        pd.to_numeric(df_agrup["Rendimiento"], errors="coerce")
        - pd.to_numeric(df_agrup["Productividad_estimada"], errors="coerce")
    )

    df_agrup["Indice_actividad_vs_estimado"] = (
        pd.to_numeric(df_agrup["Rendimiento"], errors="coerce")
        / pd.to_numeric(df_agrup["Productividad_estimada"], errors="coerce")
    ).replace([np.inf, -np.inf], np.nan)

    _desvio_tmp = pd.to_numeric(
        df_agrup["Desvio_vs_productividad_estimada"],
        errors="coerce",
    )

    df_agrup["Alerta_baja_actividad"] = np.where(
        _desvio_tmp <= -1.0,
        1.0,
        0.0,
    )

    df_agrup["Gap_vs_target_operativo"] = np.nan
    df_agrup["Gap_vs_maximo_historico"] = np.nan
    df_agrup["Presion_operacion_OK"] = np.where(_presion_ok, 1.0, 0.0)
    df_agrup["Periodo_base_productividad_2024"] = pd.Series(
        _indicador_base,
        index=df_agrup.index,
    )

    with sidebar_modelo:
        with st.expander("Diagnóstico rendimiento estimado", expanded=False):
            st.write(f"**Método:** {_modelo_status}")
            st.write(f"**Puntos entrenamiento:** {_n_train:,}")
            st.write(f"**Variables usadas:** {len(_features)}")
            if _features:
                st.code("\\n".join(_features), language="text")
            if _n_train < int(_min_puntos_entrenamiento):
                st.warning(
                    "Hay pocos puntos de entrenamiento con los filtros actuales. "
                    "Bajá el mínimo o revisá exclusiones/filtros."
                )


# Mascara inicial: rango de fechas
mask = (
    (df_agrup["Fecha_y_hora"].dt.date >= desde)
    & (df_agrup["Fecha_y_hora"].dt.date <= hasta)
)

# Filtros recomendados.
# Se aplican con el tilde "Filtros recomendados" del panel lateral.
# Polimerizacion:
# - 60 < Nivel_R2301 < 70
# - 8000 < Calor_reaccion_URA < 13000
# - 30 < Presion_R2301 < 31
# - MFI_polvo > 0
# - Slurry > 50
# Destilacion:
# - Presion_C1001 > 28
# - Ingreso_splitter > 36
# - P_succion_K1001 > 10
if filtrar_caudales_negativos:
    if unidad.startswith("Polimer"):
        if "Nivel_R2301" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Nivel_R2301"], errors="coerce")
            mask = mask & ((serie > 60.0) & (serie < 70.0))

        if "Calor_reaccion_URA" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Calor_reaccion_URA"], errors="coerce")
            mask = mask & ((serie > 8000.0) & (serie < 13000.0))

        if "Presion_R2301" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Presion_R2301"], errors="coerce")
            mask = mask & ((serie > 30.0) & (serie < 31.0))

        if "MFI_polvo" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["MFI_polvo"], errors="coerce")
            mask = mask & (serie > 0.0)

        if "Slurry" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Slurry"], errors="coerce")
            mask = mask & (serie > 50.0)

    elif unidad.startswith("Destil"):
        if "Presion_C1001" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Presion_C1001"], errors="coerce")
            mask = mask & (serie > 28.0)

        if "Ingreso_splitter" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["Ingreso_splitter"], errors="coerce")
            mask = mask & (serie > 36.0)

        if "P_succion_K1001" in df_agrup.columns:
            serie = pd.to_numeric(df_agrup["P_succion_K1001"], errors="coerce")
            mask = mask & (serie > 10.0)

# Rangos que se deben mostrar en "Variables a filtrar" cuando está activo
# el tilde "Filtros recomendados".
# El tercer elemento indica el tipo de comparación:
# - "between": min < variable < max
# - "gt": variable > min
# - "lt": variable < max
filtros_recomendados_rangos = {}
if filtrar_caudales_negativos:
    if unidad.startswith("Polimer"):
        filtros_recomendados_rangos = {
            "Nivel_R2301": (60.0, 70.0, "between"),
            "Calor_reaccion_URA": (8000.0, 13000.0, "between"),
            "Presion_R2301": (30.0, 31.0, "between"),
            "MFI_polvo": (0.0, None, "gt"),
            "Slurry": (50.0, None, "gt"),
        }
    elif unidad.startswith("Destil"):
        filtros_recomendados_rangos = {
            "Presion_C1001": (28.0, None, "gt"),
            "Ingreso_splitter": (36.0, None, "gt"),
            "P_succion_K1001": (10.0, None, "gt"),
        }

# Filtros manuales por variable
for var in vars_filtro:
    serie = pd.to_numeric(df_agrup[var], errors="coerce").dropna()
    if len(serie) == 0:
        continue

    vmin_real = float(serie.min())
    vmax_real = float(serie.max())

    vmin = vmin_real
    vmax = vmax_real

    rango_recomendado = filtros_recomendados_rangos.get(var)

    if rango_recomendado is not None:
        if len(rango_recomendado) == 3:
            rec_min, rec_max, rec_tipo = rango_recomendado
        else:
            rec_min, rec_max = rango_recomendado
            rec_tipo = "between" if rec_min is not None and rec_max is not None else ("gt" if rec_min is not None else "lt")

        if rec_min is not None:
            vmin = min(vmin, float(rec_min))
            vmax = max(vmax, float(rec_min))

        if rec_max is not None:
            vmin = min(vmin, float(rec_max))
            vmax = max(vmax, float(rec_max))

    if vmin == vmax:
        continue

    min_key = f"min_{unidad}_{var}"
    max_key = f"max_{unidad}_{var}"

    # Cuando los filtros recomendados están activos, fuerza los valores visibles
    # para que coincidan con el criterio recomendado.
    if rango_recomendado is not None:
        if len(rango_recomendado) == 3:
            rec_min, rec_max, rec_tipo = rango_recomendado
        else:
            rec_min, rec_max = rango_recomendado
            rec_tipo = "between" if rec_min is not None and rec_max is not None else ("gt" if rec_min is not None else "lt")

        st.session_state[min_key] = float(rec_min) if rec_min is not None else vmin
        st.session_state[max_key] = float(rec_max) if rec_max is not None else vmax
    else:
        # Inicializa o corrige valores guardados si cambia el rango por agrupacion/filtro.
        if min_key not in st.session_state or not (vmin <= float(st.session_state[min_key]) <= vmax):
            st.session_state[min_key] = vmin
        if max_key not in st.session_state or not (vmin <= float(st.session_state[max_key]) <= vmax):
            st.session_state[max_key] = vmax

    with st.sidebar.container(border=True):
        st.markdown(f"**{nombres_legibles[var]}**")

        if rango_recomendado is not None:
            if len(rango_recomendado) == 3:
                rec_min, rec_max, rec_tipo = rango_recomendado
            else:
                rec_min, rec_max = rango_recomendado
                rec_tipo = "between" if rec_min is not None and rec_max is not None else ("gt" if rec_min is not None else "lt")

            if rec_tipo == "between" and rec_min is not None and rec_max is not None:
                st.caption(
                    f"Filtro recomendado aplicado: {rec_min:.3f} < valor < {rec_max:.3f} | "
                    f"Rango real: {vmin_real:.3f} a {vmax_real:.3f}"
                )
            elif rec_tipo == "gt" and rec_min is not None:
                st.caption(
                    f"Filtro recomendado aplicado: valor > {rec_min:.3f} | "
                    f"Rango real: {vmin_real:.3f} a {vmax_real:.3f}"
                )
            elif rec_tipo == "lt" and rec_max is not None:
                st.caption(
                    f"Filtro recomendado aplicado: valor < {rec_max:.3f} | "
                    f"Rango real: {vmin_real:.3f} a {vmax_real:.3f}"
                )
        else:
            st.caption(f"Rango disponible: {vmin:.3f} a {vmax:.3f}")

        if st.button("Reset", key=f"reset_{unidad}_{var}", width="stretch"):
            st.session_state[min_key] = vmin
            st.session_state[max_key] = vmax
            st.rerun()

        c1, c2 = st.columns(2)
        step = max((vmax - vmin) / 100, 0.0001)
        with c1:
            r_min = st.number_input(
                "Minimo",
                min_value=vmin,
                max_value=vmax,
                step=step,
                key=min_key,
                format="%.3f",
            )
        with c2:
            r_max = st.number_input(
                "Maximo",
                min_value=vmin,
                max_value=vmax,
                step=step,
                key=max_key,
                format="%.3f",
            )

    if r_min > r_max:
        st.sidebar.error(f"Filtro invalido en {nombres_legibles[var]}: minimo mayor que maximo.")
        st.stop()

    if rango_recomendado is not None:
        serie_filtro = pd.to_numeric(df_agrup[var], errors="coerce")

        if len(rango_recomendado) == 3:
            rec_min, rec_max, rec_tipo = rango_recomendado
        else:
            rec_min, rec_max = rango_recomendado
            rec_tipo = "between" if rec_min is not None and rec_max is not None else ("gt" if rec_min is not None else "lt")

        if rec_tipo == "between":
            mask = mask & ((serie_filtro > float(rec_min)) & (serie_filtro < float(rec_max)))
        elif rec_tipo == "gt":
            mask = mask & (serie_filtro > float(rec_min))
        elif rec_tipo == "lt":
            mask = mask & (serie_filtro < float(rec_max))
        else:
            mask = mask & serie_filtro.between(r_min, r_max)
    elif r_min > vmin or r_max < vmax:
        mask = mask & pd.to_numeric(df_agrup[var], errors="coerce").between(r_min, r_max)

df_f = df_agrup[mask].copy()

# Línea de tiempo de productos para marcas de cambio de grado.
# Se calcula sobre el dataset previo al filtro por producto para que,
# aun filtrando un grado, se puedan ver los cambios/campañas del período.
df_producto_marcas_visible = pd.DataFrame()

if unidad.startswith("Polimer") and df_producto_marcas is not None:
    try:
        _tmp_marcas = df_producto_marcas.copy()
        _tmp_marcas["Fecha_y_hora"] = pd.to_datetime(
            _tmp_marcas["Fecha_y_hora"],
            errors="coerce",
        )
        _desde_dt = pd.Timestamp(desde)
        _hasta_dt = pd.Timestamp(hasta) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        df_producto_marcas_visible = _tmp_marcas[
            (_tmp_marcas["Fecha_y_hora"] >= _desde_dt)
            & (_tmp_marcas["Fecha_y_hora"] <= _hasta_dt)
        ].copy()
    except Exception:
        df_producto_marcas_visible = pd.DataFrame()


# ==============================================================================
# CUERPO PRINCIPAL - ENCABEZADO
# ==============================================================================

st.markdown(
    f"**Unidad:** {unidad} | "
    f"**Periodo:** {desde.strftime('%d/%m/%Y')} a {hasta.strftime('%d/%m/%Y')} | "
    f"**Agrupacion:** {agrupacion} | "
    f"**Registros filtrados:** {len(df_f):,}"
)

if len(df_f) == 0:
    st.warning("Sin datos para ese filtro.")
    st.stop()

with st.expander("Resumen rapido de datos filtrados", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.metric("Registros", f"{len(df_f):,}")
    c2.metric("Desde", df_f["Fecha_y_hora"].min().strftime("%d/%m/%Y %H:%M"))
    c3.metric("Hasta", df_f["Fecha_y_hora"].max().strftime("%d/%m/%Y %H:%M"))

    if variables_sel:
        resumen = df_f[variables_sel].describe().T
        resumen.index = [nombres_legibles[v] for v in resumen.index]
        st.dataframe(resumen.round(3), width="stretch")


# ==============================================================================
# TABS
# ==============================================================================

tab1, tab2, tab3, tab4 = st.tabs([
    "Series temporales",
    "Correlaciones",
    "Desfase temporal",
    "Datos filtrados",
])


# ==============================================================================
# TAB 1 - SERIES TEMPORALES
# ==============================================================================

with tab1:
    st.subheader(f"{unidad} - Tendencias ({agrupacion.lower()})")

    if grafico_rendimiento_target:
        variables_grafico = ["Rendimiento", "Productividad_estimada"]
        modo_grafico_efectivo = "Combinados"
    else:
        variables_grafico = variables_sel
        modo_grafico_efectivo = modo_grafico

    variables_grafico = [v for v in variables_grafico if v in df_f.columns]

    if not variables_grafico:
        st.info("Selecciona al menos una variable en el panel lateral.")
    else:
        if grafico_rendimiento_target:
            st.caption("Vista rapida: Rendimiento real vs rendimiento estimado por producto. Esta opcion no cambia la seleccion general de variables.")

        if modo_grafico_efectivo == "Combinados":
            usar_y2_efectivo = len(variables_grafico) >= 2
            variables_y2 = [v for v in variables_y2_sidebar if v in variables_grafico] if usar_y2_efectivo else []
            variables_y1 = [v for v in variables_grafico if v not in variables_y2] if usar_y2_efectivo else list(variables_grafico)

            # Si por algún motivo el eje izquierdo queda vacío, forzar la primera variable allí.
            if usar_y2_efectivo and not variables_y1 and variables_grafico:
                variables_y1 = [variables_grafico[0]]
                variables_y2 = [v for v in variables_y2 if v != variables_grafico[0]]

            def _rango_ampliado(series, factor=2.0):
                datos = pd.to_numeric(series, errors="coerce").dropna()
                if len(datos) == 0:
                    return None

                y_min = float(datos.min())
                y_max = float(datos.max())

                if not np.isfinite(y_min) or not np.isfinite(y_max):
                    return None

                if y_min == y_max:
                    delta = max(abs(y_min) * 0.50, 5.0)
                else:
                    delta = (y_max - y_min)

                # Rango ampliado para que el usuario pueda abrir bastante la escala.
                min_total = y_min - factor * delta
                max_total = y_max + factor * delta

                if min_total == max_total:
                    max_total = min_total + 1.0

                step = max((max_total - min_total) / 500.0, 0.0001)

                return min_total, max_total, y_min, y_max, step

            escala_y1_manual = None
            escala_y2_manual = None

            if usar_escalas_manual:
                with escala_y_sidebar_container:
                    st.caption("Ingresá mínimo y máximo visible de cada eje.")

                    if variables_y1:
                        datos_y1 = pd.concat(
                            [
                                pd.to_numeric(df_f[v], errors="coerce")
                                for v in variables_y1
                                if v in df_f.columns
                            ],
                            axis=0,
                        )

                        rango_y1 = _rango_ampliado(datos_y1, factor=8.0)

                        if rango_y1 is not None:
                            y1_min_total, y1_max_total, y1_min_auto, y1_max_auto, y1_step = rango_y1

                            st.markdown("**Eje Y izquierdo**")
                            c_y1_min, c_y1_max = st.columns(2)
                            with c_y1_min:
                                y1_min = st.number_input(
                                    "Min",
                                    min_value=float(y1_min_total),
                                    max_value=float(y1_max_total),
                                    value=float(y1_min_auto),
                                    step=float(y1_step),
                                    key=(
                                        f"num_y1_min_{unidad}_{agrupacion}_"
                                        f"{abs(hash(tuple(variables_y1)))}_"
                                        f"{round(y1_min_total, 3)}_{round(y1_max_total, 3)}"
                                    ),
                                    format="%.4f",
                                )
                            with c_y1_max:
                                y1_max = st.number_input(
                                    "Max",
                                    min_value=float(y1_min_total),
                                    max_value=float(y1_max_total),
                                    value=float(y1_max_auto),
                                    step=float(y1_step),
                                    key=(
                                        f"num_y1_max_{unidad}_{agrupacion}_"
                                        f"{abs(hash(tuple(variables_y1)))}_"
                                        f"{round(y1_min_total, 3)}_{round(y1_max_total, 3)}"
                                    ),
                                    format="%.4f",
                                )

                            if y1_min < y1_max:
                                escala_y1_manual = (float(y1_min), float(y1_max))
                            else:
                                st.warning("Y izquierdo: Min debe ser menor que Max.")

                    if usar_y2_efectivo and variables_y2:
                        datos_y2 = pd.concat(
                            [
                                pd.to_numeric(df_f[v], errors="coerce")
                                for v in variables_y2
                                if v in df_f.columns
                            ],
                            axis=0,
                        )

                        rango_y2 = _rango_ampliado(datos_y2, factor=8.0)

                        if rango_y2 is not None:
                            y2_min_total, y2_max_total, y2_min_auto, y2_max_auto, y2_step = rango_y2

                            st.markdown("**Eje Y derecho**")
                            c_y2_min, c_y2_max = st.columns(2)
                            with c_y2_min:
                                y2_min = st.number_input(
                                    "Min ",
                                    min_value=float(y2_min_total),
                                    max_value=float(y2_max_total),
                                    value=float(y2_min_auto),
                                    step=float(y2_step),
                                    key=(
                                        f"num_y2_min_{unidad}_{agrupacion}_"
                                        f"{abs(hash(tuple(variables_y2)))}_"
                                        f"{round(y2_min_total, 3)}_{round(y2_max_total, 3)}"
                                    ),
                                    format="%.4f",
                                )
                            with c_y2_max:
                                y2_max = st.number_input(
                                    "Max ",
                                    min_value=float(y2_min_total),
                                    max_value=float(y2_max_total),
                                    value=float(y2_max_auto),
                                    step=float(y2_step),
                                    key=(
                                        f"num_y2_max_{unidad}_{agrupacion}_"
                                        f"{abs(hash(tuple(variables_y2)))}_"
                                        f"{round(y2_min_total, 3)}_{round(y2_max_total, 3)}"
                                    ),
                                    format="%.4f",
                                )

                            if y2_min < y2_max:
                                escala_y2_manual = (float(y2_min), float(y2_max))
                            else:
                                st.warning("Y derecho: Min debe ser menor que Max.")

            if usar_y2_efectivo:
                fig = make_subplots(specs=[[{"secondary_y": True}]])
            else:
                fig = go.Figure()

            for v in variables_grafico:
                nombre_traza = nombres_legibles[v]

                if grafico_rendimiento_target:
                    if v == "Rendimiento":
                        nombre_traza = "Rendimiento real"
                    elif v == "Productividad_estimada":
                        nombre_traza = "Rendimiento estimado"

                modo_traza = "lines+markers"

                if unidad.startswith("Polimer") and "Producto" in df_f.columns:
                    producto_hover = (
                        df_f["Producto"]
                        .map(normalizar_producto_operativo)
                        .astype("string")
                        .fillna("")
                    )

                    if filtro_producto_activo:
                        x_plot, y_plot = preparar_serie_producto_filtrado_para_grafico(
                            df_f,
                            v,
                            agrupacion,
                        )
                    else:
                        x_plot = df_f["Fecha_y_hora"]
                        y_plot = df_f[v]

                    traza = go.Scatter(
                        x=x_plot,
                        y=y_plot,
                        name=nombre_traza,
                        mode=modo_traza,
                        connectgaps=False,
                        customdata=producto_hover,
                        hovertemplate=(
                            "%{fullData.name}<br>"
                            "Fecha: %{x}<br>"
                            "Valor: %{y}<br>"
                            "Producto: %{customdata}"
                            "<extra></extra>"
                        ),
                    )
                else:
                    traza = go.Scatter(
                        x=df_f["Fecha_y_hora"],
                        y=df_f[v],
                        name=nombre_traza,
                        mode=modo_traza,
                        connectgaps=False,
                    )

                if usar_y2_efectivo:
                    fig.add_trace(traza, secondary_y=(v in variables_y2))
                else:
                    fig.add_trace(traza)

            fig.update_layout(
                height=660,
                hovermode="x unified",
                margin=dict(
                    t=35,
                    b=75,
                    l=70,
                    r=110 if usar_y2_efectivo else 50,
                ),
                legend=dict(
                    orientation="h",
                    y=-0.13,
                    x=0.5,
                    xanchor="center",
                    yanchor="top",
                ),
            )

            if usar_y2_efectivo:
                titulo_y1 = " / ".join([nombres_legibles[v] for v in variables_y1[:2]]) if variables_y1 else "Eje izquierdo"
                titulo_y2 = " / ".join([nombres_legibles[v] for v in variables_y2[:2]]) if variables_y2 else "Eje derecho"

                fig.update_yaxes(title_text=titulo_y1, secondary_y=False, showgrid=True, fixedrange=False)
                fig.update_yaxes(title_text=titulo_y2, secondary_y=True, showgrid=False, fixedrange=False)

                if escala_y1_manual is not None:
                    fig.update_yaxes(range=list(escala_y1_manual), secondary_y=False)

                if escala_y2_manual is not None:
                    fig.update_yaxes(range=list(escala_y2_manual), secondary_y=True)
            else:
                fig.update_yaxes(fixedrange=False)
                if escala_y1_manual is not None:
                    fig.update_yaxes(range=list(escala_y1_manual))

            if unidad.startswith("Polimer") and "Producto" in df_f.columns:
                if mostrar_cambios_producto:
                    _df_marcas = (
                        df_producto_marcas_visible
                        if isinstance(df_producto_marcas_visible, pd.DataFrame)
                        and len(df_producto_marcas_visible) > 0
                        else df_f
                    )
                    fig = agregar_marcas_producto_a_figura(
                        fig,
                        _df_marcas,
                        en_subplots=False,
                        mostrar_etiquetas=True,
                    )
                    _n_marcas = len(obtener_cambios_producto_para_marcas(_df_marcas))
                    if filtro_producto_activo:
                        st.caption(
                            f"Filtro por producto activo: se muestran {_n_marcas} marcas "
                            "calculadas con la línea de tiempo completa de productos del período visible."
                        )
                    else:
                        st.caption(
                            f"Las líneas verticales punteadas muestran {_n_marcas} cambios de producto/campaña del período visible, con el nombre del grado arriba de la gráfica."
                        )
                elif filtro_producto_activo:
                    st.caption(
                        "Filtro por producto activo: se unen solo puntos consecutivos "
                        "y se corta la línea cuando hay saltos entre campañas."
                    )

            st.caption(
                "En modo combinado, las variables del eje derecho y los rangos Y se ajustan desde el panel izquierdo."
            )
            st.plotly_chart(fig, width="stretch")
        else:
            n = len(variables_grafico)
            altura_por_grafico = 300
            altura_total = altura_por_grafico * n
            spacing = min(0.1, 80 / altura_total) if n > 1 else 0

            fig = make_subplots(
                rows=n,
                cols=1,
                shared_xaxes=True,
                subplot_titles=[nombres_legibles[v] for v in variables_grafico],
                vertical_spacing=spacing,
            )

            modo_traza = "lines+markers"

            for i, v in enumerate(variables_grafico, 1):
                if unidad.startswith("Polimer") and "Producto" in df_f.columns:
                    producto_hover = (
                        df_f["Producto"]
                        .map(normalizar_producto_operativo)
                        .astype("string")
                        .fillna("")
                    )

                    if filtro_producto_activo:
                        x_plot, y_plot = preparar_serie_producto_filtrado_para_grafico(
                            df_f,
                            v,
                            agrupacion,
                        )
                    else:
                        x_plot = df_f["Fecha_y_hora"]
                        y_plot = df_f[v]

                    fig.add_trace(
                        go.Scatter(
                            x=x_plot,
                            y=y_plot,
                            mode=modo_traza,
                            connectgaps=False,
                            name=nombres_legibles[v],
                            showlegend=False,
                            customdata=producto_hover,
                            hovertemplate=(
                                "%{fullData.name}<br>"
                                "Fecha: %{x}<br>"
                                "Valor: %{y}<br>"
                                "Producto: %{customdata}"
                                "<extra></extra>"
                            ),
                        ),
                        row=i,
                        col=1,
                    )
                else:
                    fig.add_trace(
                        go.Scatter(
                            x=df_f["Fecha_y_hora"],
                            y=df_f[v],
                            mode=modo_traza,
                            connectgaps=False,
                            name=nombres_legibles[v],
                            showlegend=False,
                        ),
                        row=i,
                        col=1,
                    )

            fig.update_layout(
                height=altura_total,
                margin=dict(t=40, b=40, l=20, r=20),
                hovermode="x unified",
            )
            fig.update_xaxes(showticklabels=True)

            if unidad.startswith("Polimer") and "Producto" in df_f.columns:
                if mostrar_cambios_producto:
                    _df_marcas = (
                        df_producto_marcas_visible
                        if isinstance(df_producto_marcas_visible, pd.DataFrame)
                        and len(df_producto_marcas_visible) > 0
                        else df_f
                    )
                    fig = agregar_marcas_producto_a_figura(
                        fig,
                        _df_marcas,
                        en_subplots=True,
                        mostrar_etiquetas=True,
                    )
                    _n_marcas = len(obtener_cambios_producto_para_marcas(_df_marcas))
                    if filtro_producto_activo:
                        st.caption(
                            f"Filtro por producto activo: se muestran {_n_marcas} marcas "
                            "calculadas con la línea de tiempo completa de productos del período visible."
                        )
                    else:
                        st.caption(
                            f"Las líneas verticales punteadas muestran {_n_marcas} cambios de producto/campaña del período visible, con el nombre del grado arriba de la gráfica."
                        )
                elif filtro_producto_activo:
                    st.caption(
                        "Filtro por producto activo: se unen solo puntos consecutivos "
                        "y se corta la línea cuando hay saltos entre campañas."
                    )

            st.plotly_chart(fig, width="stretch")


# ==============================================================================
# TAB 2 - CORRELACIONES
# ==============================================================================

with tab2:
    st.subheader("Matriz de correlacion")

    if len(variables_sel) < 2:
        st.info("Selecciona al menos 2 variables para ver correlaciones.")
    else:
        col_a, col_b = st.columns([1, 1])
        with col_a:
            metodo_corr = st.radio(
                "Metodo de correlacion:",
                options=["pearson", "spearman"],
                horizontal=True,
                index=0,
            )
        with col_b:
            min_puntos = st.number_input(
                "Minimo de puntos validos por par:",
                min_value=3,
                max_value=max(3, len(df_f)),
                value=min(10, max(3, len(df_f))),
                step=1,
            )

        corr = df_f[variables_sel].corr(method=metodo_corr, min_periods=int(min_puntos)).round(3)

        fig_corr = go.Figure(data=go.Heatmap(
            z=corr.values,
            x=[nombres_legibles[v] for v in variables_sel],
            y=[nombres_legibles[v] for v in variables_sel],
            colorscale="RdBu",
            zmid=0,
            zmin=-1,
            zmax=1,
            text=corr.values.round(2),
            texttemplate="%{text}",
            colorbar=dict(title="r"),
        ))
        fig_corr.update_layout(height=650, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_corr, width="stretch")

        st.subheader("Ranking de correlaciones")
        ranking = top_correlaciones(df_f, variables_sel, nombres_legibles, metodo_corr, int(min_puntos))
        if ranking.empty:
            st.info("No hay pares con suficientes puntos validos para calcular correlacion.")
        else:
            st.dataframe(
                ranking.head(30),
                width="stretch",
                hide_index=True,
            )

        st.subheader("Dispersion entre dos variables")
        col1, col2 = st.columns(2)
        with col1:
            var_x = st.selectbox(
                "Eje X:",
                options=variables_sel,
                format_func=lambda x: nombres_legibles[x],
                key="scatter_x",
            )
        with col2:
            var_y = st.selectbox(
                "Eje Y:",
                options=variables_sel,
                index=min(1, len(variables_sel) - 1),
                format_func=lambda x: nombres_legibles[x],
                key="scatter_y",
            )

        mostrar_tendencia = st.checkbox("Agregar tendencia lineal OLS", value=True)

        if var_x == var_y:
            st.info("Selecciona dos variables distintas.")
        else:
            df_sc = df_f[[var_x, var_y]].dropna()
            if len(df_sc) < 3:
                st.warning("No hay suficientes puntos validos para el scatter.")
            else:
                try:
                    fig_sc = px.scatter(
                        df_sc,
                        x=var_x,
                        y=var_y,
                        trendline="ols" if mostrar_tendencia else None,
                        labels={var_x: nombres_legibles[var_x], var_y: nombres_legibles[var_y]},
                        title=f"{nombres_legibles[var_x]} vs {nombres_legibles[var_y]}",
                    )
                    st.plotly_chart(fig_sc, width="stretch")
                except Exception as exc:
                    st.warning(
                        "No se pudo calcular la tendencia OLS. "
                        "Verifica que statsmodels este instalado. "
                        f"Detalle: {exc}"
                    )
                    fig_sc = px.scatter(
                        df_sc,
                        x=var_x,
                        y=var_y,
                        labels={var_x: nombres_legibles[var_x], var_y: nombres_legibles[var_y]},
                        title=f"{nombres_legibles[var_x]} vs {nombres_legibles[var_y]}",
                    )
                    st.plotly_chart(fig_sc, width="stretch")


# ==============================================================================
# TAB 3 - DESFASE TEMPORAL
# ==============================================================================

with tab3:
    st.subheader("Analisis de desfase temporal")
    st.caption(
        "Se calcula la correlacion entre X(t) e Y(t + lag). "
        "Lag positivo significa que X precede a Y. Lag negativo significa que Y precede a X."
    )

    col1, col2 = st.columns(2)
    with col1:
        lag_x = st.selectbox(
            "Variable X - posible causa o variable que precede:",
            options=todas_variables,
            index=todas_variables.index(variables_sel[0]) if variables_sel else 0,
            format_func=lambda x: nombres_legibles[x],
            key="lag_x",
        )
    with col2:
        default_y_index = todas_variables.index(variables_sel[1]) if len(variables_sel) > 1 else min(1, len(todas_variables) - 1)
        lag_y = st.selectbox(
            "Variable Y - posible efecto o variable que responde:",
            options=todas_variables,
            index=default_y_index,
            format_func=lambda x: nombres_legibles[x],
            key="lag_y",
        )

    col3, col4, col5 = st.columns(3)
    with col3:
        metodo_lag = st.radio(
            "Metodo:",
            options=["pearson", "spearman"],
            horizontal=True,
            key="metodo_lag",
        )
    with col4:
        max_lag = st.slider("Maximo lag a evaluar:", min_value=1, max_value=50, value=10, step=1)
    with col5:
        min_puntos_lag = st.number_input(
            "Minimo de puntos validos:",
            min_value=3,
            max_value=max(3, len(df_f)),
            value=min(10, max(3, len(df_f))),
            step=1,
            key="min_puntos_lag",
        )

    if lag_x == lag_y:
        st.info("Selecciona dos variables distintas.")
    else:
        df_lags = calcular_lags(df_f, lag_x, lag_y, metodo_lag, int(max_lag), int(min_puntos_lag))
        df_lags_validos = df_lags.dropna(subset=["Correlacion"])

        if df_lags_validos.empty:
            st.warning("No hay suficientes puntos validos para calcular lags.")
        else:
            mejor = df_lags_validos.sort_values("abs_r", ascending=False).iloc[0]
            lag_optimo = int(mejor["Lag"])
            r_optimo = float(mejor["Correlacion"])
            puntos_optimos = int(mejor["Puntos validos"])

            if lag_optimo > 0:
                interpretacion = f"{nombres_legibles[lag_x]} precede a {nombres_legibles[lag_y]} por {lag_optimo} intervalo(s)."
            elif lag_optimo < 0:
                interpretacion = f"{nombres_legibles[lag_y]} precede a {nombres_legibles[lag_x]} por {abs(lag_optimo)} intervalo(s)."
            else:
                interpretacion = "La mayor correlacion se da sin desfase temporal."

            c1, c2, c3 = st.columns(3)
            c1.metric("Lag optimo", lag_optimo)
            c2.metric("r", f"{r_optimo:.3f}")
            c3.metric("Puntos validos", puntos_optimos)
            st.info(interpretacion)

            fig_lag = go.Figure()
            fig_lag.add_trace(go.Scatter(
                x=df_lags["Lag"],
                y=df_lags["Correlacion"],
                mode="lines+markers",
                name="Correlacion",
            ))
            fig_lag.add_hline(y=0, line_dash="dash")
            fig_lag.update_layout(
                title=f"Desfase: {nombres_legibles[lag_x]} vs {nombres_legibles[lag_y]}",
                xaxis_title="Lag",
                yaxis_title="Correlacion",
                height=500,
                margin=dict(t=50, b=40, l=20, r=20),
            )
            st.plotly_chart(fig_lag, width="stretch")

            st.dataframe(
                df_lags[["Lag", "Correlacion", "Puntos validos"]].round(3),
                width="stretch",
                hide_index=True,
            )


# ==============================================================================
# TAB 4 - DATOS FILTRADOS
# ==============================================================================

with tab4:
    st.subheader("Datos filtrados")

    if variables_sel:
        columnas_mostrar = ["Fecha_y_hora"] + variables_sel
    else:
        columnas_mostrar = ["Fecha_y_hora"] + todas_variables

    # En Polimerización, agregar siempre el producto/grado en la tabla,
    # aunque no esté seleccionado como variable numérica.
    if unidad.startswith("Polimer") and "Producto" in df_f.columns:
        if "Producto" not in columnas_mostrar:
            columnas_mostrar.insert(1, "Producto")

    columnas_mostrar = [c for c in columnas_mostrar if c in df_f.columns]

    df_tabla = df_f[columnas_mostrar].copy()

    if unidad.startswith("Polimer") and "Producto" in df_tabla.columns:
        df_tabla["Producto"] = df_tabla["Producto"].map(normalizar_producto_operativo)

    st.dataframe(df_tabla, width="stretch")

    csv = df_tabla.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Descargar datos filtrados en CSV",
        data=csv,
        file_name=f"datos_filtrados_{unidad}_{agrupacion}.csv".replace(" ", "_"),
        mime="text/csv",
    )


