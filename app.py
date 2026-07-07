"""
App Streamlit - Reactor R-100
Planta Mendoza - NOVOLEN (Petrocuyo)
Fuente de datos: Excel cargado manualmente / Reactor R-100.xlsx (hoja IP21)

Mejoras incluidas:
- Fuente de datos seleccionable: cargar Excel manualmente o usar Excel de la carpeta.
- Actualizacion desde Excel por fecha de modificacion.
- Variables calculadas con panel simple: Nombre + Variable A + Operacion + Variable B.
- Agrupacion: sin agrupacion, horario, diario, semanal y mensual.
- Correlaciones Pearson/Spearman, ranking de correlaciones y scatter.
- Analisis de desfase temporal entre variables.
- Filtros por producto, fecha y variables.
- Descarga de datos filtrados a CSV.
"""

import glob
import io
import os
from itertools import combinations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

st.set_page_config(page_title="Reactor R-100 - Mendoza", layout="wide")
st.title("Reactor R-100 - Analisis de Variables")

CARPETA = os.path.dirname(os.path.abspath(__file__))

VARIABLES = {
    "Presion_R100": "Presion R-100 [PRS1009.MS]",
    "Nivel1_R100": "Nivel 1 R-100 [LT1004-1.MS]",
    "Nivel2_R100": "Nivel 2 R-100 [LT1004-2.MS]",
    "Temperatura_R100": "Temperatura R-100 [TRCS1006.MS]",
    "Caudal_fresco": "Caudal fresco [FRC1004.MS]",
    "Caudal_etileno": "Caudal etileno [FC1013.MS]",
    "Recirculacion_domos": "Recirculacion domos [FC1012.MS]",
    "Recirculacion_fondo": "Recirculacion fondo [FC1014.MS]",
    "Catalizador_R106": "Catalizador R106 [F10020.MS]",
    "Catalizador_R206": "Catalizador R206 [F10021.MS]",
    "Caudal_H2": "Caudal H2 [F1006TG.MS]",
    "Consumo_agitador": "Consumo agitador [ERSRRM100.MS]",
    "Conversion": "Conversion [XM1]",
    "Nivel_W100": "Nivel W100 [LSLL1001.MS]",
    "Temperatura_W100": "Temperatura reciclo W100 [T1005.MS]",
    "Caudal_reciclo": "Caudal reciclo [FR1003.MS]",
    "DP_CCF_R100": "DP entre CCF y R100 [PD10020.MS]",
    "Caudal_silano": "Caudal silano [F1030B.MS]",
    "Caudal_TEA_A": "Caudal TEA A [FIRCSA1020A.MS]",
    "Caudal_TEA_B": "Caudal TEA B [FIRCSA1020B.MS]",
    "MFI_Polvo": "MFI Polvo [MFI/1200]",
    "MFI_Pellets": "MFI Pellets [MFI_S140.MS]",
    "Nivel_R106": "Nivel R106 [L10004.MS]",
    "Temp_ambiente": "Temperatura ambiente [T0340.MS]",
}

DEFAULT_VARS = [
    "Presion_R100",
    "Nivel1_R100",
    "Nivel2_R100",
    "Temperatura_R100",
    "Caudal_fresco",
    "Conversion",
]

CAUDALES = {
    "Caudal_fresco",
    "Caudal_etileno",
    "Recirculacion_domos",
    "Recirculacion_fondo",
    "Catalizador_R106",
    "Catalizador_R206",
    "Caudal_H2",
    "Caudal_reciclo",
    "Caudal_silano",
    "Caudal_TEA_A",
    "Caudal_TEA_B",
}

OPERACIONES_CALC = [
    "+",
    "-",
    "×",
    "÷",
    "A/B × 100",
    "(A-B)/B × 100",
    "|A - B|",
    "Promedio",
]


def nombre_corto(var, variables_todas):
    return variables_todas.get(var, var).split(" [")[0]


def buscar_archivo_excel():
    archivos = [
        a for a in glob.glob(os.path.join(CARPETA, "Reactor*.xlsx"))
        if not os.path.basename(a).startswith("~$")
    ]
    if not archivos:
        return None
    return max(archivos, key=os.path.getmtime)


def obtener_mtime_archivo(ruta):
    if not ruta or not os.path.exists(ruta):
        return None
    return os.path.getmtime(ruta)


def preparar_dataframe(df_raw):
    nombres = ["descartar", "Fecha_y_hora"] + list(VARIABLES.keys()) + ["Producto"]

    if df_raw.shape[1] > len(nombres):
        extras = [f"Extra_{i}" for i in range(1, df_raw.shape[1] - len(nombres) + 1)]
        nombres = nombres + extras

    df = df_raw.copy()
    df.columns = nombres[:df.shape[1]]

    if "descartar" in df.columns:
        df = df.drop(columns=["descartar"])

    if "Fecha_y_hora" not in df.columns:
        st.error("No se encontro la columna de fecha y hora. Revisar estructura del Excel.")
        st.stop()

    df["Fecha_y_hora"] = pd.to_datetime(df["Fecha_y_hora"], errors="coerce")
    df = df.dropna(subset=["Fecha_y_hora"])

    for col in df.columns:
        if col not in ["Fecha_y_hora", "Producto"]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "Producto" in df.columns:
        df["Producto"] = df["Producto"].astype("string")

    df = df.sort_values("Fecha_y_hora").reset_index(drop=True)
    return df


@st.cache_data(show_spinner="Cargando Excel de la carpeta...")
def cargar_datos_desde_carpeta(archivo, excel_mtime):
    _ = excel_mtime

    if not archivo or not os.path.exists(archivo):
        st.error("No se encontro ningun archivo Reactor*.xlsx en la carpeta de la app.")
        st.stop()

    try:
        df_raw = pd.read_excel(archivo, sheet_name="IP21", header=None, skiprows=6)
    except ValueError:
        st.error("No existe la hoja 'IP21' en el archivo Excel.")
        st.stop()
    except PermissionError:
        st.error("No se puede leer el Excel. Cerralo en Excel y volve a correr la app.")
        st.stop()
    except Exception as exc:
        st.error(f"Error leyendo el Excel: {exc}")
        st.stop()

    return preparar_dataframe(df_raw)


@st.cache_data(show_spinner="Cargando Excel subido...")
def cargar_datos_desde_upload(file_bytes):
    try:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name="IP21", header=None, skiprows=6)
    except ValueError:
        st.error("No existe la hoja 'IP21' en el archivo cargado.")
        st.stop()
    except Exception as exc:
        st.error(f"Error leyendo el Excel cargado: {exc}")
        st.stop()

    return preparar_dataframe(df_raw)

def calcular_variable_personalizada(df, definicion):
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
        elif operacion == "(A-B)/B × 100":
            resultado = ((a - b) / b) * 100
        elif operacion == "|A - B|":
            resultado = (a - b).abs()
        elif operacion == "Promedio":
            resultado = (a + b) / 2
        else:
            resultado = pd.Series(np.nan, index=df.index)

    resultado = pd.to_numeric(resultado, errors="coerce")
    return resultado.replace([np.inf, -np.inf], np.nan)


def texto_formula(definicion):
    a = definicion["var_a"]
    b = definicion["var_b"]
    op = definicion["operacion"]

    if op == "A/B × 100":
        return f"({a} / {b}) × 100"
    if op == "(A-B)/B × 100":
        return f"(({a} - {b}) / {b}) × 100"
    if op == "|A - B|":
        return f"|{a} - {b}|"
    if op == "Promedio":
        return f"({a} + {b}) / 2"
    return f"{a} {op} {b}"


def aplicar_variables_calculadas_guardadas(df, variables_todas, todas_variables, key_defs):
    for definicion in st.session_state.get(key_defs, []):
        clave = definicion["clave"]
        try:
            df[clave] = calcular_variable_personalizada(df, definicion)
            variables_todas[clave] = definicion["nombre"] + f" [{clave}]"
            if clave not in todas_variables:
                todas_variables.append(clave)
        except Exception:
            pass


def producto_moda(serie):
    s = serie.dropna()
    if len(s) == 0:
        return None
    moda = s.mode()
    return moda.iloc[0] if len(moda) else s.iloc[0]


def aplicar_agrupacion(df, agrupacion, variables_numericas):
    if agrupacion == "Sin agrupación":
        return df[["Fecha_y_hora"] + variables_numericas + (["Producto"] if "Producto" in df.columns else [])].copy()

    freq_map = {
        "Horario": "h",
        "Diario": "D",
        "Semanal": "W",
        "Mensual": "ME",
    }
    freq = freq_map[agrupacion]

    df_num = (
        df.set_index("Fecha_y_hora")[variables_numericas]
        .resample(freq)
        .mean(numeric_only=True)
        .asfreq(freq)
        .reset_index()
    )

    if "Producto" in df.columns:
        df_prod = (
            df.set_index("Fecha_y_hora")["Producto"]
            .resample(freq)
            .agg(producto_moda)
            .reset_index()
        )
        df_num = df_num.merge(df_prod, on="Fecha_y_hora", how="left")

    return df_num


def top_correlaciones(df, variables, variables_todas, metodo, min_puntos):
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
            "Variable 1": nombre_corto(var1, variables_todas),
            "Variable 2": nombre_corto(var2, variables_todas),
            "r": round(float(r), 3),
            "abs(r)": round(abs(float(r)), 3),
            "Puntos validos": int(n),
        })

    if not filas:
        return pd.DataFrame(columns=["Variable 1", "Variable 2", "r", "abs(r)", "Puntos validos"])

    return pd.DataFrame(filas).sort_values("abs(r)", ascending=False).reset_index(drop=True)


def calcular_lags(df, var_x, var_y, metodo, max_lag, min_puntos):
    filas = []
    base = df[["Fecha_y_hora", var_x, var_y]].copy()

    for lag in range(-max_lag, max_lag + 1):
        y_desfasada = base[var_y].shift(-lag)
        tmp = pd.concat([base[var_x], y_desfasada], axis=1)
        tmp.columns = [var_x, var_y]
        tmp = tmp.dropna()
        n = len(tmp)
        r = np.nan if n < min_puntos else tmp[var_x].corr(tmp[var_y], method=metodo)
        filas.append({
            "Lag": lag,
            "Correlacion": r,
            "abs_r": abs(r) if pd.notna(r) else np.nan,
            "Puntos validos": n,
        })

    return pd.DataFrame(filas)


st.sidebar.header("Fuente de datos")

fuente_datos = st.sidebar.radio(
    "Seleccionar fuente:",
    options=["Cargar Excel manualmente", "Usar Excel de la carpeta"],
    index=0,
)

archivo_nombre = None

if fuente_datos == "Cargar Excel manualmente":
    archivo_subido = st.sidebar.file_uploader(
        "Cargar archivo Excel",
        type=["xlsx"],
        accept_multiple_files=False,
    )

    if archivo_subido is None:
        st.info("Cargá un archivo Excel para comenzar el análisis.")
        st.stop()

    archivo_nombre = archivo_subido.name
    st.sidebar.caption(f"Archivo cargado: {archivo_nombre}")

    if st.sidebar.button("Recargar archivo cargado", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    df = cargar_datos_desde_upload(archivo_subido.getvalue())

else:
    archivo_excel = buscar_archivo_excel()
    excel_mtime = obtener_mtime_archivo(archivo_excel)

    if archivo_excel:
        archivo_nombre = os.path.basename(archivo_excel)
        st.sidebar.caption(f"Archivo: {archivo_nombre}")
    else:
        st.sidebar.warning("No se encontro archivo Reactor*.xlsx en la carpeta.")

    if st.sidebar.button("Actualizar datos desde Excel", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    df = cargar_datos_desde_carpeta(archivo_excel, excel_mtime)

todas_variables_base = [v for v in VARIABLES.keys() if v in df.columns]
variables_todas = {k: v for k, v in VARIABLES.items() if k in df.columns}
todas_variables = list(variables_todas.keys())

default_vars = [v for v in DEFAULT_VARS if v in todas_variables]
if not default_vars:
    default_vars = todas_variables[:min(6, len(todas_variables))]

st.sidebar.caption(
    f"Registros crudos: {len(df):,} | "
    f"{df['Fecha_y_hora'].min().strftime('%d/%m/%Y')} a "
    f"{df['Fecha_y_hora'].max().strftime('%d/%m/%Y')}"
)

key_defs = "variables_calculadas_R100"
if key_defs not in st.session_state:
    st.session_state[key_defs] = []

aplicar_variables_calculadas_guardadas(df, variables_todas, todas_variables, key_defs)

st.sidebar.markdown("---")
st.sidebar.subheader("Variables calculadas")

with st.sidebar.expander("＋ Agregar variable calculada", expanded=False):
    st.caption("Armá una variable nueva combinando dos variables existentes.")

    nombre_nueva = st.text_input(
        "Nombre",
        value="",
        placeholder="Ejemplo: Relación catalizador/fresco",
        key="calc_nombre_nueva_R100",
    )

    opciones_calc = list(variables_todas.keys())

    var_a_nueva = st.selectbox(
        "Variable A",
        options=opciones_calc,
        format_func=lambda x: nombre_corto(x, variables_todas),
        key="calc_var_a_nueva_R100",
    )

    operacion_nueva = st.selectbox(
        "Operación",
        options=OPERACIONES_CALC,
        key="calc_operacion_nueva_R100",
    )

    var_b_nueva = st.selectbox(
        "Variable B",
        options=opciones_calc,
        index=min(1, len(opciones_calc) - 1),
        format_func=lambda x: nombre_corto(x, variables_todas),
        key="calc_var_b_nueva_R100",
    )

    definicion_preview = {
        "var_a": var_a_nueva,
        "operacion": operacion_nueva,
        "var_b": var_b_nueva,
    }

    st.caption("Vista previa")
    st.code(texto_formula(definicion_preview))

    if st.button("Agregar", use_container_width=True, key="btn_agregar_calc_R100"):
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
    with st.sidebar.expander("Variables creadas", expanded=False):
        for idx, definicion in enumerate(st.session_state[key_defs]):
            st.markdown(f"**{definicion['nombre']}**")
            st.code(f"{definicion['clave']} = {texto_formula(definicion)}")
            clave = definicion["clave"]
            if clave in df.columns:
                st.caption(f"Puntos válidos: {int(df[clave].notna().sum()):,}")
            if st.button("Eliminar", key=f"delete_calc_R100_{idx}", use_container_width=True):
                st.session_state[key_defs].pop(idx)
                st.rerun()
            st.markdown("---")

st.sidebar.markdown("---")
st.sidebar.header("Filtros")

fecha_min = df["Fecha_y_hora"].min().date()
fecha_max = df["Fecha_y_hora"].max().date()

desde = st.sidebar.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max)
hasta = st.sidebar.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max)

if desde > hasta:
    st.sidebar.error("La fecha 'Desde' no puede ser posterior a 'Hasta'.")
    st.stop()

productos_sel = []
if "Producto" in df.columns:
    productos_disp = sorted([
        str(p) for p in df["Producto"].dropna().unique()
        if str(p) not in ["Sin MFI", "nan", "<NA>", "Verificar"]
    ])
    productos_sel = st.sidebar.multiselect("Producto:", options=productos_disp, default=[])

st.sidebar.markdown("---")
st.sidebar.subheader("Variables a graficar")
variables_sel = st.sidebar.multiselect(
    "Selecciona variables:",
    options=todas_variables,
    default=default_vars,
    format_func=lambda x: nombre_corto(x, variables_todas),
)

st.sidebar.markdown("---")
st.sidebar.subheader("Modo de grafico")
modo_grafico = st.sidebar.radio("Modo:", options=["Separados", "Combinados"], index=0)

st.sidebar.markdown("---")
st.sidebar.subheader("Agrupacion temporal")
agrupacion = st.sidebar.selectbox(
    "Frecuencia:",
    options=["Sin agrupación", "Horario", "Diario", "Semanal", "Mensual"],
    index=2,
)

st.sidebar.markdown("---")
st.sidebar.subheader("Limpieza de datos")
filtrar_caudales_negativos = st.sidebar.checkbox(
    "Descartar caudales negativos en variables seleccionadas/filtradas",
    value=True,
)

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros por variable")
vars_filtro = st.sidebar.multiselect(
    "Variables a filtrar:",
    options=todas_variables,
    default=[],
    format_func=lambda x: nombre_corto(x, variables_todas),
)

mask_raw = (
    (df["Fecha_y_hora"].dt.date >= desde)
    & (df["Fecha_y_hora"].dt.date <= hasta)
)
if productos_sel and "Producto" in df.columns:
    mask_raw = mask_raw & df["Producto"].isin(productos_sel)

df_filtrado = df[mask_raw].copy()

if df_filtrado.empty:
    st.warning("Sin datos para ese filtro inicial.")
    st.stop()

df_agrup = aplicar_agrupacion(df_filtrado, agrupacion, todas_variables)

mask = pd.Series([True] * len(df_agrup), index=df_agrup.index)

if filtrar_caudales_negativos:
    vars_relevantes = set(variables_sel) | set(vars_filtro)
    for var_caudal in CAUDALES:
        if var_caudal in vars_relevantes and var_caudal in df_agrup.columns:
            mask = mask & ((df_agrup[var_caudal] >= 0) | df_agrup[var_caudal].isna())

for var in vars_filtro:
    if var not in df_agrup.columns:
        continue

    serie = df_agrup[var].dropna()
    if len(serie) == 0:
        continue

    vmin = float(serie.min())
    vmax = float(serie.max())
    if vmin == vmax:
        continue

    min_key = f"min_R100_{agrupacion}_{var}"
    max_key = f"max_R100_{agrupacion}_{var}"

    if min_key not in st.session_state or not (vmin <= float(st.session_state[min_key]) <= vmax):
        st.session_state[min_key] = vmin
    if max_key not in st.session_state or not (vmin <= float(st.session_state[max_key]) <= vmax):
        st.session_state[max_key] = vmax

    with st.sidebar.container(border=True):
        st.markdown(f"**{nombre_corto(var, variables_todas)}**")
        st.caption(f"Rango disponible: {vmin:.3f} a {vmax:.3f}")

        if st.button("Reset", key=f"reset_R100_{agrupacion}_{var}", use_container_width=True):
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
        st.sidebar.error(f"Filtro invalido en {nombre_corto(var, variables_todas)}: minimo mayor que maximo.")
        st.stop()

    if r_min > vmin or r_max < vmax:
        mask = mask & df_agrup[var].between(r_min, r_max)

df_f = df_agrup[mask].copy()

prod_txt = ", ".join(productos_sel) if productos_sel else "Todos"
st.markdown(
    f"**Reactor:** R-100 Mendoza | "
    f"**Fuente:** {fuente_datos} | "
    f"**Archivo:** {archivo_nombre} | "
    f"**Periodo:** {desde.strftime('%d/%m/%Y')} a {hasta.strftime('%d/%m/%Y')} | "
    f"**Agrupacion:** {agrupacion} | "
    f"**Registros filtrados:** {len(df_f):,} | "
    f"**Producto:** {prod_txt}"
)

if df_f.empty:
    st.warning("Sin datos para ese filtro.")
    st.stop()

if not variables_sel:
    st.warning("Selecciona al menos una variable en el panel lateral.")
    st.stop()

with st.expander("Resumen rapido de datos filtrados", expanded=False):
    c1, c2, c3 = st.columns(3)
    c1.metric("Registros", f"{len(df_f):,}")
    c2.metric("Desde", df_f["Fecha_y_hora"].min().strftime("%d/%m/%Y %H:%M"))
    c3.metric("Hasta", df_f["Fecha_y_hora"].max().strftime("%d/%m/%Y %H:%M"))

    resumen = df_f[variables_sel].describe().T
    resumen.index = [nombre_corto(v, variables_todas) for v in resumen.index]
    st.dataframe(resumen.round(3), use_container_width=True)

cols_m = st.columns(min(len(variables_sel), 4))
for i, var in enumerate(variables_sel[:4]):
    with cols_m[i]:
        st.metric(
            label=nombre_corto(var, variables_todas),
            value=f"{df_f[var].mean():.3f}",
            delta=f"σ {df_f[var].std():.3f}",
        )

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs([
    "Series temporales",
    "Correlaciones",
    "Desfase temporal",
    "Datos filtrados",
])

with tab1:
    st.subheader(f"Tendencias ({agrupacion.lower()})")

    if modo_grafico == "Combinados":
        fig = go.Figure()
        for v in variables_sel:
            fig.add_trace(go.Scatter(
                x=df_f["Fecha_y_hora"],
                y=df_f[v],
                name=nombre_corto(v, variables_todas),
                mode="lines+markers",
                connectgaps=False,
            ))
        fig.update_layout(
            height=550,
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=60, b=20, l=20, r=20),
        )
        st.plotly_chart(fig, use_container_width=True)

    else:
        n = len(variables_sel)
        altura_por_grafico = 300
        altura_total = altura_por_grafico * n
        spacing = min(0.1, 80 / altura_total) if n > 1 else 0
        fig = make_subplots(
            rows=n,
            cols=1,
            shared_xaxes=True,
            subplot_titles=[nombre_corto(v, variables_todas) for v in variables_sel],
            vertical_spacing=spacing,
        )
        for i, v in enumerate(variables_sel, 1):
            fig.add_trace(
                go.Scatter(
                    x=df_f["Fecha_y_hora"],
                    y=df_f[v],
                    mode="lines+markers",
                    connectgaps=False,
                    name=nombre_corto(v, variables_todas),
                    showlegend=False,
                ),
                row=i,
                col=1,
            )
        fig.update_layout(
            height=altura_total,
            hovermode="x unified",
            margin=dict(t=40, b=40, l=20, r=20),
        )
        fig.update_xaxes(showticklabels=True)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("Tabla resumen")
    resumen_tabla = pd.DataFrame({
        "Variable": [nombre_corto(v, variables_todas) for v in variables_sel],
        "Media": [round(df_f[v].mean(), 3) for v in variables_sel],
        "Desv.Est": [round(df_f[v].std(), 3) for v in variables_sel],
        "Minimo": [round(df_f[v].min(), 3) for v in variables_sel],
        "Maximo": [round(df_f[v].max(), 3) for v in variables_sel],
        "Puntos validos": [int(df_f[v].notna().sum()) for v in variables_sel],
    })
    st.dataframe(resumen_tabla, use_container_width=True, hide_index=True)

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
            x=[nombre_corto(v, variables_todas) for v in variables_sel],
            y=[nombre_corto(v, variables_todas) for v in variables_sel],
            colorscale="RdBu",
            zmid=0,
            zmin=-1,
            zmax=1,
            text=corr.values.round(2),
            texttemplate="%{text}",
            colorbar=dict(title="r"),
        ))
        fig_corr.update_layout(height=650, margin=dict(t=20, b=20, l=20, r=20))
        st.plotly_chart(fig_corr, use_container_width=True)

        st.subheader("Ranking de correlaciones")
        ranking = top_correlaciones(df_f, variables_sel, variables_todas, metodo_corr, int(min_puntos))
        if ranking.empty:
            st.info("No hay pares con suficientes puntos validos para calcular correlacion.")
        else:
            st.dataframe(ranking.head(30), use_container_width=True, hide_index=True)

        st.subheader("Dispersion entre dos variables")
        col1, col2, col3 = st.columns([3, 3, 2])
        with col1:
            var_x = st.selectbox(
                "Eje X:",
                options=variables_sel,
                format_func=lambda x: nombre_corto(x, variables_todas),
                key="scatter_x_R100",
            )
        with col2:
            var_y = st.selectbox(
                "Eje Y:",
                options=variables_sel,
                index=min(1, len(variables_sel) - 1),
                format_func=lambda x: nombre_corto(x, variables_todas),
                key="scatter_y_R100",
            )
        with col3:
            mostrar_tendencia = st.checkbox("Linea de tendencia", value=True)
            color_producto = st.checkbox("Color por producto", value=False)

        if var_x == var_y:
            st.info("Selecciona dos variables distintas.")
        else:
            cols_ejes = st.columns(4)
            df_sc = df_f[["Fecha_y_hora", var_x, var_y] + (["Producto"] if "Producto" in df_f.columns else [])].dropna(subset=[var_x, var_y])

            if df_sc.empty:
                st.warning("No hay suficientes puntos validos para el scatter.")
            else:
                with cols_ejes[0]:
                    x_min = st.number_input("X minimo", value=float(df_sc[var_x].min()), step=0.1, format="%.3f")
                with cols_ejes[1]:
                    x_max = st.number_input("X maximo", value=float(df_sc[var_x].max()), step=0.1, format="%.3f")
                with cols_ejes[2]:
                    y_min = st.number_input("Y minimo", value=float(df_sc[var_y].min()), step=0.1, format="%.3f")
                with cols_ejes[3]:
                    y_max = st.number_input("Y maximo", value=float(df_sc[var_y].max()), step=0.1, format="%.3f")

                color_col = "Producto" if color_producto and "Producto" in df_sc.columns else None
                trendline = "ols" if mostrar_tendencia else None

                try:
                    fig_sc = px.scatter(
                        df_sc,
                        x=var_x,
                        y=var_y,
                        color=color_col,
                        trendline=trendline,
                        labels={
                            var_x: nombre_corto(var_x, variables_todas),
                            var_y: nombre_corto(var_y, variables_todas),
                        },
                        title=f"{nombre_corto(var_x, variables_todas)} vs {nombre_corto(var_y, variables_todas)}",
                    )
                except Exception as exc:
                    st.warning(f"No se pudo calcular la tendencia OLS. Se muestra scatter sin tendencia. Detalle: {exc}")
                    fig_sc = px.scatter(
                        df_sc,
                        x=var_x,
                        y=var_y,
                        color=color_col,
                        labels={
                            var_x: nombre_corto(var_x, variables_todas),
                            var_y: nombre_corto(var_y, variables_todas),
                        },
                        title=f"{nombre_corto(var_x, variables_todas)} vs {nombre_corto(var_y, variables_todas)}",
                    )

                if x_min < x_max:
                    fig_sc.update_xaxes(range=[x_min, x_max])
                if y_min < y_max:
                    fig_sc.update_yaxes(range=[y_min, y_max])

                st.plotly_chart(fig_sc, use_container_width=True)

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
            format_func=lambda x: nombre_corto(x, variables_todas),
            key="lag_x_R100",
        )
    with col2:
        default_y_index = (
            todas_variables.index(variables_sel[1])
            if len(variables_sel) > 1
            else min(1, len(todas_variables) - 1)
        )
        lag_y = st.selectbox(
            "Variable Y - posible efecto o variable que responde:",
            options=todas_variables,
            index=default_y_index,
            format_func=lambda x: nombre_corto(x, variables_todas),
            key="lag_y_R100",
        )

    col3, col4, col5 = st.columns(3)
    with col3:
        metodo_lag = st.radio("Metodo:", options=["pearson", "spearman"], horizontal=True, key="metodo_lag_R100")
    with col4:
        max_lag = st.slider("Maximo lag a evaluar:", min_value=1, max_value=50, value=10, step=1)
    with col5:
        min_puntos_lag = st.number_input(
            "Minimo de puntos validos:",
            min_value=3,
            max_value=max(3, len(df_f)),
            value=min(10, max(3, len(df_f))),
            step=1,
            key="min_puntos_lag_R100",
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
                interpretacion = (
                    f"{nombre_corto(lag_x, variables_todas)} precede a "
                    f"{nombre_corto(lag_y, variables_todas)} por {lag_optimo} intervalo(s)."
                )
            elif lag_optimo < 0:
                interpretacion = (
                    f"{nombre_corto(lag_y, variables_todas)} precede a "
                    f"{nombre_corto(lag_x, variables_todas)} por {abs(lag_optimo)} intervalo(s)."
                )
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
                title=f"Desfase: {nombre_corto(lag_x, variables_todas)} vs {nombre_corto(lag_y, variables_todas)}",
                xaxis_title="Lag",
                yaxis_title="Correlacion",
                height=500,
                margin=dict(t=50, b=40, l=20, r=20),
            )
            st.plotly_chart(fig_lag, use_container_width=True)
            st.dataframe(df_lags[["Lag", "Correlacion", "Puntos validos"]].round(3), use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Datos filtrados")

    columnas_mostrar = ["Fecha_y_hora"] + variables_sel
    if "Producto" in df_f.columns:
        columnas_mostrar.append("Producto")

    st.dataframe(df_f[columnas_mostrar], use_container_width=True)

    csv = df_f[columnas_mostrar].to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Descargar datos filtrados en CSV",
        data=csv,
        file_name=f"datos_filtrados_R100_{agrupacion}.csv".replace(" ", "_"),
        mime="text/csv",
    )
