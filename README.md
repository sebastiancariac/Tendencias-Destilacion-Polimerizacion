# Streamlit - Tendencias Destilación / Polimerización

Backup final de la app de análisis de tendencias para **Destilación / Polimerización LIPP**.

## Archivos

- `app.py`: aplicación principal de Streamlit.
- `README.md`: este instructivo/resumen del backup.

## Requisitos

La app espera encontrar el archivo Excel en la misma carpeta que `app.py` con el nombre:

```text
Tendencias.xlsx
```

El Excel debe contener las hojas:

```text
DESTILACIÓN
POLIMERIZACIÓN
```

## Principales funcionalidades

- Lectura de datos desde Excel.
- Validación de columnas por **TAG IP21**, evitando errores por corrimiento de columnas.
- Selector de unidad: **Destilación** / **Polimerización**.
- Filtros por fecha.
- Filtros por variable.
- Agrupación temporal:
  - Hora
  - Día
  - Semana
  - Mes
- Gráficos separados o combinados.
- Variables calculadas desde la interfaz.
- Correlaciones Pearson / Spearman.
- Análisis de desfase temporal.
- Descarga de datos filtrados en CSV.


## Uso en Streamlit Cloud

La app está configurada para **carga manual del Excel**.

1. Entrar al link de Streamlit.
2. En la barra lateral, dejar seleccionada la opción **Cargar Excel manualmente**.
3. Subir el archivo `Tendencias.xlsx` o el Excel equivalente.
4. Seleccionar la unidad: **Destilación** o **Polimerización**.
5. Usar filtros, gráficos, correlaciones, variables calculadas y análisis de desfase.

No es necesario subir el Excel a GitHub. Cada usuario carga el archivo manualmente desde el navegador.

## Filtros recomendados

El tilde **Filtros recomendados** está activo por defecto y cambia según la unidad seleccionada.

### Polimerización

```text
60 < Nivel reactor R-2301 [R-2301.23LRC004] < 70
8000 < Calor de reacción URA [E-2301A/B.23URA001] < 13000
30 < Presión R-2301 [R-2301.23PCZ005] < 31
MFI Polvo [ENS.MFI_POLVO] > 0
Slurry [R-2301.23LIS005] > 50
```

### Destilación

```text
Presión C-1001 [C-1001.10PIC004] > 28
Ingreso a splitter [C-1003.10FIC006] > 36
Presión succión K-1001 [V-1003.10PIC021] > 10
```

Cuando el tilde está activo, estas variables aparecen automáticamente en **Variables a filtrar** con los valores recomendados visibles. Al actualizar el Excel, los filtros se aplican automáticamente sobre los nuevos datos siempre que las hojas, tags y nombres de variables se mantengan.

## Target de productividad / rendimiento - Polimerización

Esta versión incluye un target de productividad/rendimiento por familia operativa, sin usar nombres de producto.

La familia de comparación se define con variables de proceso/calidad:

```text
Catalizador activo + MFI + H2 + XS + Propano
```

Variables nuevas agregadas:

```text
Target productividad por familia [cat+MFI+H2+XS+propano]
Desvio vs target por familia [Rendimiento - Target]
Distancia a target por familia [menor = mas similar]
Confiabilidad target por familia [%]
Indicador periodo target optimo [1=top historico]
Percentil rendimiento dentro de base historica [%]
Familia productividad codigo [cat+MFI+H2+XS+propano]
Tipo catalizador productividad [0=ZN306, 1=mixto, 2=ZN389]
Familia MFI [0=bajo, 1=medio, 2=alto]
Familia H2 [0=bajo, 1=medio, 2=alto]
Familia XS [0=bajo, 1=medio, 2=alto]
Familia propano [0=bajo, 1=medio, 2=alto]
```

## Criterios usados para el target

### Base histórica candidata

```text
28/04/2024 a 06/10/2024
01/05/2025 a 31/12/2025
```

### Período excluido

```text
01/11/2024 a 30/04/2025
```

Motivo: posible afectación por evento de baja productividad asociado a bajo nivel del reactor.

### Presión normal de operación

Para construir la base de referencia se considera:

```text
30.0 bar <= Presión R-2301 <= 31.0 bar
```

### Selección del target

Para cada familia operativa, la app usa el **top 25% histórico de rendimiento** dentro de esa familia. Si la familia tiene pocos puntos, relaja el criterio a mismo catalizador y luego a la base confiable completa.

## Interpretación

```text
Desvío vs target = Rendimiento real - Target
```

- Desvío negativo: rendimiento real por debajo del target histórico comparable.
- Desvío cercano a cero: operación cercana al mejor histórico comparable.
- Confiabilidad alta: comparación más defendible.
- Confiabilidad baja: condición actual poco representada en la base histórica.

## Cómo ejecutar

Desde PowerShell, ubicarse en la carpeta de la app:

```powershell
cd "C:\Users\scariac\OneDrive - Petroquímica CUYO S.A.I.C\SCARIAC\Tendencias Destilación-LIPP"
py -m streamlit run app.py
```

## Cómo volver a este backup

Copiar este `app.py` en la carpeta de trabajo de Streamlit, reemplazando el archivo actual.

Antes de reemplazar, se recomienda guardar una copia del archivo existente:

```powershell
Copy-Item ".\app.py" ".\app_BACKUP_ANTES_RESTAURAR.py" -Force
```

Luego copiar el `app.py` de este backup a la carpeta de la app y ejecutar nuevamente Streamlit.

## Ajustes de interfaz

- Las variables auxiliares usadas internamente para calcular el target de productividad ya no se muestran en el selector principal.
- En **Variables a graficar**, las variables seleccionadas aparecen debajo del selector con un botón para quitarlas rápidamente.
- En Polimerización se agregó la opción **Gráfico combinado: Rendimiento vs Target**, que grafica solo `Rendimiento` y `Target productividad` sin agregar el resto de las variables auxiliares.
