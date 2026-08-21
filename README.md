# Streamlit - Tendencias Destilación / Polimerización

Versión con rendimiento estimado por **Gradient Boosting** sobre variables críticas.

## Criterio

El rendimiento estimado no depende del cambio de lecho ni se ajusta contra la campaña actual.

```text
Desvío = Rendimiento real - Rendimiento estimado
```

Un desvío negativo sostenido puede indicar pérdida de actividad del catalizador, contaminante o necesidad de revisar/cambiar lechos.

## Variables críticas usadas si existen

```text
MFI polvo
XS
Concentración H2
Concentración propano
Slurry
Relación TEA/C-Donor
C-Donor
Sin C-Donor
Temperatura R-2301
Tipo catalizador / ZN389 activo
```

## Cambio principal

Se reemplazó el estimador por vecinos críticos como método default porque generaba una curva demasiado plana.
El nuevo default es:

```text
Gradient Boosting quantile
Percentil benchmark default = 0.65
Entrenar solo con datos anteriores al rango visible = activado
```


## Fix selector de fechas

Se reemplazó `st.date_input` por campos de texto `YYYY/MM/DD`.

Motivo: en algunas versiones de Streamlit el calendario desplegable no muestra cómodamente años futuros en el selector, aunque la fecha sea válida. Con texto se evita esa limitación visual.

Formatos aceptados:

```text
YYYY/MM/DD
YYYY-MM-DD
DD/MM/YYYY
DD-MM-YYYY
```


## Selector mensual para informes

Se reemplazó el botón `Últimos 60 días` por un selector mensual.

Uso:

```text
1. Elegir mes en "Mes para informe"
2. Presionar "Aplicar mes"
3. La app completa Desde/Hasta con el mes completo
```

Ejemplo:

```text
Julio 2026 → Desde 2026/07/01 - Hasta 2026/07/31
```

También se dejó el botón `Usar todo`.


## Fix cambios de grado

Se corrigió la opción `Mostrar cambios de grado`.

Antes, si había filtro por producto activo, la app no dibujaba las líneas de cambio de grado porque la línea de tiempo ya estaba filtrada a un único producto.

Ahora:

```text
1. Se guarda la línea de tiempo completa de Producto antes de filtrar.
2. Las marcas se calculan con esa línea completa dentro del período visible.
3. Si el filtro por producto está activo, igual se muestran las marcas/campañas.
4. Se usa add_shape en lugar de add_vline para que las líneas sean más robustas con fechas, subplots y eje secundario.
```
