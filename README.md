# Streamlit - Tendencias Destilación / Polimerización

## Rendimiento ideal por máximo histórico validado

Se reemplazó la estimación basada en correlaciones incompletas por un criterio de **rendimiento ideal**.

Criterio principal:

```text
Rendimiento ideal = máximo histórico validado del grado producido correctamente
Desvío = Rendimiento real - Rendimiento ideal
```

El cálculo usa el producto/grado normalizado de la app:

```text
- agrupa transiciones TE con el grado base;
- agrupa KFM como KFM6110;
- compara cada punto contra el máximo validado de su propio grado.
```

Para evitar máximos falsos, antes de calcular el ideal se aplican filtros de operación normal cuando existen:

```text
30 < Presión R-2301 < 31
60 < Nivel R-2301 < 70
8000 < Calor reacción URA < 13000
MFI polvo > 0
Slurry > 50
Rendimiento entre 0 y 40
Exclusión opcional del evento de bajo nivel
Filtro opcional de outliers por IQR
```

Si un grado tiene pocos puntos válidos, la app no deja la curva en blanco: usa fallback por familia/grupo y luego fallback global.

Opciones disponibles en la barra lateral:

```text
Modelo rendimiento ideal → Rendimiento ideal histórico
- Criterio de ideal: máximo validado / promedio top 3 / percentil 95
- Mínimo puntos por grado
- Excluir outliers IQR
- Usar filtros de operación normal
- Excluir evento bajo nivel reactor
```


## Fix de comparación real vs ideal

Correcciones aplicadas:

```text
1. Rendimiento real e ideal/estimado se fuerzan al mismo eje Y.
2. El eje secundario ya no se usa para comparar variables con la misma unidad de rendimiento.
3. El rendimiento ideal tiene un piso de consistencia: no puede ser menor que el máximo real observado del mismo grado dentro del dataset disponible.
4. Se mantiene la limpieza de outliers para la tabla de referencia, pero el valor llamado máximo histórico no queda por debajo de un real observado.
```
