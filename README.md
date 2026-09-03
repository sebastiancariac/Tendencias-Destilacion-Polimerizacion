# Streamlit - Tendencias Destilación / Polimerización

Versión corregida del rendimiento estimado para Polimerización.

## Criterio

El rendimiento estimado ya no usa Producto / Grado pellet.

Se calcula exclusivamente con variables críticas de polvo y reactor:

```text
Concentración de propano
Concentración de H2
Caudal de catalizador activo
Producción de PP
MFI del polvo
XS
```

El caudal de catalizador activo se calcula con la lógica operativa:

```text
P-2209B = ZN-306 activo cuando presión descarga P-2209B >= 30
P-2209A = ZN-389 activo cuando presión descarga P-2209A >= 30
Caudal catalizador activo = caudal ZN-306 activo + caudal ZN-389 activo
```

## Estimación

Para cada punto visible, la app busca antecedentes históricos comparables por distancia normalizada en esas variables críticas. Luego calcula el rendimiento esperable/ideal con alguno de estos criterios:

```text
Promedio top comparables  [recomendado]
Percentil 90 comparables
Máximo comparable
```

El resultado se guarda como:

```text
Productividad_estimada = Rendimiento estimado [polvo/reactor]
Desvio_vs_productividad_estimada = Rendimiento real - Rendimiento estimado
```

## Importante

La fecha de cambio de lecho, el producto pellet y el grado no participan del cálculo del estimado.

El producto se mantiene solo como filtro visual y para marcar campañas en los gráficos.
