# Streamlit - Tendencias Destilación / Polimerización

Versión con interfaz limpia para rendimiento estimado.

## Criterio visible

Se deja una sola variable visible para el usuario:

```text
Rendimiento estimado
```

Las columnas auxiliares del modelo quedan internas y no aparecen en el selector de variables.

## Estimación

El rendimiento estimado se calcula como promedio de antecedentes comparables en períodos OK usando solo variables de polvo/reactor:

```text
Concentración de propano
Concentración de H2
Caudal de catalizador activo
Producción de PP
MFI del polvo
XS
```

No se usa producto/grado pellet. No se usa máximo histórico.

## Diagnóstico interno

La app conserva un diagnóstico en el panel del modelo para revisar:

```text
Puntos visibles con estimado
Variables críticas usadas
Promedio estimado
Rango estimado
Desvío promedio visible
```
