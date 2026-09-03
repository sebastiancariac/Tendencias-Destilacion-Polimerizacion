# Streamlit - Tendencias Destilación / Polimerización

Versión con **rendimiento estimado promedio OK por polvo/reactor**.

## Criterio del modelo

El producto/grado pellet no participa de la estimación.

El estimado se calcula comparando cada punto contra antecedentes de períodos definidos como operación OK, usando únicamente:

```text
Concentración de propano
Concentración de H2
Caudal de catalizador activo
Producción de PP
MFI del polvo
XS
```

## Interpretación

```text
Desvío = Rendimiento real - Rendimiento estimado promedio OK
```

- Desvío cercano a cero: rendimiento acorde al promedio esperado para esas condiciones.
- Desvío negativo sostenido: posible pérdida de actividad / contaminante / revisar lechos.
- Desvío positivo: rendimiento por encima del promedio OK comparable.

## Períodos OK configurables

Por defecto quedan cargados:

```text
Abril 2025: 2025/04/01 a 2025/04/30
Noviembre-Diciembre 2025: 2025/11/01 a 2025/12/31
Período manual opcional: desactivado por defecto
```

Estos períodos se editan desde:

```text
Modelo rendimiento estimado → Estimación promedio por polvo/reactor
```

## Cambio respecto a la versión anterior

Se eliminó el concepto de máximo/ideal histórico como curva principal. Ahora la curva representa el **promedio esperado en períodos OK** para condiciones similares de polvo/reactor.
