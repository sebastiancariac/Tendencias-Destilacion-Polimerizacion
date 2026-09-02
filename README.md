# Streamlit - Tendencias Destilación / Polimerización

Versión con **rendimiento estimado por correlación de referencia** tomada del archivo `Curvas de productividad (2025 - 2026).xlsx`.

## Criterio

El rendimiento estimado queda alineado con la hoja `DATOS`, columna `V = Estimado` del Excel de referencia.

No se usa Gradient Boosting ni vecinos históricos. Tampoco se ajusta contra el rendimiento real de la campaña actual.

```text
Desvío = Rendimiento real - Rendimiento estimado
```

## Correlaciones implementadas

```text
K, XS 3-5: Est = 24.47 - 0.2198·Propano
K, XS 2-4: Est = 21.94394 - 0.20900·Propano - 0.011743·Slurry + 0.33498·MFI
L:         Est = 26.838 - 0.3041·Propano
H:         Est = 59.93305 - 0.16619·Propano - 0.60735·Slurry - 1.06518·MFI
```

## Familias sin correlación

La planilla de referencia no estima las familias `S`, `T`, `H KFM`, `H FFM` ni `T ZN389`. Para esos casos la app deja `Rendimiento estimado` en blanco, para no inventar una correlación no validada.
