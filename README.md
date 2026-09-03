# Streamlit - Tendencias Destilación / Polimerización

Versión corregida del modelo **Rendimiento estimado**.

## Criterio

El rendimiento estimado se calcula como promedio de antecedentes comparables en períodos OK, usando solamente variables del polvo/reactor:

- concentración de propano;
- concentración de H2;
- caudal ZN-306 activo;
- caudal ZN-389 activo;
- producción de PP;
- MFI del polvo;
- XS.

No usa producto/grado pellet como variable de cálculo.

## Corrección KFM / ZN-389

La versión anterior subestimaba KFM porque la base de comparación no estaba separando correctamente el modo de catalizador y podía mezclar referencias ZN-306 con ZN-389.

Ahora:

- se incluye por defecto abril-mayo 2026 como período OK para condición ZN-389/KFM;
- se compara ZN-389 contra ZN-389 y ZN-306 contra ZN-306, siempre que existan suficientes puntos;
- el usuario puede editar los períodos OK desde la barra lateral.

## Períodos OK default

- abril 2025;
- noviembre-diciembre 2025;
- abril-mayo 2026 para ZN-389/KFM.
