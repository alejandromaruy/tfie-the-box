# The Box: análisis de la tasa de conversión

Código en R del Trabajo Final de Ingeniería Empresarial "Proyecto de mejora de la tasa de conversión en las tiendas físicas de The Box" (Universidad del Pacífico, 2026).

## Contenido

| Archivo | Descripción |
|---|---|
| `the_box_analisis.R` | Programa completo, en 27 bloques |
| `the_box_salida_esperada.txt` | Resultado de una ejecución completa del programa |

## Datos

Las bases de datos son información interna de la empresa y no se incluyen en el repositorio. El programa usa cuatro archivos de Excel:

1. `BD The Box Presentacion 30.08.xlsx`: ventas, visitantes y metas por tienda, fecha y hora (obligatorio).
2. `Stock y Venta B2C 26-09-2026.xlsx`: venta y margen por modelo.
3. `Stock y Ventas B2C Semana 38 ano 2026.xlsx`: inventario por modelo y talla.
4. `Horarios_y_Headcount_S34-S38.xlsx`: horario del personal. Solo se leen las hojas agregadas por tienda, día y hora; no se usan datos de personas.

## Ejecución

```r
setwd("ruta/a/la/carpeta")
source("the_box_analisis.R")
```

Requiere R 4.3 o superior y los paquetes `readxl`, `dplyr`, `tidyr`, `fixest` y `lpSolve`.

## Bloques

| Bloques | Contenido | Dónde se usa en el trabajo |
|---|---|---|
| 1 a 3 | Limpieza y calidad de los datos | Apartados 1.2.6 y 2.1.1 |
| 4 a 8 | Resultados por formato, año, hora, día y tienda; contador de San Isidro | Capítulo I y Anexo B |
| 9, 12, 15 y 16 | Valor ideal, mejores semanas, efectos y margen | Apartados 2.1 y 2.2 |
| 10 y 11 | Relación entre visitantes y conversión | Apartado 2.3.1 y Anexo J |
| 13, 14 y 24 | Momentos de la semana, metas y método de metas | Apartados 2.3.3 y 2.3.4, Anexos G e I |
| 17 | Observación de campo | Anexo C |
| 18 | Disponibilidad de tallas | Apartado 2.3.1 y Anexo J |
| 19 | Personal programado frente a visitantes | Apartado 2.3.1 |
| 20 | Metas del árbol de objetivos | Capítulo III |
| 21 y 25 | Modelo de turnos y ejemplo de Jockey Plaza | Apartado 4.1.3 |
| 22 y 26 | Transferencias de tallas y ejemplo | Apartado 4.2 |
| 23 | Gráfico de control del contador | Apartado 4.3 y Anexo H |
| 27 | Evaluación económica y simulación | Capítulo V |
