# The Box: análisis de la tasa de conversión (TFIE 262, Universidad del Pacífico)

Código de replicación de las cifras del Trabajo Final de Ingeniería Empresarial sobre el proceso de atención y venta en las tiendas físicas de The Box (Ola y Montaña S.A.C.), ciclo 2026-2.

Autor: Alejandro Maruy Samaniego.

## Contenido

| Archivo | Descripción |
|---|---|
| `the_box_analisis.R` | Script completo: depuración, frontera interna, estimaciones econométricas, metas, distribución de la brecha, magnitud económica y valores esperados del trabajo de campo |
| `the_box_salida_esperada.txt` | Salida completa de una ejecución de referencia, para comparar |

## Datos

Los datos son información interna de la empresa y **no se incluyen en el repositorio**. Para replicar, coloque en el directorio de trabajo:

1. `BD The Box Presentacion 30.08.xlsx` (obligatorio).
2. `Stock y Venta B2C 26-09-2026.xlsx` (opcional; margen bruto y curva de tallas, bloques 16 y 18).
3. `Stock y Ventas B2C Semana 38 ano 2026.xlsx` (opcional; curva de tallas, bloque 18).
4. `Horarios_y_Headcount_S34-S38.xlsx` (opcional; dotación planificada, bloque 19). El script solo lee las hojas agregadas "Por hora", "Cadena por semana", "Turnos por semana" y "S38" y no usa datos individuales del personal.

Si falta un archivo opcional, el bloque correspondiente se omite y el resto se ejecuta normalmente.

## Ejecución

```r
setwd("ruta/a/la/carpeta")
source("the_box_analisis.R")
```

Requiere R 4.3 o superior y los paquetes `readxl`, `dplyr`, `tidyr`, `fixest` y `lpSolve`, que el script instala si faltan. La ejecución toma menos de un minuto.

## Verificación

Cada bloque contrasta sus resultados con las cifras del documento, al número de decimales con que allí se reportan, y los marca como `[OK]` o `[REVISAR]`. Al final se imprime la tabla de verificación completa. La ejecución de referencia, con R 4.3.3 y fixest 0.14.2 y los cuatro archivos, obtiene 445 de 445 coincidencias. Sin los archivos opcionales, los bloques 16, 18 y 19 se omiten y el total de verificaciones es menor.

## Bloques

| Bloque | Contenido | Sección del documento |
|---|---|---|
| 1 a 3 | Lectura, calidad de los datos y muestra depurada | 1.4.4 y 2.1.1 |
| 4 a 8 | Formatos, descomposición de la venta, hora, día, tienda y rupturas de medición (dieciséis tiendas) | Capítulo I y Anexo B |
| 9 | Valor ideal por frontera interna (catorce tiendas) | 2.1.2 y 2.1.3 |
| 10 | Panel con efectos fijos y estimación con variable instrumental | 2.3.1 |
| 11 | Corte transversal y descomposición de la varianza | 2.3.1 |
| 12 y 13 | Semanas frontera, composición del tráfico, temporalidad y distribución de la brecha por día y franja | 2.1.3, 2.3.1, 2.3.4 y Anexo G |
| 14 | Sistema de metas y ejercicio contrafactual | 2.3.3 |
| 15 y 16 | Magnitud económica y margen bruto | 1.3.3, 2.1.4 y 2.2 |
| 17 | Valores esperados de la observación de campo | Anexo C |
| 18 | Curva de tallas: cobertura de demanda, quiebres, validación con la venta de la semana siguiente y semanas de cobertura por tienda | 2.3.1 (disponibilidad de tallas) |
| 19 | Dotación planificada por hora frente al flujo esperado, perfil horario, turnos y cruce con la semana 34 | 2.3.1 (carga y dotación) |
| 20 | Horizonte gestionable de 7,38% (brecha menos composición del tráfico) | 3.2.1 |
| 21 | Programación entera de turnos (misma dotación y refuerzo en piloto) | 4.1.3 |
| 22 | Redistribución de tallas entre tiendas (problema de transporte) | 4.2 |
| 23 | Gráficos de control p' de Laney y reglas de calidad de datos | 4.3 y Anexo H |
| 24 | Prueba retrospectiva del método de metas | 5.4 y Anexo I |
| 25 | Ejemplo de horario: Jockey Plaza, miércoles y sábado | 4.1.3 |
| 26 | Ejemplo de transferencia de talla hacia Mall Plaza Trujillo | 4.2.1 |
| 27 | Evaluación económica: objetivo acotado, flujo mensual por fases, escenarios, punto de equilibrio (en efecto y en unidades transferidas), sensibilidad (tasa, sueldo, IGV, volumen) y Monte Carlo | 3.2.1 y 5.1 |
| 28 | Tabla de verificación | Anexo A |
