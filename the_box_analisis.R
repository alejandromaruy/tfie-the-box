# ==============================================================================
# THE BOX - OLA Y MONTANA S.A.C.
# Script de replicacion de las cifras del Trabajo Final de Ingenieria
# Empresarial (TFIE 262, Universidad del Pacifico, ciclo 2026-2)
#
# Autor: Alejandro Maruy Samaniego
# Version: 09/11/2026 (base de catorce tiendas; diagnostico, modelos de la solucion y
#          evaluacion economica con flujo mensual por fases)
#
# Archivos de entrada (no incluidos en el repositorio por ser informacion
# interna de la empresa):
#   1. BD The Box Presentacion 30.08.xlsx         (obligatorio)
#   2. Stock y Venta B2C 26-09-2026.xlsx          (opcional, bloques 16, 18 y 22)
#   3. Stock y Ventas B2C Semana 38 ano 2026.xlsx (opcional, bloques 18 y 22)
#   4. Horarios_y_Headcount_S34-S38.xlsx          (opcional, bloques 19 y 21)
#
# Cada bloque imprime sus resultados y los contrasta con el valor citado en
# el documento. Al final se imprime la tabla de verificacion completa.
#   [OK]      el resultado coincide con el documento al numero de decimales
#             con que alli se reporta
#   [REVISAR] el resultado no coincide
#
# Estimaciones con fixest: se fija fixef.rm = "none" para conservar las
# observaciones singleton. Desde la version 0.14 el valor por defecto las
# elimina; los coeficientes no cambian, pero si el numero de observaciones y,
# marginalmente, los errores estandar.
#
# Convenciones:
#   conversion = transacciones / flujo
#   periodo base = semanas 2 a 34 de cada ejercicio
#   Capitulo I  = red completa de dieciseis tiendas
#   Capitulo II = catorce tiendas con medicion valida (sin San Isidro y
#                 Mega Plaza Chimbote, por rupturas de nivel del contador)
# ==============================================================================


# ==============================================================================
# 0. PAQUETES Y RUTAS
# ==============================================================================
paquetes <- c("readxl", "dplyr", "tidyr", "fixest", "lpSolve")
faltan <- paquetes[!paquetes %in% rownames(installed.packages())]
if (length(faltan) > 0) install.packages(faltan)
invisible(lapply(paquetes, library, character.only = TRUE))
options(scipen = 999, dplyr.summarise.inform = FALSE, width = 120)

RUTA       <- "BD The Box Presentacion 30.08.xlsx"
RUTA_STOCK <- "Stock y Venta B2C 26-09-2026.xlsx"
RUTA_STOCK_SEM <- "Stock y Ventas B2C Semana 38 ano 2026.xlsx"
RUTA_HOR   <- "Horarios_y_Headcount_S34-S38.xlsx"
# Si el archivo se descargo con otro sufijo (por ejemplo "30.08 1.xlsx"),
# se busca por el inicio del nombre en el directorio de trabajo
buscar <- function(ruta, patron) {
  if (file.exists(ruta)) return(ruta)
  cand <- list.files(pattern = patron)
  if (length(cand) >= 1) return(cand[1])
  ruta
}
RUTA       <- buscar(RUTA, "^BD The Box Presentacion.*\\.xlsx$")
RUTA_STOCK <- buscar(RUTA_STOCK, "^Stock y Venta B2C 26-09-2026.*\\.xlsx$")
RUTA_STOCK_SEM <- buscar(RUTA_STOCK_SEM, "^Stock y Ventas B2C Semana 38.*\\.xlsx$")
RUTA_HOR   <- buscar(RUTA_HOR, "^Horarios.*\\.xlsx$")
if (!file.exists(RUTA)) stop("No se encuentra el archivo: ", RUTA,
                             ". Fije el directorio de trabajo con setwd().")
cat("Base horaria:", RUTA, "\n")

EXCLUIDAS <- c("TB PCSI", "TB MGCHM")   # San Isidro y Mega Plaza Chimbote

# ---- Registro de verificaciones ----------------------------------------------
verif <- data.frame(bloque = character(), indicador = character(),
                    obtenido = character(), esperado = character(),
                    estado = character(), stringsAsFactors = FALSE)

registrar <- function(bloque, indicador, obtenido, esperado, ok) {
  ok <- isTRUE(ok)
  verif[nrow(verif) + 1, ] <<- list(bloque, indicador, obtenido, esperado,
                                    ifelse(ok, "OK", "REVISAR"))
  cat(sprintf("  [%-7s] %-62s obtenido: %-14s esperado: %s\n",
              ifelse(ok, "OK", "REVISAR"), indicador, obtenido, esperado))
  invisible(ok)
}

# Compara al numero de decimales con que la cifra aparece en el documento
chk <- function(bloque, indicador, obtenido, esperado, dec) {
  ok <- abs(round(obtenido, dec) - esperado) < 10^(-dec) / 2 + 1e-9
  registrar(bloque, indicador, formatC(obtenido, format = "f", digits = dec + 2),
            formatC(esperado, format = "f", digits = dec), ok)
}

# Para valores p y condiciones de signo
chk_menor <- function(bloque, indicador, obtenido, umbral) {
  registrar(bloque, indicador, formatC(obtenido, format = "g", digits = 4),
            paste0("< ", umbral), obtenido < umbral)
}

# Para columnas completas de una tabla
chk_vec <- function(bloque, indicador, obtenidos, esperados, dec) {
  ok <- abs(round(obtenidos, dec) - esperados) < 10^(-dec) / 2 + 1e-9
  registrar(bloque, indicador, paste0(sum(ok), " de ", length(ok)),
            paste0(length(ok), " de ", length(ok)), all(ok))
  if (!all(ok)) print(data.frame(obtenido = round(obtenidos, dec + 2),
                                 esperado = esperados)[!ok, ])
}

titulo <- function(x) cat("\n", strrep("=", 78), "\n ", x, "\n", strrep("=", 78), "\n", sep = "")

conv_de <- function(d) sum(d$transacciones) / sum(d$flujo)


# ==============================================================================
# 1. LECTURA DE LA BASE HORARIA
# ==============================================================================
titulo("1. LECTURA DE LA BASE HORARIA (hoja BD Venta por hora)")

raw <- read_excel(RUTA, sheet = "BD Venta por hora")
cat("Encabezados leidos (deben ser Anio, Semana, Fecha, InicioSemana, Hora,\n",
    "Categoria, Nro Tienda, Tienda, Nombre Tienda, Flujo, Transacciones,\n",
    "Venta, Unds Vendidas):\n")
print(names(raw)[1:13])

raw <- raw[, 1:13]
names(raw) <- c("anio", "semana", "fecha", "inicio_semana", "hora", "categoria",
                "nro_tienda", "tienda", "nombre_tienda", "flujo",
                "transacciones", "venta", "unidades")
raw <- raw %>%
  mutate(across(c(anio, semana, hora, flujo, transacciones, venta, unidades),
                as.numeric),
         fecha = as.Date(fecha),
         across(c(categoria, tienda, nombre_tienda), as.character))

tb <- raw %>% filter(categoria == "THE BOX")
cb <- raw %>% filter(categoria == "COOL BRANDS")

chk("1", "Registros del libro (ambas categorias)", nrow(raw), 122329, 0)
chk("1", "Registros de la categoria The Box",      nrow(tb),  103790, 0)


# ==============================================================================
# 2. CALIDAD DE LOS DATOS (apartado 1.4.4)
# ==============================================================================
titulo("2. CALIDAD DE LOS DATOS (apartado 1.4.4)")

calidad <- tb %>% group_by(anio) %>%
  summarise(filas = n(),
            trx_flujo_cero = sum(transacciones[flujo == 0]),
            flujo_h22 = sum(flujo[hora == 22]),
            trx_h22 = sum(transacciones[hora == 22]),
            flujo_total = sum(flujo), trx_total = sum(transacciones)) %>%
  mutate(conv_bruta = trx_total / flujo_total,
         pct_trx_flujo_cero = trx_flujo_cero / trx_total)
print(as.data.frame(calidad))

chk("2", "% transacciones con flujo cero, 2026",
    100 * calidad$pct_trx_flujo_cero[calidad$anio == 2026], 4.84, 2)
chk("2", "% transacciones con flujo cero, 2025",
    100 * calidad$pct_trx_flujo_cero[calidad$anio == 2025], 7.11, 2)
chk("2", "Hora 22 de 2026: transacciones", calidad$trx_h22[calidad$anio == 2026], 1532, 0)
chk("2", "Hora 22 de 2026: visitas",       calidad$flujo_h22[calidad$anio == 2026], 25, 0)


# ==============================================================================
# 3. MUESTRA DEPURADA (apartado 2.1.1)
#    Horario 10:00 a 21:59, flujo >= transacciones y flujo > 0
# ==============================================================================
titulo("3. MUESTRA DEPURADA (apartado 2.1.1)")

depurar <- function(d) d %>% filter(hora >= 10, hora <= 21,
                                    flujo >= transacciones, flujo > 0)
dep    <- depurar(tb)
dep_cb <- depurar(cb)

chk("3", "Registros de la muestra depurada", nrow(dep), 97834, 0)
chk("3", "% de los registros de la categoria", 100 * nrow(dep) / nrow(tb), 94.3, 1)

# Sesgo del tablero: todo el ejercicio 2026, base sin depurar frente a depurada
conv_tablero <- conv_de(tb  %>% filter(anio == 2026))
conv_dep26   <- conv_de(dep %>% filter(anio == 2026))
chk("3", "Conversion 2026 sin depurar (tablero), %", 100 * conv_tablero, 6.81, 2)
chk("3", "Conversion 2026 depurada, %",              100 * conv_dep26,   6.47, 2)
chk("3", "Sesgo al alza del tablero, pp", 100 * (conv_tablero - conv_dep26), 0.34, 2)

# Subconjuntos de uso recurrente
d26   <- dep %>% filter(anio == 2026, semana >= 2, semana <= 34)   # 16 tiendas
d14   <- d26 %>% filter(!tienda %in% EXCLUIDAS)                    # 14 tiendas
dep14 <- dep %>% filter(!tienda %in% EXCLUIDAS)                    # 14 tiendas, ambos anios


# ==============================================================================
# 4. COMPARACION ENTRE FORMATOS (Tabla de formatos, Capitulo I)
# ==============================================================================
titulo("4. COMPARACION ENTRE FORMATOS, SEMANAS 2 A 34 DE 2026")

resumen <- function(d) d %>%
  summarise(tiendas = n_distinct(tienda), flujo = sum(flujo),
            trx = sum(transacciones), venta = sum(venta), unid = sum(unidades)) %>%
  mutate(conv = trx / flujo, ticket = venta / trx, upt = unid / trx,
         ppu = venta / unid)

fmt <- bind_rows(
  resumen(d26) %>% mutate(formato = "The Box"),
  resumen(dep_cb %>% filter(anio == 2026, semana >= 2, semana <= 34)) %>%
    mutate(formato = "Coolbrands"))
print(as.data.frame(fmt))

chk("4", "The Box: tiendas",               fmt$tiendas[1], 16, 0)
chk("4", "The Box: flujo",                 fmt$flujo[1], 893219, 0)
chk("4", "The Box: conversion, %",         100 * fmt$conv[1], 6.48, 2)
chk("4", "The Box: ticket, S/",            fmt$ticket[1], 354.90, 2)
chk("4", "The Box: UPT",                   fmt$upt[1], 1.431, 3)
chk("4", "The Box: precio por unidad, S/", fmt$ppu[1], 248.02, 2)
chk("4", "Coolbrands: tiendas",               fmt$tiendas[2], 3, 0)
chk("4", "Coolbrands: flujo",                 fmt$flujo[2], 150864, 0)
chk("4", "Coolbrands: conversion, %",         100 * fmt$conv[2], 6.45, 2)
chk("4", "Coolbrands: ticket, S/",            fmt$ticket[2], 177.40, 2)
chk("4", "Coolbrands: UPT",                   fmt$upt[2], 2.083, 3)
chk("4", "Coolbrands: precio por unidad, S/", fmt$ppu[2], 85.18, 2)


# ==============================================================================
# 5. COMPARACION INTERANUAL Y DESCOMPOSICION DE LA VENTA (apartado 1.3.3)
#    Venta = Flujo x Conversion x UPT x Precio por unidad
# ==============================================================================
titulo("5. COMPARACION INTERANUAL, SEMANAS 2 A 34, DIECISEIS TIENDAS")

comp <- dep %>% filter(semana >= 2, semana <= 34) %>% group_by(anio) %>%
  resumen() %>% arrange(anio)
print(as.data.frame(comp))
a <- comp[1, ]; b <- comp[2, ]
dlv <- log(b$venta / a$venta)

chk("5", "Flujo 2025",                  a$flujo, 703705, 0)
chk("5", "Conversion 2025, %",          100 * a$conv, 5.5610, 4)
chk("5", "Conversion 2026, %",          100 * b$conv, 6.4815, 4)
chk("5", "Variacion de la conversion, pp", 100 * (b$conv - a$conv), 0.92, 2)
chk("5", "Ticket 2025, S/",             a$ticket, 259.13, 2)
chk("5", "UPT 2025",                    a$upt, 1.6309, 4)
chk("5", "UPT 2026",                    b$upt, 1.4310, 4)
chk("5", "Precio por unidad 2025, S/",  a$ppu, 158.89, 2)
chk("5", "Venta 2025, S/",              a$venta, 10140566, 0)
chk("5", "Venta 2026, S/",              b$venta, 20546688, 0)
chk("5", "Crecimiento del flujo, %",    100 * (b$flujo / a$flujo - 1), 26.9, 1)
chk("5", "Crecimiento del ticket, %",   100 * (b$ticket / a$ticket - 1), 37.0, 1)
chk("5", "Variacion de la UPT, %",      100 * (b$upt / a$upt - 1), -12.3, 1)
chk("5", "Crecimiento del precio por unidad, %", 100 * (b$ppu / a$ppu - 1), 56.1, 1)
chk("5", "Crecimiento de la venta, %",  100 * (b$venta / a$venta - 1), 102.6, 1)
chk("5", "Aporte del flujo, %",      100 * log(b$flujo  / a$flujo)  / dlv, 33.8, 1)
chk("5", "Aporte de la conversion, %", 100 * log(b$conv / a$conv)   / dlv, 21.7, 1)
chk("5", "Aporte del ticket, %",     100 * log(b$ticket / a$ticket) / dlv, 44.5, 1)

# Mismas tiendas con registro en las semanas 2 a 34 de ambos anios (14)
tiendas_en <- function(y) dep %>% filter(anio == y, semana >= 2, semana <= 34) %>%
  pull(tienda) %>% unique()
comunes <- intersect(tiendas_en(2025), tiendas_en(2026))
cat("\nTiendas comparables:", length(comunes), "\n")

interanual <- function(excluir) {
  d <- dep %>% filter(semana >= 2, semana <= 34, tienda %in% comunes)
  if (excluir) d <- d %>% filter(!tienda %in% EXCLUIDAS)
  d %>% group_by(anio) %>%
    summarise(venta = sum(venta), conv = sum(transacciones) / sum(flujo)) %>%
    arrange(anio)
}
ia_todas <- interanual(FALSE); ia_sin <- interanual(TRUE)
print(as.data.frame(ia_todas)); print(as.data.frame(ia_sin))

chk("5", "Tiendas comparables", length(comunes), 14, 0)
chk("5", "Comparables: crecimiento de la venta, %",
    100 * (ia_todas$venta[2] / ia_todas$venta[1] - 1), 85.2, 1)
chk("5", "Comparables: conversion 2026, %", 100 * ia_todas$conv[2], 6.5464, 4)
chk("5", "Mejora interanual con todas las comparables, pp",
    100 * diff(ia_todas$conv), 0.99, 2)
chk("5", "Mejora interanual sin tiendas con ruptura, pp",
    100 * diff(ia_sin$conv), 1.16, 2)


# ==============================================================================
# 6. DESEMPENO POR FRANJA HORARIA Y POR DIA (apartados 1.4.1 y 1.4.2)
# ==============================================================================
titulo("6. DESEMPENO POR HORA Y POR DIA, SEMANAS 2 A 34 DE 2026, DIECISEIS TIENDAS")

por_hora <- d26 %>% group_by(hora) %>%
  summarise(flujo = sum(flujo), trx = sum(transacciones), venta = sum(venta)) %>%
  mutate(part = flujo / sum(flujo), conv = trx / flujo, ticket = venta / trx)
print(as.data.frame(por_hora))

resto10 <- por_hora %>% filter(hora != 10)
chk("6", "Participacion del flujo de 18:00 a 20:59, %",
    100 * sum(por_hora$part[por_hora$hora %in% 18:20]), 37.7, 1)
chk("6", "Participacion de las cinco horas de mayor flujo, %",
    100 * sum(sort(por_hora$part, decreasing = TRUE)[1:5]), 58.2, 1)
chk("6", "Participacion del flujo de las 10:00, %",
    100 * por_hora$part[por_hora$hora == 10], 2.54, 2)
chk("6", "Conversion de las 10:00, %", 100 * por_hora$conv[por_hora$hora == 10], 4.21, 2)
chk("6", "Conversion del resto de la jornada, %",
    100 * sum(resto10$trx) / sum(resto10$flujo), 6.54, 2)
chk("6", "Segunda hora de menor conversion, %",
    100 * sort(por_hora$conv)[2], 6.18, 2)
chk("6", "Ticket de las 10:00, S/", por_hora$ticket[por_hora$hora == 10], 323.47, 2)

por_dia <- d26 %>% mutate(dow = as.integer(format(fecha, "%u"))) %>%  # 1 = lunes
  group_by(dow) %>%
  summarise(flujo = sum(flujo), trx = sum(transacciones), venta = sum(venta)) %>%
  mutate(part = flujo / sum(flujo), conv = trx / flujo, ticket = venta / trx)
print(as.data.frame(por_dia))

chk_vec("6", "Tabla por dia: participacion del flujo (7 dias)", 100 * por_dia$part,
        c(11.28, 10.53, 10.92, 11.98, 13.21, 19.67, 22.41), 2)
chk_vec("6", "Tabla por dia: conversion (7 dias)", 100 * por_dia$conv,
        c(6.33, 6.63, 6.61, 6.63, 6.89, 6.75, 5.86), 2)
chk_vec("6", "Tabla por dia: ticket (7 dias)", por_dia$ticket,
        c(350.41, 350.09, 348.68, 357.41, 353.69, 354.92, 362.63), 2)


# ==============================================================================
# 7. DESEMPENO POR TIENDA (apartado 1.4.3 y Anexo B)
# ==============================================================================
titulo("7. DESEMPENO POR TIENDA, SEMANAS 2 A 34 DE 2026 (Anexo B)")

anexo_b <- d26 %>% group_by(tienda, nombre_tienda) %>%
  summarise(semanas = n_distinct(semana), flujo = sum(flujo),
            trx = sum(transacciones), venta = sum(venta), unid = sum(unidades)) %>%
  mutate(flujo_semanal = flujo / semanas, conv = trx / flujo,
         ticket = venta / trx, upt = unid / trx,
         cap_II = ifelse(tienda %in% EXCLUIDAS, "No", "Si")) %>%
  arrange(desc(conv))
print(as.data.frame(anexo_b %>% select(nombre_tienda, semanas, flujo_semanal,
                                        conv, ticket, upt, cap_II)))

chk_vec("7", "Anexo B: conversion por tienda (16)", 100 * anexo_b$conv,
        c(8.58, 7.61, 7.56, 7.42, 7.40, 7.10, 6.97, 6.96, 6.54, 6.44, 6.14,
          6.13, 5.62, 5.45, 5.28, 4.88), 2)
chk_vec("7", "Anexo B: semanas por tienda (16)", anexo_b$semanas,
        c(33, 33, 33, 33, 33, 33, 33, 33, 33, 33, 33, 33, 33, 29, 33, 10), 0)
chk_vec("7", "Anexo B: flujo semanal por tienda (16)", anexo_b$flujo_semanal,
        c(918, 1209, 1039, 716, 1057, 1377, 1851, 1477, 975, 4434, 4577,
          1820, 1954, 992, 2194, 1973), 0)
chk_vec("7", "Anexo B: ticket por tienda (16)", anexo_b$ticket,
        c(324.55, 333.66, 319.72, 310.76, 317.78, 345.04, 377.44, 300.60,
          290.14, 396.53, 387.31, 373.60, 359.87, 313.21, 328.92, 405.96), 2)
chk_vec("7", "Anexo B: UPT por tienda (16)", anexo_b$upt,
        c(1.396, 1.383, 1.610, 1.534, 1.327, 1.503, 1.671, 1.439, 1.433,
          1.389, 1.358, 1.490, 1.497, 1.392, 1.201, 1.563), 3)
chk("7", "Factor entre la mayor y la menor conversion",
    max(anexo_b$conv) / min(anexo_b$conv), 1.76, 2)

r16 <- resumen(d26); r14 <- resumen(d14)
chk("7", "Red de 16 tiendas: flujo semanal", r16$flujo / 33, 27067, 0)
chk("7", "Catorce tiendas: flujo semanal",   r14$flujo / 33, 25479, 0)
chk("7", "Catorce tiendas: ticket, S/",      r14$ticket, 357.52, 2)
chk("7", "Catorce tiendas: UPT",             r14$upt, 1.429, 3)


# ==============================================================================
# 8. RUPTURAS DE MEDICION (apartado 1.4.4)
# ==============================================================================
titulo("8. RUPTURAS DE MEDICION: SAN ISIDRO")

si <- d26 %>% filter(tienda == "TB PCSI") %>% group_by(semana) %>%
  summarise(flujo = sum(flujo), trx = sum(transacciones), venta = sum(venta)) %>%
  mutate(conv = trx / flujo)
print(as.data.frame(si))

tramos <- si %>% filter(semana <= 19 | semana >= 22) %>%
  mutate(tramo = ifelse(semana <= 19, "Semanas 2 a 19", "Semanas 22 a 34")) %>%
  group_by(tramo) %>%
  summarise(flujo_sem = mean(flujo), trx_sem = mean(trx), venta_sem = mean(venta),
            conv = sum(trx) / sum(flujo))
print(as.data.frame(tramos))

chk("8", "San Isidro, sem. 2-19: flujo semanal",  tramos$flujo_sem[1], 555, 0)
chk("8", "San Isidro, sem. 2-19: transacciones",  tramos$trx_sem[1], 60, 0)
chk("8", "San Isidro, sem. 2-19: venta, S/",      tramos$venta_sem[1], 17508, 0)
chk("8", "San Isidro, sem. 2-19: conversion, %",  100 * tramos$conv[1], 10.79, 2)
chk("8", "San Isidro, sem. 22-34: flujo semanal", tramos$flujo_sem[2], 938, 0)
chk("8", "San Isidro, sem. 22-34: transacciones", tramos$trx_sem[2], 44, 0)
chk("8", "San Isidro, sem. 22-34: venta, S/",     tramos$venta_sem[2], 14948, 0)
chk("8", "San Isidro, sem. 22-34: conversion, %", 100 * tramos$conv[2], 4.71, 2)

exc <- d26 %>% filter(tienda %in% EXCLUIDAS)
chk("8", "Tiendas excluidas: % del flujo",  100 * sum(exc$flujo) / sum(d26$flujo), 5.9, 1)
chk("8", "Tiendas excluidas: % de la venta", 100 * sum(exc$venta) / sum(d26$venta), 5.0, 1)


# ==============================================================================
# 9. VALOR IDEAL: FRONTERA INTERNA (apartado 2.1.3)
#    Percentil de la conversion semanal propia de cada tienda (interpolacion
#    lineal, type = 7), ponderado por el flujo de la tienda. Se excluyen las
#    semanas-tienda con menos de 100 visitas o menos de 10 transacciones.
# ==============================================================================
titulo("9. VALOR IDEAL, CATORCE TIENDAS, SEMANAS 2 A 34 DE 2026")

semanal <- function(d) d %>% group_by(tienda, semana) %>%
  summarise(f = sum(flujo), t = sum(transacciones), v = sum(venta),
            u = sum(unidades)) %>% mutate(conv = t / f) %>% ungroup()

valor_ideal <- function(d, p) {
  sw <- semanal(d) %>% filter(f >= 100, t >= 10)
  fl <- d %>% group_by(tienda) %>% summarise(F = sum(flujo))
  sw %>% group_by(tienda) %>% summarise(q = quantile(conv, p, type = 7)) %>%
    inner_join(fl, by = "tienda") %>% summarise(v = sum(q * F) / sum(F)) %>% pull(v)
}

conv_real <- conv_de(d14)
p75 <- valor_ideal(d14, 0.75); p90 <- valor_ideal(d14, 0.90); pmax <- valor_ideal(d14, 1)
brecha <- p90 - conv_real

chk("9", "Conversion real, catorce tiendas, %", 100 * conv_real, 6.4905, 4)
chk("9", "Percentil 75 ponderado, %",           100 * p75, 7.1550, 4)
chk("9", "Percentil 90 ponderado (valor ideal), %", 100 * p90, 7.6119, 4)
chk("9", "Mejor semana propia ponderada, %",    100 * pmax, 8.2468, 4)
chk("9", "Brecha contra el percentil 75, pp",   100 * (p75 - conv_real), 0.66, 2)
chk("9", "Brecha contra el percentil 90, pp",   100 * brecha, 1.12, 2)
chk("9", "Brecha contra la mejor semana, pp",   100 * (pmax - conv_real), 1.76, 2)
chk("9", "Brecha contra el referente de 10%, pp", 100 * (0.10 - conv_real), 3.51, 2)

# Semanas-tienda sobre 10% (apartado 2.1.2)
sw14 <- semanal(d14) %>% filter(f >= 100, t >= 10)
sw16 <- semanal(d26) %>% filter(f >= 100, t >= 10)
cat("\nTiendas con semanas sobre 10% (14 tiendas):",
    paste(unique(sw14$tienda[sw14$conv > 0.10]), collapse = ", "), "\n")
chk("9", "Semanas-tienda validas, catorce tiendas", nrow(sw14), 439, 0)
chk("9", "Semanas-tienda sobre 10%, catorce tiendas", sum(sw14$conv > 0.10), 10, 0)
chk("9", "Semanas-tienda sobre 10%, % de las catorce", 100 * mean(sw14$conv > 0.10), 2.3, 1)
# Red completa, mismo periodo y mismo filtro: las semanas sobre 10% que no
# pertenecen a las catorce tiendas corresponden a las dos tiendas excluidas.
chk("9", "Semanas-tienda validas, dieciseis tiendas", nrow(sw16), 499, 0)
chk("9", "Semanas-tienda sobre 10%, dieciseis tiendas", sum(sw16$conv > 0.10), 20, 0)


# ==============================================================================
# 10. RELACION ENTRE FLUJO Y CONVERSION (apartado 2.3.1)
#     conversion = b * ln(flujo) + efectos fijos tienda x hora + efectos fijos
#     fecha; errores agrupados por tienda; catorce tiendas, 2025 y 2026, horas
#     con al menos 20 visitas.
# ==============================================================================
titulo("10. PANEL CON EFECTOS FIJOS Y VARIABLE INSTRUMENTAL")

h <- dep14 %>% filter(flujo >= 20) %>%
  mutate(conv = transacciones / flujo, lf = log(flujo),
         tienda_hora = paste0(tienda, "_", hora))

m1 <- feols(conv ~ lf | tienda_hora + fecha, data = h, cluster = ~tienda, fixef.rm = "none")
print(summary(m1))
b1 <- coef(m1)["lf"]; se1 <- se(m1)["lf"]; p1 <- pvalue(m1)["lf"]

chk("10", "Panel: observaciones", nobs(m1), 37546, 0)
chk("10", "Panel: coeficiente de ln(flujo)", b1, -0.01285, 5)
chk("10", "Panel: error estandar", se1, 0.00126, 5)
chk_menor("10", "Panel: valor p", p1, 0.001)
chk("10", "Efecto de +10% de flujo, pp", 100 * b1 * log(1.1), -0.12, 2)

# Robustez frente al sesgo de division (el flujo esta en el denominador de la
# variable dependiente): se instrumenta ln(flujo) con el flujo de la misma
# tienda y hora siete dias antes.
rezago <- dep14 %>% transmute(tienda, hora, fecha = fecha + 7, flujo_7 = flujo)
hiv <- h %>% inner_join(rezago, by = c("tienda", "hora", "fecha")) %>%
  filter(flujo_7 > 0) %>% mutate(lf7 = log(flujo_7))

m_ols  <- feols(conv ~ lf | tienda_hora + fecha, data = hiv, cluster = ~tienda, fixef.rm = "none")
m_fs   <- feols(lf ~ lf7 | tienda_hora + fecha, data = hiv, cluster = ~tienda, fixef.rm = "none")
m_iv   <- feols(conv ~ 1 | tienda_hora + fecha | lf ~ lf7, data = hiv, cluster = ~tienda, fixef.rm = "none")
print(etable(m_ols, m_fs, m_iv))

chk("10", "VI: observaciones", nobs(m_iv), 36086, 0)
chk("10", "VI: coeficiente de primera etapa", coef(m_fs)["lf7"], 0.1322, 4)
chk("10", "VI: estadistico t de primera etapa", tstat(m_fs)["lf7"], 7.62, 2)
chk("10", "VI: coeficiente de ln(flujo)", coef(m_iv)["fit_lf"], -0.01984, 5)
chk("10", "VI: error estandar", se(m_iv)["fit_lf"], 0.00471, 5)
chk("10", "VI: valor p", pvalue(m_iv)["fit_lf"], 0.001, 3)
chk("10", "OLS en la muestra de la VI: coeficiente", coef(m_ols)["lf"], -0.01228, 5)

# Elasticidad de las transacciones al flujo (menor que 1 si la conversion cae)
ht <- hiv %>% filter(transacciones > 0) %>% mutate(lt = log(transacciones))
e_ols <- feols(lt ~ lf | tienda_hora + fecha, data = ht, cluster = ~tienda, fixef.rm = "none")
e_iv  <- feols(lt ~ 1 | tienda_hora + fecha | lf ~ lf7, data = ht, cluster = ~tienda, fixef.rm = "none")
chk("10", "Elasticidad transacciones-flujo, MCO", coef(e_ols)["lf"], 0.677, 3)
chk("10", "Elasticidad transacciones-flujo, VI",  coef(e_iv)["fit_lf"], 0.660, 3)


# ==============================================================================
# 11. CORTE TRANSVERSAL Y DESCOMPOSICION DE LA VARIANZA (apartado 2.3.1)
# ==============================================================================
titulo("11. CORTE TRANSVERSAL Y VARIANZA, CATORCE TIENDAS")

ct <- d14 %>% group_by(tienda) %>%
  summarise(flujo_sem = sum(flujo) / n_distinct(semana),
            conv = sum(transacciones) / sum(flujo))
sp <- suppressWarnings(cor.test(ct$flujo_sem, ct$conv, method = "spearman"))
pe <- cor.test(ct$flujo_sem, ct$conv, method = "pearson")
print(sp); print(pe)
chk("11", "Spearman flujo semanal - conversion", sp$estimate, -0.758, 3)
chk("11", "Spearman: valor p", sp$p.value, 0.003, 3)
chk("11", "Pearson flujo semanal - conversion", pe$estimate, -0.438, 3)
chk("11", "Pearson: valor p", pe$p.value, 0.117, 3)

# Varianza de la conversion semanal (todas las semanas-tienda)
vr <- semanal(d14)
r2_t <- summary(lm(conv ~ factor(tienda), data = vr))$r.squared
r2_s <- summary(lm(conv ~ factor(semana), data = vr))$r.squared
r2_a <- summary(lm(conv ~ factor(tienda) + factor(semana), data = vr))$r.squared
sd_entre  <- vr %>% group_by(tienda) %>% summarise(c = sum(t) / sum(f)) %>% pull(c) %>% sd()
sd_dentro <- vr %>% group_by(tienda) %>% summarise(s = sd(conv)) %>% pull(s) %>% mean()

chk("11", "R2 de la tienda", r2_t, 0.442, 3)
chk("11", "R2 de la semana calendario", r2_s, 0.159, 3)
chk("11", "R2 de tienda y semana", r2_a, 0.622, 3)
chk("11", "Desviacion entre tiendas, pp", 100 * sd_entre, 1.00, 2)
chk("11", "Desviacion dentro de cada tienda, pp", 100 * sd_dentro, 0.99, 2)


# ==============================================================================
# 12. SEMANAS FRONTERA FRENTE A LAS DEMAS (apartados 2.1.3, 2.3.1 y 2.3.4)
#     Semana frontera: conversion de la tienda >= su propio percentil 90
# ==============================================================================
titulo("12. SEMANAS FRONTERA, COMPOSICION DEL TRAFICO Y TEMPORALIDAD")

p90_tienda <- sw14 %>% group_by(tienda) %>%
  summarise(p90 = quantile(conv, 0.90, type = 7))
wk <- semanal(d14) %>% inner_join(p90_tienda, by = "tienda") %>%
  mutate(top = conv >= p90)
d14t <- d14 %>% inner_join(wk %>% select(tienda, semana, top),
                           by = c("tienda", "semana"))

ft <- d14t %>% group_by(top) %>% summarise(conv = sum(transacciones) / sum(flujo))
chk("12", "Conversion de las semanas frontera, %", 100 * ft$conv[ft$top], 7.93, 2)
chk("12", "Conversion de las demas semanas, %",    100 * ft$conv[!ft$top], 6.32, 2)

# Composicion del trafico: flujo medio de las semanas frontera / demas semanas
ratio <- wk %>% group_by(tienda, top) %>% summarise(f = mean(f)) %>%
  tidyr::pivot_wider(names_from = top, values_from = f) %>%
  mutate(r = `TRUE` / `FALSE`) %>% pull(r) %>% median()
comp_ols <- 100 * b1 * log(ratio)
comp_iv  <- 100 * coef(m_iv)["fit_lf"] * log(ratio)
chk("12", "Flujo de las semanas frontera / demas (mediana)", ratio, 0.8906, 4)
chk("12", "Componente de composicion con MCO, pp", comp_ols, 0.15, 2)
chk("12", "Componente de composicion con VI, pp",  comp_iv, 0.23, 2)
chk("12", "Brecha atribuible a la gestion, MCO, pp", 100 * brecha - comp_ols, 0.97, 2)
chk("12", "Brecha atribuible a la gestion, VI, pp",  100 * brecha - comp_iv, 0.89, 2)

# Temporalidad de la frontera
chk("12", "Semanas frontera", sum(wk$top), 53, 0)
chk("12", "Semanas frontera desde la semana 11", sum(wk$top & wk$semana >= 11), 52, 0)
racha <- wk %>% arrange(tienda, semana) %>% group_by(tienda) %>%
  summarise(max_racha = { r <- rle(top); max(c(0, r$lengths[r$values])) })
chk("12", "Racha maxima de semanas frontera consecutivas", max(racha$max_racha), 3, 0)

bloque_conv <- function(y, a, b) conv_de(dep14 %>% filter(anio == y, semana >= a, semana <= b))
chk("12", "Conversion 2026, semanas 2 a 10, %",  100 * bloque_conv(2026, 2, 10), 5.98, 2)
chk("12", "Conversion 2026, semanas 11 a 34, %", 100 * bloque_conv(2026, 11, 34), 6.68, 2)
chk("12", "Conversion 2025, semanas 2 a 10, %",  100 * bloque_conv(2025, 2, 10), 5.29, 2)
chk("12", "Conversion 2025, semanas 11 a 34, %", 100 * bloque_conv(2025, 11, 34), 5.44, 2)

d14b <- d14 %>% filter(semana >= 11)
chk("12", "Brecha recalculada sobre semanas 11 a 34, pp",
    100 * (valor_ideal(d14b, 0.90) - conv_de(d14b)), 1.01, 2)
red_sem <- d14 %>% group_by(semana) %>% summarise(c = 100 * sum(transacciones) / sum(flujo))
chk("12", "Tendencia de la conversion de red, pp por semana",
    coef(lm(c ~ semana, data = red_sem))["semana"], 0.023, 3)


# ==============================================================================
# 13. DISTRIBUCION DE LA BRECHA POR DIA Y FRANJA (apartado 2.3.4 y Anexo G)
#     Aporte = (conv. frontera - conv. demas) x participacion del segmento en
#     el flujo de las semanas no frontera
# ==============================================================================
titulo("13. DISTRIBUCION DE LA VENTAJA DE LA FRONTERA (Tabla de segmentos)")

d14t <- d14t %>%
  mutate(dow = as.integer(format(fecha, "%u")),
         dia = case_when(dow == 7 ~ "Domingo", dow == 6 ~ "Sabado",
                         TRUE ~ "Lunes a viernes"),
         franja = case_when(hora == 10 ~ "Apertura (10:00)",
                            hora <= 17 ~ "Valle (11:00 a 17:59)",
                            TRUE ~ "Punta (18:00 a 21:59)"))

seg <- d14t %>% group_by(dia, franja, top) %>%
  summarise(F = sum(flujo), T = sum(transacciones)) %>%
  tidyr::pivot_wider(names_from = top, values_from = c(F, T)) %>% ungroup() %>%
  mutate(share = F_FALSE / sum(F_FALSE),
         aporte = 100 * share * (T_TRUE / F_TRUE - T_FALSE / F_FALSE))

apertura <- seg %>% filter(franja == "Apertura (10:00)") %>%
  summarise(dia = "Tres dias", franja = "Apertura (10:00)",
            F_FALSE = sum(F_FALSE), F_TRUE = sum(F_TRUE),
            T_FALSE = sum(T_FALSE), T_TRUE = sum(T_TRUE),
            share = sum(share), aporte = sum(aporte))
tabla_seg <- bind_rows(seg %>% filter(franja != "Apertura (10:00)") %>%
                         arrange(desc(aporte)), apertura) %>%
  mutate(conv_demas = 100 * T_FALSE / F_FALSE, conv_frontera = 100 * T_TRUE / F_TRUE,
         acumulado = 100 * cumsum(aporte) / sum(aporte))
print(as.data.frame(tabla_seg %>% select(dia, franja, conv_demas, conv_frontera,
                                          aporte, acumulado)))

chk_vec("13", "Segmentos: conversion de las demas semanas", tabla_seg$conv_demas,
        c(6.67, 6.45, 6.72, 5.71, 5.71, 6.41, 4.16), 2)
chk_vec("13", "Segmentos: conversion de las semanas frontera", tabla_seg$conv_frontera,
        c(8.73, 7.95, 8.35, 7.36, 7.04, 7.87, 4.60), 2)
chk_vec("13", "Segmentos: aporte, pp", tabla_seg$aporte,
        c(0.57, 0.42, 0.17, 0.17, 0.16, 0.13, 0.01), 2)
chk_vec("13", "Segmentos: participacion acumulada, %", tabla_seg$acumulado,
        c(35.1, 61.0, 71.3, 81.5, 91.4, 99.4, 100.0), 1)
chk("13", "Suma de los aportes, pp", sum(tabla_seg$aporte), 1.63, 2)
chk("13", "Diferencia simple frontera - demas, pp",
    100 * (ft$conv[ft$top] - ft$conv[!ft$top]), 1.61, 2)

# Tiendas de alto flujo (Jockey Plaza y Mall Plaza Trujillo) frente a las demas
alto <- d14t %>% filter(franja != "Apertura (10:00)") %>%
  mutate(grupo = ifelse(tienda %in% c("TB JPSU", "TB MPTRU"), "alto", "otras"),
         banda = ifelse(grepl("Punta", franja), "punta", "valle")) %>%
  group_by(grupo, banda, top) %>%
  summarise(c = sum(transacciones) / sum(flujo)) %>%
  tidyr::pivot_wider(names_from = top, values_from = c) %>%
  mutate(ventaja = 100 * (`TRUE` - `FALSE`))
print(as.data.frame(alto))
ven <- function(g, b) alto$ventaja[alto$grupo == g & alto$banda == b]
chk("13", "Alto flujo: ventaja de la frontera en punta, pp", ven("alto", "punta"), 1.80, 2)
chk("13", "Alto flujo: ventaja de la frontera en valle, pp", ven("alto", "valle"), 0.87, 2)
chk("13", "Otras doce: ventaja de la frontera en punta, pp", ven("otras", "punta"), 1.83, 2)
chk("13", "Otras doce: ventaja de la frontera en valle, pp", ven("otras", "valle"), 1.77, 2)
chk("13", "Alto flujo: % de la venta de las catorce tiendas",
    100 * sum(d14$venta[d14$tienda %in% c("TB JPSU", "TB MPTRU")]) / sum(d14$venta), 37.6, 1)
chk("13", "Alto flujo: % de la venta de la red de dieciseis",
    100 * sum(d14$venta[d14$tienda %in% c("TB JPSU", "TB MPTRU")]) / sum(d26$venta), 35.7, 1)

# Escenario de la hora de apertura
h10 <- d14 %>% filter(hora == 10); h_resto <- d14 %>% filter(hora != 10)
part10 <- sum(h10$flujo) / sum(d14$flujo)
gan10  <- part10 * (conv_de(h_resto) - conv_de(h10))
chk("13", "Apertura: conversion de las 10:00, %", 100 * conv_de(h10), 4.21, 2)
chk("13", "Apertura: conversion del resto del dia, %", 100 * conv_de(h_resto), 6.55, 2)
chk("13", "Apertura: ganancia si iguala al resto, pp", 100 * gan10, 0.055, 3)
chk("13", "Apertura: % de la brecha", 100 * gan10 / brecha, 4.9, 1)


# ==============================================================================
# 14. SISTEMA DE METAS (apartado 2.3.3)
#     readxl descarta la columna A vacia de la hoja, por eso se toman 1:10
# ==============================================================================
titulo("14. SISTEMA DE METAS, CATORCE TIENDAS")

meta <- read_excel(RUTA, sheet = "BD Cumplimiento de Meta de Vta")
meta <- meta[, 1:10]
names(meta) <- c("anio", "mes", "semana", "categoria", "nro_tienda", "tienda",
                 "nombre_tienda", "meta", "venta", "cumplimiento")
meta <- meta %>%
  mutate(across(c(anio, semana, meta, venta), as.numeric),
         across(c(categoria, tienda, nombre_tienda), as.character)) %>%
  filter(!is.na(anio), categoria == "THE BOX", !tienda %in% EXCLUIDAS)
print(head(as.data.frame(meta), 3))   # nombre_tienda debe ser texto

m26 <- meta %>% filter(anio == 2026, semana >= 2, semana <= 34, !is.na(meta), meta > 0)
v25 <- meta %>% filter(anio == 2025, semana >= 2, semana <= 34) %>%
  select(semana, tienda, v25 = venta)
mj  <- m26 %>% inner_join(v25, by = c("semana", "tienda")) %>% filter(v25 > 0) %>%
  mutate(cump = venta / meta)

exige <- sum(mj$meta) / sum(mj$v25) - 1
logro <- sum(mj$venta) / sum(mj$v25) - 1
chk("14", "Observaciones semana-tienda", nrow(mj), 396, 0)
chk("14", "Crecimiento exigido por la meta 2026, %", 100 * exige, 145.2, 1)
chk("14", "Crecimiento logrado, %",                  100 * logro, 84.5, 1)
chk("14", "Meta vigente: % semanas-tienda >= 80%",  100 * mean(mj$cump >= 0.8), 38.4, 1)
chk("14", "Meta vigente: % semanas-tienda >= 100%", 100 * mean(mj$cump >= 1.0), 5.8, 1)

contrafactual <- function(g) {
  c <- mj$venta / (mj$v25 * (1 + g)); c(100 * mean(c >= 0.8), 100 * mean(c >= 1))
}
esperado_cf <- list(list(logro, 64.6, 45.2), list(0.65, 75.5, 52.8),
                    list(0.60, 77.5, 56.3), list(0.40, 87.6, 69.2))
for (e in esperado_cf) {
  r <- contrafactual(e[[1]])
  chk("14", sprintf("Meta con crecimiento de %.1f%%: >= 80%%", 100 * e[[1]]), r[1], e[[2]], 1)
  chk("14", sprintf("Meta con crecimiento de %.1f%%: >= 100%%", 100 * e[[1]]), r[2], e[[3]], 1)
}

cm <- m26 %>% group_by(tienda) %>% summarise(cumpl = sum(venta) / sum(meta))
jj <- ct %>% select(tienda, conv) %>% inner_join(cm, by = "tienda")
cc <- cor.test(jj$cumpl, jj$conv)
print(cc)
chk("14", "Tiendas cruzadas", nrow(jj), 14, 0)
chk("14", "Correlacion cumplimiento - conversion", cc$estimate, 0.085, 3)
chk("14", "Correlacion: valor p", cc$p.value, 0.773, 3)


# ==============================================================================
# 15. MAGNITUD ECONOMICA (apartados 2.1.4 y 2.2)
#     Factor de estacionalidad = flujo total 2025 / flujo semanas 2 a 34 de
#     2025, sobre la muestra depurada de la red
# ==============================================================================
titulo("15. MAGNITUD ECONOMICA, CATORCE TIENDAS")

d25 <- dep %>% filter(anio == 2025)
factor_est <- sum(d25$flujo) / sum(d25$flujo[d25$semana >= 2 & d25$semana <= 34])
sem25 <- d25 %>% group_by(semana) %>% summarise(f = sum(flujo))
chk("15", "Factor de estacionalidad", factor_est, 1.9012, 4)
chk("15", "Flujo de la semana 51 de 2025", sem25$f[sem25$semana == 51], 66775, 0)
chk("15", "Flujo de la semana 52 de 2025", sem25$f[sem25$semana == 52], 90763, 0)
chk("15", "Promedio semanal de flujo de 2025", mean(sem25$f), 25728, 0)

flujo_anual <- r14$flujo * factor_est
trx_adic    <- brecha * flujo_anual
fact_adic   <- trx_adic * r14$ticket
venta_anual <- r14$venta * factor_est
vpv         <- r14$venta / r14$flujo
chk("15", "Flujo anual equivalente", flujo_anual, 1598512, 0)
chk("15", "Transacciones anuales adicionales", trx_adic, 17926, 0)
chk("15", "Facturacion no capturada, S/ millones", fact_adic / 1e6, 6.41, 2)
chk("15", "Venta por visitante actual, S/", vpv, 23.20, 2)
chk("15", "Venta por visitante en la frontera, S/", vpv * p90 / conv_real, 27.21, 2)
chk("15", "Brecha de venta por visitante, S/", vpv * (p90 / conv_real - 1), 4.01, 2)
chk("15", "1 pp de conversion con perfil estacional, S/ millones",
    0.01 * flujo_anual * r14$ticket / 1e6, 5.71, 2)
chk("15", "1 pp de conversion sin ajuste estacional, S/ millones",
    0.01 * r14$flujo * 52 / 33 * r14$ticket / 1e6, 4.74, 2)
chk("15", "Venta anual del canal, catorce tiendas, S/ millones", venta_anual / 1e6, 37.09, 2)
chk("15", "Facturacion no capturada, % de la venta anual", 100 * fact_adic / venta_anual, 17.3, 1)

# Productividad: dotacion declarada en la hoja BD Analisis Final
hoja_af <- grep("^BD An.lisis Final$", excel_sheets(RUTA), value = TRUE)
staff <- read_excel(RUTA, sheet = hoja_af, range = "B2:I18")
names(staff)[c(1, 8)] <- c("tienda_nombre", "staff")
staff14 <- sum(staff$staff[!staff$tienda_nombre %in% c("San Isidro", "Mega Plaza Chimbote")])
trx_asesor <- r14$trx / 33 / staff14
chk("15", "Asesores declarados, catorce tiendas", staff14, 63, 0)
chk("15", "Transacciones por asesor por semana", trx_asesor, 26.2, 1)
chk("15", "Transacciones por asesor en la frontera", trx_asesor * p90 / conv_real, 30.8, 1)
chk("15", "Brecha de transacciones por asesor", trx_asesor * (p90 / conv_real - 1), 4.5, 1)
chk("15", "Brecha de venta semanal por asesor, S/",
    trx_asesor * (p90 / conv_real - 1) * r14$ticket, 1621, 0)

# Oportunidad de las unidades por transaccion con el mismo criterio de
# frontera (apartado 1.3.3): percentil 90 de la UPT semanal de cada tienda,
# ponderado por sus transacciones
sw14u <- sw14 %>% mutate(upt = u / t)
w_t <- d14 %>% group_by(tienda) %>% summarise(T = sum(transacciones))
upt_ideal <- sw14u %>% group_by(tienda) %>% summarise(q = quantile(upt, 0.9, type = 7)) %>%
  inner_join(w_t, by = "tienda") %>% summarise(v = sum(q * T) / sum(T)) %>% pull(v)
chk("15", "Oportunidad de la UPT, S/ millones",
    venta_anual * (upt_ideal / r14$upt - 1) / 1e6, 3.69, 2)


# ==============================================================================
# 16. MARGEN BRUTO (apartado 2.2). Requiere el archivo de stock y venta.
#     Venta y utilidad de los ultimos 7 dias (20 al 26/09/2026), sin IGV.
# ==============================================================================
titulo("16. MARGEN BRUTO (base de stock y venta, semana del 20 al 26/09/2026)")

if (file.exists(RUTA_STOCK)) {
  stock <- read_excel(RUTA_STOCK, sheet = "Sheet1")
  stock <- stock %>% filter(Canal == "ST-The Box") %>%
    transmute(tienda = StoreCode, venta7 = as.numeric(`Venta Ult 7 Dias (S/)`),
              util7 = as.numeric(`Utilidad Ult 7 Dias`))
  mg <- stock %>% group_by(tienda) %>%
    summarise(venta7 = sum(venta7), util7 = sum(util7)) %>% mutate(margen = util7 / venta7)
  print(as.data.frame(mg))
  margen16 <- sum(mg$util7) / sum(mg$venta7)
  margen14 <- with(mg[!mg$tienda %in% EXCLUIDAS, ], sum(util7) / sum(venta7))
  chk("16", "Tiendas The Box en la base de stock", nrow(mg), 16, 0)
  chk("16", "Margen bruto, dieciseis tiendas, %", 100 * margen16, 41.8, 1)
  chk("16", "Utilidad bruta no capturada, S/ millones", fact_adic * margen16 / 1e6, 2.68, 2)
  cat(sprintf("  Referencia: margen de las catorce tiendas %.2f%%; utilidad no capturada S/%.2f millones\n",
              100 * margen14, fact_adic * margen14 / 1e6))
} else {
  cat("  Archivo de stock no encontrado; bloque omitido.\n")
}


# ==============================================================================
# 17. VALORES ESPERADOS DE LA OBSERVACION DE CAMPO (Anexo C)
#     Domingos de las semanas 2 a 34 de 2026, misma tienda y ventana, cada hora
#     ponderada por la fraccion cubierta por la observacion
# ==============================================================================
titulo("17. OBSERVACION DE CAMPO DEL 27/09/2026 (Anexo C)")

ventana <- function(cod, pesos, obs_f, obs_t, asesores, horas) {
  x <- d26 %>% filter(tienda == cod, format(fecha, "%u") == "7",
                      hora %in% as.numeric(names(pesos))) %>%
    mutate(w = pesos[as.character(hora)]) %>% group_by(fecha) %>%
    summarise(f = sum(flujo * w), t = sum(transacciones * w))
  ic <- prop.test(obs_t, obs_f, correct = FALSE)$conf.int   # intervalo de Wilson
  list(domingos = nrow(x), f_esp = mean(x$f), t_esp = mean(x$t),
       conv_esp = sum(x$t) / sum(x$f), pct_f = 100 * mean(x$f < obs_f),
       pct_t = 100 * mean(x$t < obs_t), max_f = max(x$f), min_t = min(x$t),
       ic = 100 * ic, carga = mean(x$f) / horas / asesores)
}
# Larcomar 2: 15:10 a 16:55 ; Jockey Plaza: 16:53 a 19:11
lar <- ventana("TB LMMIR 2", c("15" = 50/60, "16" = 55/60), 64, 4, 4, 105/60)
jck <- ventana("TB JPSU", c("16" = 7/60, "17" = 1, "18" = 1, "19" = 11/60), 552, 4, 5, 138/60)
str(lar); str(jck)

chk("17", "Larcomar 2: domingos de referencia", lar$domingos, 33, 0)
chk("17", "Larcomar 2: flujo esperado", lar$f_esp, 83.2, 1)
chk("17", "Larcomar 2: percentil del flujo observado", lar$pct_f, 21, 0)
chk("17", "Larcomar 2: ventas esperadas", lar$t_esp, 4.4, 1)
chk("17", "Larcomar 2: percentil de las ventas observadas", lar$pct_t, 36, 0)
chk("17", "Larcomar 2: conversion esperada, %", 100 * lar$conv_esp, 5.26, 2)
chk("17", "Larcomar 2: limite inferior Wilson, %", lar$ic[1], 2.46, 2)
chk("17", "Larcomar 2: limite superior Wilson, %", lar$ic[2], 15.00, 2)
chk("17", "Larcomar 2: ingresos por asesor y hora, sistema", lar$carga, 11.9, 1)
chk("17", "Jockey Plaza: flujo esperado", jck$f_esp, 331.2, 1)
chk("17", "Jockey Plaza: ventas esperadas", jck$t_esp, 19.7, 1)
chk("17", "Jockey Plaza: conversion esperada, %", 100 * jck$conv_esp, 5.95, 2)
chk("17", "Jockey Plaza: limite inferior Wilson, %", jck$ic[1], 0.28, 2)
chk("17", "Jockey Plaza: limite superior Wilson, %", jck$ic[2], 1.85, 2)
chk("17", "Jockey Plaza: ingresos por asesor y hora, sistema", jck$carga, 28.8, 1)
registrar("17", "Jockey Plaza: flujo observado (552) sobre el maximo",
          formatC(jck$max_f, format = "f", digits = 1), "< 552", jck$max_f < 552)
registrar("17", "Jockey Plaza: ventas observadas (4) bajo el minimo",
          formatC(jck$min_t, format = "f", digits = 1), "> 4", jck$min_t > 4)
chk("17", "Carga Jockey / Larcomar 2 (ambas segun sistema)", jck$carga / lar$carga, 2.4, 1)


# ==============================================================================
# 18. CURVA DE TALLAS (disponibilidad de tallas, apartado 2.3.1)
#     Stock disponible al cierre del 20/09/2026 (archivo semanal, S38) y venta
#     del 21 al 26/09/2026 (archivo diario). Zapatillas en tiendas The Box.
#     Cobertura de demanda de un modelo-color (PADRE) en una tienda = suma, sobre
#     sus tallas con stock disponible >= 1, de la participacion de cada talla en
#     la demanda de la cadena (genero x sistema de talla, unidades S35-S38),
#     reescalada a las tallas que existen para ese modelo-color.
#     Las columnas se toman por posicion para no depender de los acentos.
# ==============================================================================
titulo("18. CURVA DE TALLAS (stock al 20/09/2026 y venta del 21 al 26/09/2026)")

if (file.exists(RUTA_STOCK_SEM) && file.exists(RUTA_STOCK)) {
  num <- function(x) suppressWarnings(as.numeric(trimws(x)))
  sistema <- function(t) {
    v <- suppressWarnings(as.numeric(gsub(",", ".", t)))
    ifelse(!is.na(v) & v >= 30, "EU", "US")
  }

  ws <- read_excel(RUTA_STOCK_SEM, sheet = "Sheet1", col_types = "text")
  pos_w <- c(4, 9, 18, 19, 30, 34, 42, 47, 48, 55, 58, 61)
  cat("Columnas del archivo semanal (ALU, STK DISPO, Venta Ult 4 Sem Und y S/, TIENDA,\n",
      "CANAL, PADRE, Talla, TempVenta, CATEGORIA, GENERO_F, MF):\n")
  print(names(ws)[pos_w])
  ws <- ws[, pos_w]
  names(ws) <- c("alu", "disp", "u4", "v4", "tienda", "canal", "padre", "talla",
                 "temp_venta", "categoria", "genero", "mf")
  ws <- ws %>% mutate(across(c(disp, u4, v4), num), talla = trimws(talla))

  ds <- read_excel(RUTA_STOCK, sheet = "Sheet1", col_types = "text")
  pos_d <- c(3, 8, 21:26, 35, 39, 47, 52)
  cat("Columnas del archivo diario (ALU, QtyDispo, D2 a D7 Und, StoreCode, Canal, Padre, SizeCode):\n")
  print(names(ds)[pos_d])
  ds <- ds[, pos_d]
  names(ds) <- c("alu", "disp26", paste0("d", 2:7), "tienda", "canal", "padre", "talla")
  ds <- ds %>% mutate(across(c(disp26, d2:d7), num), talla = trimws(talla))
  ds$u6 <- rowSums(ds[, paste0("d", 2:7)], na.rm = TRUE)

  tb_stk <- ws %>% filter(canal == "ST-The Box")
  sn <- tb_stk %>% filter(categoria == "SNEAKER") %>% mutate(sis = sistema(talla))
  t14 <- tb_stk %>% filter(!tienda %in% EXCLUIDAS)
  sn14 <- sn %>% filter(!tienda %in% EXCLUIDAS)
  chk("18", "Zapatillas: % de la venta en S/ de S35-S38 (14 tiendas)",
      100 * sum(sn14$v4, na.rm = TRUE) / sum(t14$v4, na.rm = TRUE), 90.5, 1)

  # Participacion de cada talla en la demanda de la cadena
  share <- sn %>% group_by(genero, sis, talla) %>% summarise(u = sum(u4, na.rm = TRUE)) %>%
    mutate(u = pmax(u, 0)) %>% group_by(genero, sis) %>% mutate(sh = u / sum(u)) %>% ungroup()

  # Curva de cada modelo-color: todas sus tallas en cualquier ubicacion
  gen <- sn %>% distinct(padre, genero)
  cur <- bind_rows(ws %>% select(padre, talla, alu), ds %>% select(padre, talla, alu)) %>%
    distinct(alu, .keep_all = TRUE)
  cp <- cur %>% filter(padre %in% sn$padre) %>% left_join(gen, by = "padre") %>%
    mutate(sis = sistema(talla)) %>%
    left_join(share %>% select(genero, sis, talla, sh), by = c("genero", "sis", "talla")) %>%
    mutate(sh = coalesce(sh, 0)) %>% group_by(padre) %>% mutate(p = sh / sum(sh)) %>%
    ungroup() %>% filter(!is.na(p))
  chk("18", "Modelos-color de zapatilla con curva valida", n_distinct(cp$padre), 462, 0)

  # Malla tienda x modelo-color exhibido (>= 1 unidad disponible) x talla
  st <- sn %>% select(tienda, alu, disp, u4)
  exh <- sn %>% filter(disp >= 1) %>% distinct(tienda, padre)
  grid <- exh %>%
    inner_join(cp %>% select(padre, alu, talla, p), by = "padre", relationship = "many-to-many") %>%
    left_join(st, by = c("tienda", "alu")) %>%
    left_join(ds %>% filter(canal == "ST-The Box") %>% select(tienda, alu, u6, disp26),
              by = c("tienda", "alu")) %>%
    mutate(across(c(disp, u4, u6, disp26), ~ coalesce(.x, 0)), av = as.integer(disp >= 1))

  u4p <- sn %>% group_by(padre) %>% summarise(dp = pmax(sum(u4, na.rm = TRUE), 0))
  shs <- sn %>% group_by(tienda) %>% summarise(s = sum(u4, na.rm = TRUE)) %>% mutate(s = s / sum(s))
  sp <- grid %>% group_by(tienda, padre) %>%
    summarise(cov = sum(p * av), n_tallas = n(), n_av = sum(av), u4 = sum(u4), u6 = sum(u6)) %>%
    ungroup() %>% left_join(u4p, by = "padre") %>% left_join(shs, by = "tienda") %>%
    mutate(ed = dp * s) %>%
    left_join(sn %>% distinct(padre, temp_venta, mf), by = "padre")
  s14 <- sp %>% filter(!tienda %in% EXCLUIDAS)

  chk("18", "Tienda x modelo-color exhibidos, catorce tiendas", nrow(s14), 1782, 0)
  chk("18", "Cobertura de demanda ponderada, %", 100 * sum(s14$cov * s14$ed) / sum(s14$ed), 73.6, 1)
  chk("18", "% de modelos exhibidos con curva incompleta", 100 * mean(s14$cov < 0.999), 80.8, 1)
  chk("18", "% de modelos exhibidos con cobertura < 50%", 100 * mean(s14$cov < 0.5), 31.5, 1)
  chk("18", "% de la demanda esperada sin talla disponible",
      100 * sum((1 - s14$cov) * s14$ed) / sum(s14$ed), 26.4, 1)

  # Tallas faltantes: agotadas tras venderse o sin venta reciente; reposicion
  miss <- grid %>% filter(!tienda %in% EXCLUIDAS, av == 0) %>%
    left_join(s14 %>% select(tienda, padre, ed), by = c("tienda", "padre")) %>% mutate(w = p * ed)
  chk("18", "Demanda faltante en tallas vendidas en S35-S38, %",
      100 * sum(miss$w[miss$u4 > 0]) / sum(miss$w), 52.0, 1)
  chk("18", "Demanda faltante repuesta al 26/09, %",
      100 * sum(miss$w[miss$disp26 >= 1]) / sum(miss$w), 20.9, 1)

  # Definicion alternativa: tallas nucleo (80% de la demanda de genero x sistema)
  core <- share %>% group_by(genero, sis) %>% arrange(desc(sh), .by_group = TRUE) %>%
    mutate(prev = lag(cumsum(sh), default = 0)) %>% filter(prev < 0.8) %>%
    ungroup() %>% select(genero, sis, talla) %>% mutate(core = TRUE)
  cc <- grid %>% filter(!tienda %in% EXCLUIDAS) %>% left_join(gen, by = "padre") %>%
    mutate(sis = sistema(talla)) %>% inner_join(core, by = c("genero", "sis", "talla")) %>%
    group_by(tienda, padre) %>% summarise(nc = n(), nav = sum(av))
  chk("18", "Tallas nucleo: modelos exhibidos con al menos una", nrow(cc), 1764, 0)
  chk("18", "Tallas nucleo: % al que le falta al menos una", 100 * mean(cc$nav < cc$nc), 73.2, 1)
  chk("18", "Tallas nucleo: % al que le faltan dos o mas", 100 * mean(cc$nc - cc$nav >= 2), 49.0, 1)

  t26 <- s14 %>% filter(substr(temp_venta, 1, 2) == "26")
  chk("18", "Temporada 2026: % de la demanda esperada", 100 * sum(t26$ed) / sum(s14$ed), 54.4, 1)
  chk("18", "Temporada 2026: cobertura ponderada, %", 100 * sum(t26$cov * t26$ed) / sum(t26$ed), 68.3, 1)

  # Validacion hacia adelante: venta del 21 al 26/09 del modelo en la tienda
  # frente a su cobertura al 20/09, con efectos fijos de tienda y de modelo y
  # control por su venta en la tienda en S35-S38. Se depuran antes los grupos
  # sin informacion (modelos en una sola tienda o sin venta) para que el
  # resultado no dependa del tratamiento por defecto de cada version de fixest.
  est <- s14 %>% mutate(u6 = pmax(u6, 0), u4 = pmax(u4, 0),
                        lcov = log(pmax(cov, 0.01)), lu4 = log1p(u4),
                        temp26 = substr(temp_venta, 1, 2) == "26")
  depurar_fe <- function(d) {
    repeat {
      n0 <- nrow(d)
      d <- d %>% group_by(padre) %>% filter(n() >= 2, sum(u6) > 0) %>%
        group_by(tienda) %>% filter(sum(u6) > 0) %>% ungroup()
      if (nrow(d) == n0) return(d)
    }
  }
  est <- depurar_fe(est)
  mp  <- fepois(u6 ~ lcov + lu4 | tienda + padre, data = est, cluster = ~tienda, fixef.rm = "none")
  print(summary(mp))
  b_cov <- coef(mp)["lcov"]
  chk("18", "Poisson: observaciones", nobs(mp), 1457, 0)
  chk("18", "Poisson: elasticidad venta - cobertura", b_cov, 0.421, 3)
  chk("18", "Poisson: error estandar", se(mp)["lcov"], 0.094, 3)
  chk_menor("18", "Poisson: valor p", pvalue(mp)["lcov"], 0.001)
  mp2 <- fepois(u6 ~ lcov + lu4 | tienda + padre, data = est, cluster = ~padre, fixef.rm = "none")
  chk_menor("18", "Poisson con error agrupado por modelo: valor p", pvalue(mp2)["lcov"], 0.01)
  e26 <- depurar_fe(est %>% filter(temp26))
  mp3 <- fepois(u6 ~ lcov + lu4 | tienda + padre, data = e26, cluster = ~tienda, fixef.rm = "none")
  chk("18", "Poisson, solo temporada 2026: elasticidad", coef(mp3)["lcov"], 0.525, 3)
  eREG <- depurar_fe(est %>% filter(mf == "REGULAR"))
  mp4 <- fepois(u6 ~ lcov + lu4 | tienda + padre, data = eREG, cluster = ~tienda, fixef.rm = "none")
  chk("18", "Poisson, solo productos regulares: elasticidad", coef(mp4)["lcov"], 0.405, 3)

  est$tramo <- cut(est$cov, c(-0.01, 0.5, 0.75, 0.999, 1.01),
                   labels = c("menor a 50%", "50 a 75%", "75 a 99%", "completa"))
  mt <- fepois(u6 ~ i(tramo, ref = "completa") + lu4 | tienda + padre, data = est,
               cluster = ~tienda, fixef.rm = "none")
  print(coeftable(mt))
  b50 <- coef(mt)["tramo::menor a 50%"]
  chk("18", "Cobertura < 50%: coeficiente frente a curva completa", b50, -0.407, 3)
  chk_menor("18", "Cobertura < 50%: valor p", pvalue(mt)["tramo::menor a 50%"], 0.001)
  registrar("18", "Cobertura 50 a 99%: sin diferencia significativa (p > 0,05)",
            formatC(min(pvalue(mt)[c("tramo::50 a 75%", "tramo::75 a 99%")]), format = "f", digits = 3),
            "> 0.05", all(pvalue(mt)[c("tramo::50 a 75%", "tramo::75 a 99%")] > 0.05))

  # Orden de magnitud en conversion. Supone que la venta perdida de un modelo
  # no se recupera en otro modelo, por lo que cada cifra es una cota superior.
  perd_b   <- sum(s14$ed * (1 - s14$cov^b_cov)) / sum(s14$ed)
  perd_50  <- (1 - exp(b50)) * sum(s14$ed[s14$cov < 0.5]) / sum(s14$ed)
  u_sem    <- sum(sn14$u4, na.rm = TRUE) / 4
  pp <- function(perd, und_trx) 100 * (perd / (1 - perd) * u_sem / und_trx) / (r14$flujo / 33)
  chk("18", "Venta de zapatillas perdida con la elasticidad estimada, %", 100 * perd_b, 14.1, 1)
  chk("18", "Venta perdida usando solo el tramo < 50%, %", 100 * perd_50, 6.5, 1)
  chk("18", "Cota en conversion, elasticidad, 1 unidad por transaccion, pp", pp(perd_b, 1), 0.99, 2)
  chk("18", "Cota en conversion, elasticidad, UPT de red, pp", pp(perd_b, r14$upt), 0.69, 2)
  chk("18", "Cota en conversion, tramo < 50%, 1 unidad por transaccion, pp", pp(perd_50, 1), 0.42, 2)
  chk("18", "Cota en conversion, tramo < 50%, UPT de red, pp", pp(perd_50, r14$upt), 0.29, 2)

  # Corte transversal: cobertura de la tienda frente a su conversion
  ci <- s14 %>% group_by(tienda) %>% summarise(cob = sum(cov * ed) / sum(ed)) %>%
    inner_join(ct %>% select(tienda, conv), by = "tienda")
  sc <- cor.test(ci$cob, ci$conv, method = "spearman")
  chk("18", "Spearman cobertura de tienda - conversion", sc$estimate, -0.090, 3)
  registrar("18", "Spearman cobertura - conversion: no significativo",
            formatC(sc$p.value, format = "f", digits = 3), "> 0.05", sc$p.value > 0.05)

  # Asignacion de stock por tienda: semanas de cobertura de zapatillas
  sem_cob <- sn14 %>% group_by(tienda) %>%
    summarise(stock = sum(disp, na.rm = TRUE), u4 = sum(u4, na.rm = TRUE)) %>%
    mutate(semanas = stock / (u4 / 4)) %>% arrange(semanas)
  print(as.data.frame(sem_cob))
  sw_ <- cor.test(sem_cob$semanas, sem_cob$u4, method = "spearman")
  chk("18", "Semanas de cobertura, minimo (Mall Plaza Trujillo)", min(sem_cob$semanas), 5.07, 2)
  chk("18", "Semanas de cobertura, maximo (La Rambla Brasil)", max(sem_cob$semanas), 22.99, 2)
  chk("18", "Semanas de cobertura, Jockey Plaza", sem_cob$semanas[sem_cob$tienda == "TB JPSU"], 5.82, 2)
  chk("18", "Spearman semanas de cobertura - venta", sw_$estimate, -0.802, 3)
  chk_menor("18", "Spearman semanas de cobertura - venta: valor p", sw_$p.value, 0.001)
} else {
  cat("  Archivos de stock no encontrados; bloque omitido.\n")
}


# ==============================================================================
# 19. DOTACION PLANIFICADA FRENTE AL FLUJO (causa "dotacion no ajustada a la
#     franja punta", apartado 2.3.1). Hoja "Por hora" del archivo de horarios:
#     personas planificadas por tienda, semana ISO, dia y hora (semanas 34 a
#     38), y hojas "Cadena por semana" y "Turnos por semana". Es horario
#     planificado, no asistencia. Solo se usan esas hojas agregadas; el script
#     no lee datos individuales del personal.
#     Flujo esperado = promedio de la misma tienda, dia de la semana y hora en
#     las semanas 2 a 34 de 2026.
# ==============================================================================
titulo("19. DOTACION PLANIFICADA FRENTE AL FLUJO (semanas 34 a 38 de 2026)")

if (file.exists(RUTA_HOR)) {
  ph <- read_excel(RUTA_HOR, sheet = "Por hora")
  hcols <- grep("^[0-9]{2}:00$", names(ph), value = TRUE)
  ph <- ph[, c(names(ph)[c(1, 3, 4)], hcols)]
  names(ph)[1:3] <- c("tienda", "semana", "dia")
  ph <- ph %>% tidyr::pivot_longer(all_of(hcols), names_to = "h", values_to = "staff") %>%
    mutate(hora = as.numeric(substr(h, 1, 2)), staff = as.numeric(staff),
           dow = match(dia, c("L", "M", "X", "J", "V", "S", "D"))) %>%
    filter(grepl("^TB", tienda), !tienda %in% EXCLUIDAS, hora >= 10, hora <= 21)

  d14d <- d14 %>% mutate(dow = as.integer(format(fecha, "%u")))
  nd <- d14d %>% group_by(tienda, dow) %>% summarise(nd = n_distinct(fecha))
  fe <- d14d %>% group_by(tienda, dow, hora) %>% summarise(f = sum(flujo)) %>%
    left_join(nd, by = c("tienda", "dow")) %>% mutate(fe = f / nd)
  X <- ph %>% left_join(fe %>% select(tienda, dow, hora, fe), by = c("tienda", "dow", "hora")) %>%
    mutate(franja = case_when(hora == 10 ~ "Apertura", hora <= 17 ~ "Valle", TRUE ~ "Punta"))

  fr <- X %>% group_by(semana, franja) %>% summarise(staff = sum(staff), fe = sum(fe)) %>%
    group_by(semana) %>% mutate(pct_staff = 100 * staff / sum(staff),
                                pct_flujo = 100 * fe / sum(fe), vis_persona = fe / staff) %>% ungroup()
  print(as.data.frame(fr))
  v <- function(sem, fj, col) fr[[col]][fr$semana == sem & fr$franja == fj]
  chk("19", "Punta: % del flujo esperado", v(34, "Punta", "pct_flujo"), 46.88, 2)
  chk("19", "Punta: % de las horas-persona, semana 34", v(34, "Punta", "pct_staff"), 32.22, 2)
  chk("19", "Punta: % de las horas-persona, semana 38", v(38, "Punta", "pct_staff"), 33.04, 2)
  chk("19", "Visitas por persona y hora en punta, semana 34", v(34, "Punta", "vis_persona"), 11.26, 2)
  chk("19", "Visitas por persona y hora en valle, semana 34", v(34, "Valle", "vis_persona"), 6.22, 2)
  chk("19", "Visitas por persona y hora en punta, semana 38", v(38, "Punta", "vis_persona"), 11.18, 2)
  chk("19", "Visitas por persona y hora en valle, semana 38", v(38, "Valle", "vis_persona"), 6.57, 2)
  pt <- X %>% filter(semana == 34, franja != "Apertura") %>% group_by(tienda, franja) %>%
    summarise(v = sum(fe) / sum(staff)) %>% tidyr::pivot_wider(names_from = franja, values_from = v)
  chk("19", "Tiendas con mas visitas por persona en punta que en valle", sum(pt$Punta > pt$Valle), 14, 0)
  chk("19", "Valle: % del flujo esperado", v(34, "Valle", "pct_flujo"), 50.75, 2)
  chk("19", "Punta: % de las horas-persona, minimo de las semanas 34 a 38",
      min(fr$pct_staff[fr$franja == "Punta"]), 31.98, 2)
  chk("19", "Punta: % de las horas-persona, maximo de las semanas 34 a 38",
      max(fr$pct_staff[fr$franja == "Punta"]), 33.59, 2)
  registrar("19", "Valle: % de las horas-persona supera 60% en todas las semanas",
            formatC(min(fr$pct_staff[fr$franja == "Valle"]), format = "f", digits = 2), "> 60",
            min(fr$pct_staff[fr$franja == "Valle"]) > 60)

  # Perfil por hora, semanas 34 a 38 en conjunto (Figura de dotacion)
  hh <- X %>% group_by(hora) %>% summarise(staff = sum(staff), fe = sum(fe)) %>%
    mutate(pct_flujo = 100 * fe / sum(fe), pct_staff = 100 * staff / sum(staff), vis_persona = fe / staff)
  print(as.data.frame(hh))
  chk("19", "Hora de mayor flujo esperado", hh$hora[which.max(hh$fe)], 19, 0)
  registrar("19", "Maximo de horas-persona entre las 14:00 y las 16:00",
            hh$hora[which.max(hh$staff)], "14 a 16", hh$hora[which.max(hh$staff)] %in% 14:16)
  chk("19", "19:00: % del flujo esperado del dia", hh$pct_flujo[hh$hora == 19], 13.18, 2)
  chk("19", "19:00: % de las horas-persona del dia", hh$pct_staff[hh$hora == 19], 8.00, 2)
  chk("19", "19:00: visitas por persona", hh$vis_persona[hh$hora == 19], 12.81, 2)

  # Hojas agregadas del mismo archivo (sin datos individuales)
  cs <- read_excel(RUTA_HOR, sheet = "Cadena por semana")
  pct11 <- as.numeric(cs[[ncol(cs)]])
  chk("19", "Turnos de 11 horas o mas, % del total, semana 34", pct11[cs[[1]] == 34], 4.7, 1)
  chk("19", "Turnos de 11 horas o mas, % del total, semana 38", pct11[cs[[1]] == 38], 33.1, 1)
  ts <- read_excel(RUTA_HOR, sheet = "Turnos por semana")
  top2 <- ts[[1]][order(-as.numeric(ts[[ncol(ts)]]))][1:2]
  registrar("19", "Dos turnos mas frecuentes, semanas 34 a 38", paste(top2, collapse = " y "),
            "14:00-22:00 y 10:00-19:00", identical(sort(top2), c("10:00-19:00", "14:00-22:00")))

  # Semana 34: unica semana completa comun a ambas bases
  a34 <- dep14 %>% filter(anio == 2026, semana == 34) %>%
    mutate(dow = as.integer(format(fecha, "%u"))) %>%
    inner_join(ph %>% filter(semana == 34) %>% select(tienda, dow, hora, staff),
               by = c("tienda", "dow", "hora")) %>%
    filter(staff > 0) %>% mutate(vpp = flujo / staff, conv = transacciones / flujo)
  chk("19", "Semana 34: celdas tienda-dia-hora cruzadas", nrow(a34), 1150, 0)
  terc <- a34 %>% group_by(tienda) %>% mutate(t = ntile(vpp, 3)) %>% group_by(t) %>%
    summarise(vpp = sum(flujo) / sum(staff), conv = 100 * sum(transacciones) / sum(flujo))
  print(as.data.frame(terc))
  chk_vec("19", "Semana 34: conversion por tercio de visitas por persona", terc$conv,
          c(7.60, 7.21, 6.99), 2)
  r34 <- a34 %>% filter(flujo >= 20) %>% mutate(th = paste(tienda, hora), lvpp = log(vpp))
  m34 <- feols(conv ~ lvpp | th + fecha, data = r34, cluster = ~tienda, fixef.rm = "none")
  print(summary(m34))
  chk("19", "Semana 34: coeficiente de ln(visitas por persona)", coef(m34)["lvpp"], -0.0112, 4)
  chk("19", "Semana 34: valor p (no significativo)", pvalue(m34)["lvpp"], 0.297, 3)
} else {
  cat("  Archivo de horarios no encontrado; bloque omitido.\n")
}


# ==============================================================================
# 20. METAS DEL ARBOL DE OBJETIVOS (apartado 3.1)
#     Objetivo central = conversion real + brecha atribuible a la gestion (brecha
#     frente al percentil 90 menos el componente de composicion del trafico con
#     variable instrumental). Fines: venta por visitante, productividad del
#     asesor y facturacion adicional con ese objetivo.
# ==============================================================================
titulo("20. METAS DEL ARBOL DE OBJETIVOS (Capitulo III)")

gestion   <- brecha - comp_iv / 100
meta_conv <- conv_real + gestion
fact_meta <- gestion * flujo_anual * r14$ticket
chk("20", "Objetivo central: conversion meta, %", 100 * meta_conv, 7.38, 2)
chk("20", "Objetivo central: mejora sobre la linea base, pp", 100 * gestion, 0.89, 2)
chk("20", "Fin: venta por visitante con la meta, S/", vpv * meta_conv / conv_real, 26.39, 2)
chk("20", "Fin: transacciones por asesor por semana con la meta", trx_asesor * meta_conv / conv_real, 29.9, 1)
chk("20", "Fin: transacciones anuales adicionales", gestion * flujo_anual, 14253, 0)
chk("20", "Fin: facturacion adicional, S/ millones", fact_meta / 1e6, 5.10, 2)
chk("20", "Fin: facturacion adicional, % de la venta anual", 100 * fact_meta / venta_anual, 13.7, 1)
if (exists("margen16")) chk("20", "Fin: utilidad bruta adicional, S/ millones", fact_meta * margen16 / 1e6, 2.13, 2)


# ==============================================================================
# 21. PROGRAMACION DE TURNOS POR FRANJA (apartado 4.2)
#     Para cada tienda y dia de la semana se elige cuantos turnos de cada tipo
#     programar, con los tipos de turno que la empresa ya usa (inicio desde las
#     10:00, hoja "Turnos por semana"). Programacion entera (lpSolve):
#       min  sum_h deficit_h + 0,001 * horas-persona
#       s.a. cobertura_h + deficit_h >= flujo esperado_h / 9      (carga objetivo)
#            cobertura_h >= min(2, cobertura actual_h)            (minimo operativo)
#            horas-persona de 10:00 a 21:59 <= las del plan de la semana 38
#            numero de turnos <= personas programadas ese dia en la semana 38
#     Una persona cuenta en la hora X si su turno cubre [X:00, X+1:00).
#     La asignacion de personas a turnos (jornada semanal, descansos) queda
#     a cargo del jefe de tienda y no forma parte del modelo.
# ==============================================================================
titulo("21. PROGRAMACION DE TURNOS POR FRANJA (semana 38 como base)")

if (file.exists(RUTA_HOR) && requireNamespace("lpSolve", quietly = TRUE)) {
  CARGA <- 9
  horas <- 10:21
  cat_turnos <- c("10:00-14:00", "10:00-17:00", "10:00-18:00", "10:00-19:00", "10:00-20:00",
                  "10:00-21:00", "10:00-22:00", "11:00-20:00", "11:00-21:00", "11:00-22:00",
                  "12:00-21:00", "12:00-22:00", "13:00-21:00", "13:00-22:00", "14:00-21:00",
                  "14:00-22:00", "15:00-21:00", "15:00-22:00", "16:00-22:00", "18:00-22:00")
  ts_hoja <- read_excel(RUTA_HOR, sheet = "Turnos por semana")
  registrar("21", "Tipos de turno del modelo presentes en la hoja de turnos",
            paste0(sum(cat_turnos %in% ts_hoja[[1]]), " de ", length(cat_turnos)),
            paste0(length(cat_turnos), " de ", length(cat_turnos)), all(cat_turnos %in% ts_hoja[[1]]))
  ini <- as.numeric(substr(cat_turnos, 1, 2)); fin <- as.numeric(substr(cat_turnos, 7, 8))
  A <- sapply(seq_along(cat_turnos), function(k) as.numeric(horas >= ini[k] & horas < fin[k]))
  largo <- colSums(A)

  hc <- read_excel(RUTA_HOR, sheet = "S38")
  hc <- hc %>% select(tienda = 1, L, M, X, J, V, S, D) %>%
    tidyr::pivot_longer(-tienda, names_to = "dia", values_to = "personas") %>%
    mutate(dow = match(dia, c("L", "M", "X", "J", "V", "S", "D")))

  base38 <- X %>% filter(semana == 38) %>% select(tienda, dow, hora, staff, fe)

  # Escenario A: misma dotacion diaria (turnos <= personas del dia) y horas-persona
  # no mayores a las del plan de la semana 38.
  e_p <- as.numeric(horas >= 18)
  resolver <- function(d, personas, refuerzo = FALSE) {
    d <- d[match(horas, d$hora), ]
    req <- d$fe / CARGA
    minc <- pmin(2, d$staff)
    K <- length(cat_turnos); H <- length(horas)
    if (!refuerzo) {
      # deficit frente a la carga de 9 (peso 1) y frente a una carga de 7 (peso 0,1):
      # el segundo termino reparte las horas restantes hacia las horas mas cargadas
      req7 <- d$fe / 7
      obj <- c(rep(0, K), rep(1, H), rep(0.1, H))
      M <- rbind(cbind(A, diag(H), matrix(0, H, H)), cbind(A, matrix(0, H, H), diag(H)),
                 cbind(A, matrix(0, H, 2 * H)),
                 c(largo, rep(0, 2 * H)), c(rep(1, K), rep(0, 2 * H)))
      dir <- c(rep(">=", 3 * H), "<=", "<="); rhs <- c(req, req7, minc, sum(d$staff), personas)
      sol <- lpSolve::lp("min", obj, M, dir, rhs, int.vec = 1:K)
      x <- sol$solution[1:K]; y <- 0
    }
    data.frame(hora = horas, staff_nuevo = as.vector(A %*% x) + y * e_p, fe = d$fe,
               staff = d$staff, refuerzo = y, horas_base = sum(largo * x), estado = sol$status)
  }
  correr <- function(refuerzo) base38 %>% group_by(tienda, dow) %>% group_modify(function(d, k) {
    resolver(d, hc$personas[hc$tienda == k$tienda & hc$dow == k$dow], refuerzo)
  }) %>% ungroup()
  planA <- correr(FALSE)
  est_ok <- function(p) sum(tapply(p$estado, paste(p$tienda, p$dow), max) == 0)
  chk("21", "Escenario A: problemas tienda-dia con solucion optima", est_ok(planA), 98, 0)

  franja <- function(h) ifelse(h == 10, "Apertura", ifelse(h <= 17, "Valle", "Punta"))
  resumen_plan <- function(p) p %>% mutate(franja = franja(hora)) %>% group_by(franja) %>%
    summarise(fe = sum(fe), actual = sum(staff), propuesto = sum(staff_nuevo)) %>%
    mutate(carga_actual = fe / actual, carga_propuesta = fe / propuesto,
           pct_actual = 100 * actual / sum(actual), pct_propuesto = 100 * propuesto / sum(propuesto))
  resA <- resumen_plan(planA); print(as.data.frame(resA))
  hA <- planA %>% group_by(hora) %>% summarise(fe = sum(fe), actual = sum(staff), propuesto = sum(staff_nuevo)) %>%
    mutate(carga_actual = fe / actual, carga_propuesta = fe / propuesto)
  print(as.data.frame(hA))
  g <- function(f, col) resA[[col]][resA$franja == f]
  chk("21", "Horas-persona de atencion, plan de la semana 38", sum(planA$staff), 3411, 0)
  chk("21", "Escenario A: horas-persona de atencion", sum(planA$staff_nuevo), 3330, 0)
  chk("21", "Punta: carga actual", g("Punta", "carga_actual"), 11.18, 2)
  chk("21", "Valle: carga actual", g("Valle", "carga_actual"), 6.57, 2)
  chk("21", "Escenario A: carga en la punta", g("Punta", "carga_propuesta"), 9.82, 2)
  chk("21", "Escenario A: carga en el valle", g("Valle", "carga_propuesta"), 7.32, 2)
  chk("21", "Escenario A: % de las horas-persona en la punta", g("Punta", "pct_propuesto"), 38.5, 1)
  chk("21", "19:00: carga actual", hA$carga_actual[hA$hora == 19], 12.78, 2)
  chk("21", "Escenario A: carga a las 19:00", hA$carga_propuesta[hA$hora == 19], 10.51, 2)
  chk("21", "Carga total del dia, plan actual", sum(planA$fe) / sum(planA$staff), 7.88, 2)
  chk("21", "Carga total del dia, escenario A", sum(planA$fe) / sum(planA$staff_nuevo), 8.07, 2)
  chk("21", "Escenario A: carga entre las 14:00 y las 16:00, minimo", min(hA$carga_propuesta[hA$hora %in% 14:16]), 6.59, 2)
  chk("21", "Escenario A: carga entre las 14:00 y las 16:00, maximo", max(hA$carga_propuesta[hA$hora %in% 14:16]), 7.82, 2)
  porT <- planA %>% mutate(franja = franja(hora)) %>% filter(franja == "Punta") %>%
    group_by(tienda) %>% summarise(antes = sum(fe) / sum(staff), despues = sum(fe) / sum(staff_nuevo))
  print(as.data.frame(porT))
  chk("21", "Escenario A: tiendas con menor carga en la punta", sum(porT$despues < porT$antes), 14, 0)

  # Sensibilidad: solo turnos de hasta nueve horas de presencia
  corto <- (fin - ini) <= 9
  factible_corto <- base38 %>% group_by(tienda, dow) %>% group_modify(function(d, k) {
    d <- d[match(horas, d$hora), ]; H <- length(horas); K9 <- sum(corto)
    M <- rbind(cbind(A[, corto], diag(H)), cbind(A[, corto], matrix(0, H, H)),
               c(largo[corto], rep(0, H)), c(rep(1, K9), rep(0, H)))
    sol <- lpSolve::lp("min", c(rep(0, K9), rep(1, H)), M, c(rep(">=", 2 * H), "<=", "<="),
                       c(d$fe / CARGA, pmin(2, d$staff), sum(d$staff),
                         hc$personas[hc$tienda == k$tienda & hc$dow == k$dow]), int.vec = 1:K9)
    data.frame(ok = sol$status == 0)
  }) %>% ungroup()
  chk("21", "Solo turnos de hasta 9 horas: dias-tienda sin solucion factible", sum(!factible_corto$ok), 33, 0)

  # Escenario B: refuerzo con turnos de 18:00 a 22:00 (tipo que la empresa ya usa)
  # en las tiendas piloto. Para cada dia se resuelve el modelo del escenario A con
  # r refuerzos fijos (r = 0 a 6) y, por programacion dinamica sobre los siete
  # dias, se reparte un total semanal de R refuerzos de forma optima.
  dia_r <- function(d, personas, r) {
    d <- d[match(horas, d$hora), ]
    req <- pmax(0, d$fe / CARGA - r * e_p); req7 <- pmax(0, d$fe / 7 - r * e_p)
    minc <- pmin(2, d$staff)
    K <- length(cat_turnos); H <- length(horas)
    M <- rbind(cbind(A, diag(H), matrix(0, H, H)), cbind(A, matrix(0, H, H), diag(H)),
               cbind(A, matrix(0, H, 2 * H)), c(largo, rep(0, 2 * H)), c(rep(1, K), rep(0, 2 * H)))
    sol <- lpSolve::lp("min", c(rep(0, K), rep(1, H), rep(0.1, H)), M,
                       c(rep(">=", 3 * H), "<=", "<="), c(req, req7, minc, sum(d$staff), personas),
                       int.vec = 1:K)
    list(valor = sol$objval, cob = as.vector(A %*% sol$solution[1:K]) + r * e_p, estado = sol$status)
  }
  semana_tienda <- function(tda, Rs, rmax = 6) {
    dias <- lapply(1:7, function(dd) base38 %>% filter(tienda == tda, dow == dd))
    pers <- sapply(1:7, function(dd) hc$personas[hc$tienda == tda & hc$dow == dd])
    tab <- lapply(1:7, function(dd) lapply(0:rmax, function(r) dia_r(dias[[dd]], pers[dd], r)))
    V <- sapply(1:7, function(dd) sapply(tab[[dd]], `[[`, "valor"))       # (rmax+1) x 7
    est <- all(sapply(tab, function(t) all(sapply(t, `[[`, "estado") == 0)))
    do.call(rbind, lapply(Rs, function(R) {
      # programacion dinamica: f[k, q] = mejor valor con los dias 1..k y q refuerzos
      f <- matrix(Inf, 8, R + 1); f[1, 1] <- 0; arg <- matrix(NA, 8, R + 1)
      for (k in 1:7) for (q in 0:R) for (r in 0:min(rmax, q)) {
        v <- f[k, q - r + 1] + V[r + 1, k]
        if (v < f[k + 1, q + 1]) { f[k + 1, q + 1] <- v; arg[k + 1, q + 1] <- r }
      }
      q <- which.min(f[8, ]) - 1; rs <- integer(7)
      for (k in 7:1) { rs[k] <- arg[k + 1, q + 1]; q <- q - rs[k] }
      cob <- unlist(lapply(1:7, function(dd) tab[[dd]][[rs[dd] + 1]]$cob))
      fe <- unlist(lapply(dias, function(d) d$fe[match(horas, d$hora)]))
      hh <- rep(horas, 7)
      data.frame(tienda = tda, refuerzos = R, usados = sum(rs),
                 carga_punta = sum(fe[hh >= 18]) / sum(cob[hh >= 18]),
                 carga_max = max(fe[hh >= 18] / cob[hh >= 18]), estado = est)
    }))
  }
  piloto_t <- c("TB JPSU", "TB MPTRU")
  curva <- do.call(rbind, lapply(piloto_t, function(tt) semana_tienda(tt, c(0, 7, 14, 21, 28))))
  print(curva)
  chk("21", "Escenario B: problemas resueltos con solucion optima", sum(curva$estado), 10, 0)
  cv <- function(tt, R, col) curva[[col]][curva$tienda == tt & curva$refuerzos == R]
  actual_p <- function(tt) { p <- base38 %>% filter(tienda == tt, hora >= 18); sum(p$fe) / sum(p$staff) }
  chk("21", "Jockey Plaza: carga actual en la punta", actual_p("TB JPSU"), 15.30, 2)
  chk("21", "Mall Plaza Trujillo: carga actual en la punta", actual_p("TB MPTRU"), 16.57, 2)
  esp_j <- c(13.69, 11.52, 10.04, 9.11, 8.23); esp_t <- c(15.06, 12.89, 11.26, 9.91, 9.10)
  for (i in 1:5) {
    R <- c(0, 7, 14, 21, 28)[i]
    chk("21", sprintf("Jockey Plaza: carga en la punta con %d refuerzos", R), cv("TB JPSU", R, "carga_punta"), esp_j[i], 2)
    chk("21", sprintf("Mall Plaza Trujillo: carga en la punta con %d refuerzos", R), cv("TB MPTRU", R, "carga_punta"), esp_t[i], 2)
  }
  chk("21", "Jockey Plaza: horas del plan de la semana 38 (10:00 a 21:59)", sum(base38$staff[base38$tienda == "TB JPSU"]), 382, 0)
  chk("21", "Mall Plaza Trujillo: horas del plan de la semana 38 (10:00 a 21:59)", sum(base38$staff[base38$tienda == "TB MPTRU"]), 410, 0)
  chk("21", "Jockey Plaza: 21 refuerzos, % de horas adicionales", 100 * 84 / sum(base38$staff[base38$tienda == "TB JPSU"]), 22.0, 1)
  chk("21", "Mall Plaza Trujillo: 21 refuerzos, % de horas adicionales", 100 * 84 / sum(base38$staff[base38$tienda == "TB MPTRU"]), 20.5, 1)
} else {
  cat("  Archivo de horarios o paquete lpSolve no disponible; bloque omitido.\n")
}


# ==============================================================================
# 22. REDISTRIBUCION DE TALLAS ENTRE TIENDAS (apartado 4.2)
#     Para cada articulo y talla (ALU), las tiendas con dos o mas unidades
#     disponibles ceden las que exceden de una, de modo que no rompen su propia
#     curva; cada tienda que exhibe el modelo sin esa talla requiere una unidad.
#     Regla logistica (definida con la gerencia):
#       Paso 1. Transferencia directa entre tiendas del mismo departamento.
#       Paso 2. El excedente que queda en las tiendas de Lima retorna al centro
#               de distribucion y viaja a las tiendas de otros departamentos en
#               el despacho semanal que ya existe. No hay envios directos entre
#               tiendas de departamentos distintos.
#     En cada paso se atiende primero a la tienda con mayor demanda esperada.
#     Se limita a las catorce tiendas y al stock al 20/09/2026.
# ==============================================================================
titulo("22. REDISTRIBUCION DE TALLAS ENTRE TIENDAS (corte al 20/09/2026)")

if (file.exists(RUTA_STOCK_SEM) && file.exists(RUTA_STOCK)) {
  dpto <- data.frame(
    tienda = c("TB JPSU", "TB LMMIR 2", "TB PCSI", "TB RBSB", "TB RBBR", "TB MGIND", "TB RPPUR 1", "TB MASA",
               "TB MAAQP", "TB MPAQP", "TB MACIX", "TB RPCIX 1", "TB MPTRU", "TB RPCUS", "TB RPPIU2", "TB MGCHM"),
    dpto = c(rep("Lima", 8), "Arequipa", "Arequipa", "Lambayeque", "Lambayeque", "La Libertad", "Cusco", "Piura", "Ancash"))
  chk("22", "Tiendas The Box en Lima", sum(dpto$dpto == "Lima"), 8, 0)
  chk("22", "Tiendas The Box fuera de Lima", sum(dpto$dpto != "Lima"), 8, 0)

  g14 <- grid %>% filter(!tienda %in% EXCLUIDAS) %>%
    left_join(s14 %>% select(tienda, padre, ed), by = c("tienda", "padre")) %>%
    mutate(w = p * ed)
  falt <- g14 %>% filter(av == 0) %>% select(tienda, alu, padre, w) %>% left_join(dpto, by = "tienda")
  exced <- sn14 %>% group_by(tienda, alu) %>% summarise(disp = sum(disp, na.rm = TRUE)) %>%
    ungroup() %>% filter(disp >= 2) %>% mutate(excedente = disp - 1) %>% left_join(dpto, by = "tienda")
  S <- exced %>% group_by(alu, dpto) %>% summarise(S = sum(excedente)) %>% ungroup()

  # Paso 1: mismo departamento
  p1 <- falt %>% inner_join(S, by = c("alu", "dpto")) %>% group_by(alu, dpto) %>%
    arrange(desc(w), .by_group = TRUE) %>% mutate(cubierto = row_number() <= S) %>% ungroup()
  c1 <- p1 %>% filter(cubierto) %>% mutate(via = "Mismo departamento")
  # Paso 2: excedente restante de Lima hacia otros departamentos, por el despacho semanal
  usado_lima <- c1 %>% filter(dpto == "Lima") %>% count(alu, name = "usado")
  resto <- S %>% filter(dpto == "Lima") %>% left_join(usado_lima, by = "alu") %>%
    mutate(R = S - coalesce(usado, 0L)) %>% filter(R > 0) %>% select(alu, R)
  p2 <- falt %>% filter(dpto != "Lima") %>% anti_join(c1, by = c("tienda", "alu")) %>%
    inner_join(resto, by = "alu") %>% group_by(alu) %>% arrange(desc(w), .by_group = TRUE) %>%
    mutate(cubierto = row_number() <= R) %>% ungroup()
  c2 <- p2 %>% filter(cubierto) %>% mutate(via = "Despacho semanal desde Lima")
  cubiertos <- bind_rows(c1 %>% select(tienda, alu, padre, w, dpto, via), c2 %>% select(tienda, alu, padre, w, dpto, via))

  chk("22", "Demanda faltante cubierta, mismo departamento, %", 100 * sum(c1$w) / sum(falt$w), 12.4, 1)
  chk("22", "Demanda faltante cubierta, ambos pasos, %", 100 * sum(cubiertos$w) / sum(falt$w), 28.5, 1)
  chk("22", "Unidades transferidas dentro del mismo departamento", nrow(c1), 427, 0)
  chk("22", "Unidades por el despacho semanal desde Lima", nrow(c2), 314, 0)
  chk("22", "Unidades totales", nrow(cubiertos), 741, 0)
  print(as.data.frame(cubiertos %>% count(dpto, via)))

  # Cobertura despues de la redistribucion
  cobertura <- function(cc) {
    g <- g14 %>% left_join(cc %>% transmute(tienda, alu, nuevo = 1L), by = c("tienda", "alu")) %>%
      mutate(av2 = pmax(av, coalesce(nuevo, 0L)))
    g %>% group_by(tienda, padre) %>% summarise(cov = sum(p * av), cov2 = sum(p * av2), ed = first(ed)) %>% ungroup()
  }
  s14a <- cobertura(c1); s14b <- cobertura(cubiertos)
  cob_antes <- sum(s14b$cov * s14b$ed) / sum(s14b$ed)
  cob_p1 <- sum(s14a$cov2 * s14a$ed) / sum(s14a$ed)
  cob_desp <- sum(s14b$cov2 * s14b$ed) / sum(s14b$ed)
  chk("22", "Cobertura ponderada antes, %", 100 * cob_antes, 73.6, 1)
  chk("22", "Cobertura ponderada solo con el paso 1, %", 100 * cob_p1, 76.8, 1)
  chk("22", "Cobertura ponderada despues de ambos pasos, %", 100 * cob_desp, 81.1, 1)
  chk("22", "Modelos con cobertura menor a 50%, antes, %", 100 * mean(s14b$cov < 0.5), 31.5, 1)
  chk("22", "Modelos con cobertura menor a 50%, despues, %", 100 * mean(s14b$cov2 < 0.5), 24.0, 1)
  cobt <- s14b %>% group_by(tienda) %>% summarise(a = sum(cov * ed) / sum(ed), b = sum(cov2 * ed) / sum(ed))
  print(as.data.frame(cobt %>% left_join(cubiertos %>% count(tienda, name = "recibe"), by = "tienda")))
  chk("22", "Mejor cobertura de tienda antes de redistribuir, %", 100 * max(cobt$a), 80.1, 1)
  chk("22", "Peor cobertura de tienda antes de redistribuir, %", 100 * min(cobt$a), 67.4, 1)
  chk("22", "Peor cobertura de tienda despues de redistribuir, %", 100 * min(cobt$b), 76.6, 1)

  # Asignacion por semanas de cobertura
  red_sem <- sum(sem_cob$stock) / (sum(sem_cob$u4) / 4)
  chk("22", "Semanas de cobertura de la red", red_sem, 9.72, 2)
} else {
  cat("  Archivos de stock no encontrados; bloque omitido.\n")
}


# ==============================================================================
# 23. CONTROL ESTADISTICO DE LA CONVERSION (apartado 4.4)
#     Grafico p' de Laney por tienda sobre la conversion semanal: ajusta los
#     limites del grafico p a la variacion observada entre semanas (sigma z). Fase I (limites): semanas 2 a 19 de 2026. Fase II
#     (vigilancia): semanas 20 a 34. Regla de alerta de medicion: conversion bajo
#     el limite inferior y flujo semanal sobre su propio limite superior en la
#     misma semana, la firma de una ruptura del contador.
# ==============================================================================
titulo("23. CONTROL ESTADISTICO DE LA CONVERSION (graficos p' de Laney)")

laney <- function(d) {
  f1 <- d %>% filter(semana >= 2, semana <= 19)
  pbar <- sum(f1$t) / sum(f1$f)
  z <- (f1$conv - pbar) / sqrt(pbar * (1 - pbar) / f1$f)
  sz <- mean(abs(diff(z))) / 1.128
  lf <- log(f1$f); mlf <- mean(lf); slf <- mean(abs(diff(lf))) / 1.128
  d %>% mutate(pbar = pbar, sz = sz,
               lcl = pbar - 3 * sz * sqrt(pbar * (1 - pbar) / f),
               ucl = pbar + 3 * sz * sqrt(pbar * (1 - pbar) / f),
               ucl_f = exp(mlf + 3 * slf),
               bajo = conv < lcl, alerta = conv < lcl & f > ucl_f)
}
sw_all <- semanal(d26) %>% filter(f >= 100, t >= 10) %>% arrange(tienda, semana)
cc <- sw_all %>% group_by(tienda) %>% group_modify(~ laney(.x)) %>% ungroup()
f2 <- cc %>% filter(semana >= 20)
res_cc <- f2 %>% group_by(tienda) %>% summarise(semanas = n(), bajo = sum(bajo), alertas = sum(alerta),
                                                 sz = first(sz)) %>% filter(!is.na(sz))
# Real Plaza Piura 2 abrio en la semana 20 y no tiene fase I
print(as.data.frame(res_cc))
chk("23", "Fase II: alertas de medicion en San Isidro", res_cc$alertas[res_cc$tienda == "TB PCSI"], 10, 0)
chk("23", "Fase II: alertas de medicion en Mega Plaza Chimbote", res_cc$alertas[res_cc$tienda == "TB MGCHM"], 0, 0)
chk("23", "Fase II: alertas de medicion en las catorce tiendas", sum(res_cc$alertas[!res_cc$tienda %in% EXCLUIDAS]), 0, 0)
chk("23", "Fase II: tiendas de las catorce con fase I", sum(!res_cc$tienda %in% EXCLUIDAS), 13, 0)
chk("23", "Fase II: semanas de San Isidro bajo el limite inferior", res_cc$bajo[res_cc$tienda == "TB PCSI"], 14, 0)
chk("23", "Fase II: semanas-tienda vigiladas, catorce tiendas", sum(res_cc$semanas[!res_cc$tienda %in% EXCLUIDAS]), 195, 0)
chk("23", "Sobredispersion mediana (sigma z) de las catorce tiendas", median(res_cc$sz[!res_cc$tienda %in% EXCLUIDAS]), 1.16, 2)
primera <- f2 %>% filter(tienda == "TB PCSI", alerta) %>% summarise(s = min(semana)) %>% pull(s)
chk("23", "San Isidro: primera semana con alerta", primera, 21, 0)

# Reglas de calidad de datos que se proponen como control diario
q26 <- tb %>% filter(anio == 2026, !tienda %in% EXCLUIDAS)
chk("23", "Catorce tiendas, 2026: % de transacciones en horas con flujo cero",
    100 * sum(q26$transacciones[q26$flujo == 0]) / sum(q26$transacciones), 3.91, 2)
chk("23", "Catorce tiendas, 2026: % de horas con mas transacciones que visitas",
    100 * mean(q26$transacciones > q26$flujo & q26$flujo > 0), 0.08, 2)


# ==============================================================================
# 24. METODO DE FIJACION DE METAS (apartado 4.5)
#     Meta semanal de cada tienda = venta de la misma semana del anio anterior x
#     (1 + crecimiento interanual de la tienda en las ocho semanas previas).
#     Se compara, en las semanas 10 a 34 de 2026, con la meta vigente.
# ==============================================================================
titulo("24. METODO DE FIJACION DE METAS (prueba retrospectiva 2026)")

vm <- meta %>% filter(semana >= 1, semana <= 34) %>% select(anio, semana, tienda, meta, venta)
v25w <- vm %>% filter(anio == 2025) %>% select(semana, tienda, v25 = venta)
v26w <- vm %>% filter(anio == 2026) %>% select(semana, tienda, meta, venta)
pan <- v26w %>% inner_join(v25w, by = c("semana", "tienda")) %>% filter(v25 > 0, venta > 0) %>%
  arrange(tienda, semana)
crec8 <- function(tt, s) {
  x <- pan %>% filter(tienda == tt, semana >= s - 8, semana <= s - 1)
  if (nrow(x) < 6) return(NA_real_)
  sum(x$venta) / sum(x$v25) - 1
}
bt <- pan %>% filter(semana >= 10) %>% rowwise() %>% mutate(g8 = crec8(tienda, semana)) %>% ungroup() %>%
  filter(!is.na(g8), !is.na(meta), meta > 0) %>%
  mutate(meta_m = v25 * (1 + g8), c_vig = venta / meta, c_met = venta / meta_m)
ev <- function(cmp) c(n = length(cmp), s80 = 100 * mean(cmp >= 0.8), s100 = 100 * mean(cmp >= 1),
                      med = 100 * median(cmp), err = 100 * mean(abs(cmp - 1)))
print(rbind(vigente = ev(bt$c_vig), metodo = ev(bt$c_met), metodo_5 = ev(bt$venta / (bt$meta_m * 1.05))))
chk("24", "Prueba retrospectiva: semanas-tienda evaluadas", nrow(bt), 300, 0)
chk("24", "Meta vigente: % semanas-tienda >= 80%", ev(bt$c_vig)["s80"], 36.3, 1)
chk("24", "Metodo: % semanas-tienda >= 80%", ev(bt$c_met)["s80"], 72.3, 1)
chk("24", "Meta vigente: % semanas-tienda >= 100%", ev(bt$c_vig)["s100"], 4.7, 1)
chk("24", "Metodo: % semanas-tienda >= 100%", ev(bt$c_met)["s100"], 46.3, 1)
chk("24", "Meta vigente: cumplimiento mediano, %", ev(bt$c_vig)["med"], 71.6, 1)
chk("24", "Metodo: cumplimiento mediano, %", ev(bt$c_met)["med"], 97.2, 1)
chk("24", "Meta vigente: desviacion absoluta media del cumplimiento, pp", ev(bt$c_vig)["err"], 28.4, 1)
chk("24", "Metodo: desviacion absoluta media del cumplimiento, pp", ev(bt$c_met)["err"], 25.4, 1)
chk("24", "Metodo con exigencia de 5%: % semanas-tienda >= 80%", ev(bt$venta / (bt$meta_m * 1.05))["s80"], 68.7, 1)
chk("24", "Metodo con exigencia de 5%: % semanas-tienda >= 100%", ev(bt$venta / (bt$meta_m * 1.05))["s100"], 38.0, 1)
chk("24", "Metodo con exigencia de 5%: cumplimiento mediano, %", ev(bt$venta / (bt$meta_m * 1.05))["med"], 92.6, 1)
chk("24", "Metodo con exigencia de 5%: desviacion absoluta media, pp", ev(bt$venta / (bt$meta_m * 1.05))["err"], 24.9, 1)


# ==============================================================================
# 25. EJEMPLO DE HORARIO: JOCKEY PLAZA, MIERCOLES Y SABADO (apartado 4.1.3)
#     Turnos que propone el modelo del bloque 21 para el sabado de la semana 38,
#     con la misma dotacion y con los refuerzos que la programacion dinamica
#     asigna al sabado cuando el total semanal es de 21.
# ==============================================================================
titulo("25. EJEMPLO DE HORARIO: JOCKEY PLAZA, MIERCOLES Y SABADO DE LA SEMANA 38")

if (exists("dia_r") && exists("semana_tienda")) {
  turnos_dia <- function(d, personas, r) {
    d <- d[match(horas, d$hora), ]
    req <- pmax(0, d$fe / CARGA - r * e_p); req7 <- pmax(0, d$fe / 7 - r * e_p)
    K <- length(cat_turnos); H <- length(horas)
    M <- rbind(cbind(A, diag(H), matrix(0, H, H)), cbind(A, matrix(0, H, H), diag(H)),
               cbind(A, matrix(0, H, 2 * H)), c(largo, rep(0, 2 * H)), c(rep(1, K), rep(0, 2 * H)))
    sol <- lpSolve::lp("min", c(rep(0, K), rep(1, H), rep(0.1, H)), M,
                       c(rep(">=", 3 * H), "<=", "<="), c(req, req7, pmin(2, d$staff), sum(d$staff), personas),
                       int.vec = 1:K)
    list(x = round(sol$solution[1:K]), cob = as.vector(A %*% round(sol$solution[1:K])) + r * e_p)
  }
  # refuerzos asignados a cada dia con 21 por semana (misma programacion dinamica del bloque 21)
  reparto <- function(tda, R, rmax = 6) {
    dias <- lapply(1:7, function(dd) base38 %>% filter(tienda == tda, dow == dd))
    pers <- sapply(1:7, function(dd) hc$personas[hc$tienda == tda & hc$dow == dd])
    V <- sapply(1:7, function(dd) sapply(0:rmax, function(r) dia_r(dias[[dd]], pers[dd], r)$valor))
    f <- matrix(Inf, 8, R + 1); f[1, 1] <- 0; arg <- matrix(NA, 8, R + 1)
    for (k in 1:7) for (q in 0:R) for (r in 0:min(rmax, q)) {
      v <- f[k, q - r + 1] + V[r + 1, k]
      if (v < f[k + 1, q + 1]) { f[k + 1, q + 1] <- v; arg[k + 1, q + 1] <- r }
    }
    q <- which.min(f[8, ]) - 1; rs <- integer(7)
    for (k in 7:1) { rs[k] <- arg[k + 1, q + 1]; q <- q - rs[k] }
    rs
  }
  rs_j <- reparto("TB JPSU", 21)
  registrar("25", "Jockey Plaza: refuerzos por dia (L a D) con 21 por semana",
            paste(rs_j, collapse = " "), "3 2 1 2 2 5 6", identical(as.integer(rs_j), c(3L, 2L, 1L, 2L, 2L, 5L, 6L)))
  ejemplo <- function(dd) {
    d <- base38 %>% filter(tienda == "TB JPSU", dow == dd)
    p <- hc$personas[hc$tienda == "TB JPSU" & hc$dow == dd]
    a <- turnos_dia(d, p, 0); b <- turnos_dia(d, p, rs_j[dd])
    ej <- data.frame(hora = horas, flujo = d$fe[match(horas, d$hora)], actual = d$staff[match(horas, d$hora)],
                     misma = a$cob, refuerzo = b$cob)
    cat(sprintf("\n  Dia %d: %d personas, %d refuerzos\n", dd, p, rs_j[dd])); print(ej, digits = 3)
    cat("  Turnos con la misma dotacion:", paste(cat_turnos[a$x > 0], a$x[a$x > 0], sep = " x"), "\n")
    cat("  Turnos con refuerzo:", paste(cat_turnos[b$x > 0], b$x[b$x > 0], sep = " x"), "\n")
    pj <- horas >= 18
    list(ej = ej, p = p, horas_act = sum(ej$actual), horas_misma = sum(ej$misma),
         c_act = sum(ej$flujo[pj]) / sum(ej$actual[pj]), c_misma = sum(ej$flujo[pj]) / sum(ej$misma[pj]),
         c_ref = sum(ej$flujo[pj]) / sum(ej$refuerzo[pj]), xa = a$x, xb = b$x)
  }
  mie <- ejemplo(3); sab <- ejemplo(6)
  chk("25", "Jockey miercoles: personas programadas", mie$p, 5, 0)
  chk("25", "Jockey miercoles: horas-persona, plan y modelo", mie$horas_misma, mie$horas_act, 0)
  chk("25", "Jockey miercoles: horas-persona del plan", mie$horas_act, 48, 0)
  chk("25", "Jockey miercoles: carga en la punta, plan actual", mie$c_act, 14.37, 2)
  chk("25", "Jockey miercoles: carga en la punta, misma dotacion", mie$c_misma, 10.59, 2)
  chk("25", "Jockey miercoles: carga en la punta, con un refuerzo", mie$c_ref, 9.15, 2)
  chk("25", "Jockey miercoles: personas a las 19:00, plan actual", mie$ej$actual[mie$ej$hora == 19], 3, 0)
  chk("25", "Jockey miercoles: personas a las 19:00, modelo", mie$ej$misma[mie$ej$hora == 19], 5, 0)
  chk("25", "Jockey sabado: personas programadas", sab$p, 7, 0)
  chk("25", "Jockey sabado: carga en la punta, plan actual", sab$c_act, 14.15, 2)
  chk("25", "Jockey sabado: carga en la punta, misma dotacion", sab$c_misma, 14.15, 2)
  chk("25", "Jockey sabado: carga en la punta, con cinco refuerzos", sab$c_ref, 8.43, 2)
} else {
  cat("  Requiere el bloque 21; bloque omitido.\n")
}


# ==============================================================================
# 26. EJEMPLO DE TRANSFERENCIA DE TALLA (apartado 4.2)
#     Transferencia de mayor demanda esperada con destino Jockey Plaza desde
#     otra tienda de Lima.
# ==============================================================================
titulo("26. EJEMPLO DE TRANSFERENCIA DE TALLA")

if (exists("cubiertos")) {
  dest_j <- cubiertos %>% filter(tienda == "TB JPSU") %>% arrange(desc(w))
  ej_t <- dest_j[1, ]
  info <- sn14 %>% filter(alu == ej_t$alu) %>% left_join(dpto, by = "tienda") %>%
    select(tienda, dpto, padre, talla, disp, genero) %>% arrange(desc(disp))
  print(as.data.frame(ej_t)); print(as.data.frame(info))
  orig <- info %>% filter(dpto == "Lima", disp >= 2)
  registrar("26", "Ejemplo: modelo y talla", paste(ej_t$padre, info$talla[1]), "3ME10101430-M-WHT 8",
            ej_t$padre == "3ME10101430-M-WHT" && info$talla[1] == "8")
  registrar("26", "Ejemplo: tienda de Lima con excedente", orig$tienda[1], "TB LMMIR 2", orig$tienda[1] == "TB LMMIR 2")
  chk("26", "Ejemplo: unidades disponibles en el origen", orig$disp[1], 2, 0)
  cobT <- cobt %>% filter(tienda %in% c("TB MPTRU", "TB JPSU"))
  chk("26", "Unidades que recibe Jockey Plaza", sum(cubiertos$tienda == "TB JPSU"), 63, 0)
  chk("26", "Unidades que recibe Mall Plaza Trujillo", sum(cubiertos$tienda == "TB MPTRU"), 97, 0)
  chk("26", "Jockey Plaza: cobertura antes, %", 100 * cobT$a[cobT$tienda == "TB JPSU"], 76.9, 1)
  chk("26", "Jockey Plaza: cobertura despues, %", 100 * cobT$b[cobT$tienda == "TB JPSU"], 82.3, 1)
  chk("26", "Mall Plaza Trujillo: cobertura antes, %", 100 * cobT$a[cobT$tienda == "TB MPTRU"], 72.1, 1)
  chk("26", "Mall Plaza Trujillo: cobertura despues, %", 100 * cobT$b[cobT$tienda == "TB MPTRU"], 81.1, 1)
} else {
  cat("  Requiere el bloque 22; bloque omitido.\n")
}


# ==============================================================================
# 27. EVALUACION ECONOMICA (Capitulo V)
#     Flujo de caja incremental mensual a tres anios, construido segun las fases
#     del plan de implementacion:
#       Mes 1 (semanas 1 a 4): medicion. Regla de conteo y auditoria. Sin beneficio.
#       Meses 2 a 4 (semanas 5 a 16): piloto. Solucion 1 (turnos, protocolo y 21
#         refuerzos semanales de 18:00 a 22:00) en Jockey Plaza y Mall Plaza
#         Trujillo; solucion 2 (transferencias y reposicion por talla) en las
#         catorce tiendas.
#       Meses 5 a 36 (desde la semana 17): extension de turnos y protocolo a las
#         otras doce tiendas, sin refuerzo. Su efecto sobre la venta NO se valora
#         (supuesto conservador); solo se carga la capacitacion de sus asesores.
#     El efecto de cada solucion sobre la venta no esta medido en The Box; se
#     toma de la literatura y se evalua por escenarios, por punto de equilibrio
#     y por simulacion de Monte Carlo.
#     Se distinguen tres tipos de parametro: DATO de la empresa, RESULTADO de un
#     modelo de este trabajo y SUPUESTO (debe reemplazarse por el costo real).
# ==============================================================================
titulo("27. EVALUACION ECONOMICA (flujo mensual por fases, VAN, simulacion)")

venta_piloto <- venta_anual * sum(d14$venta[d14$tienda %in% c("TB JPSU", "TB MPTRU")]) / sum(d14$venta)
conv_piloto  <- with(d14[d14$tienda %in% c("TB JPSU", "TB MPTRU"), ], sum(transacciones) / sum(flujo))
share_zap <- if (exists("sn14")) sum(sn14$v4, na.rm = TRUE) / sum(t14$v4, na.rm = TRUE) else 0.905
mgn <- if (exists("margen16")) margen16 else 0.418
chk("27", "Venta anual de Jockey Plaza y Mall Plaza Trujillo, S/ millones", venta_piloto / 1e6, 13.93, 2)
chk("27", "Conversion de Jockey Plaza y Mall Plaza Trujillo, %", 100 * conv_piloto, 6.29, 2)

# Costo laboral por hora efectiva (regimen general): 12 sueldos + 2 gratificaciones
# + bonificacion extraordinaria de 9% sobre ellas + CTS (1 sueldo + 1/6 de gratificacion)
# + EsSalud 9% de 12 sueldos; horas efectivas = 48 h x (52 - 30/7) semanas.
SUELDO <- 2200                      # DATO: sueldo bruto mensual del asesor, sin comisiones (gerencia, 4/09/2026)
fac_sueldo <- 1.0
costo_anual_trab <- function(S) 12 * S + 2 * S + 0.09 * 2 * S + (S + S / 6) + 0.09 * 12 * S
h_ef <- 40 * (52 - 30 / 7)          # DATO: base de 40 horas semanales (gerencia)
c_hora <- function(f = fac_sueldo) costo_anual_trab(SUELDO * f) / h_ef
chk("27", "Costo laboral anual de un asesor, S/", costo_anual_trab(SUELDO), 36139, 0)
chk("27", "Horas efectivas al anio", h_ef, 1909, 0)
chk("27", "Costo laboral por hora efectiva, S/", c_hora(), 18.93, 2)
# La venta de la base horaria y la del archivo de stock estan sin IGV: la venta por unidad
# del archivo de stock equivale al precio de lista dividido entre 1,18
chk("27", "Venta por unidad / precio de lista en el archivo de stock (New Balance 530)", 355.09 / 419.90, 0.846, 3)

unid_modelo <- if (exists("cubiertos")) nrow(cubiertos) else 741
par <- list(
  h_ref_mes = 21 * 4 * 2 * 52 / 12,        # RESULTADO (bloque 21): 21 turnos de 4 h por semana y tienda piloto
  comision = 0.02,                          # DATO: comision del asesor sobre la venta (Estructura de Comisiones)
  c_unid = 8,                               # SUPUESTO no cotizado: costo por unidad transferida
  unid_ini = unid_modelo,                   # RESULTADO (bloque 22): unidades de la primera redistribucion
  unid_mes = unid_modelo,                   # SUPUESTO: unidades al mes en regimen
  analista_mes = 0, auditoria_mes = 0, observacion_mes = 0,  # personal que ya trabaja en la empresa: sin costo adicional
  inversion = 120 * 45 + 2 * 2 * 20 * c_hora(),     # SUPUESTO no cotizado: 120 h de desarrollo a S/45; capacitacion de 20 asesores
  capac_ext = 2 * 2 * 43 * c_hora(),                # capacitacion de los 43 asesores de las otras doce tiendas (mes 5)
  ir = 0.295, tasa = 0.15, meses = 36)
chk("27", "Inversion inicial, S/", par$inversion, 6915, 0)
chk("27", "Horas de refuerzo por anio, dos tiendas", par$h_ref_mes * 12, 8736, 0)
chk("27", "Costo anual del refuerzo, S/", par$h_ref_mes * 12 * c_hora(), 165416, 0)
chk("27", "Costo anual de transferencias, S/", par$c_unid * par$unid_mes * 12, 71136, 0)
chk("27", "Capacitacion de la extension (mes 5), S/", par$capac_ext, 3257, 0)

# Flujo mensual. Mes 0 = inversion.
flujo_m <- function(eA, eB, p = par, m = mgn, cu = p$c_unid, um = p$unid_mes, frente = "AB") {
  t <- 1:p$meses
  act <- as.numeric(t >= 2)                                   # soluciones activas desde el mes 2
  benA <- if (grepl("A", frente)) eA * venta_piloto / 12 * act else 0 * t
  benB <- if (grepl("B", frente)) eB * venta_anual * share_zap / 12 * act else 0 * t
  ben <- (benA + benB) * (m - p$comision)
  cA <- if (grepl("A", frente)) p$h_ref_mes * c_hora() * act + ifelse(t == 5, p$capac_ext, 0) else 0 * t
  unid <- ifelse(t == 2, p$unid_ini, um) * act
  cB <- if (grepl("B", frente)) cu * unid else 0 * t
  op <- ben - cA - cB
  neto <- op * (1 - p$ir)
  inv <- if (frente == "B") 40 * 45 else if (frente == "A") p$inversion - 40 * 45 else p$inversion
  c(-inv, neto)
}
tasa_mes <- function(r) (1 + r)^(1 / 12) - 1
van <- function(f, r = par$tasa) sum(f / (1 + tasa_mes(r))^(seq_along(f) - 1))
payback <- function(f) { ac <- cumsum(f); if (all(ac < 0)) NA else min(which(ac >= 0)) - 1 }

# Efecto de la solucion 2: elasticidad de la venta a la cobertura (bloque 18) aplicada al alza
# de cobertura que produce la redistribucion (bloque 22). Base: 0,42; optimista: 0,53 (extremo
# superior de las especificaciones alternativas).
cb0 <- if (exists("cob_antes")) cob_antes else 0.736
cb1 <- if (exists("cob_desp")) cob_desp else 0.811
eB_base <- 0.421 * log(cb1 / cb0)
eB_alto <- 0.53 * log(cb1 / cb0)
chk("27", "Efecto base de la solucion 2 sobre la venta de zapatillas, %", 100 * eB_base, 4.1, 1)
chk("27", "Efecto optimista de la solucion 2, %", 100 * eB_alto, 5.2, 1)
esc <- data.frame(escenario = c("Sin efecto", "Conservador", "Base", "Optimista"),
                  eA = c(0, 0.01, 0.033, 0.045), eB = c(0, 0.01, eB_base, eB_alto))
esc$van  <- mapply(function(a, b) van(flujo_m(a, b)), esc$eA, esc$eB)
esc$vanA <- mapply(function(a, b) van(flujo_m(a, b, frente = "A")), esc$eA, esc$eB)
esc$vanB <- mapply(function(a, b) van(flujo_m(a, b, frente = "B")), esc$eA, esc$eB)
print(esc, digits = 6)
chk("27", "VAN, escenario sin efecto, S/", esc$van[1], -401633, 0)
chk("27", "VAN, escenario conservador, S/", esc$van[2], -87840, 0)
chk("27", "VAN, escenario base, S/", esc$van[3], 812652, 0)
chk("27", "VAN, escenario optimista, S/", esc$van[4], 1158839, 0)
chk("27", "VAN de la solucion 1, escenario base, S/", esc$vanA[3], 21853, 0)
chk("27", "VAN de la solucion 2, escenario base, S/", esc$vanB[3], 790799, 0)
chk("27", "VAN de la solucion 1, escenario conservador, S/", esc$vanA[2], -189773, 0)
chk("27", "VAN de la solucion 2, escenario conservador, S/", esc$vanB[2], 101933, 0)

# Flujo mensual del escenario base, por fase
fb <- flujo_m(0.033, eB_base)
cat("Flujo neto mensual del escenario base, meses 0 a 12:\n"); print(round(fb[1:13]))
chk("27", "Escenario base: flujo neto del mes 1 (medicion), S/", fb[2], 0, 0)
chk("27", "Escenario base: flujo neto mensual en el piloto (meses 2 a 4), S/", fb[3], 29092, 0)
chk("27", "Escenario base: flujo neto del mes 5 (inicio de la extension), S/", fb[6], 26796, 0)
chk("27", "Escenario base: flujo neto del anio 1, S/", sum(fb[2:13]), 317711, 0)
chk("27", "Escenario base: flujo neto de los anios 2 y 3, S/", sum(fb[14:25]), 349099, 0)
chk("27", "Escenario base: mes en que se recupera la inversion", payback(fb), 2, 0)

detalle <- function(eA, eB, t) {
  act <- as.numeric(t >= 2)
  benA <- eA * venta_piloto / 12 * act; benB <- eB * venta_anual * share_zap / 12 * act
  c(ventaA = benA, ventaB = benB, margen = (benA + benB) * mgn, comision = (benA + benB) * par$comision,
    refuerzo = par$h_ref_mes * c_hora() * act,
    transf = par$c_unid * ifelse(t == 2, par$unid_ini, par$unid_mes) * act,
    capac = ifelse(t == 5, par$capac_ext, 0))
}
op_mes <- function(d) unname(d["margen"] - sum(d[c("comision", "refuerzo", "transf", "capac")]))
m2 <- detalle(0.033, eB_base, 2); m5 <- detalle(0.033, eB_base, 5)
print(round(rbind(mes2 = m2, mes5 = m5)))
chk("27", "Escenario base: venta adicional mensual de la solucion 1, S/", m2["ventaA"], 38310, 0)
chk("27", "Escenario base: venta adicional mensual de la solucion 2, S/", m2["ventaB"], 114898, 0)
chk("27", "Escenario base: venta adicional mensual total, S/", m2["ventaA"] + m2["ventaB"], 153209, 0)
chk("27", "Escenario base: margen bruto mensual adicional, S/", m2["margen"], 64041, 0)
chk("27", "Escenario base: comision mensual, S/", m2["comision"], 3064, 0)
chk("27", "Escenario base: refuerzo mensual, S/", m2["refuerzo"], 13785, 0)
chk("27", "Escenario base: transferencias mensuales, S/", m2["transf"], 5928, 0)
chk("27", "Escenario base: utilidad operativa mensual en regimen, S/", op_mes(m2), 41265, 0)
chk("27", "Escenario base: utilidad operativa del mes 5, S/", op_mes(m5), 38008, 0)
chk("27", "Escenario base: impuesto a la renta mensual en regimen, S/", op_mes(m2) * par$ir, 12173, 0)
chk("27", "Escenario base: impuesto a la renta del mes 5, S/", op_mes(m5) * par$ir, 11212, 0)
D1 <- sapply(1:12, function(t) detalle(0.033, eB_base, t)); y1 <- rowSums(D1)
y1["operativa"] <- y1["margen"] - sum(y1[c("comision", "refuerzo", "transf", "capac")])
chk("27", "Anio 1: venta adicional de la solucion 1, S/", y1["ventaA"], 421415, 0)
chk("27", "Anio 1: venta adicional de la solucion 2, S/", y1["ventaB"], 1263879, 0)
chk("27", "Anio 1: margen bruto adicional, S/", y1["margen"], 704456, 0)
chk("27", "Anio 1: comision, S/", y1["comision"], 33706, 0)
chk("27", "Anio 1: refuerzo, S/", y1["refuerzo"], 151631, 0)
chk("27", "Anio 1: transferencias, S/", y1["transf"], 65208, 0)
chk("27", "Anio 1: utilidad operativa, S/", y1["operativa"], 450655, 0)
chk("27", "Anio 1: impuesto a la renta, S/", y1["operativa"] * par$ir, 132943, 0)

# Anio completo en regimen
dvA_b <- 0.033 * venta_piloto; dvB_b <- eB_base * venta_anual * share_zap
er <- c(venta = dvA_b + dvB_b, margen = (dvA_b + dvB_b) * mgn, comision = (dvA_b + dvB_b) * par$comision,
        refuerzo = par$h_ref_mes * 12 * c_hora(), transferencias = par$c_unid * par$unid_mes * 12)
er["operativa"] <- er["margen"] - sum(er[c("comision", "refuerzo", "transferencias")])
er["ir"] <- er["operativa"] * par$ir; er["neto"] <- er["operativa"] - er["ir"]
print(round(er))
chk("27", "Anio en regimen: venta adicional de la solucion 1, S/", dvA_b, 459726, 0)
chk("27", "Anio en regimen: venta adicional de la solucion 2, S/", dvB_b, 1378777, 0)
chk("27", "Anio en regimen: margen bruto adicional, S/", er["margen"], 768498, 0)
chk("27", "Anio en regimen: comision, S/", er["comision"], 36770, 0)
chk("27", "Anio en regimen: costos recurrentes (refuerzo y transferencias), S/", er["refuerzo"] + er["transferencias"], 236552, 0)
chk("27", "Anio en regimen: utilidad operativa, S/", er["operativa"], 495176, 0)
chk("27", "Anio en regimen: impuesto a la renta, S/", er["ir"], 146077, 0)
chk("27", "Anio en regimen: flujo neto, S/", er["neto"], 349099, 0)

# Objetivo central del Capitulo III
dv_base <- dvA_b + dvB_b
alza <- dv_base / venta_anual
conv_obj <- conv_real * (1 + alza)
chk("27", "Objetivo: venta adicional anual, S/ millones", dv_base / 1e6, 1.84, 2)
chk("27", "Objetivo: venta adicional, % de la venta del canal", 100 * alza, 5.0, 1)
chk("27", "Objetivo: conversion equivalente, %", 100 * conv_obj, 6.81, 2)
chk("27", "Objetivo: alza equivalente, pp", 100 * (conv_obj - conv_real), 0.32, 2)
chk("27", "Objetivo: alza como % de la brecha de 1,12 pp", 100 * (conv_obj - conv_real) / 0.0112, 29, 0)
chk("27", "Objetivo: utilidad bruta adicional, S/ millones", dv_base * mgn / 1e6, 0.77, 2)
chk("27", "Objetivo: venta por visitante, S/", vpv * (1 + alza), 24.35, 2)
chk("27", "Objetivo: transacciones por asesor y semana", trx_asesor * (1 + alza), 27.6, 1)
dv_opt <- 0.045 * venta_piloto + eB_alto * venta_anual * share_zap
dv_con <- 0.01 * venta_piloto + 0.01 * venta_anual * share_zap
chk("27", "Escenario conservador: conversion equivalente, %", 100 * conv_real * (1 + dv_con / venta_anual), 6.57, 2)
chk("27", "Escenario optimista: conversion equivalente, %", 100 * conv_real * (1 + dv_opt / venta_anual), 6.90, 2)

# Punto de equilibrio
pe <- function(fr) uniroot(function(e) van(flujo_m(e, e, frente = fr)), c(0, 0.5))$root
peA <- pe("A"); peB <- pe("B")
chk("27", "Equilibrio de la solucion 1, % de la venta piloto", 100 * peA, 3.06, 2)
chk("27", "Equilibrio de la solucion 1, pp de conversion", 100 * peA * conv_piloto, 0.19, 2)
chk("27", "Equilibrio de la solucion 2, % de la venta de zapatillas", 100 * peB, 0.54, 2)
chk("27", "Equilibrio de la solucion 2, pp de conversion", 100 * peB * conv_real, 0.04, 2)
unid_eq <- uniroot(function(u) van(flujo_m(0.033, eB_base, um = u, frente = "B")), c(100, 200000))$root
chk("27", "Unidades al mes que anulan el VAN de la solucion 2, escenario base", unid_eq, 5883, 0)
chk("27", "Ese volumen en veces la primera redistribucion", unid_eq / par$unid_ini, 7.9, 1)

# Sensibilidad
chk("27", "Sensibilidad: VAN base con tasa de 10%, S/", van(fb, 0.10), 869281, 0)
chk("27", "Sensibilidad: VAN base con tasa de 20%, S/", van(fb, 0.20), 762866, 0)
chk("27", "Sensibilidad: VAN base con el doble de unidades transferidas, S/", van(flujo_m(0.033, eB_base, um = 2 * par$unid_mes)), 698687, 0)
chk("27", "Sensibilidad: VAN base con costo de S/15 por unidad transferida, S/", van(flujo_m(0.033, eB_base, cu = 15)), 709360, 0)
chk("27", "Sensibilidad: VAN base con margen bruto de 38%, S/", van(flujo_m(0.033, eB_base, m = 0.38)), 696710, 0)

# Monte Carlo
set.seed(2026)
rtri <- function(n, a, m, b) { u <- runif(n); fc <- (m - a) / (b - a)
  ifelse(u < fc, a + sqrt(u * (b - a) * (m - a)), b - sqrt((1 - u) * (b - a) * (b - m))) }
N <- 10000
sim <- data.frame(eA = rtri(N, 0, 0.033, 0.045), eB = rtri(N, 0, eB_base, eB_alto),
                  m = rtri(N, 0.38, mgn, 0.45),
                  cu = rtri(N, 4, 8, 15), um = rtri(N, par$unid_mes / 2, par$unid_mes, 2 * par$unid_mes))
sim$van <- with(sim, mapply(function(a, b, mm, c1, u1) van(flujo_m(a, b, m = mm, cu = c1, um = u1)), eA, eB, m, cu, um))
q05 <- quantile(sim$van, 0.05)
mc <- c(media = mean(sim$van), p_neg = mean(sim$van < 0), var5 = unname(q05), cvar5 = mean(sim$van[sim$van <= q05]),
        p50 = median(sim$van))
print(round(mc, 4))
chk("27", "Monte Carlo: VAN medio, S/", mc["media"], 480081, 0)
chk("27", "Monte Carlo: VAN mediano, S/", mc["p50"], 508247, 0)
chk("27", "Monte Carlo: probabilidad de VAN negativo, %", 100 * mc["p_neg"], 5.1, 1)
chk("27", "Monte Carlo: VaR al 5%, S/", mc["var5"], -1217, 0)
chk("27", "Monte Carlo: CVaR al 5%, S/", mc["cvar5"], -95711, 0)
write.csv(data.frame(van = sim$van), "fig_mc.csv", row.names = FALSE)


# ==============================================================================
# 28. TABLA DE VERIFICACION
# ==============================================================================
titulo("28. TABLA DE VERIFICACION")
print(verif, right = FALSE, row.names = FALSE, max = 100000)
cat(sprintf("\n%d de %d valores coinciden con el documento.\n",
            sum(verif$estado == "OK"), nrow(verif)))
if (any(verif$estado != "OK")) {
  cat("Valores a revisar:\n")
  print(verif[verif$estado != "OK", ], right = FALSE, row.names = FALSE, max = 100000)
}
cat("\nEjecutado con", R.version.string, "| fixest", as.character(packageVersion("fixest")), "\n")
