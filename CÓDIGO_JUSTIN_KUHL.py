# ============================================================
# INFORME ELÉCTRICO - genera el Excel con circuitos y materiales
# ============================================================
# Este algoritmo le pregunta al usuario los datos de una instalación
# eléctrica (empalme, tablero, circuitos, climatización, agua caliente...)
# y al final arma un Excel con 3 hojas:
#   - Informe: resumen de la instalación y de cada circuito
#   - Materiales: la lista completa de materiales a comprar
#   - Base Normativa: qué artículo del RIC justifica cada material
#
# Cómo está armado el archivo, de arriba hacia abajo:
#   1) Tabla con la normativa RIC (BLOQUES_NORMATIVA) y la función que
#      arma la hoja "Base Normativa" con los links.
#   2) Funciones de cálculo sacadas de las tablas del RIC: conduit o
#      canaleta, alimentador, acometida y corrección por temperatura.
#      Ojo con lo que NO está acá: la sección de los conductores de cada
#      circuito se calcula más abajo, en el algoritmo final y con seccion_A1
#      (que está dentro de build_materiales_df), y los ferrules y los
#      tornillos también se cuentan dentro de build_materiales_df.
#   3) Funciones para preguntar datos por consola (pedir_float_positivo...).
#   4) build_materiales_df(): la función más grande de todas. Recorre los
#      circuitos y arma, línea por línea, todo el listado de materiales.
#   5) De la línea 8690 aprox. hacia abajo ya no hay más funciones: es el algoritmo
#      que se ejecuta al correr el código, en esete orden:
#      Parte 1 y 2: pregunta por los ambientes,cargas y datos del empalme
#      Parte 2.1: reparte las cargas y arma los circuitos
#      Parte 3 y 3.1: elige el termomagnético, el empalme y los diferenciales
#      Parte 3.2: sección de cada conductor y caída de tensión
#      Parte 3.3: arma las tablas y llama build_materiales_df()
#      Parte 4: da forma a las hojas y guarda el Excel
#
# Las variables que más se repiten en todo el archivo (para no perderse):
#   circuitos ............... lista con cada circuito armado; cada uno es un
#                             diccionario con nombre, longitud, potencia, sus
#                             items, etc. Es la base de todo lo demás.
#   circuitos_df ............ los mismos circuitos pero ya pasados a tabla de
#                             pandas, que es lo que se escribe en el Excel.
#   ambientes_detalle ....... lista con los datos de cada ambiente que ingresó
#                             el usuario (nombre, área, enchufes, luminarias).
#   ambientes_df ............ lo mismo pero como tabla de pandas.
#   items_por_nombre ........ diccionario {nombre del circuito: [items]}. Cada
#                             item es un dict con un enchufe, una luminaria o
#                             un equipo. Se usa mucho para contar materiales.
#   group_info .............. diccionario {gid: datos del grupo}: qué circuitos
#                             cuelgan de cada diferencial y con qué calibre.
#                             "gid" es el número correlativo del grupo (0,1,2...).
#   tm_corrientes ........... corriente del termomagnético de cada circuito,
#                             en el mismo orden que los circuitos.
#   texto_omni .............. texto del interruptor general omnipolar.
#   dif_calibre / max_sum_tm  calibre del diferencial y tope de suma de TM que
#                             puede colgar de él.
#   tensión / tension_nominal  220V, y fp es el factor de potencia.
#   temperatura ............. temperatura ambiente, para corregir la corriente
#                             que aguanta el cable.
#   zona .................... "húmeda" o "seca". Ojo: NO cambia la tabla de
#                             corriente (las dos ramas usan tabla_70); lo único
#                             que cambia es el cable por defecto: THWN-2 en
#                             húmeda y H07Z1-K en seca (tipo_cable_default).
#   tipo_canalizacion ....... "embutida" o "sobrepuesta".
#   material_forrado_exterior  fibrocemento, madera o siding; define el tipo de
#                             tornillo y tarugo.
#   longitud_alimentador .... metros entre el empalme y el tablero.
#   longitud_transformador_empalme  metros de la acometida (transformador al
#                             empalme). Las demás "longitud_..." y "dist_..."
#                             son tramos puntuales según cómo quedó el empalme.
#   circuitos_climatizacion / circuitos_agua_caliente  equipos de clima y de
#                             agua caliente, con su cálculo ya hecho.
#   Las que van en MAYÚSCULAS (SECCIONES_MM2_ALIM, TABLA_AMPACIDAD_ALIM,
#   CALIBRES_TM_CLIMA, etc.) son tablas escritas a mano: unas copian una tabla
#   del RIC y otras son medidas o calibres comerciales de catálogo. Se cargan
#   una vez al importar el archivo y no cambian durante la ejecución.
#
# Si es primera vez que se lee: las funciones de cálculo del principio
# son fáciles de entender (reciben datos, devuelven un resultado).
# build_materiales_df es la más grande y compleja, por eso está separada
# en bloques con títulos tipo "# ===== NOMBRE SECCIÓN =====".
# El bloque final es el que realmente interactúa con el usuario.
# ============================================================

import pandas as pd  # pandas: para armar tablas de datos (DataFrames) y exportarlas a Excel
from datetime import datetime  # datetime: para poner la fecha/hora en el informe
import os, sys  # os: rutas de archivos; sys: interactuar con la consola/sistema
import math  # math: funciones matemáticas (raíces, redondeos, etc.)
import numpy as np  # numpy: cálculos numericos y manejo de arreglos

# =========================================================
# BASE NORMATIVA + HIPERVÍNCULOS PARA HOJA "MATERIALES"
# =========================================================

BASE_NORMATIVA_ROWS = []  # generado dinámicamente desde BLOQUES_NORMATIVA

# Tabla grande con la normativa RIC de cada material. Cada elemento es una
# tupla con 2 datos: (nombre del material, artículos del RIC que lo respaldan).
# Esta agrupada por secciones del RIC (empalme, tablero, climatización, etc.)
# marcadas con líneas "-- RIC N - TEMA --". No se comenta material por
# material porque todas las filas tienen la misma forma: nombre + norma.
BLOQUES_NORMATIVA = [
    # ── RIC 1 - EMPALME ──────────────────────────────────────────────────────
    ("Disyuntor termomagnético empalme",             "RIC 1 (5.3, 8.1, 8.2, Anexo 1.3) / RIC 5 (8.5, 8.7.1, 8.7.2, 8.7.3, 8.7.6.3, 8.7.7.4, 8.7.7.5) / RIC 10 (5.2.1)"),  # el TM que protege el empalme, va antes del medidor
    ("Portafusible de loza empalme",                 "RIC 1 (8.1, 8.2, Anexo 1.3) / RIC 5 (8.7.6.3)"),  # portafusible de loza del empalme, junto a la unidad de medida
    ("Sellador de roscas con teflón",                "RIC 1 (6.1) / RIC 4 (5.18)"),  # sella las roscas del conduit del empalme para que no le entre agua
    ("Abrazaderas tipo caddy",                       "RIC 1 (6.2) / RIC 4 (5.12.4, 5.44, 7.1.3, 7.1.10, 7.1.15, 7.16.2.1, 7.16.2.2, 7.16.2.3, 7.16.2.4, 7.16.2.5, 7.16.4.3, 7.16.4.4)"),  # abrazaderas caddy que afirman la canalización del empalme al muro
    ("Conector HUB acero galvanizado",               "RIC 1 (6.2) / RIC 2 (6.1.21.6) / RIC 4 (5.6, 5.12.4, 5.14, 5.24, 5.44, 7.1.3, 7.1.10, 7.1.15, 7.16.1.11, 7.16.2.1, 7.16.2.2, 7.16.2.3, 7.16.2.4, 7.16.2.5) / RIC 5 (8.6.4) / RIC 6 (7.2) / RIC 11 (10.3.5)"),  # el hub va sobre la caja de empalme, entrada estanca del conduit
    ("Terminal ferrul acometida",                    "RIC 1 (6.2) / RIC 4 (5.9, 5.11.4)"),  # ferrul de la punta del cable de acometida, para que no se abra el hilo
    ("Terminal ferrul alimentador",                  "RIC 1 (6.2) / RIC 4 (5.9, 5.11.4)"),  # ferrul del alimentador, mismo criterio de conexión prolija
    ("Terminal ferrul doble alimentador",            "RIC 1 (6.2) / RIC 4 (5.9, 5.11.4)"),  # ferrul doble cuando entran dos conductores al mismo borne
    ("Terminal ferrul tablero interior",             "RIC 4 (5.9, 5.11.4)"),  # ferrules del cableado interno del tablero
    ("Terminal de compresión tipo ojo",              "RIC 1 (6.2) / RIC 4 (5.13) / RIC 5 (8.6.4)"),  # terminal de ojo para apernar a barra, típico de la tierra
    ("Cable acometida concéntrico aéreo",            "RIC 1 (6.2) / RIC 3 (5.1.3) / RIC 4 (5.2, 5.5, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4)"),  # el concéntrico aéreo que baja del poste al empalme
    ("Cable acometida concéntrico subterráneo",      "RIC 1 (6.2, 7.20) / RIC 3 (5.1.3) / RIC 4 (5.2, 5.5, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4)"),  # el mismo concéntrico pero enterrado, cambia el artículo del RIC 1
    ("Mordaza para alimentador aéreo",               "RIC 1 (6.2)"),  # mordaza que tensa el cable aéreo en el poste o la muralla
    ("Cabeza de servicio",                           "RIC 1 (6.2)"),  # cabeza de servicio arriba del conduit, para que no entre lluvia
    ("Caja de empalme metálica",                     "RIC 1 (6.3, 6.5, Anexo 1.1) / RIC 2 (6.1.21.3) / RIC 4 (5.38, 7.16.1.11) / RIC 5 (8.6.4) / RIC 6 (7.2) / RIC 10 (5.2.4) / RIC 11 (10.3.5)"),  # la caja metálica donde va el medidor con sus protecciones
    ("Abrazadera conduit PVC acometida",             "RIC 1 (6.3) / RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2, 7.16.4.3, 7.16.4.5)"),  # abrazaderas que fijan el conduit de la acometida
    ("Terminal PVC conduit con 2 tuercas",           "RIC 1 (6.3) / RIC 2 (6.1.21.6) / RIC 4 (5.14, 5.24, 7.16.1.11)"),  # terminal con dos tuercas para entrar el conduit a la caja
    ("Unidad de medida monofásica",                  "RIC 1 (6.6, 7.1, 7.2, 7.3, 7.5) / RIC 10 (5.2.5)"),  # el medidor monofásico propiamente tal
    ("Conduit PVC alimentador",                      "RIC 1 (7.15, 7.19) / RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.34, 5.42, 5.43, 5.44, 5.46, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.15.1.1, 7.15.1.2, 7.16.1.2, 7.16.1.3, 7.16.1.4, 7.16.1.5, 7.16.1.6, 7.16.1.7, 7.16.1.9, 7.16.1.10, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2, Tabla 4.17, Tabla 4.18, Tabla 4.19, Tabla 4.20, Tabla 4.23)"),  # el conduit por donde va el alimentador del empalme al tablero
    ("Abrazadera conduit PVC alimentador",           "RIC 1 (7.15, 7.19) / RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2, 7.16.4.3, 7.16.4.5)"),  # abrazaderas de ese mismo conduit del alimentador
    ("Conduit PVC acometida subterránea",            "RIC 1 (7.20) / RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.34, 5.42, 5.43, 5.44, 5.46, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.9.1, 7.9.2, 7.9.3, 7.9.4, 7.9.5, 7.9.6, 7.9.7.1, 7.9.7.2, 7.9.7.3, 7.9.7.4, 7.9.7.5, 7.16.7, 7.16.7.1, 7.16.7.2, 7.16.7.3, Tabla 4.29)"),  # conduit enterrado de la acometida, con su profundidad mínima
    # ── RIC 2 - TABLERO ──────────────────────────────────────────────────────
    ("Tablero embutido de PVC",                      "RIC 2 (5.2, 5.3, 5.3.6, 6.1.1, 6.1.3, 6.1.4, 6.1.8, 6.1.9, 6.1.10, 6.1.11, 6.1.12, 6.1.16.3, 6.1.19, 6.1.21.2, 6.1.24) / RIC 4 (5.34, 7.16.1.11) / RIC 10 (5.1.3.1, 5.1.3.2, 5.1.3.3)"),  # el tablero cuando va embutido en el muro
    ("Tablero sobrepuesto de PVC",                   "RIC 2 (5.2, 5.3, 5.3.6, 6.1.1, 6.1.3, 6.1.4, 6.1.8, 6.1.9, 6.1.10, 6.1.11, 6.1.12, 6.1.16.3, 6.1.19, 6.1.21.2, 6.1.24) / RIC 4 (5.34, 7.16.1.11) / RIC 10 (5.1.3.1, 5.1.3.2, 5.1.3.3)"),  # el mismo tablero pero montado a la vista
    ("Interruptor diferencial",                      "RIC 2 (5.3.7, 6.2.6) / RIC 5 (5.6, 7.8.2, 8.5, 8.7.1, 8.7.2, 8.7.3, 8.7.6.3, 8.7.7.4) / RIC 6 (7.9) / RIC 7 (7.3.7, 7.4.5) / RIC 10 (5.1.3.5, 5.1.3.6, 5.1.3.7, 5.1.3.8, 5.2.1)"),  # el diferencial general, 30 mA para proteger a las personas
    ("Interruptor diferencial 10mA agua caliente",   "RIC N°11 (6, 6.4.3, Tabla Vol.1) / RIC N°07 (7.4.5, 7.6.5.4) — Volumen 1 baño: sensibilidad ≤10mA"),  # diferencial de 10 mA, exigido para ducha y calefón en el baño
    ("Tablero desconexión externo agua caliente",    "RIC N°07 (7.2.8, 7.3.3, 7.4.1, 7.4.2) / RIC N°02 (5.2, 6.1) / RIC N°11 (6, Tabla Vol.3) — Fuera Vol.0,1,2 a la vista del equipo"),  # tablerito de desconexión a la vista del equipo, fuera de los volúmenes del baño
    ("TM bipolar desconexión externo agua caliente", "RIC N°07 (7.2.8, 7.3.4, 7.4.2, 7.4.5) — Interruptor de desconexión a la vista del equipo; conductor ≥ I×1,25"),  # el TM bipolar que va dentro de ese tablero de desconexión
    ("Bornera PE tablero externo agua caliente",     "RIC N°02 (6.2.7) / RIC N°06 (5.11, 5.14) — Continuidad conductor de protección"),  # bornera de tierra del tablero externo, para no cortar el PE
    ("Prensaestopa tablero externo agua caliente",   "RIC N°04 (5.15, 5.24) / RIC N°07 (5.2.8) — Fijación canalización en tablero"),  # prensaestopa que entra el cable a ese tablero sin dañarlo
    # ── RIC 7 - CLIMATIZACIÓN ────────────────────────────────────────────────
    ("Enchufe climatización",                        "RIC 7 (7.1.1, 7.1.3, 7.3.1, 7.3.2, 7.4.4) — Circuito exclusivo climatización; enchufe coordinado con TM; capacidad nominal 16/20/25/32 A"),  # el enchufe del split, circuito exclusivo y coordinado con su TM
    ("Disyuntor termomagnético climatización",         "RIC 7 (7.1.2, 7.2.8, 7.3.3, 7.3.4, 7.3.6, 7.4.5) — Circuito exclusivo; conductor ≥ I×1,25; mín 2,5 mm^2; TM coordinado con conductor; diferencial exclusivo"),  # TM del circuito de clima, dimensionado a 1,25 veces la corriente
    ("Interruptor diferencial climatización",        "RIC 7 (7.1.2, 7.3.7, 7.4.5, 7.6.5.4) — Diferencial exclusivo por circuito de climatización; ≤ 30 mA"),  # diferencial propio del clima, no se comparte con otros circuitos
    ("Canalización climatización",                   "RIC 7 (7.5.1, 7.5.2) / RIC 4 (5.2, 5.5, 5.6, 5.34, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15) — Canalización según RIC 4 sec.7; corrección temperatura si T>30°C"),  # la canalización del circuito de climatización
    ("Conductor climatización",                      "RIC 7 (7.3.4, 7.5.1, 7.5.2) / RIC 3 (5.1.3) / RIC 4 (5.2, 5.4, 5.5, 5.34, 6.2.5) — Conductor ≥ I×1,25; mín 2,5 mm^2; factor corrección si T>30°C"),  # conductor del clima, mínimo 2,5 mm^2 y corregido por temperatura
    ("Interruptor general omnipolar",                "RIC 2 (5.3.7, 6.5.3, 6.6.1, 6.6.2) / RIC 3 (5.2.5) / RIC 5 (8.5, 8.7.1, 8.7.2) / RIC 6 (6.5) / RIC 10 (5.1.3.3, 5.2.1)"),  # el general omnipolar que corta todo el tablero de una
    ("Supresor de transiente (SPD)",                 "RIC 5 (8.7.7) / RIC 6 (6.2.2)"),  # supresor de transientes, para las sobretensiones que vienen de la red
    ("Protector sobrevoltaje y corriente",            "RIC 6 (6.6.2)"),  # protector de sobrevoltaje, corta si la tensión se va de rango
    ("Barra repartidora bipolar",                     "RIC 2 (6.2.1, 6.2.4, 6.2.7)"),  # peineta bipolar que reparte fase y neutro entre los TM
    ("Disyuntor termomagnético",                       "RIC 2 (5.3.7) / RIC 5 (8.5, 8.7.1, 8.7.2, 8.7.3, 8.7.6.3, 8.7.7.4, 8.7.7.5) / RIC 7 (7.2.8, 7.3.1, 7.3.6, 7.4.2, 7.4.4, 7.4.5) / RIC 10 (5.1.2.9, 5.1.3.4, 5.1.3.6, 5.1.3.7, 5.1.3.8, 5.1.4.1, 5.1.4.2, 5.2.1)"),  # el TM de cada circuito interior
    ("Riel DIN",                                     "RIC 2 (6.1.15, 6.1.23)"),  # el riel donde se montan TM y diferenciales
    ("Salida de caja conduit de PVC",                "RIC 2 (6.1.21.6) / RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.14, 5.24, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.16.1.11)"),  # salida de caja para sacar el conduit del tablero
    ("Conductor alimentador llegada tablero",         "RIC 2 (6.2.1) / RIC 10 (5.2.1)"),  # el tramo de alimentador que llega a la barra del tablero
    ("Barra repartidora tetrapolar",                 "RIC 2 (6.2.1, 6.2.4, 6.2.7)"),  # peineta tetrapolar, para cuando el tablero es trifásico
    ("Conductor interior tablero ≤6mm^2",             "RIC 2 (6.2.2, 6.2.11) / RIC 4 (5.2, 5.4, 5.5, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4)"),  # cableado interno del tablero hasta 6 mm^2
    ("Conductor interior tablero >6mm^2",             "RIC 2 (6.2.2, 6.2.11) / RIC 4 (5.2, 5.4, 5.5, 5.8, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4)"),  # el mismo cableado interno pero sobre 6 mm^2
    ("Barra unipolar verde PE",                      "RIC 2 (6.2.7) / RIC 6 (5.11, 5.14)"),  # barra verde donde se juntan todas las tierras
    ("Bornera de conexión",                          "RIC 2 (6.2.12)"),  # borneras para ordenar las conexiones dentro del tablero
    ("Luz piloto",                                   "RIC 2 (6.2.14, 6.2.15)"),  # luz piloto que avisa que el tablero está energizado
    ("Portafusible tablero",                         "RIC 2 (6.2.15, 6.3.6) / RIC 5 (8.7.6.3)"),  # portafusible del tablero, va con la luz piloto
    ("Fusible cilíndrico tablero",                   "RIC 2 (6.2.15, 6.3.6) / RIC 5 (8.7.6.3, 8.7.7.4, 8.7.7.5)"),  # el fusible que va dentro de ese portafusible
    ("Barra copperweld + conector bronce",           "RIC 2 (6.4.1) / RIC 5 (5.1, 5.6, 6.7.1.2, 8.7.3, 8.7.7.1) / RIC 6 (5.4, 5.5, 5.6, 6.1, 6.4, 7.1, 7.3, 7.8, 7.9, 8.1, 8.3.1, 8.3.2, 8.5, 8.11, 8.13, 8.14, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3, Tabla 6.2, Tabla 6.3, Tabla 6.4) / RIC 10 (5.2.1) / RIC 11 (10.3.5)"),  # la barra de tierra enterrada con su conector de bronce
    ("Camarilla PVC naranjo",                        "RIC 2 (6.4.1) / RIC 6 (5.4, 5.6, 5.15, 6.4, 7.8, 7.9, 8.2) / RIC 10 (5.2.1) / RIC 11 (10.3.5)"),  # camarilla que deja registrable la barra de tierra
    ("Conductor THWN-2 verde PT ≤6mm^2",               "RIC 2 (6.4.1) / RIC 3 (5.2.3) / RIC 4 (5.2, 5.4, 5.5, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 5 (5.1, 5.6, 6.7.1.2, 8.7.3, 8.7.7.1) / RIC 6 (5.4, 5.6, 6.1, 6.4, 6.6.1, 6.6.2, 6.6.3, 7.3, 7.5, 7.6, 7.7, 7.8, 7.9, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3) / RIC 10 (5.2.1) / RIC 11 (10.3.5)"),  # cable verde de tierra hasta 6 mm^2
    ("Conductor THWN-2 verde PT >6mm^2",               "RIC 2 (6.4.1) / RIC 3 (5.2.3) / RIC 4 (5.2, 5.4, 5.5, 5.8, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 5 (5.1, 5.6, 6.7.1.2, 8.7.3, 8.7.7.1) / RIC 6 (5.4, 5.6, 6.1, 6.4, 6.6.1, 6.6.2, 6.6.3, 7.3, 7.5, 7.6, 7.7, 7.8, 7.9, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3) / RIC 10 (5.2.1) / RIC 11 (10.3.5)"),  # cable verde de tierra sobre 6 mm^2
    ("Conductor THWN-2 blanco PT ≤6mm^2",              "RIC 4 (5.2, 5.4, 5.5, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 5 (6.4.1, 6.4.2, 6.4.3, 6.5.1, 6.5.2, 6.7.1.2) / RIC 6 (5.4, 5.6, 6.1, 6.3, 6.4, 6.5, 6.6.1, 6.6.2, 6.6.3, 6.6.4, 7.7, 7.8, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3) / RIC 10 (5.2.1)"),  # cable blanco del neutro hasta 6 mm^2
    ("Conductor THWN-2 blanco PT >6mm^2",              "RIC 4 (5.2, 5.4, 5.5, 5.8, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.6, 6.2.7, Tabla 4.4) / RIC 5 (6.4.1, 6.4.2, 6.4.3, 6.5.1, 6.5.2, 6.7.1.2) / RIC 6 (5.4, 5.6, 6.1, 6.3, 6.4, 6.5, 6.6.1, 6.6.2, 6.6.3, 6.6.4, 7.7, 7.8, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3) / RIC 10 (5.2.1)"),  # cable blanco del neutro sobre 6 mm^2
    ("Conductor desnudo Cu 16mm^2",                   "RIC 2 (6.4.1) / RIC 4 (Tabla 4.3) / RIC 6 (8.3.2, 8.5, 8.9)"),  # cobre desnudo de 16 mm^2 para la puesta a tierra
    # ── RIC 3 - CONDUCTORES ──────────────────────────────────────────────────
    ("Alimentador RV-K Cu ≤6mm^2",                   "RIC 3 (5.1.2, 5.1.3, 5.2.2, 5.2.6, 6.1, 6.2, 6.3, 6.4.1) / RIC 4 (5.2, 5.4, 5.5, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4)"),  # alimentador RV-K hasta 6 mm^2; el RIC 3 es el de conductores
    ("Alimentador RV-K Cu >6mm^2",                   "RIC 3 (5.1.2, 5.1.3, 5.2.2, 5.2.6, 6.1, 6.2, 6.3, 6.4.1) / RIC 4 (5.2, 5.4, 5.5, 5.8, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4)"),  # el mismo alimentador pero sobre 6 mm^2
    ("Conductores circuitos interiores ≤6mm^2",      "RIC 3 (5.1.3) / RIC 4 (5.2, 5.4, 5.5, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4) / RIC 7 (7.1.2, 7.1.3, 7.3.4, 7.5.1, 7.5.2) / RIC 10 (5.1.1.1, 5.1.1.3, 5.1.3.4, 5.1.3.5, 5.1.4.1, 5.1.5.1, 5.1.5.2, 5.1.5.3) / RIC 11 (10.3.4)"),  # los conductores de los circuitos de la casa hasta 6 mm^2
    ("Conductores circuitos interiores >6mm^2",      "RIC 3 (5.1.3) / RIC 4 (5.2, 5.4, 5.5, 5.8, 5.32, 5.34, 5.37.4, 6.1.1, 6.1.2, 6.2.1, 6.2.2, 6.2.5, 6.2.6, 6.2.7, Tabla 4.4) / RIC 6 (6.6.4) / RIC 7 (7.1.2, 7.1.3, 7.3.4, 7.5.1, 7.5.2) / RIC 10 (5.1.1.1, 5.1.1.3, 5.1.3.4, 5.1.3.5, 5.1.4.1, 5.1.5.1, 5.1.5.2, 5.1.5.3) / RIC 11 (10.3.4)"),  # los mismos conductores de circuitos sobre 6 mm^2
    # ── RIC 4 - CANALIZACIONES Y ACCESORIOS ──────────────────────────────────
    ("Canalización embutida conduit PVC",            "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.34, 5.42, 5.43, 5.44, 5.46, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.15.1.1, 7.15.1.2, 7.16.1.2, 7.16.1.3, 7.16.1.4, 7.16.1.5, 7.16.1.6, 7.16.1.7, 7.16.1.9, 7.16.1.10, 7.16.1.16, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2, Tabla 4.17, Tabla 4.18, Tabla 4.19, Tabla 4.20, Tabla 4.23) / RIC 7 (7.2.2, 7.5.1) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # conduit embutido en el muro, con su llenado y radios de curva
    ("Canalización sobrepuesta canaleta PVC",        "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.34, 5.42, 5.43, 5.44, 5.46, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.7.1, 7.7.2, 7.7.3, 7.7.4, 7.7.5, 7.7.6, 7.7.7, 7.7.9, 7.16.3.1, 7.16.3.2, 7.16.4.1, 7.16.4.6, 7.16.4.7, 7.16.4.8, 7.16.4.9, Tabla 4.23) / RIC 7 (7.2.2, 7.5.1) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # canaleta a la vista, cuando no se pica el muro
    ("Abrazadera conduit PVC circuitos",             "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2, 7.16.4.3, 7.16.4.5) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # abrazaderas del conduit de los circuitos interiores
    ("Boquilla conduit PVC",                         "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.16.3.1, 7.16.3.2, 7.16.3.3, 7.16.3.5, 7.16.3.6, 7.16.4.2) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # boquilla en la punta del conduit para no pelar el cable
    ("Unión copla canaleta PVC",                     "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.7.1, 7.7.2, 7.7.3, 7.7.4, 7.7.5, 7.7.6, 7.7.7, 7.7.9, 7.16.3.1, 7.16.3.2, 7.16.4.1, 7.16.4.6, 7.16.4.7, 7.16.4.8, 7.16.4.9, Tabla 4.23) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # copla para unir dos tramos de canaleta
    ("Curvas internas 90° canaleta PVC",             "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.7.1, 7.7.2, 7.7.3, 7.7.4, 7.7.5, 7.7.6, 7.7.7, 7.7.9, 7.16.3.1, 7.16.3.2, 7.16.4.1, 7.16.4.6, 7.16.4.7, 7.16.4.8, 7.16.4.9, Tabla 4.23) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # curva interna de canaleta, para el rincón cóncavo
    ("Curvas planas 90° canaleta PVC",               "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.7.1, 7.7.2, 7.7.3, 7.7.4, 7.7.5, 7.7.6, 7.7.7, 7.7.9, 7.16.3.1, 7.16.3.2, 7.16.4.1, 7.16.4.6, 7.16.4.7, 7.16.4.8, 7.16.4.9, Tabla 4.23) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # curva plana de canaleta, girando en el mismo muro
    ("Curva T canaleta PVC",                         "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.44, 7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15, 7.7.1, 7.7.2, 7.7.3, 7.7.4, 7.7.5, 7.7.6, 7.7.7, 7.7.9, 7.16.3.1, 7.16.3.2, 7.16.4.1, 7.16.4.6, 7.16.4.7, 7.16.4.8, 7.16.4.9, Tabla 4.23) / RIC 10 (5.1.2.1, 5.1.2.6) / RIC 11 (10.3.3, 10.3.4)"),  # curva T cuando la canaleta se abre en dos
    ("Caja derivación embutida PVC",                 "RIC 4 (5.2, 5.5, 5.6, 5.12.1, 5.12.3, 5.12.4, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.45, 5.48.1, 7.16.1.11, 7.16.1.13) / RIC 10 (5.1.2.2, 5.1.2.3, 5.1.2.11, 5.1.2.12, 5.1.2.14, 5.2.2) / RIC 11 (10.3.4)"),  # caja de derivación embutida de los circuitos
    ("Caja de paso estanca",                         "RIC 4 (5.2, 5.5, 5.6, 5.12.1, 5.12.3, 5.12.4, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.45, 5.48.1, 7.16.1.11, 7.16.1.13) / RIC 10 (5.1.2.2, 5.1.2.3, 5.1.2.11, 5.1.2.12, 5.1.2.14, 5.2.2) / RIC 11 (10.3.4)"),  # caja de paso estanca, para zonas húmedas o a la intemperie
    ("Caja derivación sobrepuesta PVC",              "RIC 4 (5.2, 5.5, 5.6, 5.12.1, 5.12.3, 5.12.4, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.45, 5.48.1, 7.16.1.11) / RIC 10 (5.1.2.2, 5.1.2.3, 5.1.2.14, 5.2.2) / RIC 11 (10.3.4)"),  # caja de derivación montada a la vista
    ("Caja derivación sobrepuesta chuqui PVC",       "RIC 4 (5.2, 5.5, 5.6, 5.12.1, 5.12.3, 5.12.4, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.45, 5.48.1, 7.16.1.11) / RIC 10 (5.1.2.2, 5.1.2.3, 5.1.2.14, 5.2.2) / RIC 11 (10.3.4)"),  # la caja chuqui, la sobrepuesta chica de enchufes e interruptores
    ("Caja derivación octogonal PVC",                "RIC 4 (5.2, 5.5, 5.6, 5.12.1, 5.12.3, 5.12.4, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.45, 5.48.1, 7.16.1.11) / RIC 10 (5.1.2.2, 5.1.2.3, 5.1.2.11, 5.1.2.12, 5.1.2.13, 5.1.2.14, 5.2.2) / RIC 11 (10.3.4)"),  # caja octogonal del centro de luz en el cielo
    ("Tapa ciega octogonal PVC",                     "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.17, 5.22)"),  # tapa de la octogonal cuando queda sin artefacto
    ("Tapa ciega PVC",                               "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.17, 5.22)"),  # tapa ciega de la caja de derivación
    ("Tapa ciega chuqui PVC",                        "RIC 4 (5.2, 5.5, 5.6, 5.12.4, 5.17, 5.22)"),  # tapa ciega de la caja chuqui
    ("Tubo conduit galvanizado",                     "RIC 4 (5.6, 5.12.4, 5.38, 5.42, 5.44, 5.46, 7.1.3, 7.1.10, 7.1.15, 7.16.2.1, 7.16.2.2, 7.16.2.3, 7.16.2.4, 7.16.2.5, Tabla 4.17, Tabla 4.18, Tabla 4.19, Tabla 4.20) / RIC 5 (8.6.4) / RIC 6 (7.2) / RIC 11 (10.3.5)"),  # conduit metálico, se usa donde puede haber golpes
    ("Caja derivación metálica",                     "RIC 4 (5.6, 5.12.2, 5.12.3, 5.12.4, 5.13, 5.14, 5.16, 5.21, 5.23, 5.27, 5.34, 5.38, 5.39, 5.45, 7.16.1.11) / RIC 5 (8.6.4) / RIC 6 (7.2) / RIC 11 (10.3.5)"),  # caja de derivación metálica, hay que aterrizarla
    ("Enchufe",                                       "RIC 4 (5.34, 5.45) / RIC 7 (7.2.1, 7.3.2, 7.4.4) / RIC 10 (5.1.2.7, 5.1.2.8, 5.1.2.9, 5.1.4.6, 5.2.2) / RIC 11 (10.3.4)"),  # el enchufe común de los circuitos de la casa
    ("Interruptor de circuito",                      "RIC 4 (5.34, 5.45) / RIC 10 (5.1.2.4, 5.1.2.9, 5.2.2) / RIC 11 (10.3.4)"),  # el interruptor de pared que comanda la luz
    ("Portalámpara / luminaria circuito interior",   "RIC 10 (5.1.2.11, 5.1.2.12, 5.1.2.13, 5.1.4.4, 5.1.4.5, 5.2.2) — Centro de iluminación en caja; soporte independiente del conductor; con PT"),  # el centro de luz: portalámpara o luminaria, siempre con tierra
    ("Conector cónico",                              "RIC 4 (5.11.3)"),  # conectores cónicos para empalmar los cables dentro de las cajas
    ("Cámara tipo C subterránea",                    "RIC 4 (7.9.5, 7.9.7.8, 7.9.7.9, 7.9.7.10, 7.9.7.12, 7.9.8.1, 7.9.8.2, 7.9.8.4.3, 7.9.8.5, 7.9.8.7)"),  # cámara tipo C para registrar la canalización enterrada
    ("Marco metálico cámara tipo C",                 "RIC 4 (7.9.5, 7.9.7.8, 7.9.7.9, 7.9.7.10, 7.9.7.12, 7.9.8.1, 7.9.8.2, 7.9.8.4.3, 7.9.8.5, 7.9.8.7)"),  # el marco y la tapa metálica de esa cámara
    ("Boquilla PVC cámara tipo C",                   "RIC 4 (7.9.8.9)"),  # boquilla por donde el conduit entra a la cámara
    ("Tornillo conexión tierra caja metálica",       "RIC 4 (5.13)"),  # tornillo de tierra de las cajas metálicas
    ("Prensaestopa",                                 "RIC 4 (5.15, 5.24)"),  # prensaestopa común, para entrar cables a cajas y tableros
    ("Tubo de estaño",                               "RIC 4 (4.4.1, 5.11.1)"),  # estaño para las soldaduras de conexión
    ("Pasta para soldar",                            "RIC 4 (4.4.1, 5.11.1)"),  # la pasta que va junto con el estaño
]


# Generar BASE_NORMATIVA_ROWS desde BLOQUES_NORMATIVA
# Cada entrada de BLOQUES_NORMATIVA trae 2 datos juntos: el nombre del
# material y su norma correspondiente
for _mat, _norma in BLOQUES_NORMATIVA:
    BASE_NORMATIVA_ROWS.append({  # arma una fila por material para la hoja Base Normativa
        "Material / Elemento": _mat,  # columna con el nombre del material
        "Norma": _norma,  # columna con los artículos del RIC que lo respaldan
    })  # cierra la fila de la hoja Base Normativa


# Deja la norma escrita corta: "RIC 4" en vez de "RIC 4.7.2".
# OJO con lo que se ve en el Excel: este texto casi nunca queda a la vista.
# Más abajo, en aplicar_base_normativa_e_hipervinculos() le pone el link a la
# hoja Base Normativa y le cambia el texto de la celda por "Ver normativa".
# Así que el "RIC N" que deja esta función solo se alcanza a ver en los
# materiales que NO tienen fila en BLOQUES_NORMATIVA (esos quedan con su
# texto, o con "-" si no aplica norma).
def normalizar_ric_materiales(norma, descripcion=""):
    """
    Deja limpia la columna 'Norma / RIC' en la hoja Materiales.
    Ejemplo: 'RIC 4.7.2' se deja como 'RIC 4'.
    Si no detecta RIC, conserva el texto original.
    """
    s = str(norma or "").strip()  # texto original de la norma
    d = str(descripcion or "").lower()  # nombre del material; solo se usa si la norma no trae número de RIC, para deducirlo por palabras clave

    # Si no vino ninguna norma (vacío, "nan" o "none" como texto), no hay
    # nada que limpiar.
    if s == "" or s.lower() in ["nan", "none"]:  # el material vino sin norma escrita
        return ""  # no hay norma, no hay nada que limpiar

    # Un "-" explícito significa "no aplica norma específica" — se respeta tal
    # cual, sin intentar adivinar un RIC por palabras clave de la descripción.
    if s == "-":
        return "-"  # deja el guion tal cual, es una norma válida de "no aplica"

    import re  # re solo se usa acá abajo, para pescar el número del RIC
    m = re.search(r"ric\s*(n°|nº|n|#)?\s*0?(\d+)", s, flags=re.IGNORECASE)  # busca "RIC" seguido de un número
    if m:  # encontró un RIC escrito dentro del texto
        return f"RIC {int(m.group(2))}"  # deja solo "RIC N" (sin puntos ni artículo)

    # Casos en que la norma dice "SEC" (o algo genérico) y no trae número de
    # RIC: se deduce el tomo que corresponde mirando el nombre del material.
    # Ej: si el nombre dice "alimentador" -> RIC 3, si dice "tierra" -> RIC 6.
    # Si ninguna palabra calza, más abajo devuelve el texto tal cual llegó.
    if "empalme" in d or "acometida" in d or "medidor" in d:
        return "RIC 1"  # RIC 1 es el del empalme y la acometida
    if "tablero" in d or "riel din" in d or "barra" in d or "luz piloto" in d:  # palabras del tablero: el tablero mismo, riel, barras, luz piloto
        return "RIC 2"  # RIC 2 es el del tablero y todo lo que va adentro
    if "alimentador" in d:  # si habla de alimentador, es el tramo entre empalme y tablero
        return "RIC 3"  # RIC 3 es el de alimentadores/conductores (igual que la sección "── RIC 3 - CONDUCTORES ──" de BLOQUES_NORMATIVA)
    if "conductor" in d or "canaleta" in d or "conduit" in d or "caja" in d or "prensaestopa" in d or "boquilla" in d:  # todo lo que sea conductor o canalización cae en el RIC 4
        return "RIC 4"  # RIC 4 es el de canalizaciones y conductores
    if "diferencial" in d or "automático" in d or "termomagnético" in d or "protección" in d or "supresor" in d:  # si es una protección (TM, diferencial, supresor) es RIC 5
        return "RIC 5"  # RIC 5 es el de las protecciones
    if "tierra" in d or "jabalina" in d or "copperweld" in d or "camarilla" in d:  # palabras típicas de la puesta a tierra
        return "RIC 6"  # RIC 6 es el de la puesta a tierra
    if "enchufe" in d or "interruptor" in d or "luminaria" in d or "ampolleta" in d:  # artefactos de la vivienda: enchufes, interruptores, luminarias
        return "RIC 10"  # RIC 10 es el de las instalaciones de viviendas

    return s  # ninguna palabra clave calzo, se deja el texto tal cual llegó (ej: "SEC")

# arma la hoja "Base Normativa" del Excel y deja los links desde
# cada material de la hoja Materiales a su artículo del RIC
def aplicar_base_normativa_e_hipervinculos(writer, materiales_df, sheet_materiales="Materiales", sheet_base="Base Normativa"):
    """
    Crea hoja 'Base Normativa' con 2 columnas: Material / Elemento | Norma
    Agrega hipervínculos en columna 'Norma / RIC' de la hoja Materiales.

    Qué recibe:
      writer ............ el "escritor" de pandas con el Excel abierto
      materiales_df ..... la tabla de materiales ya armada
      sheet_materiales .. nombre de la hoja de materiales (por defecto "Materiales")
      sheet_base ........ nombre de la hoja que va a crear (por defecto "Base Normativa")
    Devuelve la lista de filas que puso en la hoja nueva.
    """
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side  # estilos de openpyxl para pintar la hoja: letra, borde y relleno

    wb     = writer.book  # libro de Excel completo (todas las hojas)
    ws_mat = writer.sheets[sheet_materiales]  # hoja "Materiales", ya creada antes, para agregarle los links después

    # ── Crear hoja ────────────────────────────────────────────────────────────
    if sheet_base in wb.sheetnames:  # revisa si el libro ya trae una hoja con ese nombre
        del wb[sheet_base]  # si ya existía (por una corrida anterior), la borra primero
    ws_base = wb.create_sheet(sheet_base)  # crea la hoja Base Normativa vacía

    # Estilos reutilizables para pintar la hoja nueva: bordes finos grises,
    # fondo celeste para el título, gris para encabezados y blanco para datos.
    thin        = Side(border_style="thin", color="BFBFBF")  # línea fina gris para los bordes de las celdas
    fill_titulo = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")  # celeste del título
    fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # gris de los encabezados
    fill_blanco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # blanco de las filas de datos
    borde_gris  = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde gris por los cuatro lados

    # ── Fila 1: título ────────────────────────────────────────────────────────
    ws_base.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)  # combina A1:B1 en una sola celda ancha para el título
    tc = ws_base["A1"]  # celda donde va el título
    tc.value     = "BASE NORMATIVA RIC PARA TRAZABILIDAD DE MATERIALES"  # título que se ve arriba de todo en la hoja
    tc.font      = Font(bold=True, size=13, color="000000")  # letra negrita y más grande para el título
    tc.alignment = Alignment(horizontal="center", vertical="center")  # título centrado
    tc.fill      = fill_titulo  # fondo celeste para el título
    # pinta también la celda de al lado (columna B) con el mismo fondo y borde
    for c in range(1, 3):
        ws_base.cell(row=1, column=c).fill   = fill_titulo  # mismo fondo celeste en A1 y B1
        ws_base.cell(row=1, column=c).border = borde_gris  # y el mismo borde gris
    ws_base.row_dimensions[1].height = 22  # alto de la fila del título

    # ── Fila 2: encabezados ───────────────────────────────────────────────────
    # escribe los 2 encabezados de columna (fila 2)
    for j, h in enumerate(["Material / Elemento", "Norma"]):
        cell = ws_base.cell(row=2, column=j+1)  # celda del encabezado: j=0 material, j=1 norma
        cell.value     = h  # escribe el texto del encabezado
        cell.font      = Font(bold=True, color="000000")  # encabezados en negrita
        cell.alignment = Alignment(horizontal="center", vertical="center")  # encabezados centrados
        cell.fill      = fill_header  # fondo gris para encabezados
        cell.border    = borde_gris  # borde gris al encabezado
    ws_base.row_dimensions[2].height = 16  # alto de la fila de encabezados

    # ── Escribir filas desde BLOQUES_NORMATIVA ────────────────────────────────
    mat_anchor = {}   # {nombre del material en minúsculas: nº de fila de la hoja Base Normativa}. Más abajo, get_row_from_material() consulta este diccionario para saber a qué fila apuntar el hipervínculo
    current_row = 3   # las filas 1 y 2 son título y encabezado, los datos parten en la 3

    # recorre toda la tabla normativa y la va escribiendo fila por fila
    for mat, norma in BLOQUES_NORMATIVA:
        mat_anchor[mat.lower()] = current_row  # guarda en qué fila quedó este material (para el link después)
        for j, v in enumerate([mat, norma]):  # recorre columna material (j=0) y columna norma (j=1)
            cell = ws_base.cell(row=current_row, column=j+1)  # celda donde cae este dato
            cell.value     = v  # escribe el material o el texto de la norma
            cell.font      = Font(bold=(j==0), color="000000")  # la columna del material va en negrita
            cell.alignment = Alignment(vertical="center", wrap_text=True,  # el texto se ajusta solo si es muy largo (wrap_text)
                                       horizontal="left")  # y pegado a la izquierda
            cell.fill      = fill_blanco  # fondo blanco para las filas de datos
            cell.border    = borde_gris  # borde gris a la celda
        # fila alta porque el texto de la norma puede ser largo
        ws_base.row_dimensions[current_row].height = 50
        current_row += 1  # avanza a la siguiente fila

    # ── Anchos de columna ─────────────────────────────────────────────────────
    # ancho de columnas: A angosta (nombre del material), B ancha (texto
    # largo de la norma)
    ws_base.column_dimensions["A"].width = 45
    ws_base.column_dimensions["B"].width = 115  # la B va bien ancha porque el texto de la norma es largo
    ws_base.freeze_panes = "A3"  # deja fijas las filas 1 y 2 cuando se mueve hacia abajo en la pantalla

    # Botón volver
    ws_base["D1"].value     = "← Volver a Materiales"  # texto del botón
    ws_base["D1"].hyperlink = f"#'{sheet_materiales}'!A1"  # link interno a la otra hoja
    ws_base["D1"].font      = Font(color="0000FF", underline="single", bold=True)  # texto azul y subrayado, como un link de verdad
    ws_base["D1"].alignment = Alignment(horizontal="center")  # botón centrado
    ws_base.column_dimensions["D"].width = 22  # ancho de la columna del botón

    # ── Hipervínculos en hoja Materiales ──────────────────────────────────────
    # A partir de la descripción y el circuito, busca a qué fila de la hoja
    # normativa corresponde ese material.
    # Solo devuelve una fila si el material tiene entrada en BLOQUES_NORMATIVA.
    def get_row_from_material(desc, circ=""):
        # busca a qué fila de la hoja "Base Normativa" corresponde este material,
        # para armar el link. Compara palabras clave del texto (tipo de
        # material + tipo de circuito). Si no encuentra nada, devuelve None.
        #
        # De acá para abajo es puro "si el texto dice tal cosa, busca esta norma".
        # Está ordenado por secciones (RIC 1, RIC 2, RIC 4, etc.) y cada línea
        # ya se explica solita: el texto que busca y el nombre de la norma que
        # devuelve están en español. No hay vuelta que darle línea por línea,
        # es una lista larga de casos.
        import re  # re se usa más abajo, en extraer_sec
        d = str(desc).lower()  # descripción en minúsculas, para comparar sin importar mayúsculas
        c = str(circ).lower()  # nombre del circuito en minúsculas, por el mismo motivo


        # Función auxiliar: mira el texto del material y extrae el número de
        # sección (mm^2) que tenga escrito, para poder elegir la norma según
        # el calibre del conductor (por ejemplo, ≤6mm^2 o >6mm^2).
        def extraer_sec(texto):
            """Extrae sección en mm^2 desde el descriptor. Retorna float o None."""
            m = re.search(r"(\d+(?:[.,]\d+)?)\s*mm", texto.lower())  # busca un número seguido de "mm"
            if m:  # si encontró el patrón número + "mm" en el texto
                return float(m.group(1).replace(",", "."))  # lo convierte a número (cambia la coma por punto)
            return None  # no encontró ningún número de sección en el texto

        # A partir de aquí viene una lista larga de reglas: cada "if" revisa si
        # el texto de la descripción (d) o del nombre del circuito (c) contiene
        # ciertas palabras clave, y si calzan, devuelve el "ancla" (mat_anchor)
        # de la fila de la hoja Normativa que le corresponde a ese material.
        # Está ordenado por secciones del RIC, cada bloque con su título
        # "── RIC X ──". Son casos parecidos entre sí: se busca una palabra
        # clave y se devuelve la norma asociada, por eso no se comenta cada
        # "if" por separado.
        # ── RIC 7 - Climatización ─────────────────────────────────────────────
        # el "sin enchufe" es para saltarse los splits que van conectados directo,
        # sin enchufe de por medio
        if ("enchufe" in d or "2p+t" in d) and "sin enchufe" not in d and any(k in c for k in ("clima", "split", "aire")):
            return mat_anchor.get("enchufe climatización")  # manda a la fila del enchufe de clima
        # el TM del circuito de clima: las descripciones dicen "Disyuntor
        # termomagnético ...", por eso se busca por "disyuntor" y no por
        # "interruptor automático" (así calza con BLOQUES_NORMATIVA)
        if ("disyuntor" in d or "termomagnético" in d or "termomagnetico" in d) and any(k in c for k in ("clima", "split", "aire")):
            return mat_anchor.get("disyuntor termomagnético climatización")  # fila del TM de climatización
        if "interruptor diferencial" in d and any(k in c for k in ("clima", "split", "aire")):  # el diferencial exclusivo del circuito de clima
            return mat_anchor.get("interruptor diferencial climatización")  # fila del diferencial de climatización
        if ("canalización" in d or "canalizacion" in d or "conduit" in d or "canaleta" in d) and any(k in c for k in ("clima", "split", "aire")):  # conduit o canaleta, pero dentro de un circuito de clima
            return mat_anchor.get("canalización climatización")  # fila de la canalización de clima
        if "conductor" in d and ("rojo =" in d or "blanco =" in d) and any(k in c for k in ("clima", "split", "aire")):  # los conductores del clima se reconocen por el color escrito en la descripción
            return mat_anchor.get("conductor climatización")  # fila del conductor de clima

        # ── RIC 7 - Agua caliente ─────────────────────────────────────────────
        _kw_agua_c = ("ducha", "termo", "calefon", "calefón", "calentador", "agua caliente")  # palabras que delatan un circuito de agua caliente
        _es_agua_c = any(k in c for k in _kw_agua_c)  # queda en True si el circuito es ducha, termo o calefón
        if "diferencial" in d and "10" in d and "ma" in d.lower() and _es_agua_c:  # el diferencial de 10 mA se pesca por el "10" y el "ma" del texto
            return mat_anchor.get("interruptor diferencial 10ma agua caliente")  # fila del diferencial de 10 mA del agua caliente
        # el .replace("o","o") no cambia nada, quedó de alguna prueba
        if "tablero sobrepuesto" in d and "desconexión" in d.replace("o","o") and _es_agua_c:
            return mat_anchor.get("tablero desconexión externo agua caliente")  # fila del tablero de desconexión externo
        if ("interruptor automático" in d or "1p+n" in d) and "desconexión" in d and _es_agua_c:  # el TM bipolar que va en ese tablero de desconexión
            return mat_anchor.get("tm bipolar desconexión externo agua caliente")  # fila del TM bipolar de desconexión
        if "bornera" in d and "pe" in d and _es_agua_c:  # la bornera de tierra del tablero externo
            return mat_anchor.get("bornera pe tablero externo agua caliente")  # fila de la bornera PE
        if "prensaestopa" in d and _es_agua_c:  # prensaestopa, pero solo si el circuito es de agua caliente
            return mat_anchor.get("prensaestopa tablero externo agua caliente")  # fila del prensaestopa del tablero externo
        if ("canalización" in d or "canalizacion" in d or "conduit" in d or "canaleta" in d) and _es_agua_c:  # la canalización del circuito de agua caliente
            return mat_anchor.get("canalización climatización")  # usa la misma fila que la canalización de clima
        if "conductor" in d and ("rojo =" in d or "blanco =" in d) and _es_agua_c:  # los conductores del circuito de agua caliente
            return mat_anchor.get("conductor climatización")  # también reutiliza la fila del conductor de clima

        # Estos no tienen sección propia: son artículos generales (luminarias,
        # protectores de sobretensión, barras) que no dependen del tipo de circuito.
        if "portalámpara" in d or "portalampara" in d:
            return mat_anchor.get("portalámpara / luminaria circuito interior")  # fila del centro de iluminación
        if any(x in d for x in ["foco", "luminaria", "ampolleta", "panel led", "tubo led",  # cualquier cosa que ilumine se respalda con la misma norma
                                  "tubo fluorescente", "aplique led", "aplique"]):  # más nombres de luminaria que aparecen en el listado de materiales
            return mat_anchor.get("portalámpara / luminaria circuito interior")  # todas apuntan a la misma fila de luminaria
        if "supresor" in d or "spd" in d:  # el supresor de transientes
            return mat_anchor.get("supresor de transiente (spd)")  # fila del SPD
        if "protector sobrevoltaje" in d or "protector de sobrevoltaje" in d:  # el protector de sobrevoltaje
            return mat_anchor.get("protector sobrevoltaje y corriente")  # fila del protector de sobrevoltaje
        if "barra repartidora bipolar" in d or "barra bipolar" in d:  # la peineta bipolar del tablero
            return mat_anchor.get("barra repartidora bipolar")  # fila de la barra repartidora bipolar

        # ── RIC 1 ────────────────────────────────────────────────────────────
        # ojo: acá va el "and empalme in c". Si no, este if se traga TODOS los
        # disyuntores (los de cada circuito y el de clima) y los manda a la
        # fila del empalme, que no es la norma que les corresponde.
        if "disyuntor" in d and "empalme" in c:
            return mat_anchor.get("disyuntor termomagnético empalme")  # fila del disyuntor del empalme
        if "portafusible de loza" in d:  # el portafusible de loza del empalme
            return mat_anchor.get("portafusible de loza empalme")  # fila del portafusible de loza
        if "sellador" in d and ("teflón" in d or "teflon" in d):  # el sellador aparece escrito con y sin tilde
            return mat_anchor.get("sellador de roscas con teflón")  # fila del sellador con teflón
        if "abrazadera" in d and "caddy" in d:  # la abrazadera caddy
            return mat_anchor.get("abrazaderas tipo caddy")  # fila de las abrazaderas caddy
        if "conector hub" in d:  # el conector hub del empalme
            return mat_anchor.get("conector hub acero galvanizado")  # fila del hub galvanizado
        if "terminal ferrul doble" in d:  # el ferrul doble se revisa primero, si no lo pescaría el simple
            return mat_anchor.get("terminal ferrul doble alimentador")  # fila del ferrul doble
        if "terminal ferrul" in d and "acometida" in d:  # ferrul de la acometida
            return mat_anchor.get("terminal ferrul acometida")  # fila del ferrul de acometida
        if "terminal ferrul" in d and "alimentador" in d:  # ferrul del alimentador
            return mat_anchor.get("terminal ferrul alimentador")  # fila del ferrul de alimentador
        if "terminal ferrul" in d:  # cualquier otro ferrul es del cableado interno del tablero
            return mat_anchor.get("terminal ferrul tablero interior")  # fila del ferrul de tablero
        if "terminal" in d and ("compresión" in d or "compresion" in d or " ojo" in d):  # el terminal de compresión tipo ojo, escrito de varias formas
            return mat_anchor.get("terminal de compresión tipo ojo")  # fila del terminal de ojo
        if "cable" in d and ("concéntrico" in d or "concentrico" in d):  # el cable concéntrico de la acometida
            if "sub" in d or "subterráneo" in d or "subterraneo" in d:  # si dice subterráneo cambia el artículo
                return mat_anchor.get("cable acometida concéntrico subterráneo")  # fila del concéntrico subterráneo
            return mat_anchor.get("cable acometida concéntrico aéreo")  # si no, es el concéntrico aéreo
        if "mordaza" in d:  # la mordaza del cable aéreo
            return mat_anchor.get("mordaza para alimentador aéreo")  # fila de la mordaza
        if "cabeza de servicio" in d:  # la cabeza de servicio
            return mat_anchor.get("cabeza de servicio")  # fila de la cabeza de servicio
        if "caja de empalme" in d:  # empalme: caja de empalme metálica
            return mat_anchor.get("caja de empalme metálica")  # fila de la caja de empalme metálica
        # La abrazadera pvc cambia de nombre según el circuito: acometida,
        # alimentador, o los demás circuitos interiores.
        if "abrazadera" in d and "pvc" in d:  # abrazaderas del conduit pvc
            if "acometida" in c:  # las de la acometida tienen su propia fila
                return mat_anchor.get("abrazadera conduit pvc acometida")  # abrazadera del tramo de acometida
            if "alimentador" in c:  # las del alimentador van aparte
                return mat_anchor.get("abrazadera conduit pvc alimentador")  # abrazadera del tramo alimentador
            return mat_anchor.get("abrazadera conduit pvc circuitos")  # el resto de los circuitos usa la fila general
        if "terminal pvc conduit" in d or ("terminal" in d and "pvc" in d and "conduit" in d):  # terminal pvc de conduit, el que va con 2 tuercas
            return mat_anchor.get("terminal pvc conduit con 2 tuercas")  # fila del terminal con sus 2 tuercas
        if "salida de caja" in d:  # salida de caja para conduit pvc
            return mat_anchor.get("salida de caja conduit de pvc")  # fila de la salida de caja
        if "unidad de medida" in d:  # unidad de medida del empalme
            return mat_anchor.get("unidad de medida monofásica")  # acá el empalme siempre es monofásico
        # El conduit pvc también depende del circuito: alimentador o acometida
        # (subterránea). Si no es ninguno de los dos, no hay coincidencia.
        if ("conduit de pvc" in d or "conduit pvc" in d) and "abrazadera" not in d:  # tubo conduit pvc, ojo que no sea la abrazadera
            if "alimentador" in c:  # conduit del tramo alimentador
                return mat_anchor.get("conduit pvc alimentador")  # fila del conduit del alimentador
            if "acometida" in c:  # conduit de la acometida
                return mat_anchor.get("conduit pvc acometida subterránea")  # la acometida va enterrada, tiene su propia cita
            return None  # conduit de otro circuito: no hay cita para ese caso
        if "tubo conduit galvanizado" in d:  # conduit metálico galvanizado
            return mat_anchor.get("tubo conduit galvanizado")  # fila del tubo galvanizado
        if ("marco metálico" in d or "marco metalico" in d) and ("cámara" in d or "camara" in d):  # marco metálico de la cámara subterránea
            return mat_anchor.get("marco metálico cámara tipo c")  # fila del marco de la cámara tipo c
        if "boquilla" in d and ("cámara" in d or "camara" in d):  # boquilla de entrada a la cámara
            return mat_anchor.get("boquilla pvc cámara tipo c")  # fila de la boquilla de la cámara
        if "cámara tipo c" in d or "camara tipo c" in d:  # la cámara tipo c propiamente tal
            return mat_anchor.get("cámara tipo c subterránea")  # fila de la cámara subterránea

        # ── RIC 3 + RIC 4 conductores (con lógica de sección) ────────────────
        # Estos casos usan extraer_sec() para ver el calibre del conductor:
        # si la sección es mayor a 6mm^2, se usa una norma distinta a la de
        # los conductores más delgados (≤6mm^2).
        if "rv-k" in d and "alimentador" in c:  # conductor rv-k del alimentador
            sec = extraer_sec(d)  # saca la sección en mm^2 desde la descripción
            if sec and sec > 6:  # sobre 6mm^2 la cita del ric es otra
                return mat_anchor.get("alimentador rv-k cu >6mm^2")  # alimentador grueso
            return mat_anchor.get("alimentador rv-k cu ≤6mm^2")  # hasta 6mm^2 inclusive
        if ("flexible libre de halógenos" in d or "flexible libre de halogenos" in d) and "tablero" in c:  # conductor libre de halógenos que va dentro del tablero
            sec = extraer_sec(d)  # sección del conductor interior
            if sec and sec > 6:  # mismo corte de 6mm^2
                return mat_anchor.get("conductor interior tablero >6mm^2")  # conductor interior grueso
            return mat_anchor.get("conductor interior tablero ≤6mm^2")  # conductor interior delgado
        if ("flexible libre de halógenos" in d or "flexible libre de halogenos" in d):  # si no es para tablero, no hay coincidencia
            return None  # si no es del tablero, no se sabe qué fila usar
        if ("rojo =" in d or "blanco =" in d) and "conductor" in d:  # conductores de fase y neutro de los circuitos interiores
            sec = extraer_sec(d)  # sección del conductor del circuito
            if sec and sec > 6:  # sobre 6mm^2 cambia la cita
                return mat_anchor.get("conductores circuitos interiores >6mm^2")  # conductores interiores gruesos
            return mat_anchor.get("conductores circuitos interiores ≤6mm^2")  # conductores interiores delgados
        if "thwn-2" in d and "verde" in d:  # conductor verde: es la tierra de protección
            sec = extraer_sec(d)  # sección del conductor de tierra
            if sec and sec > 6:  # mismo corte de 6mm^2
                return mat_anchor.get("conductor thwn-2 verde pt >6mm^2")  # tierra gruesa
            return mat_anchor.get("conductor thwn-2 verde pt ≤6mm^2")  # tierra delgada
        if "thwn-2" in d and "blanco" in d:  # conductor blanco thwn-2
            sec = extraer_sec(d)  # sección del blanco
            if sec and sec > 6:  # mismo corte de 6mm^2
                return mat_anchor.get("conductor thwn-2 blanco pt >6mm^2")  # blanco grueso
            return mat_anchor.get("conductor thwn-2 blanco pt ≤6mm^2")  # blanco delgado

        # ── RIC 2 ────────────────────────────────────────────────────────────
        if "tablero embutido" in d:  # tablero embutido en el muro
            return mat_anchor.get("tablero embutido de pvc")  # fila del tablero embutido
        if "tablero sobrepuesto" in d:  # tablero montado sobre el muro
            return mat_anchor.get("tablero sobrepuesto de pvc")  # fila del tablero sobrepuesto
        if "interruptor general omnipolar" in d or "general omnipolar" in d:  # el general omnipolar del tablero
            return mat_anchor.get("interruptor general omnipolar")  # fila del interruptor general
        if "interruptor diferencial" in d:  # protección diferencial del tablero
            return mat_anchor.get("interruptor diferencial")  # fila del diferencial
        if "disyuntor" in d or "interruptor automático" in d or "interruptor automatico" in d:  # el TM de cada circuito interior
            return mat_anchor.get("disyuntor termomagnético")  # fila del TM de circuito (RIC 2 / RIC 5 / RIC 7 / RIC 10)
        if "luz piloto" in d:  # luz piloto del tablero
            return mat_anchor.get("luz piloto")  # fila de la luz piloto
        if "riel din" in d:  # riel din donde se montan las protecciones
            return mat_anchor.get("riel din")  # fila del riel din
        if "conductor" in d and "alimentador" in d and ("llegada" in d or "tablero" in d):  # el alimentador cuando ya llega al tablero
            return mat_anchor.get("conductor alimentador llegada tablero")  # fila del alimentador de llegada
        if ("flexible libre de halógenos" in d or "flexible libre de halogenos" in d) and "tablero" in c:  # código muerto: allá arriba, en el bloque RIC 3 + RIC 4, los "flexible libre de halógenos" de tablero ya devolvieron su fila y el resto ya devolvió None, así que acá nunca llega ninguno
            sec = extraer_sec(d)  # sección del conductor interior
            if sec and sec > 6:  # mismo corte de 6mm^2
                return mat_anchor.get("conductor interior tablero >6mm^2")  # conductor interior grueso
            return mat_anchor.get("conductor interior tablero ≤6mm^2")  # conductor interior delgado
        if "barra repartidora tetrapolar" in d or "barra tetrapolar" in d:  # barra que reparte las fases dentro del tablero
            return mat_anchor.get("barra repartidora tetrapolar")  # fila de la barra tetrapolar
        if "barra unipolar verde" in d:  # barra verde: es la de tierra de protección
            return mat_anchor.get("barra unipolar verde pe")  # fila de la barra pe
        if "bornera de conexión" in d or "bornera de conexion" in d:  # borneras de conexión del tablero
            return mat_anchor.get("bornera de conexión")  # fila de la bornera
        if "portafusible" in d and ("1p" in d or "32a" in d or "10x38" in d):  # portafusible del tablero, formato 10x38
            return mat_anchor.get("portafusible tablero")  # fila del portafusible
        if "fusible cilíndrico" in d or ("fusible" in d and "10x38" in d):  # el fusible que va dentro del portafusible
            return mat_anchor.get("fusible cilíndrico tablero")  # fila del fusible cilíndrico
        if "copperweld" in d:  # barra copperweld de la puesta a tierra
            return mat_anchor.get("barra copperweld + conector bronce")  # fila de la barra con su conector de bronce
        if "camarilla" in d and "naranjo" in d:  # camarilla naranja donde queda la barra de tierra
            return mat_anchor.get("camarilla pvc naranjo")  # fila de la camarilla
        if "conductor desnudo cu" in d:  # conductor desnudo de la puesta a tierra
            return mat_anchor.get("conductor desnudo cu 16mm^2")  # fila del desnudo de 16mm^2

        # ── RIC 4 canalizaciones y accesorios ────────────────────────────────
        if "canalización embutida" in d or "canalizacion embutida" in d:  # canalización embutida: va con conduit pvc
            return mat_anchor.get("canalización embutida conduit pvc")  # fila de la canalización embutida
        if "canalización sobrepuesta" in d or "canalizacion sobrepuesta" in d:  # canalización sobrepuesta: va con canaleta pvc
            return mat_anchor.get("canalización sobrepuesta canaleta pvc")  # fila de la canalización sobrepuesta
        if "boquilla" in d and "conduit" in d:  # boquillas de las puntas del conduit
            return mat_anchor.get("boquilla conduit pvc")  # fila de la boquilla de conduit
        if "unión copla" in d or "union copla" in d:  # copla para unir tramos de canaleta
            return mat_anchor.get("unión copla canaleta pvc")  # fila de la unión copla
        if "curvas internas" in d:  # curvas internas de la canaleta
            return mat_anchor.get("curvas internas 90° canaleta pvc")  # fila de las curvas internas
        if "curvas planas" in d:  # curvas planas de la canaleta
            return mat_anchor.get("curvas planas 90° canaleta pvc")  # fila de las curvas planas
        if "curva t" in d:  # curva t, para derivar la canaleta
            return mat_anchor.get("curva t canaleta pvc")  # fila de la curva t
        # Algunas cajas y tapas tienen variantes (octogonal, chuqui) además
        # del modelo estándar; el "if" de adentro elige la variante correcta.
        if "caja de derivación embutida" in d or "caja de derivacion embutida" in d:  # cajas de derivación embutidas
            if "octogon" in d:  # la octogonal es la del centro de luz
                return mat_anchor.get("caja derivación octogonal pvc")  # fila de la caja octogonal
            return mat_anchor.get("caja derivación embutida pvc")  # fila de la caja embutida común
        if "caja de paso" in d:  # cajas de paso del recorrido
            # Cajas de paso (interior o de alimentador, estancas): tienen su
            # propia fila en BLOQUES_NORMATIVA ("Caja de paso estanca"), con
            # exactamente el mismo texto de norma que "Caja derivación embutida
            # PVC" — la parte que las justifica es RIC 4 (7.16.1.11, 7.16.1.13).
            return mat_anchor.get("caja de paso estanca")  # todas apuntan a la fila de la caja de paso estanca
        if "caja de derivación sobrepuesta" in d or "caja de derivacion sobrepuesta" in d:  # cajas de derivación sobrepuestas
            if "chuqui" in d:  # variante chuqui
                return mat_anchor.get("caja derivación sobrepuesta chuqui pvc")  # fila de la sobrepuesta chuqui
            return mat_anchor.get("caja derivación sobrepuesta pvc")  # fila de la sobrepuesta común
        if "caja de derivación metálica" in d or "caja de derivacion metalica" in d:  # cajas de derivación metálicas
            return mat_anchor.get("caja derivación metálica")  # fila de la caja metálica
        if "tapa ciega" in d:  # tapas ciegas de las cajas
            if "octogon" in d:  # tapa de la caja octogonal
                return mat_anchor.get("tapa ciega octogonal pvc")  # fila de la tapa octogonal
            if "chuqui" in d:  # tapa de la caja chuqui
                return mat_anchor.get("tapa ciega chuqui pvc")  # fila de la tapa chuqui
            return mat_anchor.get("tapa ciega pvc")  # fila de la tapa común
        if ("enchufe" in d or "2p+t" in d) and "sin enchufe" not in d:  # módulos de enchufe 2p+t, salvo que diga sin enchufe
            return mat_anchor.get("enchufe")  # fila del enchufe
        if "interruptor 9/" in d:  # interruptores de circuito 9/12, 9/15, 9/24, 9/32
            return mat_anchor.get("interruptor de circuito")  # fila del interruptor
        if "conector cónico" in d or "conector conico" in d:  # conectores cónicos de las uniones de cables
            return mat_anchor.get("conector cónico")  # fila del conector cónico
        if "tierra" in d and ("caja metálica" in d or "caja metalica" in d):  # tornillo de tierra de las cajas metálicas
            return mat_anchor.get("tornillo conexión tierra caja metálica")  # fila del tornillo de tierra
        if "prensaestopa" in d:  # prensaestopa de las cajas estancas
            return mat_anchor.get("prensaestopa")  # fila del prensaestopa
        if "tubo de estaño" in d or "tubo de estano" in d:  # estaño para las soldaduras de la puesta a tierra
            return mat_anchor.get("tubo de estaño")  # fila del tubo de estaño
        if "pasta para soldar" in d:  # pasta para soldar
            return mat_anchor.get("pasta para soldar")  # fila de la pasta

        # Si no hay entrada en BLOQUES_NORMATIVA, no lleva hipervínculo
        return None  # sin cita, la celda de norma queda sin link

    # busca en qué columna de la hoja Materiales está cada dato que necesitamos
    try:
        col_desc = list(materiales_df.columns).index("Descripción técnica") + 1  # columna con la descripción técnica del material
    except ValueError:  # si la hoja no trae esa columna, pandas tira ValueError
        col_desc = None  # esa columna no existe en esta hoja

    try:  # ubica la columna Norma / RIC, que es donde va el link
        col_norma = list(materiales_df.columns).index("Norma / RIC") + 1  # columna donde se va a poner el link a la norma
    except ValueError:  # no está la columna de norma en esta hoja
        col_norma = None  # sin esta columna no se pueden poner links

    try:  # ubica la columna Sello SEC
        col_sello = list(materiales_df.columns).index("Sello SEC") + 1  # columna que indica si el material lleva Sello SEC
    except ValueError:  # no está la columna de sello
        col_sello = None  # columna opcional, si no está no pasa nada

    try:  # ubica la columna Circuito
        col_circ = list(materiales_df.columns).index("Circuito") + 1  # columna con el nombre del circuito
    except ValueError:  # no está la columna de circuito
        col_circ = None  # sin esta columna no se puede filtrar por circuito

    # solo agrega los links si existen esas dos columnas en la hoja
    if col_norma and col_desc:
        header_row_mat = 3  # fila 3 = encabezado de la hoja Materiales; los datos parten en la 4
        last_row_mat   = header_row_mat + len(materiales_df)  # última fila con datos en la hoja Materiales
        # recorre cada fila de materiales y le pone el link si corresponde
        for r in range(header_row_mat + 1, last_row_mat + 1):
            val_desc   = str(ws_mat.cell(row=r, column=col_desc).value or "")  # texto del material en esta fila
            val_circ   = str(ws_mat.cell(row=r, column=col_circ).value or "") if col_circ else ""  # nombre del circuito en esta fila (si existe la columna)
            cell_norma = ws_mat.cell(row=r, column=col_norma)  # celda donde va a ir el link a la norma
            destino    = get_row_from_material(val_desc, val_circ)  # busca a qué fila de la norma corresponde
            if destino:  # si encontró una norma para este material
                cell_norma.value     = "Ver normativa"  # pisa el "RIC N" que dejo normalizar_ric_materiales: en la celda se ve "Ver normativa" y el RIC queda en la hoja Base Normativa
                cell_norma.hyperlink = f"#'{sheet_base}'!A{destino}"  # link a la hoja Base Normativa
                cell_norma.font      = Font(color="0000FF", underline="single")  # lo pinta como un link (azul y subrayado)
                cell_norma.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)  # centra horizontal, pega arriba y deja que el texto salte de línea
            else:  # si no encontró norma, no hay link pero igual se centra el texto
                cell_norma.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)  # mismo formato, sin link
            # Centrar columna Sello SEC
            if col_sello:  # solo si esa columna existe en la hoja
                cell_sello = ws_mat.cell(row=r, column=col_sello)  # celda de la columna Sello SEC
                cell_sello.alignment = Alignment(horizontal="center", vertical="top")  # centra el valor

# =========================
# HELPERS (NO BORRAN NADA)
# =========================


# ===== Tabla de diámetros de conduit según sección y N° de conductores (Tabla RIC N°4.17) =====
def conduit_por_tabla(seccion_mm2, n_cond):
    """
    Tabla N°4.17 (H07V-U/R/K, etc.) -> Ø nominal ducto mm según sección y Nº conductores (1..5)
    Si n_cond > 5, lo limita a 5 (puedes ampliarlo después si quieres).
    """
    # Tabla: sección : {n_cond: ducto_mm}
    tabla = {  # se copia igual que la tabla del ric, para no calcular nada
        1.5: {1:16, 2:16, 3:16, 4:20, 5:20},  # 1,5mm^2: se queda en 16mm salvo con 4 o 5 conductores
        2.5: {1:16, 2:16, 3:20, 4:20, 5:20},  # 2,5mm^2: la sección típica de enchufes
        4.0: {1:16, 2:16, 3:20, 4:20, 5:25},   # según Tabla N°4.17 RIC (2 cond. = 16mm)
        6.0: {1:16, 2:20, 3:20, 4:25, 5:25},  # 6mm^2: con 4 o 5 conductores ya pide 25mm
        10.0:{1:16, 2:20, 3:25, 4:32, 5:32},  # 10mm^2: sube a 32mm
        16.0:{1:20, 2:25, 3:32, 4:40, 5:40},  # 16mm^2: la sección más gruesa que trae la tabla
        # si después necesitas más secciones, las agregamos
    }

    # elegir la sección "igual o superior" disponible en tabla
    secciones = sorted(tabla.keys())  # lista de secciones (mm^2) que existen en la tabla, de menor a mayor
    sec = None  # todavía no se encontró una sección que alcance
    for s in secciones:  # recorre las secciones de menor a mayor
        if float(seccion_mm2) <= float(s):  # ¿la sección real cabe en esta sección de tabla?
            sec = s  # ya encontró la sección de tabla que alcanza a cubrir la sección real
            break  # ya la encontró, no sigue buscando
    if sec is None:  # no encontró ninguna sección que alcance
        sec = secciones[-1]  # la sección real es más grande que todo lo que hay en la tabla, usa la máxima

    n = int(n_cond)  # nº de conductores como entero
    if n < 1: n = 1   # mínimo 1 conductor
    if n > 5: n = 5   # la tabla solo llega hasta 5 conductores

    return tabla[sec][n]  # busca en la tabla el ducto para esa sección y ese N° de conductores


# El TIPO de canalización (embutida o sobrepuesta) ya lo eligió el usuario y
# llega en tipo_canalizacion; acá sólo se decide la MEDIDA:
#   - embutida ..... diámetro del conduit por Tabla N°4.17 (conduit_por_tabla)
#   - sobrepuesta .. tamaño de canaleta por área ocupada (canalizacion_recomendada)
# Quien la llama es el bloque de cálculo de sección del script final, y el texto
# que devuelve queda guardado en la columna "Canalización" de cada circuito
# (circuitos_df); build_materiales_df después lee esa misma columna.
def canalizacion_recomendada_por_conductores(tipo_canalizacion, seccion_mm2, n_cond):
    """
    - Si es embutida: usa la tabla de conduit según sección y N° de conductores.
    - Si es sobrepuesta: usa el cálculo de canaleta por área real del conductor.
    """
    tipo = (tipo_canalizacion or "").strip().lower()  # normaliza el texto para comparar (sin mayúsculas ni espacios)

    if "embut" in tipo:  # ¿el usuario pidió canalización embutida?
        mm = conduit_por_tabla(seccion_mm2, n_cond)  # busca el diámetro del conduit en la tabla
        return f"Embutida (PVC conduit {mm}mm)"  # arma el texto final con el diámetro de conduit

    # sobrepuesta: usa cálculo dinámico por área real del conductor
    return canalizacion_recomendada(tipo_canalizacion, seccion_mm2, n_cond)  # arma el texto final con la canaleta que corresponda

# Calcula cuántos conductores necesita un circuito de iluminación para
# dimensionar su canalización, según cuántas luminarias hay y cuántas son conmutadas.
def n_conductores_iluminacion_para_circuito(circuito_items, ambientes_df):
    """
    Decide el N° de conductores para dimensionar canalización en iluminación,
    usando tu regla:
      - conmutado => 3
      - no conmutado: simple=3, doble=4, triple=5 (según luminarias del ambiente en ese circuito)
    Retorna el MÁXIMO requerido dentro del circuito.
    """
    if not isinstance(circuito_items, list) or len(circuito_items) == 0:  # circuito vacío o mal formado
        return 3  # sin items, se asume el caso más simple (3 conductores)

    # mapa conmutadas por ambiente desde ambientes_df
    ncon_map = {}  # ambiente -> cuántas luminarias conmutadas tiene
    if ambientes_df is not None and "Ambiente" in ambientes_df.columns:  # solo si la tabla de ambientes trae esa columna
        amb_col = ambientes_df["Ambiente"].astype(str).str.strip().str.lower()  # nombres de ambiente normalizados (sin mayúsculas ni espacios)
        if "N_conmutadas_924 (u)" in ambientes_df.columns:  # columna con la cantidad de luminarias conmutadas por ambiente
            vals = pd.to_numeric(ambientes_df["N_conmutadas_924 (u)"], errors="coerce").fillna(0).astype(int)  # convierte a número, lo que no se puede convertir queda en 0
            for a, v in zip(amb_col, vals):  # arma el diccionario ambiente -> cantidad conmutadas
                ncon_map[a] = int(max(0, v))  # cuántas luminarias conmutadas tiene cada ambiente

    # contar luminarias por ambiente dentro del circuito (desde _items)
    lum_by_amb = {}  # {ambiente: cuántas luminarias tiene en este circuito}. Deja fuera enchufes y cargas especiales, pero las conmutadas SI se cuentan acá; se descuentan más abajo con ncon_map
    for it in circuito_items:  # recorre cada item (luminaria, enchufe, etc.) del circuito
        if not isinstance(it, dict):  # descarta items que no sean diccionarios
            continue  # salta lo que no venga como diccionario
        # excluir enchufes y especiales
        if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it) or ("nombre" in it):  # esto no es una luminaria, no se cuenta
            continue  # los enchufes y especiales no cuentan acá
        amb = str(it.get("amb", "")).strip().lower() or "sin_amb"  # ambiente de este item, o 'sin_amb' si no tiene
        lum_by_amb[amb] = lum_by_amb.get(amb, 0) + 1  # suma una luminaria más a este ambiente

    n_max = 3  # mínimo posible (3 conductores)

    # recorre cada ambiente y ve cuántos conductores necesita, se queda con el peor caso
    for amb, n_lum in lum_by_amb.items():  # recorre cada ambiente que tuvo luminarias en este circuito
        n_conmutadas = int(ncon_map.get(amb, 0))  # cuántas luminarias conmutadas hay en este ambiente
        n_conmutadas = max(0, min(n_conmutadas, int(n_lum)))  # no puede haber más conmutadas que luminarias
        n_restantes = int(n_lum) - n_conmutadas  # luminarias que NO son conmutadas
        # Base: 9/12 => 3 conductores
        n_req = 3  # parte en 3, que es el caso base 9/12
        # Si hay conmutado (9/24)
        if n_conmutadas > 0:  # si hay conmutado (9/24)
            n_req = max(n_req, 3)  # el conmutado igual se resuelve con 3, no sube nada
        # Para el resto de luminarias
        if n_restantes > 0:  # ahora las luminarias que no son conmutadas
            c12, c15, c32 = descomponer_interruptores(int(n_restantes))  # cómo se agrupan las luminarias restantes
            if c32 > 0:  # si hay un 9/32 se necesitan 5 conductores
                n_req = max(n_req, 5)  # existe 9/32
            elif c15 > 0:  # si hay un 9/15 bastan 4
                n_req = max(n_req, 4)  # existe 9/15
            else:  # puros 9/12
                n_req = max(n_req, 3)  # solo 9/12
        n_max = max(n_max, int(n_req))  # se queda con el máximo requerido entre todos los ambientes
    return int(n_max)  # el mayor N° de conductores que pidió algún ambiente


# Reparte una lista de cargas (items con potencia) en varios circuitos (bins),
# de modo que ningún circuito supere la potencia máxima permitida (Pmax).
def binpack_items(items, Pmax):
    """
    Agrupa items (cargas) en bins/subcircuitos para que cada bin no supere Pmax [W].
    items: lista de dicts, mínimo {"amb": "...", "potencia": float}
    """
    # ordena las cargas de mayor a menor potencia (First-Fit Decreasing)
    items_sorted = sorted(items, key=lambda x: float(x.get("potencia", 0.0)), reverse=True)  # las cargas más grandes se ubican primero
    bins = []  # lista de circuitos (bins) ya armados
    # recorre cada carga ya ordenada
    for it in items_sorted:
        p = float(it.get("potencia", 0.0))  # potencia de la carga
        placed = False  # ¿ya se asignó a algún circuito?
        for b in bins:  # recorre los circuitos ya creados
            if b["potencia_total"] + p <= Pmax:  # ¿cabe sin superar el máximo?
                b["items"].append(it)  # agrega la carga al circuito
                b["potencia_total"] += p  # actualiza la potencia total
                placed = True  # marca que esta carga ya quedó ubicada
                break  # sale al encontrar el primer circuito que sirve
        if not placed:  # si no cupo en ninguno, abre un circuito nuevo
            bins.append({"items": [it], "potencia_total": p})  # no cupo en ninguno: crea un circuito nuevo con esta sola carga
    return bins  # cada bin trae sus items y la potencia total que acumuló

def pedir_longitud_sub(nombre, default_L):  # pide por consola el largo real de un tramo, en metros
    # pregunta la longitud real de un tramo (en metros). Si el usuario
    # deja vacío (solo ENTER), usa default_L. Insiste hasta que le den
    # un número válido.
    while True:  # sigue preguntando hasta obtener un valor válido
        s = input(
            f"       • Longitud REAL del '{nombre}' en metros (ENTER = usar {default_L} m)\n"
            f"         (En tramos horizontales: recorridos a 0,30m del cielo y 0,20m del piso - RIC N°4 7.16.1.16): "
        ).strip()  # lee lo que escribió el usuario
        if s == "":  # el usuario no escribió nada
            return float(default_L)  # dejó vacío, usa el valor por defecto
        try:  # intenta leer el largo como número; si no lo es, salta al except
            v = float(s)  # intenta convertir a número
            if v <= 0:  # un largo de 0 o negativo no sirve
                print("         ! Debe ser mayor que 0.")  # le avisa y vuelve a preguntar
                continue  # no sirve, vuelve a preguntar
            return v  # número válido
        except:  # lo que escribió no era un número
            print("         ! Ingrese un número válido.")  # no era un número

def pedir_float_positivo(prompt):  # pide un número obligatorio, siempre mayor que 0
    # pregunta un número que tiene que ser mayor que 0. Si escriben
    # cualquier cosa que no sirva, vuelve a preguntar.
    while True:  # sigue preguntando hasta obtener un valor válido
        s = input(prompt).strip()  # lee lo que escribió el usuario
        try:  # intenta leer el número; si viene basura, avisa y repregunta
            v = float(s)  # intenta convertirlo a número
            if v <= 0:  # tiene que ser mayor que 0
                print("         ! Debe ser un número mayor que 0.")  # le avisa y vuelve a preguntar
                continue  # vuelve a preguntar
            return v  # número válido, listo
        except:  # lo que escribió no era un número
            print("         ! Ingrese un número válido (ej: 7.8).")  # no era un número

def pedir_float_opcional(prompt, valor_defecto=0.0, min_val=0.0):  # pide un número opcional, con valor por defecto y mínimo
    # pregunta un número, pero es opcional: si dejan vacío usa "valor_defecto".
    # Si escriben algo, tiene que ser >= min_val o vuelve a preguntar.
    # qué recibe:
    #   prompt ......... la pregunta que se le muestra al usuario
    #   valor_defecto .. qué usar si el usuario aprieta ENTER sin escribir nada
    #   min_val ........ el mínimo aceptado
    # devuelve el número (float) que quedó.
    while True:  # sigue preguntando hasta obtener un valor válido
        s = input(prompt).strip()  # lee lo que escribió el usuario
        if s == "":  # apretó ENTER sin escribir nada
            return float(valor_defecto)  # dejó vacío, usa el valor por defecto
        try:  # intenta leer el número; si no se puede, vuelve al inicio del ciclo
            v = float(s)  # intenta convertirlo a número
            if v < min_val:  # no llega al mínimo pedido
                print(f"         ! Debe ser >= {min_val}.")  # le muestra cuál es el mínimo
                continue  # no cumple el mínimo, vuelve a preguntar
            return v  # número válido
        except:  # lo que escribió no era un número
            print("         ! Ingrese un número válido.")  # no era un número

def limpiar_nombre_circuito(txt):  # deja prolijo el nombre del circuito antes de usarlo
    # limpia el nombre del circuito que escribió el usuario: saca espacios
    # de más y arregla el típico error de tipeo "eenc..." (lo deja como "enc...").
    t = (txt or "").strip()  # saca espacios al inicio/final
    low = t.lower()  # versión en minúsculas, para comparar
    if low.startswith("eenc"):  # error típico de tipeo
        t = t[1:].strip()  # saca la "e" repetida
    t = " ".join(t.split())  # deja un solo espacio entre palabras
    return t  # nombre ya limpio

def sugerir_nombre_circuito(nombre):  # revisa si el nombre del circuito es uno de los 3 tipos, mal escrito
    """Solo sugiere corrección si el nombre parece ser iluminacion,
    enchufes o climatizacion pero está mal escrito."""
    import difflib, unicodedata  # difflib compara textos parecidos, unicodedata saca las tildes

    # pasa el texto a minúsculas y le saca las tildes, para poder comparar
    def normalizar(s):
        # pasa a minúsculas y saca las tildes, así "Iluminación" e "iluminación"
        # se pueden comparar como si fueran lo mismo.
        s = s.lower().strip()  # a minúsculas y sin espacios de sobra
        s = unicodedata.normalize("NFD", s)  # separa cada letra de su tilde (ej: 'a' + acento)
        return "".join(c for c in s if unicodedata.category(c) != "Mn")  # arma el texto de nuevo, sin los caracteres de tilde (categoría Mn)

    tipos_clave = ["ilumin", "enchufe", "clima", "aire", "split", "ac ", "a/c"]  # palabras que ya indican bien el tipo de circuito
    base_norm = normalizar(nombre)  # nombre del circuito sin tildes ni mayúsculas

    # Si ya contiene una clave reconocida, no hacer nada
    if any(k in base_norm for k in tipos_clave):  # si ya trae una palabra clave, el nombre está bien
        return nombre  # está bien escrito, no hay que sugerir nada

    # Solo comparar contra los 3 tipos que importan
    tipos_canonicos = ["iluminacion", "enchufes", "climatizacion"]  # los 3 tipos de circuito válidos
    matches = difflib.get_close_matches(base_norm, tipos_canonicos, n=1, cutoff=0.55)  # busca el más parecido
    if matches:  # encontró algo parecido
        sugerido = matches[0]  # el tipo más parecido al nombre escrito
        # pregunta si el nombre que escribió en realidad era el tipo sugerido
        resp = input(
            f"       ! No se reconoció '{nombre}' como tipo de circuito. "
            f"¿Quisiste decir '{sugerido}'? (si/no): "
        ).strip().lower()  # respuesta del usuario, en minúsculas
        if resp == "si":  # el usuario confirma que sí era ese tipo
            return sugerido  # el usuario aceptó la corrección
    return nombre  # se deja el nombre tal cual lo escribió

# Arma el texto de la columna Canalización: si va embutida calcula el conduit
# de PVC según la sección, y si va sobrepuesta elige la canaleta que cumpla
# con el 40% de ocupación.
# OJO: en la práctica sólo corre la rama sobrepuesta. El único que llama a esta
# función es canalizacion_recomendada_por_conductores(), y ese ya resolvió antes
# el caso embutido con conduit_por_tabla(). Por eso conduit_map (acá abajo) hoy
# queda sin uso real.
def canalizacion_recomendada(tipo_canalizacion, seccion_mm2, n_cond=3):
    """
    Devuelve el texto final para la columna Canalización:
    - Embutida (PVC conduit XXmm)
    - Sobrepuesta (Canaleta de PVC AAxBBmm)

    Qué recibe:
      tipo_canalizacion .. "embutida" o "sobrepuesta" (lo que eligió el usuario)
      seccion_mm2 ........ sección del conductor del circuito, en mm^2
      n_cond ............. cuántos conductores van juntos por ese ducto (por defecto 3)
    """
    conduit_map = {  # cada sección en mm^2 con el conduit de PVC que le corresponde
        1.5: 16, 2.08: 16,  # 1.5 y 2.08 mm^2 caben en conduit de 16 mm
        2.5: 20, 3.31: 20,  # 2.5 y 3.31 mm^2 ya piden 20 mm
        4.0: 25, 5.26: 25, 6.0: 25,  # de 4 a 6 mm^2 se van a conduit de 25 mm
        8.37: 32, 10.0: 32,  # 8.37 y 10 mm^2 necesitan 32 mm
        13.3: 40, 16.0: 40  # 13.3 y 16 mm^2 son los más gruesos de la tabla, 40 mm
    }  # para cada sección en mm^2, su diámetro de conduit en mm
    # Área real del conductor con aislación H07Z1-K, valores aproximados de catálogo
    AREA_CON_AISLACION = {
        1.5:  9.08,  # 1.5 mm^2 ocupa 9.08 mm^2 con la aislación puesta
        2.08: 5.31,  # 2.08 mm^2 (AWG 14); ojo: los valores AWG de esta tabla no siguen el mismo criterio que los metricos
        2.5:  13.20,  # 2.5 mm^2 ocupa 13.2 mm^2
        3.31: 7.16,  # 3.31 mm^2 (AWG 12) ocupa 7.16 mm^2
        4.0:  18.10,  # 4 mm^2 ocupa 18.1 mm^2
        5.26: 11.40,  # 5.26 mm^2 (AWG 10) ocupa 11.4 mm^2
        6.0:  22.06,  # 6 mm^2 ocupa 22.06 mm^2
        8.37: 23.67,  # 8.37 mm^2 (AWG 8) ocupa 23.67 mm^2
        10.0: 36.32,  # 10 mm^2 ocupa 36.32 mm^2
        13.3: 33.59,  # 13.3 mm^2 (AWG 6) ocupa 33.59 mm^2
        16.0: 51.53,  # 16 mm^2 es el más gordo, 51.53 mm^2
    }

    # Canaletas disponibles: para cada dimensión, su área total en mm^2
    CANALETAS = [
        ('20x10',  200),  # canaleta chica: 20x10 mm da 200 mm^2 de área
        ('32x12',  384),  # 32x12 mm da 384 mm^2
        ('40x16',  640),  # 40x16 mm da 640 mm^2
        ('60x25', 1500),  # 60x25 mm es la más grande, 1500 mm^2
    ]
    FACTOR_OCUPACION = 0.40  # solo se puede ocupar el 40% del área de la canaleta

    def seleccionar_canaleta(seccion_mm2, n_conductores):  # elige la canaleta más chica que aguante los conductores
        """Selecciona la canaleta más pequeña que cumpla con el factor de ocupación."""
        # seccion_mm2 = sección del conductor (ej: 2.5)
        # n_conductores = cuántos conductores van juntos dentro de la misma canaleta

        # Buscar área del conductor más cercana
        # claves = las secciones que existen en la tabla AREA_CON_AISLACION (1.5, 2.5, 4.0, etc.)
        claves = sorted(AREA_CON_AISLACION.keys())  # ordena las secciones de la tabla de áreas
        key = claves[-1]  # por si seccion_mm2 es más grande que todo lo que hay en la tabla
        for k in claves:  # recorre las secciones de la tabla, de menor a mayor
            if float(seccion_mm2) <= float(k):  # busca la primera que cubra la sección real del conductor
                key = k  # primera sección de la tabla que alcanza a cubrir la real
                break  # ya la encontró, no sigue buscando
        area_cond = AREA_CON_AISLACION[key]  # área de UN conductor (mm^2)
        area_necesaria = area_cond * n_conductores  # área total ocupada por todos los conductores juntos
        # Seleccionar canaleta más pequeña que cumpla
        # dim = nombre de la canaleta (ej: "20x10"), area_total = su área en mm^2
        for dim, area_total in CANALETAS:
            if area_total * FACTOR_OCUPACION >= area_necesaria:  # la canaleta solo se puede llenar hasta el 40% de su área
                return dim  # esta canaleta ya alcanza con el 40% de ocupación
        return CANALETAS[-1][0]  # ninguna alcanzó: usa la más grande de la lista

    tipo = (tipo_canalizacion or "").strip().lower()  # normaliza el texto para comparar
    claves = sorted(AREA_CON_AISLACION.keys())  # vuelve a ordenar las secciones (mismo cálculo de más arriba)

    key = claves[-1]  # parte con la sección más grande por si nada calza
    for k in claves:  # recorre las secciones de la tabla de áreas
        if float(seccion_mm2) <= float(k):  # la primera que cubre la sección real del circuito
            key = k  # sección de la tabla que corresponde a la sección real
            break  # ya la encontró, corta la búsqueda

    # decide si el circuito va embutido (dentro de un conduit) o sobrepuesto (dentro de una canaleta)
    if "embut" in tipo:
        mm = conduit_map.get(key, 20)  # diámetro de conduit (20mm si no está en la tabla)
        return f"Embutida (PVC conduit {mm}mm)"  # embutida: se informa el conduit de PVC con su diámetro
    else:  # no va embutido, entonces es canaleta sobrepuesta
        dim = seleccionar_canaleta(seccion_mm2, n_cond if n_cond else 3)  # sobrepuesta: elige el tamaño de canaleta según cuántos conductores van
        return f"Sobrepuesta (Canaleta de PVC {dim}mm)"  # texto final con la medida de la canaleta

# Arma un texto resumen agrupando los items (enchufes, luminarias o cargas especiales)
# por ambiente, para mostrar en el informe. "modo" indica qué tipo de items son.
def resumen_items_por_ambiente(items, modo="enchufe"):
    """
    items: lista de dicts con al menos {"amb","potencia"}.
    Devuelve texto detallado por ambiente.
    """

    if modo == "especial":  # modo especial: cada carga se lista por separado, no se suman
        # circuitos especiales: cada item es una carga con nombre propio (ej: "horno")
        partes = []  # acá se van juntando los trozos de texto, uno por carga
        for it in items:  # cada item es una carga especial del listado
            amb = it.get("amb", "sin_amb")  # ambiente donde está la carga
            nombre = it.get("nombre", "carga")  # nombre de la carga (horno, encimera, etc.)
            potencia = int(round(float(it.get("potencia", 0.0))))  # potencia en watts, redondeada
            partes.append(f"{amb}: {nombre} ({potencia}W)")  # arma el trozo 'ambiente: carga (potencia)'
        return "; ".join(partes) if partes else ""  # junta todo con punto y coma, o vacío si no había nada

    # para enchufes o iluminación: se suma potencia y cantidad por ambiente
    res = {}  # acá se acumula potencia y cantidad por cada ambiente
    for it in items:  # recorre los enchufes o las luminarias, según el modo
        a = it.get("amb", "sin_amb")  # nombre del ambiente
        p = float(it.get("potencia", 0.0))  # potencia de este item

        if modo == "enchufe":  # modo enchufe: la cantidad viene en n_ench
            n = int(it.get("n_ench", 1))  # cantidad de enchufes de este item
        elif modo == "iluminacion":  # modo iluminación: la cantidad viene en n_lum
            n = int(it.get("n_lum", 1))  # cantidad de luminarias de este item
        else:  # cualquier otro modo
            n = 0  # modo desconocido: no cuenta unidades, solo potencia

        if a not in res:  # ambiente nuevo, hay que inicializarlo
            res[a] = {"p": 0.0, "n": 0}  # primera vez que aparece este ambiente

        res[a]["p"] += p  # va sumando la potencia del ambiente
        res[a]["n"] += n  # va sumando la cantidad del ambiente

    # arma el texto final, un trozo por ambiente
    partes = []  # trozos de texto, uno por ambiente
    for a, v in res.items():  # a = ambiente, v = sus totales acumulados
        ptxt = int(round(v["p"]))  # potencia total del ambiente, en watts enteros
        if modo == "enchufe":  # en enchufes se muestra la cantidad de enchufes
            partes.append(f"{a}: {v['n']} ench ({ptxt}W)")  # queda tipo 'cocina: 4 ench (900W)'
        elif modo == "iluminacion":  # en iluminación se muestra la cantidad de luminarias
            partes.append(f"{a}: {v['n']} lum ({ptxt}W)")  # queda tipo 'living: 3 lum (60W)'
        else:  # modo distinto, no hay unidades que contar
            partes.append(f"{a}: {ptxt}W")  # otros modos: solo la potencia del ambiente

    return ", ".join(partes) if partes else ""  # texto final separado por comas

# Lee el texto de un TM (protección termomagnética) como "1x16A curva C"
# y saca solo el número de corriente nominal (In).
def parse_in_tm(tm_text):
    # saca el número de corriente (In) de un texto tipo "1x16A curva C" (da 16).
    # Si no logra entenderlo, devuelve None.
    # ejemplo: "1x16A curva C" da 16
    if tm_text is None:  # sin texto de TM no hay nada que leer
        return None  # sin TM no hay corriente nominal que informar
    s = str(tm_text).lower().replace(" ", "")  # saca espacios y pasa a minúsculas
    try:  # si el texto viene raro, mejor no reventar
        if "x" in s and "a" in s:  # tiene forma 1x16a, 2x25a...: polos, "x", corriente y la "a" de amperes
            after_x = s.split("x", 1)[1]  # se queda con lo que hay después de la "x"
            num = ""  # acá se van juntando los dígitos de la corriente
            for ch in after_x:  # va juntando los dígitos del principio
                if ch.isdigit():  # mientras sean dígitos, es parte del número de la corriente
                    num += ch  # sigue siendo número, se suma al In
                else:  # llegó la letra a u otra cosa
                    break  # se corta apenas encuentra algo que no es número (la "a")
            return int(num) if num else None  # devuelve el In como entero, o None si no junto ningún digito
    except:  # cualquier error leyendo el texto
        return None  # texto raro, mejor no devolver corriente
    return None  # no tenía forma de TM, no se pudo sacar el In

# A partir de la cantidad de luminarias de un circuito, decide cuántos
# interruptores de cada tipo (9/12 = simple, 9/15 = doble, 9/32 = triple) usar.
def descomponer_interruptores(n_lum):
    """
    Regla (según cuántas luminarias hay):
      1 luminaria:  9/12
      2 luminarias: 9/15
      3 luminarias: 9/32
      4 luminarias: 9/32 + 9/12
      5 luminarias: 9/32 + 9/15
      6 luminarias: 9/32 + 9/32
      ... (primero se arman grupos de a 3, después de a 2, al final de a 1)
    Devuelve (c12, c15, c32)
    """
    n = int(max(0, n_lum))  # cantidad de luminarias, nunca negativo
    c32 = n // 3      # cuántos interruptores 9/32 entran (grupos de 3)
    rem = n % 3        # lo que sobra después de armar los grupos de 3
    c15 = 1 if rem == 2 else 0  # si sobran 2, va un 9/15
    c12 = 1 if rem == 1 else 0  # si sobra 1, va un 9/12
    return c12, c15, c32  # devuelve cuántos 9/12, 9/15 y 9/32 hay que comprar

def _get_amb_row(ambientes_df, amb_name):  # saca la fila del ambiente desde la tabla de ambientes
    # busca la fila del ambiente "amb_name" dentro de ambientes_df
    # (sin importar mayúsculas ni espacios). Si no la encuentra, None.
    if ambientes_df is None or "Ambiente" not in ambientes_df.columns:  # sin datos de ambientes o sin la columna, no hay dónde buscar
        return None  # no hay datos de ambientes, no se puede buscar
    key = str(amb_name).strip().lower()  # normaliza el nombre que estamos buscando
    m = ambientes_df[ambientes_df["Ambiente"].astype(str).str.strip().str.lower() == key]  # filtra la fila que calza
    if len(m) == 0:  # el filtro no encontró ninguna fila
        return None  # no existe ese ambiente
    return m.iloc[0]  # devuelve la primera (y única) coincidencia

# descripción automática SIN PREGUNTAR NADA MÁS
# Decide si una luminaria se conecta con estaño (cable) o a bornes (portalámpara/ampolleta),
# revisando palabras clave en el tipo y la descripción que escribió el usuario.
def _lum_necesita_estaño(tipo_txt, desc_txt=""):
    """Retorna True si la luminaria se conecta con estaño (LED con cable), False si va a bornes (portalámpara/ampolleta)."""
    t = str(tipo_txt or "").strip().lower()  # tipo de luminaria, en minúsculas
    d = str(desc_txt or "").strip().lower()  # descripción extra, en minúsculas
    if any(x in t or x in d for x in ["ampol", "incand", "portalamp", "portalámp", "e27", "e14"]):  # palabras típicas de ampolleta o portalámpara
        return False  # va a bornes (ampolleta/portalámpara), no lleva estaño
    if any(x in t or x in d for x in ["foco", "panel", "led", "aplique", "tubo", "fluores"]):  # palabras de luminarias que vienen con cable, no con casquillo
        return True  # va con cable, se conecta con estaño
    return False  # por defecto, no requiere estaño


# Arma automáticamente el texto descriptivo de una luminaria (para el informe y
# la lista de materiales), a partir del tipo, el montaje y la potencia que ya
# ingresó el usuario, sin preguntarle nada extra.
def desc_luminaria_auto(tipo_txt, montaje_txt, potencia_w):
    """
    No pregunta nada extra. Solo usa lo que ya ingresas:
    - tipo (texto que el usuario escribió)
    - montaje (embutida/sobrepuesta)
    - potencia (W ingresada)
    Devuelve la descripción "bonita".
    """
    t = (tipo_txt or "").strip().lower()  # tipo que escribió el usuario
    m = (montaje_txt or "").strip().lower()  # embutida o sobrepuesta
    try:  # la potencia puede venir como texto o vacía
        pw = int(round(float(potencia_w)))  # potencia redondeada, en watts
    except:  # no se pudo leer la potencia
        pw = 0  # si viene mal, se deja en 0

    # detecta qué tipo de luminaria es, buscando palabras clave en el texto
    es_foco = ("foco" in t) or ("panel" in t)  # foco o panel
    es_aplique = ("aplique" in t)  # aplique de muro
    es_tubo = ("tubo" in t) or ("fluores" in t)  # tubo o equipo fluorescente
    es_ampolleta = ("ampol" in t) or ("incand" in t)  # ampolleta o incandescente

    incand = ("incand" in t)  # ¿el texto menciona incandescente? (usada también más abajo por ampolleta)
    fluor = ("fluores" in t)  # ¿el texto menciona fluorescente?

    # tecnología para foco/panel y aplique: por defecto LED, salvo que el
    # texto que escribió el usuario diga explícitamente incandescente o
    # fluorescente (tubo y ampolleta calculan su propia tecnología más abajo)
    if incand:
        tecnologia = "incandescente"  # el usuario pidió incandescente
    elif fluor:  # el texto menciona fluorescente
        tecnologia = "fluorescente"  # queda como fluorescente
    else:  # no dijo nada de la tecnología
        tecnologia = "LED"  # hoy por defecto todo es LED

    if es_foco:  # foco/panel: el texto cambia según el montaje
        if "embut" in m:  # va embutido en el cielo
            return f"Foco panel {tecnologia} cuadrado/redondo embutido {pw} W luz cálida/fría/neutro"  # descripción del foco panel embutido
        else:  # no va embutido, entonces va sobrepuesto al cielo
            return f"Foco panel {tecnologia} cuadrado/redondo sobrepuesto {pw} W luz cálida/fría/neutro"  # descripción del foco panel sobrepuesto

    if es_aplique:  # aplique de muro
        # arma el texto indicando si el aplique va embutido o sobrepuesto
        montaje_apl = "embutido" if "embut" in m else "sobrepuesto"
        return f"Aplique {tecnologia} {montaje_apl} {pw} W luz cálida/fría/neutro"  # descripción final del aplique

    # tubo y ampolleta calculan su propia tecnología (fluorescente/incandescente)
    # más abajo, por eso no usan la variable "tecnología" definida arriba

    if es_tubo:  # tubo o equipo fluorescente
        # arma el texto indicando si el tubo va embutido o sobrepuesto
        montaje_tubo = "embutido" if "embut" in m else "sobrepuesto"
        if ("fluores" in t):  # el usuario pidió fluorescente explícitamente
            return f"Tubo fluorescente {montaje_tubo} {pw} W luz cálida/fría/neutro"  # tubo fluorescente
        else:  # no dijo fluorescente
            return f"Tubo LED {montaje_tubo} {pw} W luz cálida/fría/neutro"  # si no dijo nada, los tubos se toman como LED

    if es_ampolleta:  # ampolleta con portalámpara
        montaje_amp = "embutida" if "embut" in m else "sobrepuesta"  # en ampolleta el montaje va en femenino
        if incand:  # el texto decía incandescente
            return f"Ampolleta incandescente {pw} W {montaje_amp} luz cálida/fría/neutro"  # ampolleta incandescente
        else:  # no decía incandescente
            return f"Ampolleta LED {pw} W {montaje_amp} luz cálida/fría/neutro"  # por defecto la ampolleta se toma LED

    # no calzó con ningún tipo conocido: descripción genérica
    if "embut" in m:
        return f"Luminaria {pw} W (embutida) luz cálida/fría/neutro"  # genérica embutida
    return f"Luminaria {pw} W (sobrepuesta) luz cálida/fría/neutro"  # genérica sobrepuesta

# =========================================================
# ALIMENTADOR (selección por caída de tensión + tabla)
# =========================================================
# Lista de secciones de conductor (en mm^2) disponibles para el alimentador,
# de menor a mayor. Sirve para recorrer y elegir la sección más chica que cumpla.
SECCIONES_MM2_ALIM = [1.5, 2.08, 2.5, 3.31, 4, 5.26, 6, 8.37, 10, 13.3, 16, 21.1, 25, 26.7, 33.6, 35, 42.4, 50]
# Tabla de ampacidad (corriente máxima admisible, en A) del alimentador según
# el método de instalación (B1=ducto, D1=subterráneo, E=aéreo) y la sección (mm^2).
TABLA_AMPACIDAD_ALIM = {
    "B1": {  # ducto
        1.5: 18, 2.08: 24, 2.5: 24, 3.31: 31, 4: 37, 5.26: 39, 6: 48, 8.37: 59,  # B1, secciones chicas: 1.5 mm^2 aguanta 18 A y 8.37 mm^2 llega a 59 A
        10: 66, 13.3: 79, 16: 88, 21.1: 105, 25: 117, 26.7: 122, 33.6: 141, 35: 144,  # B1, secciones medias: de 10 mm^2 (66 A) hasta 35 mm^2 (144 A)
        42.4: 163, 50: 175  # B1, las más gruesas: 42.4 mm^2 con 163 A y 50 mm^2 con 175 A
    },
    "D1": {  # subterráneo
        1.5: 19, 2.08: 30, 2.5: 33, 3.31: 38, 4: 42, 5.26: 48, 6: 52, 8.37: 63,  # D1 enterrado, secciones chicas: 1.5 mm^2 con 19 A hasta 8.37 mm^2 con 63 A
        10: 68, 13.3: 80, 16: 89, 21.1: 103, 25: 113, 26.7: 117, 33.6: 132, 35: 136,  # D1, secciones medias: 10 mm^2 con 68 A hasta 35 mm^2 con 136 A
        42.4: 150, 50: 159  # D1, las más gruesas: 42.4 mm^2 con 150 A y 50 mm^2 con 159 A
    },
    "E": {   # aéreo
        1.5: 19, 2.08: 28, 2.5: 32, 3.31: 38, 4: 42, 5.26: 50, 6: 54, 8.37: 67,  # E al aire, secciones chicas: 1.5 mm^2 con 19 A hasta 8.37 mm^2 con 67 A
        10: 75, 13.3: 89, 16: 100, 21.1: 114, 25: 127, 26.7: 133, 33.6: 154, 35: 158,  # E, secciones medias: 10 mm^2 con 75 A hasta 35 mm^2 con 158 A
        42.4: 178, 50: 192  # E, las más gruesas: 42.4 mm^2 con 178 A y 50 mm^2 con 192 A
    }
}

# Tabla de ampacidad (en A) para el conductor de la ACOMETIDA (empalme entre
# la red y el medidor), según método de instalación (E=aéreo, D1=ducto enterrado)
# y sección del conductor (mm^2).
TABLA_AMPACIDAD_ACOM = {
    "E": {   # Método E (aéreo) — 70°C, Tabla 4.4 RIC 4
        1.5: 19,  # 1.5 mm^2 al aire aguanta 19 A
        2.08: 22,  # 2.08 mm^2 (AWG 14), 22 A
        2.5: 24,  # 2.5 mm^2, 24 A
        3.31: 30,  # 3.31 mm^2 (AWG 12), 30 A
        4: 31,  # 4 mm^2, 31 A
        5.26: 38,  # 5.26 mm^2 (AWG 10), 38 A
        6: 43,  # 6 mm^2, 43 A
        8.37: 53,  # 8.37 mm^2 (AWG 8), 53 A
        10: 60,  # 10 mm^2, 60 A
        13.3: 71,  # 13.3 mm^2 (AWG 6), 71 A
        16: 80,  # 16 mm^2, 80 A
        21.1: 91,  # 21.1 mm^2 (AWG 4), 91 A
        25: 101,  # 25 mm^2, 101 A
        26.7: 106,  # 26.7 mm^2 (AWG 3), 106 A
        33.6: 122,  # 33.6 mm^2 (AWG 2), 122 A
        35: 126,  # 35 mm^2, 126 A
        42.4: 142,  # 42.4 mm^2 (AWG 1), 142 A
        50: 159  # 50 mm^2, 159 A: el tope de la tabla
    },
    "D1": {   # Método D1 (ducto enterrado)
        1.5: 19,  # 1.5 mm^2 enterrado, 19 A
        2.08: 30,  # 2.08 mm^2, 30 A: el suelo enfria mejor que el aire
        2.5: 33,  # 2.5 mm^2, 33 A
        3.31: 38,  # 3.31 mm^2, 38 A
        4: 42,  # 4 mm^2, 42 A
        5.26: 48,  # 5.26 mm^2, 48 A
        6: 52,  # 6 mm^2, 52 A
        8.37: 63,  # 8.37 mm^2, 63 A
        10: 68,  # 10 mm^2, 68 A
        13.3: 80,  # 13.3 mm^2, 80 A
        16: 89,  # 16 mm^2, 89 A
        21.1: 103,  # 21.1 mm^2, 103 A
        25: 113,  # 25 mm^2, 113 A
        26.7: 117,  # 26.7 mm^2, 117 A
        33.6: 132,  # 33.6 mm^2, 132 A
        35: 136,  # 35 mm^2, 136 A
        42.4: 150,  # 42.4 mm^2, 150 A
        50: 159  # 50 mm^2, 159 A, el tope enterrado
    }
}

# Calcula el factor de corrección "ft" que hay que aplicar a la ampacidad de
# tabla según la temperatura ambiente y el método de instalación. Recibe la
# temperatura en °C y el método (B1/A1, D1/D2, E), y devuelve el factor ft.
def factor_temperatura_ft(temp_c, metodo):
    """Factor de corrección de ampacidad por temperatura — RIC N°4 Tabla N°4.7 / art. 6.2.6.
    Uso correcto: Ic = Iz × ft  (ft MULTIPLICA sobre la ampacidad de tabla, NO divide I_diseno)
    ft = 1.00 a 30°C · ft < 1.0 para T > 30°C · ft > 1.0 para T < 30°C
    El método E (aéreo, sol directo) usa ft_e < ft_b1 para T > 30°C.
    """
    t = float(temp_c)  # temperatura ambiente del proyecto, en °C
    # Tabla N°4.7 — tres columnas:
    # ft_b1: A1/B1 (embutido o en conducto)
    # ft_e:  E (aéreo expuesto al sol — más caliente que B1)
    # ft_d:  D1/D2 (subterráneo — temperatura del suelo)
    if t <= 10:
        ft_b1 = 1.22;  ft_e = 1.22;  ft_d = 1.07  # hasta 10 grados: 22% más en aire o ducto, 7% más enterrado
    elif t <= 15:  # entre 10 y 15 grados
        ft_b1 = 1.17;  ft_e = 1.17;  ft_d = 1.04  # un poco menos de premio, 17% arriba
    elif t <= 20:  # entre 15 y 20 grados
        ft_b1 = 1.12;  ft_e = 1.12;  ft_d = 1.00  # 12% arriba, y el enterrado ya queda en 1.00
    elif t <= 25:  # entre 20 y 25 grados
        ft_b1 = 1.06;  ft_e = 1.06;  ft_d = 0.96  # 6% arriba en aire, el suelo ya castiga un poco
    elif t <= 30:  # entre 25 y 30 grados
        ft_b1 = 1.00;  ft_e = 1.00;  ft_d = 0.93  # 30 grados es la temperatura base de la tabla, ft = 1.00
    elif t <= 35:  # entre 30 y 35 grados
        ft_b1 = 0.94;  ft_e = 0.90;  ft_d = 0.89  # sobre los 30 empieza a castigar, y el aéreo más por el sol
    elif t <= 40:  # entre 35 y 40 grados
        ft_b1 = 0.87;  ft_e = 0.82;  ft_d = 0.85  # a 40 grados el aéreo ya pierde casi un 20%
    elif t <= 45:  # entre 40 y 45 grados
        ft_b1 = 0.79;  ft_e = 0.74;  ft_d = 0.80  # ambiente caluroso, el cable aguanta bastante menos
    elif t <= 50:  # entre 45 y 50 grados
        ft_b1 = 0.71;  ft_e = 0.65;  ft_d = 0.76  # a 50 grados el aéreo queda en 0.65
    elif t <= 55:  # entre 50 y 55 grados
        ft_b1 = 0.61;  ft_e = 0.55;  ft_d = 0.71  # condiciones bien duras, el aéreo pierde casi la mitad
    else:  # sobre 55 grados
        ft_b1 = 0.50;  ft_e = 0.43;  ft_d = 0.65  # el peor caso de la tabla, ft bajisimo
    if metodo in ("D1", "D2"):  # metodos enterrados: el que manda es la temperatura del suelo
        return ft_d   # subterráneo: usa la columna de temperatura del suelo
    if metodo == "E":  # aéreo expuesto
        return ft_e   # aéreo: usa su propia columna, distinta de B1
    return ft_b1  # embutido o en conducto (A1/B1): columna por defecto

# ---------------------------------------------------------------
# Funciones para elegir la sección del ALIMENTADOR (cable entre el empalme
# y el tablero de la casa) y de la ACOMETIDA (tramo entre el transformador
# de la red y el empalme): método de instalación, sección
# mínima, caída de tensión y diámetro del ducto necesario.
# ---------------------------------------------------------------
def metodo_por_tipo_alim(tipo_alimentador: str) -> str:
    # según si el alimentador es aéreo, subterráneo o por ducto, devuelve
    # el "método de instalación" del RIC (E, D1 o B1) para usar la tabla
    # de corriente correcta.
    t = str(tipo_alimentador).strip().lower()  # normaliza el texto para poder comparar (minúsculas, sin espacios)
    if "aer" in t:  # contiene "aer": es un alimentador aéreo
        return "E"  # método E: instalación al aire libre
    if "sub" in t:  # contiene "sub": es un alimentador subterráneo
        return "D1"  # método D1: cable en ducto enterrado
    if "duc" in t:  # contiene "duc": va dentro de un ducto o conduit
        return "B1"  # método B1: embutido o en conducto
    return "B1"  # default
def siguiente_seccion_alim(smin: float) -> float:  # sube a la sección comercial más cercana
    # a partir de la sección mínima que dio el cálculo (smin), busca la
    # sección comercial más chica que alcance a cubrirla.
    for s in SECCIONES_MM2_ALIM:  # recorre las secciones comerciales de menor a mayor
        if s >= smin:  # la primera que sea igual o mayor a smin sirve
            return s  # esta ya alcanza a cubrir la sección mínima
    return SECCIONES_MM2_ALIM[-1]  # ninguna alcanzó, se usa la más grande de la lista
def dv_volts_alim(L_m: float, I_A: float, fp: float, rho: float, S_mm2: float) -> float:  # calcula la caída de tensión en volts de un tramo
    # fórmula típica de caída de tensión monofásica: ΔV = 2·L·I·fp·ρ / S
    # qué recibe:
    #   L_m .... largo del tramo, en metros
    #   I_A .... corriente que va a pasar, en Amperes
    #   fp ..... factor de potencia (0 a 1)
    #   rho .... resistividad del cobre (0.0179)
    #   S_mm2 .. sección del conductor, en mm^2
    # devuelve la caída de tensión en VOLTS (no en %).
    return (2 * L_m * I_A * fp * rho) / S_mm2

# Elige la sección del alimentador: parte de la corriente de diseño y el método
# de instalación, corrige la ampacidad por temperatura y después revisa que la
# caída de tensión quede dentro de lo permitido.
def seleccionar_alimentador(L_m: float, I_demanda_A: float, I_empalme_A: float, fp: float, V_nom: float,
                           tipo_alimentador: str,  # aéreo, subterráneo o en ducto
                           dv_max_volts: float = None,  # caída de tensión máxima permitida para el alimentador, en volts
                           dv_circuitos: list = None,  # caídas de los circuitos, para descontarlas del total
                           temp_override: float = None,  # temperatura ambiente forzada a mano, si el usuario la puso
                           ):
    # elige la sección del alimentador (y su ducto) probando calibres desde
    # el mínimo hacia arriba, hasta que cumpla: soporte la corriente (Iz),
    # la caída de tensión no pase el 3% y sumado a los circuitos de abajo
    # no pase el 5% total (límites del RIC). Devuelve la sección elegida
    # junto con todos los datos del cálculo.
    # qué recibe:
    #   L_m .............. largo del alimentador, en metros
    #   I_demanda_A ...... corriente de la casa ya con factor de demanda
    #   I_empalme_A ...... corriente del interruptor del empalme
    #   fp ............... factor de potencia
    #   V_nom ............ tensión nominal (220V)
    #   tipo_alimentador . aéreo, subterráneo o por ducto
    #   dv_max_volts ..... tope de caída de tensión en volts (si no se pasa, usa el 3%)
    #   dv_circuitos ..... lista con la caída de tensión de cada circuito de la casa
    #   temp_override .... temperatura a usar si no es la ambiente (ej: la del suelo)
    # devuelve un diccionario con: método, Smin, S (la sección elegida),
    # Iz_base, ft, Iz, dV_V, dV_pct, dV_total y, si no se pudo cumplir todo,
    # una "Advertencia".
    # Paso 1: Smin
    rho = 0.0179  # resistividad del cobre
    if dv_max_volts is None:  # si no le pasaron un tope, se usa el 3% del RIC
        dv_max_volts = V_nom * 0.03  # tope de caída de tensión: 3% de la tensión nominal
    smin = (2 * L_m * I_demanda_A * fp * rho) / dv_max_volts  # sección mínima teórica
    # Paso 2: método por tipo
    metodo = metodo_por_tipo_alim(tipo_alimentador)  # E, D1 o B1 según sea aéreo/subterráneo/ducto
    tabla = TABLA_AMPACIDAD_ALIM[metodo]  # tabla de corriente admisible de ese método
    # piso de 4 mm^2 para el alimentador (línea de abajo): aunque la caída de tensión permita
    # menos, no se baja de ahí. La sección comercial >= smin se elige recién en
    # la línea siguiente, con siguiente_seccion_alim()
    smin = max(smin, 4.0)
    S = siguiente_seccion_alim(smin)  # primera sección comercial candidata
    # Temperatura efectiva: usa temp_override si se proporcionó (ej. T° suelo para subterráneo)
    temp_calc = temp_override if temp_override is not None else temperatura
    # Paso 3 y 4: subir hasta cumplir Iz>=I, ΔV%<=3 y ΔV_total<=5% (RIC)
    _dv_circ_validos = [v for v in (dv_circuitos or []) if v is not None and not (v != v)]  # saca los None/NaN
    _dv_max_circ = max(_dv_circ_validos) if _dv_circ_validos else 0.0  # peor caída de tensión de los circuitos de abajo
    while True:  # sube de calibre y prueba de nuevo hasta que cumpla todo, o hasta llegar al más grande disponible
        Iz_base = tabla.get(S, 0)  # corriente admisible de esta sección (sin corregir)
        ft = factor_temperatura_ft(temp_calc, metodo)  # factor de corrección por temperatura
        Iz = Iz_base * ft  # corriente admisible real
        dv_v = dv_volts_alim(L_m, I_demanda_A, fp, rho, S)  # caída de tensión con esta sección
        dv_pct = (dv_v / V_nom) * 100.0  # caída de tensión en %
        cumple_ampacidad = (Iz >= I_demanda_A) and (Iz >= I_empalme_A)  # aguanta la corriente de demanda y de empalme
        cumple_caida = (dv_pct <= 3.0)  # no supera el 3% solo del alimentador
        cumple_total = (dv_pct + _dv_max_circ) <= 5.0  # sumado a los circuitos, no supera el 5% total
        if cumple_ampacidad and cumple_caida and cumple_total:  # si esta sección pasa las tres pruebas, no hay que seguir subiendo de calibre
            # esta sección cumple todo, se queda con ella
            return {"metodo": metodo, "Smin": smin, "S": S, "Iz_base": Iz_base, "ft": ft, "Iz": Iz,
                    "I_demanda_A": I_demanda_A, "I_empalme_A": I_empalme_A, "dV_V": dv_v, "dV_pct": dv_pct,  # corrientes usadas y la caída de tensión que dio
                    "dV_max_circ": _dv_max_circ, "dV_total": round(dv_pct + _dv_max_circ, 2)}  # la peor caída de los circuitos de abajo y el total sumado

        idx = SECCIONES_MM2_ALIM.index(S)  # posición de la sección actual dentro de la lista de calibres comerciales
        if idx >= len(SECCIONES_MM2_ALIM) - 1:  # ya se llegó al calibre más grande disponible y no cumplió
            # ya no hay secciones más grandes, avisa que no se pudo cumplir todo
            return {"metodo": metodo, "Smin": smin, "S": S, "Iz_base": Iz_base, "ft": ft, "Iz": Iz,
                    "I_demanda_A": I_demanda_A, "I_empalme_A": I_empalme_A, "dV_V": dv_v, "dV_pct": dv_pct,  # corrientes usadas y la caída de tensión que dio
                    "dV_max_circ": _dv_max_circ, "dV_total": round(dv_pct + _dv_max_circ, 2),  # peor caída de los circuitos y el total sumado
                    "Advertencia": "No cumple Iz > I_demanda y/o Iz >= I_empalme y/o ΔV% <= 3 y/o ΔV_total <= 5% con secciones disponibles"}  # aviso de que ninguna sección disponible alcanzo a cumplir las tres condiciones
        S = SECCIONES_MM2_ALIM[idx + 1]  # sube al siguiente calibre y prueba de nuevo

def seleccionar_acometida(L_m: float, I_empalme_A: float, fp: float, V_nom: float,  # dimensiona la acometida: largo del tramo, corriente del empalme, fp y tensión
                          tipo_acometida: str,  # aérea o subterránea, de eso depende el método de instalación
                          dv_max_volts: float = None,  # tope de caída de tensión en volts, opcional
                          temp_override: float = None):  # temperatura distinta a la ambiente, opcional (ej: la del suelo)
    # lo mismo que seleccionar_alimentador pero para la acometida (el tramo
    # entre el transformador de la red y el empalme): calcula la sección que aguante la
    # corriente y no pase la caída de tensión máxima.
    # qué recibe:
    #   L_m ............ largo de la acometida, en metros
    #   I_empalme_A .... corriente del interruptor del empalme
    #   fp ............. factor de potencia
    #   V_nom .......... tensión nominal (220V)
    #   tipo_acometida . "aéreo" o "subterráneo" (si es otra cosa, tira error)
    #   dv_max_volts ... tope de caída de tensión en volts (si no se pasa, usa el 3%)
    #   temp_override .. temperatura a usar si no es la ambiente
    # devuelve un diccionario con: método, Smin, S (la sección elegida),
    # Iz_base, ft, Iz, dV_V, dV_pct y, si no se pudo cumplir, un "warning".
    rho = 0.0179  # resistividad del cobre
    if dv_max_volts is None:  # si no le pasaron un tope, se usa el 3% del RIC
        dv_max_volts = V_nom * 0.03  # por defecto, tope de caída de tensión = 3% de la tensión nominal
    smin = (2 * L_m * I_empalme_A * fp * rho) / dv_max_volts  # sección mínima teórica según caída de tensión
    # Elegir método según tipo de acometida
    if tipo_acometida.lower() == "aereo":  # el texto indica acometida aérea
        metodo = "E"  # acometida al aire libre, método E
    elif tipo_acometida.lower() == "subterraneo":  # el texto indica acometida subterránea
        metodo = "D1"  # acometida enterrada en ducto, método D1
    else:  # cualquier otro texto no es un tipo de acometida conocido
        raise ValueError(f"Tipo de acometida no válido: {tipo_acometida}")  # corta el cálculo, sin saber como va la acometida no se puede elegir el método
    tabla = TABLA_AMPACIDAD_ACOM[metodo]  # tabla de corriente admisible según el método
    # mínimo 4mm^2 por consistencia con alimentador
    smin = max(smin, 4.0)
    S = siguiente_seccion_alim(smin)  # primera sección comercial candidata
    temp_calc = temp_override if temp_override is not None else temperatura  # temperatura para corregir la ampacidad: la del suelo si viene, si no la ambiente
    while True:  # sube de calibre hasta que aguante la corriente y no pase el 3% de caída de tensión
        Iz_base = tabla.get(S, 0)  # corriente admisible de esta sección (sin corregir)
        ft = factor_temperatura_ft(temp_calc, metodo)  # factor de corrección por temperatura
        Iz = Iz_base * ft  # corriente admisible real
        dv_v = dv_volts_alim(L_m, I_empalme_A, fp, rho, S)  # caída de tensión con esta sección
        dv_pct = (dv_v / V_nom) * 100.0  # caída de tensión en %
        if (Iz >= I_empalme_A) and (dv_pct <= 3.0):  # aguanta la corriente del empalme y no se pasa del 3% de caída
            # esta sección cumple ambos requisitos, se queda con ella
            return {"metodo": metodo,"Smin": smin,"S": S,"Iz_base": Iz_base,"ft": ft,"Iz": Iz,"dV_V": dv_v,
                    "dV_pct": dv_pct}  # caída de tensión en % de la sección elegida
        idx = SECCIONES_MM2_ALIM.index(S)  # posición de la sección actual dentro de la lista de calibres comerciales
        if idx >= len(SECCIONES_MM2_ALIM) - 1:  # ya se llegó al calibre más grande disponible y no cumplió
            # ya no hay secciones más grandes disponibles, avisa que no se pudo cumplir
            return {"metodo": metodo,"Smin": smin,"S": S,"Iz_base": Iz_base,"ft": ft,"Iz": Iz,"dV_V": dv_v,
                    "dV_pct": dv_pct,"warning": "No cumple Iz y/o ΔV% con secciones disponibles"}  # aviso de que ninguna sección disponible cumplió
        S = SECCIONES_MM2_ALIM[idx + 1]  # sube al siguiente calibre y prueba de nuevo

# =========================================================
# CANALIZACIÓN ALIMENTADOR - Tablas N°4.19 (mm^2) y N°4.20 (AWG)
# =========================================================
# Tabla N°4.19 (mm^2)
# clave principal: sección del conductor en mm^2
# clave interna (1 a 5): cantidad de conductores que van juntos dentro del ducto
# valor: diámetro nominal comercial del ducto, en mm
DUCTO_N419_MM2 = {
    1.5:  {1:16, 2:16, 3:16, 4:20, 5:25},  # 1.5mm^2. ojo: el resultado final pasa por max(32, d), así que los diámetros bajo 32mm de esta tabla nunca salen
    2.5:  {1:16, 2:20, 3:20, 4:32, 5:32},  # 2.5mm^2: 16mm para uno solo, 32mm cuando van 4 o 5
    4.0:  {1:16, 2:25, 3:25, 4:32, 5:40},  # 4mm^2: de 16mm a 40mm según cuántos conductores entren
    6.0:  {1:16, 2:25, 3:32, 4:32, 5:40},  # 6mm^2: igual que el de 4mm^2 pero ya pide 32mm con 3 conductores
    10.0: {1:20, 2:32, 3:32, 4:40, 5:50},  # 10mm^2: mínimo 20mm, y 50mm si van 5 conductores
    16.0: {1:25, 2:32, 3:40, 4:50, 5:50},  # 16mm^2: parte en 25mm, con 4 o 5 conductores se va a 50mm
    25.0: {1:25, 2:40, 3:50, 4:50, 5:63},  # 25mm^2: de 25mm a 63mm según la cantidad de conductores
    35.0: {1:32, 2:40, 3:50, 4:63, 5:63},  # 35mm^2: parte en 32mm y llega a 63mm
    50.0: {1:32, 2:50, 3:63, 4:63, 5:75},  # 50mm^2: de 32mm a 75mm
    70.0: {1:40, 2:50, 3:63, 4:75, 5:75},  # 70mm^2: parte en 40mm, con 4 o 5 conductores pide 75mm
    95.0: {1:40, 2:63, 3:75, 4:100, 5:100},  # 95mm^2: con 4 o 5 conductores ya se va a 100mm
    120.0:{1:50, 2:63, 3:75, 4:100, 5:100},  # 120mm^2: de 50mm a 100mm
    150.0:{1:50, 2:75, 3:100,4:100, 5:125},  # 150mm^2: llega a 125mm con 5 conductores
    185.0:{1:63, 2:75, 3:100,4:125, 5:125},  # 185mm^2: parte en 63mm y termina en 125mm
    240.0:{1:63, 2:100,3:125,4:125, 5:150},  # 240mm^2 es la sección más grande que cubre la tabla 4.19
}

# Tabla N°4.29 - Cables para uso en tuberías de canalizaciones subterráneas
# Solo hasta 25mm^2 y máximo 3 conductores (uso acometida y alimentador subterráneo)
# clave principal: sección del conductor en mm^2
# clave interna (1 a 3): cantidad de conductores dentro del ducto
# valor: diámetro nominal comercial del ducto, en mm
DUCTO_N429_SUBTERRANEO = {
    1.5:  {1: 25, 2: 25, 3: 25},  # 1.5mm^2: 25mm dan lo mismo 1, 2 o 3 conductores
    2.5:  {1: 25, 2: 25, 3: 32},  # 2.5mm^2: 25mm hasta 2 conductores, 32mm con 3
    4.0:  {1: 25, 2: 32, 3: 40},  # 4mm^2: de 25mm a 40mm según cuántos van
    6.0:  {1: 32, 2: 32, 3: 50},  # 6mm^2: 32mm, y 50mm si van los 3 conductores
    10.0: {1: 40, 2: 50, 3: 63},  # 10mm^2: parte en 40mm y llega a 63mm
    16.0: {1: 50, 2: 50, 3: 63},  # 16mm^2: 50mm para 1 o 2, 63mm con 3
    25.0: {1: 63, 2: 63, 3: 75},  # 25mm^2 es el tope de la tabla 4.29, más arriba no hay dato
}
# Tabla N°4.20 (AWG/kcmil) - lo que aparece en tu imagen
# clave principal: calibre AWG/kcmil del conductor
# clave interna (1 a 5): cantidad de conductores dentro del ducto
# valor: diámetro nominal comercial del ducto, en mm
DUCTO_N420_AWG = {
    "14":  {1:16, 2:16, 3:20, 4:25, 5:32},  # AWG 14, el más fino de la tabla: de 16mm a 32mm
    "12":  {1:16, 2:20, 3:25, 4:32, 5:32},  # AWG 12: 16mm para uno solo, 32mm con 4 o 5
    "10":  {1:16, 2:25, 3:32, 4:32, 5:40},  # AWG 10: parte en 16mm y llega a 40mm
    "8":   {1:20, 2:25, 3:32, 4:40, 5:40},  # AWG 8: mínimo 20mm, hasta 40mm
    "6":   {1:20, 2:32, 3:40, 4:40, 5:50},  # AWG 6: de 20mm a 50mm según cuántos conductores
    "4":   {1:25, 2:32, 3:40, 4:50, 5:50},  # AWG 4: parte en 25mm y llega a 50mm
    "2":   {1:32, 2:40, 3:50, 4:63, 5:63},  # AWG 2: de 32mm a 63mm
    "1":   {1:32, 2:40, 3:50, 4:63, 5:63},  # AWG 1: mismos ductos que el AWG 2
    "1/0": {1:32, 2:50, 3:63, 4:63, 5:75},  # AWG 1/0: parte en 32mm y llega a 75mm
    "2/0": {1:40, 2:50, 3:63, 4:75, 5:75},  # AWG 2/0: de 40mm a 75mm
    "3/0": {1:40, 2:63, 3:75, 4:75, 5:100},  # AWG 3/0: con 5 conductores ya pide 100mm
    "4/0": {1:50, 2:63, 3:75, 4:100, 5:100},  # AWG 4/0: de 50mm a 100mm
    "250": {1:50, 2:75, 3:100, 4:100, 5:125},  # 250 kcmil: parte en 50mm y llega a 125mm
    "300": {1:50, 2:75, 3:100, 4:100, 5:125},  # 300 kcmil: mismos ductos que el de 250
    "350": {1:63, 2:75, 3:100, 4:125, 5:125},  # 350 kcmil: de 63mm a 125mm
    "400": {1:63, 2:100, 3:100, 4:125, 5:125},  # 400 kcmil: parte en 63mm y llega a 125mm
    "500": {1:63, 2:100, 3:125, 4:125, 5:150},  # 500 kcmil: con 5 conductores pide 150mm
    "600": {1:75, 2:100, 3:125, 4:150, 5:150},  # 600 kcmil: parte en 75mm y llega a 150mm
    "750": {1:75, 2:125, 3:150, 4:150, 5:175},  # 750 kcmil: de 75mm a 175mm
    "1000":{1:100, 2:125, 3:150, 4:175, 5:200},  # calibre más grande que cubre la tabla 4.20
}
# Conversión de mm^2 a AWG (tabla de sección comercial)
# clave: sección en mm^2, valor: calibre AWG comercial equivalente más cercano
MM2_A_AWG = {
    2.08: "14",  # 2.08mm^2 es lo mismo que el AWG 14
    3.31: "12",  # 3.31mm^2 equivale al AWG 12
    5.26: "10",   # 5.26mm^2 equivale al AWG 10
    8.37: "8",  # 8.37mm^2 equivale al AWG 8
    13.3: "6",  # 13.3mm^2 equivale al AWG 6
    21.1: "4",  # 21.1mm^2 equivale al AWG 4
    33.6: "2",  # 33.6mm^2 equivale al AWG 2
    42.4: "1",  # 42.4mm^2 equivale al AWG 1
}
def ducto_nominal_tablas(S_mm2: float, n_cond: int, tipo_alimentador: str):  # entrega el diámetro de ducto que manda la tabla según sección, cantidad de conductores y tipo de alimentador
    """
    Devuelve Ø ducto nominal (mm) usando:
    - aereo          => None (sin canalización)
    - subterraneo    => Tabla N°4.29 (hasta 25mm^2, máx 3 conductores)
    - embutido/ducto => Tabla N°4.19 mm^2 o N°4.20 AWG si no existe en N°4.19
    """
    t = str(tipo_alimentador).strip().lower()  # normaliza el texto para poder comparar
    if "aer" in t:  # alimentador aéreo
        return None  # aéreo no lleva ducto
    n = max(1, min(5, int(n_cond)))  # la tabla solo cubre entre 1 y 5 conductores
    S_key = round(float(S_mm2), 2)  # redondea para poder buscar coincidencia exacta en las tablas (evita errores de punto flotante)

    # Si es subterráneo, se usa la Tabla N°4.29
    if "sub" in t:
        n_sub = max(1, min(3, n))   # tabla solo tiene hasta 3 conductores
        S_val = float(S_mm2)  # copia numerica de la sección, para comparar contra las claves de la tabla
        # buscar clave más cercana por arriba en la tabla
        claves = sorted(DUCTO_N429_SUBTERRANEO.keys())
        clave = None  # todavía no se encontró un umbral que alcance a cubrir la sección
        for c in claves:  # va probando los umbrales de menor a mayor
            if S_val <= c + 0.01:  # el 0.01 es tolerancia por los decimales de la sección
                clave = c  # primer umbral que alcanza a cubrir la sección
                break  # ya encontró el umbral que sirve, no sigue buscando
        if clave is None:  # la sección es más grande que todo lo que cubre la tabla
            clave = claves[-1]   # usar la mayor si supera 25mm^2
        return DUCTO_N429_SUBTERRANEO[clave][n_sub]  # diámetro de ducto según la clave encontrada y la cantidad de conductores

    # Si es embutido o va por ducto, se usa la Tabla N°4.19
    if S_key in DUCTO_N419_MM2:  # la sección esta tal cual en la tabla de mm^2
        d = DUCTO_N419_MM2[S_key][n]  # la sección está directo en la tabla
    else:  # la sección no aparece en mm^2, hay que buscarla por AWG
        # Intentar por AWG equivalente (Tabla N°4.20)
        awg = MM2_A_AWG.get(S_key)  # busca el calibre AWG equivalente a esta sección
        if awg is None:  # esa sección no tiene AWG equivalente en la tabla
            return None  # no se encontró equivalencia, no se puede calcular
        if awg not in DUCTO_N420_AWG:  # el AWG existe pero la tabla de ductos no lo trae
            return None  # el calibre AWG no está en la tabla de ductos, no se puede calcular
        d = DUCTO_N420_AWG[awg][n]  # diámetro de ducto según el AWG equivalente y la cantidad de conductores
    return max(32, d)  # nunca menor a 32mm (mínimo práctico)

# =========================
# HOJA 2: LISTADO DE MATERIALES (PRO + MARCAS + CONTEO REAL)
# =========================
def build_materiales_df(circuitos_df, texto_omni, ambientes_df, group_info, tipo_canalizacion,tipo_alimentador,  # arma toda la hoja de materiales a partir de los circuitos, los ambientes y el tipo de canalización
                        tipo_acometida, tipo_instalacion_empalme, requiere_mastil, acometida_txt, interruptor_texto,  # datos del empalme y la acometida, si lleva mástil y los textos del interruptor
                        longitud_transformador_empalme, canalizacion_txt, dist_empalme_pt1, dist_tda_pt2,  # metros de acometida, texto del ducto del alimentador y distancias a las dos puestas a tierra
                        longitud_subterraneo_medidor, longitud_mastil, altura_acometida_aerea, longitud_subterraneo_medidor2,  # largos de los tramos subterráneos, del mástil y de la subida de la acometida aérea
                        longitud_llegada_aerea_tda, longitud_alimentador, longitud_abrazaderas_alimentador,  # largo de la llegada aérea al tablero, del alimentador y del tramo con abrazaderas
                        longitud_poste_alimentador_aereo, dist_vertical_acometida, alim_txt,  # subida del alimentador por el poste, distancia vertical de la acometida y texto del alimentador
                        circuitos_climatizacion=None, items_por_nombre=None, cajas_adic_por_nombre=None,  # equipos de clima, los items de cada circuito y las cajas extra que pidió el usuario
                        tipo_dif_por_circ=None,  # dice si cada circuito va colgado del diferencial general o lleva uno propio
                        n_barras_pt1=1, n_barras_pt2=1,  # cuántas barras copperweld lleva cada puesta a tierra
                        long_cond_desnudo_pt1=0, long_cond_desnudo_pt2=0):  # metros de conductor desnudo de cada puesta a tierra
    """
    Esta es la función grande: arma toda la hoja "Materiales" del Excel
    a partir de los circuitos que ya se definieron antes.

    Recibe los circuitos (circuitos_df), los ambientes (ambientes_df),
    los datos del empalme/acometida/alimentador, la climatización, y
    algunos datos auxiliares más (los ítems de cada circuito, las cajas
    adicionales, cómo quedaron agrupados los diferenciales, etc.)

    Qué es cada dato que recibe, en el mismo orden en que van arriba:
      circuitos_df ...................... tabla con todos los circuitos ya
                                          armados (nombre, longitud, potencia,
                                          TM, conductor, etc.)
      texto_omni ........................ texto del interruptor general
                                          omnipolar, ej: "2x40A / 10kA / Curva C"
      ambientes_df ...................... tabla de ambientes (nombre, área,
                                          enchufes y luminarias de cada uno)
      group_info ........................ qué circuitos comparten cada
                                          diferencial y con qué calibre
      tipo_canalizacion ................. "embutida" o "sobrepuesta"
      tipo_alimentador .................. cómo va el alimentador: aéreo,
                                          subterráneo o por ducto
      tipo_acometida .................... cómo llega la acometida: aérea o
                                          subterránea
      tipo_instalacion_empalme .......... "fachada" o "independiente"
      requiere_mastil ................... "si"/"no", solo cuando el empalme va
                                          en fachada
      acometida_txt ..................... texto del cable de acometida, ej:
                                          "Concéntrico Cu 2x 10 mm^2"
      interruptor_texto ................. texto del interruptor del empalme,
                                          ej: "1x25A / 6kA / Curva D"
      longitud_transformador_empalme .... metros de acometida (transformador
                                          hasta el empalme)
      canalizacion_txt .................. ducto del alimentador, ej: "Ø PVC
                                          Conduit 25 mm" ("-" si es aéreo)
      dist_empalme_pt1 .................. metros del empalme a la camarilla
                                          N°1 (puesta a tierra 1)
      dist_tda_pt2 ...................... metros del tablero a la camarilla
                                          N°2 (puesta a tierra 2)
      longitud_subterraneo_medidor ...... tramo subterráneo hasta el medidor
                                          (caso fachada sin mástil)
      longitud_mastil ................... largo del mástil, si es que lleva
      altura_acometida_aerea ............ subida de la acometida aérea por el
                                          poste
      longitud_subterraneo_medidor2 ..... tramo subterráneo cuando el empalme
                                          es independiente
      longitud_llegada_aerea_tda ........ tramo desde la llegada del
                                          alimentador aéreo hasta el tablero
      longitud_alimentador .............. metros de alimentador (empalme hasta
                                          el tablero)
      longitud_abrazaderas_alimentador .. tramo subterráneo desde la salida del
                                          alimentador hasta el tablero
      longitud_poste_alimentador_aereo .. subida del alimentador aéreo por el
                                          poste
      dist_vertical_acometida ........... distancia vertical entre la caja de
                                          empalme y la llegada de la acometida
      alim_txt .......................... texto del alimentador, ej:
                                          "RV-K Cu 3x10 mm^2"
      circuitos_climatizacion ........... lista de equipos de clima ya
                                          calculados (puede venir vacía)
      items_por_nombre .................. para cada circuito, la lista de sus
                                          enchufes / luminarias / equipos
      cajas_adic_por_nombre ............. cajas de derivación extra que pidió
                                          el usuario, por circuito
      tipo_dif_por_circ ................. si cada circuito usa el diferencial
                                          "general" o uno propio
      n_barras_pt1 / n_barras_pt2 ....... cuántas barras copperweld lleva cada
                                          puesta a tierra
      long_cond_desnudo_pt1 / _pt2 ...... metros de conductor desnudo de cada
                                          puesta a tierra

    Va calculando todo circuito por circuito, en este orden (cada bloque
    está marcado en el código con "# ===== NOMBRE SECCIÓN ====="):
      1. Canalización (conduit o canaleta), abrazaderas y salidas de caja.
      2. Conductores de cada circuito y sus chicotes.
      3. Cajas de derivación, uniones, cajas para enchufes/interruptores/
         luminarias.
      4. Protecciones: TM y diferencial de cada circuito y del tablero.
      5. Ferrules, borneras y cableado interior del tablero.
      6. Materiales de empalme y de acometida/alimentador.
      7. Puesta a tierra (barra copperweld, camarilla, conductor desnudo).
      8. Tornillos, tarugos, prensaestopas.
      9. Ojo: no hay un paso final que arme las filas. Cada sección va
         llamando a add_row() a medida que calcula lo suyo.

    Al final devuelve un DataFrame de pandas: una fila por material, listo
    para escribirse en la hoja "Materiales".
    """
    filas = []  # aquí se van juntando las filas (una por material) antes de armar la tabla final
    item = 1  # número correlativo de item, para la columna del listado
    # =========================
    # ACUMULADORES TORNILLERÍA
    # =========================
    _abrazaderas_total = 0       # embutida: total de abrazaderas de conduit sumando todos los circuitos; más abajo sale de acá la tornillería (2 tornillos por abrazadera)
    _clima_items = []             # un dict por circuito de clima: {"circ","L","canal","in_tm","es_emb","mm","con_enchufe","cajas_paso"}. Se llena antes que todo, y después lo releen varias secciones (salidas de caja, cajas, tapas, conductores)
    _long_sobrepuesta_total = 0  # sobrepuesta: metros de canaleta realmente comprada (tramos de 2m); después sale 1 tornillo por metro
    _cajas_total = 0             # total de cajas de derivación de todos los circuitos (las de paso también se suman acá); de este número salen espuma, tornillos y tarugos
    _tapas_ciegas_total = 0      # total de tapas ciegas; se calcula recién en el bloque TAPAS CIEGAS y después da 2 tornillos por tapa
    _cajas_paso_total = 0        # total de cajas de paso intermedias: 1 cada 20m de tramo embutido (RIC N°4, art. 7.16.1.13), y sólo si el usuario confirmó que hay un tramo continuo de 20m o más (_tiene_tramo_20m)
    _mm_conduit_ilumin = 20      # diámetro conduit circuito iluminación (default 20mm)
    _riel_m = 0                  # largo de la tira de riel DIN, en metros: 1, o 2 si el tablero es de 56 o 72 puestos. Después define cuántos tornillos lleva el riel (7 o 14)
    _puestos_tablero = None      # tamaño del tablero comercial elegido, en puestos (2, 4, 6, 8, 12, 16, 18, 24, 36, 42, 48, 54, 56, 72). Más abajo se usa para sacar el alto y ancho del gabinete de TAB_DIMS


    # Diccionario grande: para cada tipo de material (la clave) da un
    # texto con 2 o 3 marcas sugeridas (el valor). Se usa más adelante
    # para llenar la columna "Marcas sugeridas" del listado final.
    # Es una tabla larga y repetitiva, por eso no se comenta marca por
    # marca, solo se explica acá el sentido general de la tabla.
    # ====== MARCAS SUGERIDAS ======
    marcas = {
        # protecciones y cables, lo basico de la instalación
        "Protecciones": "Legrand, Schneider Electric, Bticino",  # automáticos y diferenciales del tablero
        "Conductores": "Madeco, Cosesa, Revi",  # los cables de los circuitos
        "Conectores cónicos": "Globaltronic, Lexo, Mec",  # para empalmar conductores dentro de las cajas
        "Ferrule": "Lexo, Soliot, Genérico",  # terminal que se pone en la punta del cable antes de apretarlo
        # canalización: canaleta a la vista y conduit embutido
        "Canaleta PVC (sobrepuesta)": "Grantt, Schneider Electric, Hoffens",  # canalización a la vista
        "Accesorios canaleta PVC": "Grantt, Schneider Electric, Hoffens",  # codos, uniones y tapas de la canaleta
        "Conduit PVC (embutida)": "Hoffens, Revi, Halux",  # tuberia que va dentro del muro
        "Boquilla bordes redondos": "Halux, Hoffens, Tigre",  # protege el cable donde sale del conduit
        "Prensaestopas": "Lexo, Genérico, Soliot",  # entrada sellada de cables al tablero
        "Abrazaderas": "Hoffens, Tigre, Halux",  # fijan el conduit al muro
        # el tablero y lo que va montado adentro
        "Tablero embutido": "Lexo, Stanford, Tibox",  # gabinete que va dentro del muro
        "Tablero sobrepuesto": "Lexo, Stanford, Tibox",  # gabinete montado a la vista
        "Barra repartidora": "Lexo, DTK, Cabur",  # reparte la fase a cada automático
        "Riel DIN": "Lexo, Mec, Stanford",  # el riel donde se montan las protecciones
        # lo que queda a la vista en los ambientes
        "Enchufes": "Bticino, MEC, Schneider Electric",  # los tomacorrientes de cada ambiente
        "Interruptores": "MEC, Schneider Electric, Bticino",  # los de pared, para la iluminación
        "Iluminarias": "Dairu, Eglo, AVC",  # las luminarias de cada ambiente
        "Portalamparas": "Bticino, Schneider Electric, Genérico",  # el casquillo donde va la ampolleta
        # cajas de derivación y sus tapas
        "Cajas derivación embutidas": "Schneider Electric, Bticino, Hoffens",  # cajas que quedan dentro del muro
        "Cajas derivación sobrepuestas": "Schneider Electric, Bticino, Legrand",  # cajas montadas a la vista
        "Cajas de paso estancas": "Gewiss, Scame, Lexo",  # cajas de paso para exterior o zonas húmedas
        "Tapa ciega": "Schneider Electric, Lexo, Bticino",  # tapa para caja cuadrada sin uso
        "Tapa ciega octogonal": "Schneider Electric, Epem, Bticino",  # la misma tapa pero para caja octogonal
        # (embutida accesorios)
        "Salida de caja conduit": "Halux, Hoffens, Tigre",  # pieza que remata el conduit al llegar a la caja
        "Abrazadera conduit": "Hoffens, Tigre, Halux",  # fija el conduit embutido
        # accesorios que van dentro del tablero
        "Barra unipolar verde": "Lexo, DTK, Cabur",  # barra de tierra del tablero
        "Supresor de transiente": "Schneider Electric, Legrand, ABB",  # protege contra sobretensiones por maniobra o rayo
        "Luz Piloto": "Schneider Electric, Lexo, Bticino",  # avisa que el tablero tiene tensión
        "Portafusible": "Lexo, Legrand, Bticino",  # la base donde va el fusible
        "Fusible": "Bussmann, Mersen, Schneider Electric",  # protección de respaldo
        # fijaciones y sellos
        "Tornillos": "Fixser; Mamut; Fischer",  # fijaciones en general
        "Espuma expansiva PU": "Sika Boom; Fischer; Ceresita",  # sella las pasadas de muro
        # materiales del empalme
        "Medidor empalme": "Lexo, Metertek, DRL",  # el medidor del empalme
        "Caja empalme": "Metertek, Saime, Stanford",  # el gabinete del empalme
        "Disyuntor empalme": "Saime, Legrand, Schneider Electric",  # el automático del empalme
        # materiales de la acometida
        "Cable acometida": "Madeco, Covisa, Elexor",  # cable entre la red y el empalme
        "Tubo conduit galvanizado acometida": "Lexo, Ekoline, Rhona",  # tubo metálico por donde sube la acometida
        "Cabeza de servicio": "Ekoline, Elexor, Rhona",  # remate del tubo para que no le entre agua
        "Cancamo abierto": "Mamut, Fixser, Rhona",  # donde se amarra la acometida aérea
        "Granpa de retención": "Elexor, Tecnored, Standford",  # tensa y sujeta el cable aéreo
        "Mordaza acometida": "Generico, Elexor, Rema",  # aprieta el cable en el punto de conexión
        "Conector HUB": "Bekam, Enelux, Ekoline",  # conector estanco de la acometida
        "Terminal PVC conduit con 2 tuercas": "Hoffens, Halux, Revi",  # remate del conduit con tuerca a cada lado
        # conductores de neutro y tierra, y la puesta a tierra
        "Conductor THWN-2 blanco": "Madeco, Covisa, Nexans",  # cable del neutro
        "Conductor THWN-2 verde": "Madeco, Covisa, Nexans",  # cable de la tierra de protección
        "Barra copperweld": "Ekoline, Generico, Indutref",  # barra que se entierra en la puesta a tierra
        "Prensaestopa": "Legrand, Bticino, Generico",  # entrada sellada de cable, en singular
        # ── AGUA CALIENTE ──────────────────────────────────────────────────────
        "Diferencial 10mA agua caliente": "Legrand, Schneider Electric, Bticino",  # diferencial más sensible, 10mA, para calefont o ducha eléctrica
        "Tablero externo agua caliente":  "Lexo, Stanford, Tibox",  # tablerito aparte solo para el agua caliente
        "TM bipolar agua caliente":       "Legrand, Schneider Electric, Bticino",  # automático de dos polos de ese circuito
        "Bornera PE agua caliente":       "Legrand, Phoenix Contact, Weidmuller",  # bornera de tierra de ese tablerito
        "Bornera de conexion":            "Lexo, DTK, Cabur",  # borneras normales de conexión
        # soldadura y materiales de puesta a tierra
        "Estaño": "Golden, Macrotel, Altronics",  # para soldar uniones
        "Pasta para soldar": "Metalfer, Indepp, MCT",  # ayuda a que la soldadura pegue
        "Conductor desnudo Cu": "Madeco, Covisa, Nexans",  # cobre desnudo de la puesta a tierra
        "Camarilla PVC naranjo": "Generico, DTK, Tigre",  # cámara de registro de la puesta a tierra
        "Abrazadera tipo caddy": "Lexo, Rhona, Ekoline",  # abrazadera metálica de fijación
        "Terminal compresion tipo ojo": "Generico, Soliot, Sofamel",  # terminal para atornillar el cable a la barra
        "Sellador de roscas": "Dura, Permatex, Topex",  # sella las uniones roscadas
        # ferrules de acometida y alimentador
        "Terminal ferrul acometida": "Generico, Soliot, Lexo",  # ferrul para la punta del cable de acometida
        "Terminal ferrul alimentador": "Generico, Soliot, Lexo",  # ferrul para la punta del alimentador
        "Terminal ferrul doble alimentador": "Generico, Soliot, Lexo",  # ferrul doble, cuando entran dos cables al mismo borne
        # tornillería de fijación
        "Tirafondo hexagonal madera": "Mamut, Fixser, Fischer",  # fijación a poste o estructura de madera
        "Tornillo autoperforante hexagonal": "Mamut, Fixser, Fischer",  # fijación a fierro, sin perforar antes
        # materiales del alimentador
        "Alimentador RV-K": "Nexans, Cocesa, Madeco",  # el cable del alimentador
        "Conduit PVC": "Hoffens, Halux, Revi",  # la tuberia por donde va el alimentador
        "Abrazadera conduit PVC alimentador": "Hoffens, Tigre, Halux",  # fija ese conduit al muro o al poste
        "Tornillo punta fina madera": "Fixser, Mamut, Fischer",  # tornillo para madera
        "Tornillo autoperforante broca": "Fixser, Mamut, Fischer",  # tornillo con broca, para plancha metálica
        # cajas, tarugos y tornillos varios
        "Caja derivacion metalica": "Ectray, Generico, Standford",  # caja de derivación de fierro
        "Tarugo paloma": "Mamut, Fixser, Fischer",  # tarugo para tabique hueco
        "Tornillo para tarugo paloma": "Mamut, Fixser, Fischer",  # el tornillo que va con ese tarugo
        "Portafusible de loza": "Lexo, Fujian, Kersting",  # portafusible antiguo, de loza
        # postes, pilares y protección contra sobretensión
        "Poste madera": "Imperial, Lifewood, Sodimac",  # poste donde llega la acometida aérea
        "Pilar metalico": "Genérico, ACMA, Kupfer",  # pilar donde se monta el empalme
        "Protector sobrevoltaje": "CNC, TOMZN, Sinotimer",  # corta si la tensión se va muy arriba o muy abajo
        # cintas y cámaras del tramo subterráneo
        "Cinta autofundente goma": "3M, Lexo, Truper",  # sella empalmes a la intemperie
        "Cinta aislante PVC": "3M, Lexo, Truper",  # aisla y termina el empalme
        "Camara tipo C": "Prefabricados del Sur, Hormipret, Genérico",  # cámara de hormigón para el paso subterráneo
        "Marco metalico camara C": "Genérico, ACMA, Prefabricados del Sur",  # marco y tapa de esa cámara
        "Boquilla camara tipo C": "Hoffens, Halux, Tigre",  # boquilla por donde entra el ducto a la cámara
    }

    # Arma el texto de la columna "Descripción técnica" para la
    # canalización de un circuito, a partir del texto tipo "Embutida
    # (PVC conduit 16mm)" que se armo antes. Si es sobrepuesta, agrega
    # "2mts" porque la canaleta se vende en tramos de a 2 metros.
    def formatear_canalizacion_material(canal_txt: str, es_sobrepuesta: bool) -> str:
        """
        Ejemplo: "Embutida (PVC conduit 16mm)" queda como "Canalización embutida PVC conduit 16mm"
        Ejemplo: "Sobrepuesta (Canaleta de PVC 20x10mm)" queda como "Canalización sobrepuesta Canaleta de PVC 20x10mm, 2mts"
        """
        t = (canal_txt or "").strip()  # t: el texto de canalización, sin espacios de sobra
        # si no llegó texto, no hay nada que formatear
        if not t:
            return "Canalización no definida"  # no vino texto
        t = " ".join(t.split())  # saca espacios dobles

        # si el texto trae paréntesis, ahí viene el detalle del material (tipo + medida)
        if "(" in t and ")" in t:
            tipo = t.split("(")[0].strip().lower()  # "embutida" o "sobrepuesta"
            dentro = t.split("(", 1)[1].split(")")[0].strip()  # lo que va entre paréntesis (ej: "PVC conduit 16mm")
            base = f"Canalización {tipo} {dentro}"  # junta tipo + detalle en el texto final
        else:  # no venía el detalle entre paréntesis
            base = f"Canalización {t.lower()}"  # no tenía paréntesis, se deja tal cual

        if es_sobrepuesta:  # la canaleta sobrepuesta se cotiza por tramo, la embutida no
            return f"{base}, 2mts"  # la canaleta sobrepuesta se vende en tramos de 2 metros
        return base  # si es embutida sale sin metraje; el ", 3mts" se lo agrega después quien llama

    # Busca en el texto de canalización una medida tipo "20x10mm" para
    # que los accesorios (unión copla, curvas, curva T) usen la misma
    # medida que la canaleta real del circuito.
    def _extraer_medida_canaleta(canal_txt: str) -> str:
        """
        Saca la medida "AAxBBmm" del texto de canalización, para que los
        accesorios (unión copla, curvas, curva T) usen la MISMA medida que
        la canaleta real de ese circuito, en vez de un valor fijo.
        Ejemplo: "Sobrepuesta (Canaleta de PVC 40x16mm)" da "40x16mm"
        Si no encuentra el patrón, usa "20x10mm" como valor por defecto.
        """
        import re  # módulo para buscar patrones de texto (expresiones regulares)
        m = re.search(r'\d+x\d+mm', canal_txt or "")  # busca el patrón número x número mm dentro del texto
        return m.group(0) if m else "20x10mm"  # si lo encontró lo devuelve, si no usa 20x10mm por defecto

    # Saca el diámetro en mm del conduit desde el texto de canalización
    # (por ejemplo, de "conduit 20mm" saca el número 20).
    def _extraer_mm_conduit(canal_txt: str):
        """
        Ejemplo: "Embutida (PVC conduit 20mm)" da 20
        """
        t = (canal_txt or "").lower().replace(" ", "")  # sin espacios, para buscar más fácil
        # buscar "...conduit20mm" o "conduit16mm"
        if "conduit" in t and "mm" in t:
            try:  # por si el texto viene mal armado y no se puede leer el número
                seg = t.split("conduit", 1)[1]  # lo que viene después de la palabra "conduit"
                num = ""  # num: va a guardar el número que aparezca después de "conduit"
                for ch in seg:  # va leyendo caracter por caracter hasta que se acabe el diámetro
                    if ch.isdigit():  # mientras sean dígitos siguen siendo parte del diámetro
                        num += ch  # va juntando los dígitos del número
                    else:  # llegó a un caracter que ya no es parte del diámetro
                        break  # se acabó el número
                return int(num) if num else None  # si junto dígitos los convierte a número entero
            except:  # el texto venía mal armado, no se pudo leer el diámetro
                return None  # algo fallo al convertir, no hay diámetro
        return None  # el texto no tenía "conduit...mm"

    def add_section(nombre):  # fila de título para separar secciones del listado
        # agrega una fila "separadora" con solo el título de la sección
        # (ej: "Empalme"), para que se vea ordenado el Excel.
        # recibe: nombre = el título que se quiere mostrar en esa fila.
        # no devuelve nada, solo agrega la fila a la lista "filas".
        filas.append({  # fila con todas las columnas vacias, solo con el título de la sección
            "Ítem": "",  # la fila de título no lleva número de ítem
            "Descripción técnica": nombre,  # acá va el nombre de la sección, ej: Empalme
            "Marcas sugeridas": "",  # tampoco marcas
            "Sello SEC": "",  # ni sello SEC
            "Norma / RIC": "",  # ni artículo del RIC
            "Circuito": "",  # no pertenece a ningún circuito
            "Unidad": "",  # no se compra nada, va vacía
            "K": "",  # sin factor multiplicador
            "Longitud (m) / Unidad": "",  # sin metraje
            "Cantidad": ""  # y sin cantidad que comprar
        })  # cierra la fila separadora de sección

    def add_row(desc, marcas_txt, norma, circuito, unidad, k, longitud_m, cantidad):  # agrega una fila de material al listado final
        # agrega una fila de material al listado final: le pone el número de
        # ítem, revisa si necesita Sello SEC y guarda todos los datos.
        # qué es cada dato que recibe:
        #   desc ......... descripción técnica del material (lo que se ve en el Excel)
        #   marcas_txt ... marcas sugeridas, sale del diccionario "marcas" de más arriba
        #   norma ........ artículo del RIC que lo respalda, o "SEC" si el
        #                  material necesita Sello SEC
        #   circuito ..... a qué circuito o tramo pertenece
        #   unidad ....... "u" si se compra por unidad, "m" si se compra por metro
        #   k ............ factor multiplicador, casi siempre 1
        #   longitud_m ... texto que se muestra en la columna, ej: "12 m" o "3 unid"
        #   cantidad ..... el número final, o sea lo que hay que comprar
        # no devuelve nada: agrega la fila y sube el contador de ítems.
        nonlocal item  # usa (y modifica) el contador de items de más afuera
        sello = "SEC" if str(norma).strip() == "SEC" else "-"  # valor provisorio: al final la columna Sello SEC se recalcula con _requiere_sello_sec()
        filas.append({  # arma la fila del material con todas las columnas del listado
            "Ítem": item,  # número correlativo del item en el listado
            "Descripción técnica": desc,  # el material tal como se ve en el Excel
            "Marcas sugeridas": marcas_txt,  # marcas sugeridas para ese material
            "Sello SEC": sello,  # SEC o guión, según lo que se revisó arriba
            "Norma / RIC": norma,  # norma con artículo y todo. normalizar_ric_materiales() la recorta al tomo ("RIC 4") solo como paso intermedio: en el Excel la celda termina siempre como "Ver normativa" con link, o como "-". El artículo completo queda solo en el código
            "Circuito": circuito,  # a qué circuito o tramo pertenece
            "Unidad": unidad,  # u si se compra por unidad, m si va por metro
            "K": k,  # factor multiplicador, casi siempre 1
            "Longitud (m) / Unidad": longitud_m,  # texto de metraje o unidades, ej 12 m
            "Cantidad": cantidad  # lo que finalmente hay que comprar
        })  # cierra la fila del material
        item += 1  # sube el número de ítem para la próxima fila

    # -----------------------------
    # Conectores cónicos
    # -----------------------------
    # Saca el número de sección (en mm^2) desde el texto del conductor,
    # por ejemplo "THWN-2 2.5 mm^2" da 2.5. Sirve para elegir el
    # conector cónico correcto más abajo.
    def extraer_seccion_mm2(txt_conductor: str):
        """
        Ejemplo: "H07V 2.5 mm^2" da 2.5
        Ejemplo: "THWN-2 1.5 mm^2" da 1.5
        """
        s = (txt_conductor or "").lower().replace("mm^2", "mm2")  # pasa el "mm^2" del texto a un "mm2" simple, que es lo que se busca abajo
        toks = s.split()  # separa el texto en palabras
        for i, tok in enumerate(toks):  # recorre las palabras buscando donde dice "mm2"
            if tok == "mm2" and i > 0:  # ojo: acá se compara contra "mm2" porque la línea de arriba ya lo normalizó
                try:  # la palabra de antes debería ser el número de la sección
                    return float(toks[i-1].replace(",", "."))  # el número justo antes del "mm2"
                except:  # esa palabra no era un número, se sigue buscando
                    pass  # ese no era número, sigue buscando en el resto del texto
        # si no lo encontró así, junta cualquier número suelto en el texto
        num = ""  # num: va juntando los dígitos que encuentre
        for ch in s:  # junta el primer número que aparezca en el texto
            if ch.isdigit() or ch in [".", ","]:  # dígitos, punto y coma son parte del número
                num += ch  # va juntando dígitos y separadores decimales
            elif num:  # ya venía juntando dígitos y se cortó el número
                break  # ya encontró un número, se detiene
        try:  # convierte lo que junto a número, si es que se puede
            return float(num.replace(",", "."))  # convierte el texto juntado a número, con punto decimal
        except:  # no había ningún número aprovechable en el texto
            return None  # no se pudo sacar ningún número

    # Devuelve el conector cónico que corresponde (color, número y el
    # texto que se muestra en el listado) según la sección del cable y
    # cuántos conductores se unen en ese punto.
    def conico_por_seccion(seccion_mm2, n_cables=3):
        """
        Tabla de colores de conectores cónicos, según catálogo real, en
        función de la SECCIÓN del cable y la CANTIDAD DE CABLES que se
        unen en ese punto (2 o 3):

          Sección      | 2 cables            | 3 cables
          -------------|---------------------|---------------------
          1.5mm^2       | N°33 (P73) Naranja  | N°33 (P73) Naranja
          2.5mm^2       | N°44 (P74) Amarillo | N°44 (P74) Amarillo
          4.0mm^2       | N°66 (P75) Rojo     | N°66 (P75) Rojo
          5.26mm^2(10AWG)| N°66 (P75) Rojo    | N°88 (SP8/P78) Gris Grande
          6.0mm^2       | N°66 (P75) Rojo     | N°88 (SP8/P78) Gris Grande

        No existen conectores cónicos para secciones mayores a 6.0mm^2.
        La sección real se redondea hacia arriba a la más cercana de la tabla.
        """
        try:  # la sección puede venir como texto o vacía
            s = float(seccion_mm2)  # s: la sección real del cable, convertida a número
        except:  # la sección no venía como número
            s = None  # no se pudo convertir, sección inválida

        if s is None:  # sin sección no se puede elegir el conector
            return ("(definir)", None, "")  # sección inválida, no se puede determinar

        if s > 6.0:  # no hay conectores cónicos para secciones tan grandes
            return ("(definir)", None, f"{s} mm^2 (fuera de rango cónico)")  # queda en (definir) para revisarlo a mano, ahí se usa otro tipo de unión

        # Redondea hacia arriba a la sección comercial de la tabla
        secciones_tabla = [1.5, 2.5, 4.0, 5.26, 6.0]  # secciones comerciales disponibles, de menor a mayor
        sec_com = None  # sec_com: la sección "comercial" que le toca (la más
                         # chica de la tabla que alcanza a cubrir la real)
        for sec in secciones_tabla:  # busca la primera sección de la tabla que alcance a cubrir la real
            if s <= sec:  # esta sección comercial ya sirve
                sec_com = sec  # se queda con esa
                break  # no sigue buscando
        if sec_com is None:  # por si la sección es mayor que todas las de la lista (no debería pasar, ya se filtro antes)
            sec_com = 6.0  # se deja en la mayor de la tabla

        # tabla de 2 cables: sección (mm^2) -> (color, número, código de catálogo)
        tabla_2_cables = {
            1.5:  ("Naranja", 33, "P73"),  # 1.5mm^2 con 2 cables: cónico naranja N°33
            2.5:  ("Amarillo", 44, "P74"),  # 2.5mm^2 con 2 cables: amarillo N°44
            4.0:  ("Rojo", 66, "P75"),  # 4mm^2 con 2 cables: rojo N°66
            5.26: ("Rojo", 66, "P75"),  # 5.26mm^2 (10AWG) con 2 cables todavía entra en el rojo 66
            6.0:  ("Rojo", 66, "P75"),  # 6mm^2 con 2 cables también aguanta el rojo 66
        }
        # tabla de 3 cables: mismos datos, pero para uniones de 3 conductores
        tabla_3_cables = {
            1.5:  ("Naranja", 33, "P73"),  # con 3 cables el 1.5mm^2 sigue en el naranja 33
            2.5:  ("Amarillo", 44, "P74"),  # 2.5mm^2 con 3 cables: amarillo 44
            4.0:  ("Rojo", 66, "P75"),  # 4mm^2 con 3 cables: el rojo 66 todavía da
            5.26: ("Gris Grande", 88, "SP8/P78"),  # con 3 cables el rojo 66 ya queda chico, se pasa al gris grande
            6.0:  ("Gris Grande", 88, "SP8/P78"),  # mismo caso del 5.26: 3 cables de 6mm^2 necesitan el 88
        }

        tabla = tabla_2_cables if int(n_cables) == 2 else tabla_3_cables  # elige la tabla según si son 2 o 3 cables los que se unen
        color, num, pcode = tabla.get(sec_com, ("(definir)", None, ""))  # pcode: código de catálogo (ej. "P73")
        _sec_txt = f"{sec_com}".rstrip("0").rstrip(".") if "." in f"{sec_com}" else f"{sec_com}"  # _sec_txt: la sección sin ceros ni punto de sobra (ej: 4.0 -> "4")
        rango = f"{_sec_txt}mm^2, {int(n_cables)} cables ({pcode})"  # arma el texto que se va a mostrar en el listado
        return (color, num, rango)  # devuelve color, número de conector y el texto descriptivo

    # =========================
    # ACCESORIOS CANALETA SOBREPUESA (PVC, medida variable según sección)
    # Reglas propias (no del RIC): son criterios de cubicación del autor.
    # Acá sólo quedan escritas; el cálculo mismo está más abajo, en el bloque
    # "1) CANALIZACIONES", rama sobrepuesta, con add_row() por cada accesorio.
    #  ILUMINACIÓN (sobrepuesta):
    #    - Unión copla: una por cada unión entre tramos de 2m, o sea ceil(L/2) - 1
    #    - Curvas internas 90°: 4 por ambiente + 1 por interruptor
    #    - Curvas planas 90°: 2 por interruptor del circuito (1 salida tablero
    #      + 2 por cada caja troncal, excepto la última que lleva 1 = 2×n_int)
    #  ENCHUFES (sobrepuesta):
    #    - Unión copla: una por cada unión entre tramos de 2m, o sea ceil(L/2) - 1
    #    - Curvas internas 90°: 3 por ambiente
    #    - Curva T: (n° enchufes - 1)
    #    - Curvas planas 90°: 1 salida del circuito + 1 por ambiente
    #  CARGAS ESPECIALES (horno, lavadora... todo lo que no es iluminación ni
    #  enchufes por nombre; clima y agua caliente se cuentan en su propia
    #  sección): mismas coplas, 3 curvas internas por ambiente y
    #  1 + n_ambientes curvas planas, pero SIN curva T (hay un solo punto de
    #  conexión, no una cadena de enchufes).
    # =========================
    def _contar_ambientes_en_items(items):
        # cuenta cuántos ambientes distintos hay en los items de un circuito
        ambs = set()  # set para no repetir ambientes
        if isinstance(items, list):  # si items es una lista de diccionarios (items del circuito)
            for it in items:  # revisa cada punto del circuito
                if isinstance(it, dict):  # solo los items que vienen bien armados
                    a = str(it.get("amb", "")).strip()  # nombre del ambiente de este item
                    if a:  # si el item trae ambiente, se cuenta
                        ambs.add(a.strip().lower())  # lo agrega (sin mayúsculas, sin espacios)
        return len(ambs)  # cuántos ambientes distintos quedaron


    def _contar_enchufes_en_items(items):  # suma cuántos enchufes tiene el circuito
        # suma los enchufes del circuito (n_ench). Si un item no trae n_ench,
        # se cuenta como 1
        n = 0  # arranca el contador en 0
        if isinstance(items, list):  # solo recorre si items es una lista
            for it in items:  # revisa cada item (punto) del circuito
                if not isinstance(it, dict):  # si no es un diccionario, es un dato raro
                    continue  # item raro, se salta
                if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it):  # es un item de enchufe (tiene alguna de estas claves)
                    n += int(it.get("n_ench", 1) or 1)  # suma la cantidad de este item
        return int(n)  # devuelve el total de enchufes contados

    # ==========================================================
    # CONTEOS POR CIRCUITO PARA "Salida de caja conduit" (canalización embutida)
    #
    # Estas reglas eran la idea original. El conteo que de verdad corre está
    # más abajo, en el bloque "Salida de caja conduit: reglas por circuito",
    # y quedó bastante más fino que esta nota. Lo que hace hoy el código:
    #
    #  - Abrazaderas: NO es 1 por metro. La separación sale de la Tabla N°4.24
    #    del RIC N°4 (tuberías no metálicas): cada 1,20 m si el conduit es de
    #    16 a 25 mm, y cada 1,50 m de 32 mm para arriba. Además se calcula
    #    sobre los metros realmente comprados (tiras de 3 m completas), no
    #    sobre el largo medido del circuito.
    #  - Salidas de caja: parte en 1 (la salida del tablero) y después:
    #      * Iluminación: se recorre ambiente por ambiente e interruptor por
    #        interruptor. Cada caja troncal suma 3 (si es la última del
    #        circuito) o 4 (si el troncal sigue). La caja del interruptor suma
    #        1, salvo el conmutado 9/24 que son 2 cajas y suma 4. Cada caja
    #        octogonal de luminaria suma 2, o 1 si es la última de la cadena
    #        (el 9/24 suma 2 siempre, incluso en la última).
    #      * Enchufes y cargas especiales: 2 por enchufe intermedio y 1 por el
    #        último de cada ambiente, más 3 por cada caja de derivación
    #        adicional entre ambientes.
    #      * Cajas de paso: 2 cada una (entrada y salida).
    #
    # IMPORTANTÍSIMO (esto sí sigue vigente):
    #  - No se inventan uniones en circuitos de enchufes: si el usuario no pidió
    #    cajas adicionales, son 0.
    #  - Las uniones sólo se modelan en iluminación, con las cajas troncales.
    # ==========================================================
    # =========================
    # 1) CANALIZACIONES (UNA FILA POR CIRCUITO, NO AGRUPADO)
    #    accesorios embutida (abrazaderas + salidas de caja)
    #    accesorios sobrepuesta (canaleta PVC, medida variable según sección)
    # =========================


    # Definir clima_rows y llenar _clima_items ANTES de todas las secciones
    clima_rows = pd.DataFrame()  # arranca vacío, se llena más abajo si hay circuitos de climatización
    if circuitos_df is not None and "Circuito" in circuitos_df.columns:  # sigue solo si existe la tabla de circuitos con la columna Circuito
        # detecta los circuitos de climatización por el nombre (contiene "climatiz", "aire acond" o "split")
        mask_nombre = circuitos_df["Circuito"].astype(str).str.contains(
            "climatiz|aire.acond|split", case=False, na=False  # palabras con las que se reconoce un circuito de climatización
        )
        if "_es_climatizacion" in circuitos_df.columns:  # si existe una columna que marca climatización explícitamente
            # o también por una marca explícita que se le puso al crear el circuito
            mask_flag = circuitos_df["_es_climatizacion"].apply(
                lambda x: bool(x) if x is not None and x is not np.nan else False  # los vacíos (NaN) se toman como que no es climatización
            )
        else:  # si no existe esa columna, no hay marca explícita
            mask_flag = pd.Series(False, index=circuitos_df.index)  # entonces ninguna fila queda marcada por defecto
        clima_rows = circuitos_df[mask_nombre | mask_flag].copy()  # se queda con cualquiera de las dos formas

    # ── Llenar _agua_items (análogo a _clima_items) ──────────────────────────
    _agua_items = []  # lista donde se guardará un dict por cada equipo de agua caliente
    _KEYWORDS_AGUA = ("ducha", "termo", "calefon", "calefón", "calentador", "agua caliente")  # palabras que identifican un circuito de agua caliente por su nombre
    if circuitos_df is not None and "Circuito" in circuitos_df.columns:  # igual que en clima, solo si existe la tabla de circuitos
        # mismo criterio que climatización, pero buscando palabras de agua caliente
        mask_agua_nombre = circuitos_df["Circuito"].astype(str).str.contains(
            "|".join(_KEYWORDS_AGUA), case=False, na=False  # arma el patrón de búsqueda con todas las palabras de agua caliente
        )
        if "_es_agua_caliente" in circuitos_df.columns:  # columna que marca agua caliente explícitamente, si existe
            mask_agua_flag = circuitos_df["_es_agua_caliente"].apply(  # revisa esa columna fila por fila
                lambda x: bool(x) if x is not None and x is not np.nan else False  # los vacíos (NaN) se toman como que no es agua caliente
            )
        else:  # si no existe la columna, no hay marca explícita
            mask_agua_flag = pd.Series(False, index=circuitos_df.index)  # ninguna fila marcada por defecto
        agua_rows = circuitos_df[mask_agua_nombre | mask_agua_flag].copy()  # filas de circuitos que son de agua caliente

        # recorre cada circuito de agua caliente encontrado y arma su canalización
        for _, r_ac in agua_rows.iterrows():  # procesa cada equipo de agua caliente encontrado
            circ_ac_i  = str(r_ac.get("Circuito", "")).strip()  # nombre del circuito
            L_ac_i     = float(r_ac.get("Longitud (m)", 0) or 0)  # largo del tramo en metros
            canal_ac_i = str(r_ac.get("Canalización", "")).strip()  # texto de la canalización (ej. 'Embutida conduit 20mm')
            tm_ac_i    = str(r_ac.get("Disyuntor termomagnético", "")).strip()  # texto del TM (disyuntor termomagnético) de este circuito
            in_tm_ac_i = parse_in_tm(tm_ac_i)  # None si no se puede leer — no se asume 20A
            if L_ac_i <= 0:  # si no hay largo definido, no se puede calcular nada
                continue  # sin longitud, no hay nada que calcular para este equipo
            canal_low_ac  = canal_ac_i.lower()  # versión en minúsculas para comparar texto

            # Forzar conduit embutido si el equipo está en Vol.1
            # RIC N°11 art. 6.4.3: en Volumen 1 se exige mínimo IPX4, y la canaleta NO cumple ese requisito
            # Art. 6.5.3: cable bajo tubo aislante con IPX5 garantizado
            _vol1_ac = False  # por defecto, se asume que NO está en Volumen 1 (zona húmeda)
            if "_vol1_bano_agua" in agua_rows.columns:  # si el dato viene explícito en la tabla, se usa
                _vol1_ac = bool(r_ac.get("_vol1_bano_agua", False))  # toma el valor que venga de la tabla
            if not _vol1_ac:  # si no vino el dato, se adivina por el nombre del circuito
                # si no viene ese dato, se asume por el nombre: la ducha siempre está en Vol.1
                _vol1_ac = "ducha" in circ_ac_i.lower()  # las duchas siempre se consideran Volumen 1

            if _vol1_ac and ("canaleta" in canal_low_ac or "sobrepuesta" in canal_low_ac):  # zona húmeda con canaleta: no cumple, hay que corregir la canalización
                # está en volumen 1 pero eligieron canaleta: se fuerza a conduit embutido por norma
                canal_low_ac = "conduit embutida"  # se cambia el texto para que el resto del código lo trate como embutido
                canal_ac_i   = "Embutida (conduit forzado — RIC N°11 6.4.3 Vol.1 IPX4)"  # se deja registro de por qué se forzó el cambio

            es_emb_ac_i   = ("conduit" in canal_low_ac) or ("embutida" in canal_low_ac)  # True si la canalización es embutida (conduit)
            mm_ac_i       = _extraer_mm_conduit(canal_ac_i) or 20  # diámetro del conduit (20mm si no se pudo leer)
            # Buscar si lleva tablero externo
            _lleva_te = False  # por defecto, no lleva tablero externo
            for eq_ac_i in circuitos_agua_caliente:  # busca entre los equipos de agua caliente ingresados por el usuario
                if eq_ac_i.get("nombre_circ", "").lower() in circ_ac_i.lower():  # coincide el nombre del equipo con este circuito
                    _lleva_te = eq_ac_i.get("lleva_tablero_externo", False)  # guarda si este equipo tiene tablero externo propio
                    break  # ya se encontró, no sigue buscando
            _tiene_20m_ac = bool(r_ac.get("_tiene_tramo_20m", True))  # True si hay algún tramo continuo de 20m o más
            _cajas_paso_ac = int(L_ac_i // 20) if (es_emb_ac_i and _tiene_20m_ac) else 0  # 1 caja de paso cada 20m, solo si es embutido y hay tramo largo
            # guarda los datos de este equipo para usarlos más abajo
            _agua_items.append({
                "circ":               circ_ac_i,  # nombre del circuito del equipo
                "L":                  L_ac_i,  # largo del tramo en metros
                "canal":              canal_ac_i,  # texto de la canalización que quedó al final
                "in_tm":              in_tm_ac_i,  # corriente del TM, o None si no se pudo leer
                "es_emb":             es_emb_ac_i,  # si va embutido en conduit o no
                "mm":                 mm_ac_i,  # diámetro del conduit en mm
                "lleva_tab_ext":      _lleva_te,  # si el equipo lleva su propio tablero externo
                "cajas_paso":         _cajas_paso_ac,  # cajas de paso que le tocan por el largo del tramo
            })  # guarda todos los datos de este equipo de agua caliente
            _cajas_paso_total += _cajas_paso_ac  # cajas de paso de agua caliente también cuentan para el total
            # Este ítem se suma más abajo vía _agua_cajas_emb/_agua_cajas_sob,
            # una vez que ya se sabe si la canalización es embutida o sobrepuesta.

    def _buscar_datos_clima(nombre_circ):  # busca el equipo de clima que va con este circuito
        # OJO: esta función queda definida pero NUNCA se llama en todo el
        # archivo. Los datos del equipo de clima que sí se usan salen de
        # _clima_items, que se arma en el bucle de acá abajo.
        # busca el equipo de climatización que corresponde a este circuito,
        # para sacar sus datos técnicos (LRA, tipo de compresor, etc.)
        if not circuitos_climatizacion:  # si no hay equipos de climatización definidos
            return {}  # no hay equipos de clima definidos
        for eq in circuitos_climatizacion:  # recorre cada equipo de climatización ingresado
            if eq.get("nombre_circ", "").lower() in nombre_circ.lower():  # coincide el nombre
                return eq  # encontrado, devuelve sus datos
        return {}  # no se encontró ningún equipo para este circuito

    # recorre cada circuito de climatización encontrado y arma su canalización
    if len(clima_rows) > 0:  # solo si se detectaron circuitos de climatización
        for _, r_cl in clima_rows.iterrows():  # procesa cada circuito de climatización
            circ_cl  = str(r_cl.get("Circuito", "")).strip()  # nombre del circuito
            L_cl     = float(r_cl.get("Longitud (m)", 0) or 0)  # largo del tramo en metros
            canal_cl = str(r_cl.get("Canalización", "")).strip()  # texto de la canalización
            tm_cl    = str(r_cl.get("Disyuntor termomagnético", "")).strip()  # texto del TM de este circuito
            in_tm_cl_raw = parse_in_tm(tm_cl)  # intenta leer el valor numérico del TM (puede ser None)
            in_tm_cl = in_tm_cl_raw if in_tm_cl_raw is not None else 16  # si no se pudo leer, asume 16A
            if L_cl <= 0:  # sin largo no hay nada que calcular
                continue  # sin longitud, no hay nada que calcular
            canal_low   = canal_cl.lower()  # versión en minúsculas para comparar
            es_emb_cl   = ("conduit" in canal_low) or ("embutida" in canal_low)  # True si la canalización es embutida
            mm_cl       = _extraer_mm_conduit(canal_cl) or 20  # diámetro del conduit (20mm si no se pudo leer)
            # Si no se pudo leer el TM, no se adivina si tiene enchufe o no —
            # con_enchufe queda en None: el bloque de cónicos deja un aviso
            # "(definir)" para revisarlo a mano, y el bloque C2 de ferrules
            # simplemente no cuenta ferrules de enchufe para este equipo.
            con_enchufe = (in_tm_cl <= 16) if in_tm_cl_raw is not None else None
            _tiene_20m_cl = bool(r_cl.get("_tiene_tramo_20m", True))  # True si hay algún tramo continuo de 20m o más
            _cajas_paso_cl = int(L_cl // 20) if (es_emb_cl and _tiene_20m_cl) else 0  # 1 caja de paso cada 20m, solo si es embutido y hay tramo largo
            # guarda los datos de este circuito de climatización
            _clima_items.append({
                "circ":        circ_cl,  # nombre del circuito de climatización
                "L":           L_cl,  # largo del tramo en metros
                "canal":       canal_cl,  # texto de la canalización
                "in_tm":       in_tm_cl,  # corriente del TM del circuito
                "es_emb":      es_emb_cl,  # si va embutido en conduit
                "mm":          mm_cl,  # diámetro del conduit en mm
                "con_enchufe": con_enchufe,  # si el equipo se conecta por enchufe, None si no se sabe
                "cajas_paso":  _cajas_paso_cl,  # cajas de paso por los tramos de 20m
            })  # cierra el registro del circuito de climatización
            _cajas_paso_total += _cajas_paso_cl  # cajas de paso de climatización también cuentan para el total
            # Este ítem se suma más abajo vía _clima_cajas_emb/_clima_cajas_sob,
            # una vez que ya se sabe si la canalización es embutida o sobrepuesta.
            # este bucle NO toca _tapas_ciegas_total: ese contador se calcula
            # entero, de una sola vez, en el bloque TAPAS CIEGAS más abajo

    # Cajas de paso cada 20m (RIC N°4, art. 7.16.1.13) para enchufes,
    # iluminación y especiales. Climatización y agua caliente no entran aquí
    # porque ya cargaron sus propias cajas de paso más arriba, en _agua_items
    # y _clima_items (si entraran otra vez quedarían contadas dos veces).
    # Este bloque no emite ninguna fila de material: solo llena el contador
    # _cajas_paso_total, que después se usa en el bloque ACCESORIOS (cajas
    # rectangulares / cajas chuqui) y en el bloque TAPAS CIEGAS.
    if {"Canalización", "Longitud (m)", "Circuito"}.issubset(set(circuitos_df.columns)):  # solo si existen todas estas columnas
        _tmp_pre = circuitos_df.copy()  # copia para no modificar el DataFrame original
        _tmp_pre["Longitud (m)"] = pd.to_numeric(_tmp_pre["Longitud (m)"], errors="coerce").fillna(0)  # convierte a número; si no se puede, queda en 0
        for _, _r_pre in _tmp_pre.iterrows():  # revisa cada circuito
            _canal_pre = str(_r_pre.get("Canalización", "")).strip().lower()  # canalización en minúsculas
            _circ_pre  = str(_r_pre.get("Circuito", "")).strip().lower()  # nombre del circuito en minúsculas
            _L_pre     = float(_r_pre.get("Longitud (m)", 0))  # largo del tramo
            _es_emb_pre  = ("conduit" in _canal_pre) or ("embutida" in _canal_pre)  # True si es canalización embutida
            _es_cli_pre  = any(k in _circ_pre for k in ("climatiz", "aire acond", "split"))  # True si es un circuito de climatización
            _es_agua_pre = any(k in _circ_pre for k in _KEYWORDS_AGUA)  # True si es un circuito de agua caliente
            # Si no hay ningún tramo continuo >=20m (confirmado en el input),
            # no se cuenta ninguna caja de paso aunque el LARGO TOTAL del
            # circuito supere 20m (puede ser la suma de varios tramos cortos)
            _tiene_20m_pre = bool(_r_pre.get("_tiene_tramo_20m", True))  # True si hay algún tramo continuo de 20m o más
            if _es_emb_pre and not _es_cli_pre and not _es_agua_pre and _L_pre > 0 and _tiene_20m_pre:  # solo enchufes/iluminación/especiales embutidos, no clima ni agua (para no duplicar)
                _cajas_paso_total += int(_L_pre // 20)  # 1 caja de paso cada 20m; es división entera: 19m da 0, 20m da 1, 41m da 2

    add_section("Canalizaciones")  # abre la sección del informe/Excel para canalizaciones
    # OJO con el parámetro norma= de add_row que se usa en todo este bloque: el
    # artículo detallado ("RIC 4.7.2", "RIC 6.4.1", ...) NO llega al Excel tal
    # cual. Al final, normalizar_ric_materiales() lo recorta al tomo ("RIC 4",
    # "RIC 6") y después la celda se reemplaza por "Ver normativa" con link a la
    # hoja Base Normativa. El artículo queda solo como respaldo en el código.
    # ncon_map_local solo depende de ambientes_df, disponible desde el inicio de la función.
    ncon_map_local = {}  # diccionario: ambiente -> cuántas luminarias conmutadas tiene
    if ambientes_df is not None and "Ambiente" in ambientes_df.columns:  # solo si existe la tabla de ambientes
        amb_col_l = ambientes_df["Ambiente"].astype(str).str.strip()  # nombres de los ambientes, sin espacios sobrantes
        if "N_conmutadas_924 (u)" in ambientes_df.columns:  # solo si esa columna existe
            ncon_col_l = pd.to_numeric(ambientes_df["N_conmutadas_924 (u)"], errors="coerce").fillna(0).astype(int)  # cantidad de luminarias conmutadas por ambiente
            for a_l, nval_l in zip(amb_col_l, ncon_col_l):  # recorre cada ambiente junto con su cantidad
                ncon_map_local[a_l.lower()] = int(max(0, nval_l))  # cuántas luminarias conmutadas tiene cada ambiente

    # recorre TODOS los circuitos (uno por fila del Excel) y arma su canalización
    if {"Canalización", "Longitud (m)", "Circuito"}.issubset(set(circuitos_df.columns)):  # solo si existen todas estas columnas
        circuitos_tmp = circuitos_df.copy()  # copia para no modificar el DataFrame original
        circuitos_tmp["Longitud (m)"] = pd.to_numeric(circuitos_tmp["Longitud (m)"], errors="coerce").fillna(0)  # convierte a número; si no se puede, queda en 0
        for _row_idx, r in circuitos_tmp.iterrows():  # recorre cada circuito, uno por fila
            canal = str(r.get("Canalización", "")).strip()  # texto de la canalización de este circuito
            circ = str(r.get("Circuito", "")).strip()  # nombre del circuito
            L = float(r.get("Longitud (m)", 0))  # largo del tramo en metros

            if not canal or not circ or L <= 0:  # si falta algún dato clave, se salta este circuito
                continue  # circuito sin datos válidos, se salta
            canal_low = canal.lower()  # versión en minúsculas para comparar texto
            es_sobrepuesta = ("canaleta" in canal_low) or ("sobrepuesta" in canal_low)  # True si va por canaleta PVC a la vista
            es_embutida = ("conduit" in canal_low) or ("embutida" in canal_low)  # True si va por conduit embutido en la muralla

            # Climatización: el conduit/canaleta se genera aquí normalmente,
            # pero los accesorios (abrazaderas, salidas, coplas, curvas) se generan
            # en la sección "Climatización" con la lógica correcta, así que acá se saltan los accesorios
            _es_clima_canal = False  # por defecto, se asume que no es climatización
            if "_es_climatizacion" in circuitos_tmp.columns:  # si la columna existe, se usa su valor
                _es_clima_canal = bool(r.get("_es_climatizacion", False))  # marca explícita de climatización
            if not _es_clima_canal:  # si no vino marcado, se adivina por el nombre
                _es_clima_canal = any(k in circ.lower() for k in ("climatiz", "aire acond", "split"))  # busca palabras clave de climatización en el nombre

            # Agua caliente: igual que climatización — accesorios (boquillas, salidas de caja)
            # se generan en su propio bloque _agua_items, así que acá se saltan las salidas genéricas
            _es_agua_canal = False  # por defecto, se asume que no es agua caliente
            if "_es_agua_caliente" in circuitos_tmp.columns:  # si la columna existe, se usa su valor
                _es_agua_canal = bool(r.get("_es_agua_caliente", False))  # marca explícita de agua caliente
            if not _es_agua_canal:  # si no vino marcado, se adivina por el nombre
                _es_agua_canal = any(k in circ.lower() for k in _KEYWORDS_AGUA)  # busca palabras clave de agua caliente en el nombre
            # fila de canalización
            # elige qué artículo del RIC va en la columna 'Norma' de esta fila
            _norma_canal = (
                "RIC 7 (7.5.1, 7.5.2) / RIC 4 (7.1.3, 7.1.8, 7.1.9, 7.1.10, 7.1.15)"  # norma que se cita cuando es clima o agua caliente
                if (_es_clima_canal or _es_agua_canal)  # se usa esa norma solo si el circuito es de clima o agua caliente
                else "RIC 4.7.2"  # canalización común: basta con RIC 4.7.2
            )
            if es_sobrepuesta:  # rama para canalización sobrepuesta (canaleta PVC)
                # canaleta: se compra en tramos de 2 metros, se redondea hacia arriba
                tramo_m = 2.0  # la canaleta se vende en tramos de 2 metros
                cantidad_tramos = int(math.ceil(L / tramo_m))  # cuántos tramos de 2m hay que comprar (redondeado hacia arriba)
                # agrega la fila de canaleta al listado de materiales
                add_row(
                    desc=formatear_canalizacion_material(canal, es_sobrepuesta=True),  # texto del material tal como sale en la lista
                    marcas_txt=marcas.get("Canaleta PVC (sobrepuesta)", ""),  # marcas sugeridas de canaleta, si están cargadas
                    norma=_norma_canal,  # artículo del RIC que respalda esta canalización
                    circuito=circ,  # circuito al que se carga este material
                    unidad="u",  # se compra por tramo, no por metro
                    k=1,  # sin factor extra, la cantidad va tal cual
                    longitud_m=round(L, 4),  # largo medido del tramo, solo informativo
                    cantidad=cantidad_tramos  # tramos de 2m que hay que comprar
                )
                # acumula METROS de canaleta realmente comprada (tramos x 2m),
                # no el largo medido L: el bloque de tornillos, mucho más abajo,
                # compra 1 tornillo por cada metro de este total
                _long_sobrepuesta_total += cantidad_tramos * 2  # 1 tornillo por metro FÍSICO comprado (no el mínimo teórico)

                # Climatización: los accesorios (coplas, curvas) se generan en la sección Climatización, así que acá se saltan
                if _es_clima_canal:  # climatización ya tiene sus propios accesorios más abajo
                    continue  # no genera accesorios de canaleta genéricos para clima
                # Agua caliente: los accesorios se generan en el bloque _agua_items, así que acá se saltan
                if _es_agua_canal:  # agua caliente ya tiene sus propios accesorios más abajo
                    continue  # no genera accesorios de canaleta genéricos para agua

                # =========================
                # ACCESORIOS CANALETA PVC (SOBREPUESTA, medida variable según sección)
                # =========================
                # Los puntos del circuito se sacan del diccionario items_por_nombre
                # (que se pasa aparte a la función), no de la columna "_items" del
                # DataFrame: esa columna se borra antes de llamar a build_materiales_df
                _circ_key = str(r.get("Circuito",""))  # nombre del circuito, para buscar sus items
                items = (items_por_nombre or {}).get(_circ_key, [])  # los puntos (luminarias/enchufes) de este circuito
                if not isinstance(items, list):  # por seguridad, si no es una lista
                    items = []  # se trata como si no tuviera items
                # Unión copla: 1 copla por cada unión entre tramos de 2m, o sea n_tramos - 1
                _n_tramos = int(math.ceil(L / 2.0))  # cuántos tramos de 2m tiene este circuito
                union_copla = max(0, _n_tramos - 1)  # nunca negativo
                # contar ambientes / focos / enchufes del circuito
                n_amb = _contar_ambientes_en_items(items)  # cuántos ambientes distintos recorre este circuito
                n_ench = _contar_enchufes_en_items(items)  # cuántos enchufes tiene este circuito
                # determinar si el circuito es iluminación o enchufe (por nombre)
                circ_low = (circ or "").lower()  # nombre del circuito en minúsculas
                es_circ_ilum = ("ilumin" in circ_low)  # True si el circuito es de iluminación
                es_circ_ench = ("enchufe" in circ_low)  # True si el circuito es de enchufes
                # Unión copla (sirve para ambos: ilum/ench)
                if union_copla > 0:  # solo si hace falta al menos una unión
                    # agrega la fila de uniones copla al listado
                    add_row(
                        desc=f"Unión copla para canaleta de PVC {_extraer_medida_canaleta(canal)}",  # descripción de la copla con la medida de canaleta que corresponde
                        marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                        norma="RIC 4.7.2",  # canalización, RIC 4.7.2
                        circuito=circ,  # circuito al que se carga
                        unidad="u",  # las coplas se cuentan por unidad
                        k=1,  # sin factor extra
                        longitud_m=f"{union_copla} unid",  # cantidad en texto para el informe
                        cantidad=union_copla  # una copla por cada unión entre tramos
                    )
                # Accesorios según tipo de circuito
                if es_circ_ilum:  # reglas de accesorios para circuitos de iluminación
                    # Curvas internas 90°: 4 por ambiente + 1 por interruptor
                    # (1 curva por bajada desde caja rect hacia primera oct del grupo)
                    _n_int_circ_real = 0  # cuenta cuántos interruptores físicos hay en total en el circuito
                    _ambs_vistos = set()  # para no procesar el mismo ambiente dos veces
                    for _it in items:  # revisa cada item del circuito
                        if not isinstance(_it, dict): continue  # item raro, se salta
                        if "nombre" in _it or "id_ench" in _it or "modulos" in _it or "n_ench" in _it: continue  # salta enchufes/especiales
                        _amb_k = str(_it.get("amb", "")).strip().lower() or "sin_amb"  # nombre del ambiente de este item (o 'sin_amb' si no tiene)
                        if _amb_k in _ambs_vistos: continue  # este ambiente ya se contó
                        _ambs_vistos.add(_amb_k)  # lo marca como ya procesado
                        # cuenta cuántas luminarias tiene este ambiente en este circuito
                        _n_lum_amb_c = sum(1 for x in items if isinstance(x, dict) and
                                          str(x.get("amb","")).strip().lower() == _amb_k and  # solo los items del mismo ambiente
                                          "nombre" not in x and "id_ench" not in x and  # deja fuera cargas especiales y enchufes
                                          "modulos" not in x and "n_ench" not in x)  # y también multitomas y módulos
                        _n_conm_amb_c = int(min(_n_lum_amb_c, ncon_map_local.get(_amb_k, 0)))  # cuántas son conmutadas
                        _n_rest_amb_c = _n_lum_amb_c - _n_conm_amb_c  # el resto, no conmutadas
                        _c12c, _c15c, _c32c = descomponer_interruptores(_n_rest_amb_c)  # cómo se agrupan
                        _n_int_circ_real += _c12c + _c15c + _c32c + (2 if _n_conm_amb_c > 0 else 0)  # si el ambiente tiene conmutadas, suma 2 (un solo par 9/24 por ambiente, no 2 por luminaria)


                    curvas_int_90 = int(4 * int(n_amb) + _n_int_circ_real)  # 4 curvas por ambiente + 1 por interruptor
                    # Curvas planas 90°:
                    # 1 salida tablero + 2 por cada caja troncal (excepto última del circuito que lleva 1)
                    # = 1 + 2*(n_int-1) + 1 = 2*n_int
                    curvas_plan_90 = 2 * _n_int_circ_real  # el desglose de arriba (1 + 2*(n-1) + 1) se simplifica a 2 x n_interruptores
                    # Patrón que se repite de aquí en adelante: si la cantidad calculada es mayor a 0,
                    # se agrega una fila de material; si es 0, no se agrega nada (no se compra 0 unidades).
                    # Curvas internas en 90 grados para la canaleta de iluminación
                    if curvas_int_90 > 0:  # si no sale ninguna curva, no se agrega la fila
                        add_row(
                            desc=f"Curvas internas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # texto que va a aparecer en la lista de materiales
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas o modelos sugeridos para este material, si hay definidas
                            norma="RIC 4.7.2",  # respaldo normativo de la fila; en el Excel queda recortado a "RIC 4" y luego linkeado
                            circuito=circ,  # a que circuito pertenece esta fila
                            unidad="u",  # unidad de medida: "u" = unidad, "m" = metro
                            k=1,  # solo llena la columna "K" del Excel; NO multiplica la cantidad, esta ya viene calculada
                            longitud_m=f"{curvas_int_90} unid",  # texto informativo de cantidad que se muestra en el informe
                            cantidad=curvas_int_90  # cantidad real que se suma al total de materiales
                        )
                    # Curvas planas en 90 grados para la canaleta de iluminación
                    if curvas_plan_90 > 0:  # idem: si no hay curvas planas no se compra nada
                        add_row(
                            desc=f"Curvas planas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # nombre del accesorio con la medida de la canaleta
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curvas_plan_90} unid",  # cantidad en texto para el informe
                            cantidad=curvas_plan_90  # curvas planas que se compran
                        )
                # Circuito de enchufes (por nombre): otra cantidad de curvas/accesorios
                elif es_circ_ench:  # rama para circuitos de enchufes
                    # Curvas internas 90°: 3 por ambiente
                    curvas_int_90 = int(3 * int(n_amb))  # el enchufe no lleva bajada a interruptor, por eso 3 y no 4
                    # Curva T: n° enchufes - 1 (mínimo 0)
                    curva_t = int(max(0, int(n_ench) - 1))  # una T por cada enchufe intermedio de la cadena
                    # Curvas planas 90°: 1 salida del circuito + 1 por cada
                    # ambiente (el último enchufe de cada ambiente)
                    curvas_plan_90 = 1 + int(n_amb)  # la salida del circuito más la del último enchufe de cada ambiente
                    # fila: curvas internas 90 grados (enchufes)
                    if curvas_int_90 > 0:  # si no sale ninguna curva interna, no se agrega la fila
                        add_row(
                            desc=f"Curvas internas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # nombre del accesorio con la medida de la canaleta
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curvas_int_90} unid",  # cantidad en texto para el informe
                            cantidad=curvas_int_90  # curvas internas que se compran
                        )
                    # fila: curva T, para encadenar varios enchufes en el mismo tramo
                    if curva_t > 0:  # con un solo enchufe no hay nada que ramificar
                        add_row(
                            desc=f"Curva T para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # curva T, para sacar el ramal hacia el enchufe
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curva_t} unid",  # cantidad en texto para el informe
                            cantidad=curva_t  # curvas T que se compran
                        )
                    # fila: curvas planas 90 grados (enchufes)
                    if curvas_plan_90 > 0:  # si no sale ninguna curva plana, no se agrega la fila
                        add_row(
                            desc=f"Curvas planas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # nombre del accesorio con la medida de la canaleta
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curvas_plan_90} unid",  # cantidad en texto para el informe
                            cantidad=curvas_plan_90  # curvas planas que se compran
                        )
                else:  # circuitos especiales: ni iluminación ni enchufes por nombre
                    # Especiales genéricos (lavadora, horno, etc. — no
                    # iluminación ni enchufes por nombre, y clima/agua ya se
                    # excluyeron arriba con continue): mismo criterio que
                    # enchufes — 3 curvas internas por ambiente, sin curva T
                    # (1 solo punto de conexión, no hay cadena entre varios
                    # enchufes), 1 salida del circuito + 1 por ambiente
                    curvas_int_90 = int(3 * int(n_amb))  # 3 curvas internas por ambiente, igual que en enchufes
                    curvas_plan_90 = 1 + int(n_amb)  # 1 salida del circuito + 1 por ambiente, igual que en enchufes
                    # fila: curvas internas 90 grados
                    if curvas_int_90 > 0:  # si no sale ninguna curva interna, no se agrega la fila
                        add_row(
                            desc=f"Curvas internas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # nombre del accesorio con la medida de la canaleta
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curvas_int_90} unid",  # cantidad en texto para el informe
                            cantidad=curvas_int_90  # curvas internas que se compran
                        )
                    # fila: curvas planas 90 grados
                    if curvas_plan_90 > 0:  # si no sale ninguna curva plana, no se agrega la fila
                        add_row(
                            desc=f"Curvas planas en 90° para canaletas de PVC {_extraer_medida_canaleta(canal)}",  # nombre del accesorio con la medida de la canaleta
                            marcas_txt=marcas.get("Accesorios canaleta PVC", ""),
                            norma="RIC 4.7.2",
                            circuito=circ,
                            unidad="u",
                            k=1,
                            longitud_m=f"{curvas_plan_90} unid",  # cantidad en texto para el informe
                            cantidad=curvas_plan_90  # curvas planas que se compran
                        )
                # ya se agregaron todos los materiales de esta canaleta sobrepuesta; sigue con el siguiente circuito
                continue
            # Canalización tipo conduit embutido (va dentro de la muralla o losa, no queda a la vista)
            elif es_embutida:
                # conduit: se compra en tramos de 3 metros, se redondea hacia arriba
                tramo_m = 3.0  # el conduit viene en tiras de 3m
                cantidad_tramos = int(math.ceil(L / tramo_m))  # redondea hacia arriba: aunque sobre material, hay que comprar el tramo completo
                desc_base = formatear_canalizacion_material(canal, es_sobrepuesta=False)  # arma el texto descriptivo del conduit (diámetro, tipo)
                desc_final = f"{desc_base}, 3mts"  # agrega la medida del tramo a la descripción
                # fila: el conduit en si, comprado en tramos de 3 metros
                add_row(
                    desc=desc_final,  # descripción del conduit ya armada arriba
                    marcas_txt=marcas.get("Conduit PVC (embutida)", ""),  # marcas de conduit embutido
                    norma=_norma_canal,  # RIC de clima/agua o el genérico, según el circuito
                    circuito=circ,  # circuito al que se carga
                    unidad="u",  # se compra por tira, no por metro
                    k=1,  # sin factor extra
                    longitud_m=f"{round(L, 4)}m",  # largo real medido, solo informativo
                    cantidad=cantidad_tramos  # tiras de 3m que hay que comprar
                )
                # accesorios canalización embutida por circuito
                mm = _extraer_mm_conduit(canal) or 20  # diámetro del conduit en mm; si no se pudo extraer del texto, usa 20mm por defecto
                # Capturar mm conduit de iluminación para cajas y tapas
                if "ilumin" in circ.lower():  # solo interesa el diámetro del conduit de iluminación
                    _mm_conduit_ilumin = mm  # se guarda para las cajas octogonales y tapas de más abajo
                # el conduit se compra en tramos de 3m completos, así que la
                # cantidad real instalada es cantidad_tramos*3 (no el largo
                # medido L) — las abrazaderas van sobre esa cantidad
                L_efectiva = cantidad_tramos * tramo_m  # metros de conduit realmente comprados, no el largo medido
                # 1) Abrazaderas: separación según Tabla N°4.24 (tuberías no metálicas)
                if mm <= 25:  # conduit delgado (16 a 25mm), abrazaderas más juntas
                    sep_abraz = 1.20   # de 16 a 25mm, van cada 1,20m
                else:  # conduit más grueso (32 a 63mm), aguanta más separación
                    sep_abraz = 1.50   # de 32 a 63mm, van cada 1,50m
                cant_abraz = int(math.ceil(L_efectiva / sep_abraz))  # cantidad de abrazaderas para todo el tramo instalado
                # fila: abrazaderas para fijar el conduit a la muralla
                add_row(
                    desc=f"Abrazadera conduit de PVC de {mm}mm",  # abrazadera del mismo diámetro que el conduit
                    marcas_txt=marcas.get("Abrazadera conduit", ""),  # marcas de abrazadera
                    norma="RIC 4 (Tabla N°4.24)",  # tabla del RIC que fija la separación entre abrazaderas
                    circuito=circ,  # circuito al que se carga
                    unidad="u",  # se cuentan por unidad
                    k=1,  # sin factor extra
                    longitud_m=f"{cant_abraz} unid",  # cantidad en texto para el informe
                    cantidad=cant_abraz  # abrazaderas de este circuito
                )
                _abrazaderas_total += cant_abraz  # acumula el total de abrazaderas de todos los circuitos (para el resumen general)

                # Cajas de paso intermedias (RIC N°4, art. 7.16.1.13): 1 cada 20m
                # de tramo. Acá no se emite fila: la caja física se agrupa con las
                # demás rectangulares en el bloque ACCESORIOS (vía _cajas_paso_total),
                # y sus salidas de caja se suman unas líneas más abajo (cajas_paso_circ).

                # Climatización embutida: las salidas de caja se generan en la sección Climatización, así que acá se saltan
                if _es_clima_canal:  # clima embutida: sus salidas de caja se cuentan en la sección Climatización
                    continue  # no se cuentan salidas genéricas acá
                # Agua caliente embutida: las boquillas y salidas se generan en el bloque _agua_items, así que acá se saltan
                if _es_agua_canal:  # agua caliente: sus salidas se cuentan en _agua_items
                    continue  # no se cuentan salidas genéricas acá

                # Salida de caja conduit: reglas por circuito (con 2 por enchufe)
                # Igual que en la rama sobrepuesta: los puntos salen de
                # items_por_nombre, no del DataFrame
                _circ_key2 = str(r.get("Circuito",""))  # nombre del circuito, para buscar sus puntos (luminarias/enchufes)
                items = (items_por_nombre or {}).get(_circ_key2, [])  # lista de puntos (luminarias/enchufes/especiales) de este circuito
                if not isinstance(items, list):  # si vino algo raro en vez de una lista
                    items = []  # si no hay datos, se asume lista vacía
                # OJO: n_ench_circ y n_lum_circ se cuentan acá pero después NO se
                # usan para nada — la rama de enchufes de más abajo vuelve a contar
                # todo por su cuenta en n_ench_total. Quedan como conteo muerto.
                n_ench_circ = 0  # cantidad de enchufes del circuito
                n_lum_circ = 0   # cantidad de luminarias del circuito
                # contar enchufes/luminarias del circuito desde _items
                if isinstance(items, list):  # recorre los puntos del circuito para contarlos
                    for it in items:  # revisa punto por punto
                        if not isinstance(it, dict):  # se salta si el item no es un diccionario válido
                            continue  # item inválido, se salta
                        # enchufes
                        if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it):  # es un enchufe: simple, multitoma o módulo
                            n_ench_circ += int(it.get("n_ench", 1) or 1)  # suma la cantidad de enchufes de este item (1 si no viene definida)
                            continue  # ya quedó contado como enchufe
                        # especiales NO cuentan
                        if "nombre" in it:  # las cargas especiales (horno, lavadora) traen nombre
                            continue  # las especiales no se cuentan acá
                        # luminaria
                        n_lum_circ += 1  # lo que queda es una luminaria
                # Calcular salidas de caja iterando por tipo de caja
                # Cajas de paso intermedias (RIC N°4, art. 7.16.1.13): 1 cada 20m de
                # tramo, solo si el circuito tiene un tramo continuo >=20m
                # confirmado en el input (no basta con que el largo total
                # del circuito supere 20m)
                _tiene_20m_salidas = bool(r.get("_tiene_tramo_20m", True))  # True si el usuario confirmo que hay un tramo continuo de 20m o más
                cajas_paso_circ = int(L // 20) if _tiene_20m_salidas else 0  # 1 caja de paso por cada 20m de tramo confirmado
                # 'salidas' es el contador de salidas de caja conduit de este
                # circuito: arranca en 1 por la salida del tablero de origen, y
                # abajo se le suma lo que aporta cada caja del recorrido
                salidas = 1  # 1 salida tablero origen

                if "ilumin" in circ.lower():  # iluminación: las salidas se cuentan ambiente por ambiente
                    # Iluminación: iterar por ambiente y por interruptor
                    lum_by_amb_s = {}  # diccionario: ambiente -> lista de luminarias de ese ambiente en este circuito
                    for it in items:  # arma el diccionario de luminarias por ambiente
                        if not isinstance(it, dict): continue  # item raro, se salta
                        if "nombre" in it or "id_ench" in it or "modulos" in it or "n_ench" in it: continue  # salta enchufes y cargas especiales; aquí solo interesan luminarias
                        amb = str(it.get("amb", "")).strip().lower() or "sin_amb"  # nombre del ambiente, normalizado; si no tiene, usa "sin_amb"
                        if amb not in lum_by_amb_s: lum_by_amb_s[amb] = []  # crea la lista si es la primera luminaria de ese ambiente
                        lum_by_amb_s[amb].append(it)  # agrupa las luminarias por ambiente

                    n_amb_s = len(lum_by_amb_s)  # cuántos ambientes distintos hay en este circuito
                    # recorre cada ambiente del circuito, uno por uno
                    for i_amb_s, (amb_key_s, lums_s) in enumerate(lum_by_amb_s.items()):  # el orden importa: el último ambiente lleva una salida menos
                        es_ultimo_amb_s = (i_amb_s == n_amb_s - 1)  # ¿es el último ambiente del circuito?
                        n_lum_s = len(lums_s)  # luminarias de este ambiente
                        n_conm_s = int(min(n_lum_s, ncon_map_local.get(amb_key_s, 0)))  # cuántas conmutadas
                        n_rest_s = n_lum_s - n_conm_s  # el resto, no conmutadas
                        ints_s = []  # lista de interruptores que hay que instalar en este ambiente
                        if n_conm_s > 0: ints_s.append(('9/24', n_conm_s))  # agrupa las luminarias conmutadas en interruptores 9/24
                        if n_rest_s > 0:  # las que no son conmutadas se reparten en interruptores normales
                            c12s, c15s, c32s = descomponer_interruptores(n_rest_s)  # cuántos interruptores de 1, 2 y 3 puestos hacen falta para el resto
                            for _ in range(c32s): ints_s.append(('9/32', 3))  # interruptor de 3 puestos, controla 3 luminarias
                            for _ in range(c15s): ints_s.append(('9/15', 2))  # interruptor de 2 puestos, controla 2 luminarias
                            for _ in range(c12s): ints_s.append(('9/12', 1))  # interruptor de 1 puesto, controla 1 luminaria
                        n_ints_s = len(ints_s)  # total de interruptores que quedan en el ambiente

                        # recorre cada interruptor de este ambiente
                        for i_int_s, (tipo_s, n_lum_g_s) in enumerate(ints_s):  # cada interruptor lleva su caja troncal y su caja de interruptor
                            es_ultimo_int_s = (i_int_s == n_ints_s - 1)  # ¿es el último interruptor de este ambiente?

                            # Caja troncal: la cantidad de salidas depende de si es el último
                            # interruptor y/o el último ambiente del circuito
                            if n_ints_s == 1 and es_ultimo_amb_s:  # único interruptor y último ambiente: el troncal termina acá
                                salidas += 3  # entrada + salida hacia el interruptor + salida hacia la caja octogonal
                            elif n_ints_s == 1 and not es_ultimo_amb_s:  # único interruptor pero quedan ambientes: el troncal sigue
                                salidas += 4  # entrada + salida hacia el interruptor + salida hacia la caja octogonal + salida hacia el siguiente ambiente
                            elif es_ultimo_int_s and es_ultimo_amb_s:  # último interruptor del último ambiente: no sale nada más
                                salidas += 3  # entrada + salida hacia el interruptor + salida hacia la caja octogonal
                            elif es_ultimo_int_s and not es_ultimo_amb_s:  # último interruptor pero quedan ambientes: el troncal sigue
                                salidas += 4  # entrada + salida hacia el interruptor + salida hacia la caja octogonal + salida hacia el siguiente ambiente
                            else:  # interruptor del medio: el troncal sigue al siguiente
                                salidas += 4  # primera o intermedia: entrada + salida hacia el interruptor + salida hacia la caja octogonal + salida hacia el siguiente tramo troncal

                            # Caja interruptor: 9/24 son 2 cajas físicas (viajeros
                            # entre ambas), cada una con entrada+salida = 4 en
                            # total. Los demás (9/12, 9/15, 9/32) son 1 sola caja,
                            # el retorno comparte el mismo conduit = 1.
                            salidas += 4 if tipo_s == '9/24' else 1  # el 9/24 son 2 cajas por los viajeros, el resto 1 sola

                            # Cajas octogonales: 1 por luminaria de este grupo.
                            # 9/24: TODAS llevan 2 salidas (incluida la última,
                            # no aplica la regla de "última=1" de los demás
                            # tipos). Resto (9/12, 9/15, 9/32): +1 extra si no
                            # es la última (para seguir la cadena), la última
                            # solo lleva 1 (solo entrada).
                            for i_lum_s in range(n_lum_g_s):  # una caja octogonal por cada luminaria del grupo
                                es_ultima_lum_s = (i_lum_s == n_lum_g_s - 1)  # marca la última luminaria de la cadena
                                if tipo_s == '9/24':  # conmutadas: entra y sale siempre, incluso la última
                                    salidas += 2  # entrada + salida, siempre, incluida la última
                                elif es_ultima_lum_s:  # última de la cadena, solo entrada
                                    salidas += 1  # solo entrada
                                else:  # todavía quedan luminarias por delante
                                    salidas += 2  # entrada + salida hacia la siguiente caja octogonal

                else:  # circuitos de enchufes y cargas especiales
                    # Enchufes: intermedio=2, último=1. Si el circuito se
                    # ramifica (caja adicional entre ambientes), hay 1 último
                    # por cada ambiente, no solo 1 para todo el circuito.
                    # Horno, lavadora y otras cargas especiales también
                    # cuentan como 1 último, porque igual necesitan su caja
                    # de conexión al artefacto.
                    # cuenta el total de enchufes/cargas especiales del circuito:
                    # enchufes (con su posible multitoma) más cargas especiales (cuentan 1 cada una)
                    n_ench_total = sum(
                        int(it.get("n_ench", 1) or 1)  # cada punto puede traer más de un enchufe (multitoma)
                        for it in items  # recorre los puntos del circuito
                        if isinstance(it, dict) and ("id_ench" in it or "modulos" in it or "n_ench" in it)  # solo los que son enchufes
                    ) + sum(
                        1  # cada carga especial cuenta como 1
                        for it in items  # recorre otra vez los puntos
                        if isinstance(it, dict) and "nombre" in it  # las especiales se reconocen porque traen nombre
                    )
                    _cajas_adic_s = int((cajas_adic_por_nombre or {}).get(circ, 0))  # cuántas cajas adicionales (ramificaciones) tiene este circuito
                    if _cajas_adic_s > 0:  # circuito ramificado: cada ramal termina en su propio enchufe final
                        # el circuito se divide en un ramal por ambiente, y
                        # cada ramal termina en su propio último
                        _ambs_ench_s = set(
                            str(it.get("amb", "")).strip().lower()  # nombre del ambiente normalizado
                            for it in items  # recorre los puntos del circuito
                            if isinstance(it, dict) and ("id_ench" in it or "modulos" in it or "n_ench" in it or "nombre" in it)  # enchufes y cargas especiales, las luminarias no cuentan acá
                        )
                        n_ultimos_s = max(1, len(_ambs_ench_s))  # al menos 1 "último" por cada ambiente que tenga enchufes
                    else:  # circuito sin ramificar: un solo enchufe final
                        n_ultimos_s = 1 if n_ench_total > 0 else 0  # circuito sin ramificar: 1 solo "último" si tiene algún enchufe
                    n_intermedios_s = max(0, n_ench_total - n_ultimos_s)  # todos los enchufes menos los "últimos" son intermedios
                    salidas += 2 * n_intermedios_s + n_ultimos_s + 3 * _cajas_adic_s  # los intermedios llevan 2 salidas, los últimos 1, y cada caja adicional suma 3


                # Cajas de paso: 2 por caja (entrada + salida)
                salidas += 2 * cajas_paso_circ  # suma 2 salidas por cada caja de paso intermedia
                # Registra en materiales las salidas de caja conduit de este circuito (embutida)
                add_row(
                    desc=f"Salida de caja conduit de PVC de {mm}mm",  # salida de caja del mismo diámetro del conduit del circuito
                    marcas_txt=marcas.get("Salida de caja conduit", ""),  # marcas de salida de caja
                    norma="RIC 4.7.2",  # canalización, RIC 4.7.2
                    circuito=circ,  # circuito al que se carga
                    unidad="u",  # se cuentan por unidad
                    k=1,  # sin factor extra
                    longitud_m=f"{salidas} unid",  # cantidad en texto para el informe
                    cantidad=salidas  # total de salidas contadas arriba
                )
                continue  # ya se registró este circuito, se pasa a la siguiente canalización
            # si no calza como sobrepuesta ni como embutida, se cobra por metro lineal
            # Registra el conduit por metro lineal cuando la canalización no es sobrepuesta ni embutida
            add_row(
                desc=formatear_canalizacion_material(canal, es_sobrepuesta=False),  # descripción del conduit, sin el agregado de tramo
                marcas_txt=marcas.get("Conduit PVC (embutida)", ""),  # marcas de conduit
                norma="RIC 4.7.2",  # canalización, RIC 4.7.2
                circuito=circ,  # circuito al que se carga
                unidad="m",  # acá si se cobra por metro lineal
                k=1,  # sin factor extra
                longitud_m=math.ceil(L),  # metros redondeados hacia arriba
                cantidad=math.ceil(L)  # se compran los mismos metros que se calcularon
            )

    # ── Salidas de caja / accesorios de canaleta de CLIMATIZACIÓN ────────────
    # Acá se cierra lo que el bucle de arriba se saltó con 'continue' para los
    # circuitos de clima: si va embutido salen salidas de caja, y si va
    # sobrepuesto salen coplas y curvas. Los datos vienen de _clima_items,
    # llenado al principio de la función.
    for _cl in _clima_items:  # recorre cada tramo de climatización guardado antes
        if _cl["es_emb"]:  # ¿la canalización de este equipo es embutida (va dentro de conduit)?
            _n_salidas_cl = 2 + 2 * _cl.get("cajas_paso", 0)  # tablero + equipo, más 2 por cada caja de paso intermedia (>=20m)
            # Registra la salida de caja conduit del equipo de climatización embutido
            add_row(
                desc=f"Salida de caja conduit de PVC de {_cl['mm']}mm",  # descripción que va al listado, con el diámetro del conduit
                marcas_txt=marcas.get("Salida de caja conduit", ""),  # marcas sugeridas para salidas de caja conduit
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones embutidas
                circuito=_cl["circ"],  # circuito de climatización al que se le carga
                unidad="u",  # las salidas de caja se cuentan por unidad
                k=1,  # sin holgura extra, la cantidad va exacta
                longitud_m=f"{_n_salidas_cl} unid",  # detalle que se muestra en la columna de longitud
                cantidad=_n_salidas_cl  # cantidad final de salidas de caja
            )
        else:  # si no es embutida, es canaleta sobrepuesta
            # Canaleta sobrepuesta: coplas + curvas
            # la canaleta viene en tramos de 2m: n tramos se unen con n-1 coplas
            cant_copla_cl = max(0, int(math.ceil(_cl["L"] / 2.0)) - 1)  # 1 unión copla por cada unión entre tramos de 2m
            if cant_copla_cl > 0:  # solo se agregan coplas si el tramo mide más de 2m
                # Registra las uniones copla de la canaleta sobrepuesta
                add_row(
                    desc=f"Unión copla para canaleta de PVC {_extraer_medida_canaleta(_cl['canal'])}",  # descripción con la medida de la canaleta (20x10, 40x25, etc)
                    marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                    norma="RIC 4.7.2",  # mismo artículo del RIC de canalizaciones
                    circuito=_cl["circ"],  # circuito de climatización
                    unidad="u",  # las coplas se cuentan por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cant_copla_cl} unid",  # detalle para la columna de longitud
                    cantidad=cant_copla_cl  # total de coplas del tramo
                )
            # Curvas fijas de clima: se asumen 2 planas + 2 internas por equipo
            # (el quiebre para salir del tablero y el quiebre para llegar al
            # equipo). Es un criterio propio, no sale de una tabla del RIC.
            # Registra las curvas planas de la canaleta (2 fijas, para dar vuelta en las esquinas)
            add_row(
                desc=f"Curvas planas en 90° para canaletas de PVC {_extraer_medida_canaleta(_cl['canal'])}",  # curvas planas, para los quiebres del recorrido de la canaleta
                marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito=_cl["circ"],  # circuito de climatización
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m="2 unid",  # siempre son 2, no depende del largo del tramo
                cantidad=2  # 2 curvas planas fijas
            )
            # Registra las curvas internas de la canaleta (2 fijas)
            add_row(
                desc=f"Curvas internas en 90° para canaletas de PVC {_extraer_medida_canaleta(_cl['canal'])}",  # curvas internas, para los quiebres hacia adentro
                marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # mismas marcas de accesorios
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito=_cl["circ"],  # circuito de climatización
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m="2 unid",  # también quedan fijas en 2
                cantidad=2  # 2 curvas internas fijas
            )

    # ── Salidas de caja conduit de AGUA CALIENTE ─────────────────────────────
    # Mismo cierre que climatización: el bucle de canalizaciones se saltó estos
    # circuitos con 'continue' y acá se les cargan sus accesorios, usando los
    # datos guardados en _agua_items al principio de la función.
    # Embutida sin tab.ext: 2 salidas (tablero principal + equipo)
    # Embutida con tab.ext: 4 salidas (el circuito se parte en 2 tramos:
    #   tablero principal -> tablero externo, y tablero externo -> equipo)
    # + 2 por cada caja de paso intermedia (tramo continuo >=20m)
    # Sobrepuesta: coplas + curvas (igual que climatización)
    for _ac in _agua_items:  # recorre cada tramo de agua caliente guardado antes
        if _ac["es_emb"]:  # ¿la canalización de este equipo es embutida?
            # el nombre "bouq" viene de boquilla, pero lo que cuenta son SALIDAS
            # DE CAJA conduit (es la fila que se emite justo abajo)
            _n_bouq_ac = (4 if _ac["lleva_tab_ext"] else 2) + 2 * _ac.get("cajas_paso", 0)  # 4 salidas si lleva tablero externo (2 tramos), si no 2; más 2 por cada caja de paso
            # Registra la salida de caja conduit del equipo de agua caliente embutido
            add_row(
                desc=f"Salida de caja conduit de PVC de {_ac['mm']}mm",  # descripción con el diámetro del conduit del equipo de agua caliente
                marcas_txt=marcas.get("Salida de caja conduit", ""),  # marcas de salidas de caja conduit
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones embutidas
                circuito=_ac["circ"],  # circuito de agua caliente
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{_n_bouq_ac} unid",  # detalle para la columna de longitud
                cantidad=_n_bouq_ac  # 2 o 4 salidas según si hay tablero externo, más las cajas de paso
            )
        else:  # canaleta sobrepuesta
            # Canaleta sobrepuesta: coplas + curvas (igual que climatización)
            cant_copla_ac = max(0, int(math.ceil(_ac["L"] / 2.0)) - 1)  # misma fórmula de coplas que climatización
            if cant_copla_ac > 0:  # solo se agregan coplas si el tramo mide más de 2m
                # Registra las uniones copla de la canaleta sobrepuesta
                add_row(
                    desc=f"Unión copla para canaleta de PVC {_extraer_medida_canaleta(_ac['canal'])}",  # descripción con la medida de la canaleta del tramo
                    marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                    norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                    circuito=_ac["circ"],  # circuito de agua caliente
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cant_copla_ac} unid",  # detalle para la columna de longitud
                    cantidad=cant_copla_ac  # total de coplas del tramo
                )
            # Si hay tablero externo, el circuito se parte en 2 tramos (igual que
            # en el caso embutido, donde las salidas de caja se duplican de 2 a 4)
            _mult_ac = 2 if _ac["lleva_tab_ext"] else 1  # 2 tramos si lleva tablero externo, si no 1
            _curvas_plan_ac = 2 * _mult_ac  # 2 curvas planas por cada tramo
            _curvas_int_ac  = 2 * _mult_ac  # 2 curvas internas por cada tramo
            # Registra las curvas planas de la canaleta
            add_row(
                desc=f"Curvas planas en 90° para canaletas de PVC {_extraer_medida_canaleta(_ac['canal'])}",  # curvas planas de la canaleta del equipo de agua caliente
                marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito=_ac["circ"],  # circuito de agua caliente
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{_curvas_plan_ac} unid",  # detalle para la columna de longitud
                cantidad=_curvas_plan_ac  # 2 por tramo, se duplican si hay tablero externo
            )
            # Registra las curvas internas de la canaleta
            add_row(
                desc=f"Curvas internas en 90° para canaletas de PVC {_extraer_medida_canaleta(_ac['canal'])}",  # curvas internas de la canaleta
                marcas_txt=marcas.get("Accesorios canaleta PVC", ""),  # marcas de accesorios de canaleta
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito=_ac["circ"],  # circuito de agua caliente
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{_curvas_int_ac} unid",  # detalle para la columna de longitud
                cantidad=_curvas_int_ac  # 2 por tramo, se duplican si hay tablero externo
            )

    # =========================
    # 2) CONDUCTORES (POR CIRCUITO)
    # =========================
    add_section("Conductores")  # abre el bloque de materiales de conductores en el Excel
    if {"Conductor", "Longitud (m)", "Circuito"}.issubset(set(circuitos_df.columns)):  # sigue solo si la tabla de circuitos tiene las columnas necesarias
        circuitos_tmp = circuitos_df.copy()  # copia, para no modificar la tabla original de circuitos
        circuitos_tmp["Longitud (m)"] = pd.to_numeric(circuitos_tmp["Longitud (m)"], errors="coerce").fillna(0)  # convierte la columna a número; lo que no se puede convertir queda en 0
        # Recorre cada circuito de la instalación (una fila de la tabla por circuito)
        for _, r in circuitos_tmp.iterrows():
            cond = str(r.get("Conductor", "")).strip()  # tipo/calibre del conductor de este circuito
            circ = str(r.get("Circuito", "")).strip()  # nombre del circuito
            L_total = float(r.get("Longitud (m)", 0))  # largo total del circuito, en metros
            if (not cond) or (not circ) or L_total <= 0:  # si falta algún dato clave, no se puede calcular este circuito
                continue  # circuito sin datos válidos, se salta
            es_ilum    = ("ilumin" in circ.lower())  # ¿es un circuito de iluminación?
            es_especial_cond = any(k in circ.lower() for k in ("especial", "horno", "encimera", "lavadora",  # circuitos que son cargas especiales (electrodomésticos grandes)
                                   "lavavajilla", "climatiz", "aire", "split", "ac ", "a/c",  # sigue la lista de palabras clave de cargas especiales
                                   "ducha", "termo", "calefon", "calefón", "calentador", "agua caliente"))  # cierra la lista: ducha, termo, calefón y agua caliente también cuentan

            # --- Calcular chicotes según puntos de conexión (RIC) ---
            items_circ = (items_por_nombre or {}).get(circ, [])  # lista de artefactos (enchufes, luminarias, etc.) conectados a este circuito
            if not isinstance(items_circ, list):  # por seguridad: si no vino una lista, se usa una vacía
                items_circ = []  # queda vacío, así el circuito no suma chicotes por artefactos
            # Contar enchufes e interruptores del circuito
            n_ench_cond  = 0  # contador de enchufes/módulos del circuito
            for it in items_circ:  # recorre cada artefacto del circuito
                if not isinstance(it, dict):  # descarta datos mal formados
                    continue  # ítem mal formado, se ignora
                if "n_ench" in it or "id_ench" in it or "modulos" in it:  # ¿este ítem es un enchufe o módulo?
                    n_ench_cond += int(it.get("n_ench", 1) or 1)  # suma la cantidad de enchufes de este ítem

            if es_ilum:  # si es iluminación, los puntos de conexión se cuentan distinto (por ambiente e interruptores)
                # OJO: en iluminación este n_puntos NO llega al resultado final.
                # Sirve solo de referencia/checkeo: más abajo, el largo de cada
                # color se recalcula con chic_R y chic_N (bloque "Calcular
                # chicotes por color"), y L_con_chicotes se descarta.
                # ── Cálculo exacto de chicotes por punto de conexión ──────────────
                # Lógica validada ejemplo a ejemplo:
                #
                # CAJA RECTANGULAR por ambiente:
                #   - Intermedio: 3 entrada + 3 salida + 1 fase hacia el interruptor + 1N hacia la luminaria + 1T hacia la luminaria = 9
                #   - Último:     3 entrada + 1 fase hacia el interruptor + 1N hacia la luminaria + 1T hacia la luminaria             = 6
                #
                # CAJA INTERRUPTOR (por tipo):
                #   - 9/12: 1 fase + 1 retorno                = 2
                #   - 9/15: 1 fase + 2 retornos               = 3
                #   - 9/32: 1 fase + 3 retornos               = 4
                #   - 9/24 primer int: 1 fase + 2 viajeros    = 3
                #   - 9/24 segundo int: 2 viajeros + 1 retorno = 3
                #
                # CAJA OCTOGONAL (no conmutada):
                #   - Intermedia: 1 retorno + 3N + 3T = 7
                #   - Última:     1 retorno + 1N + 1T = 3
                #
                # CAJA OCTOGONAL (conmutada 9/24):
                #   - Intermedia: 3F + 3N + 3T = 9
                #   - Última:     1F + 1N + 1T = 3

                # Obtener luminarias por ambiente y conmutadas
                # ncon_map_local ya se construyó antes del bucle principal (ver arriba)
                lum_by_amb_cond = {}  # diccionario: ambiente -> cantidad de luminarias
                for it in items_circ:  # recorre los artefactos del circuito
                    if not isinstance(it, dict): continue  # descarta datos mal formados
                    if "n_ench" in it or "id_ench" in it or "modulos" in it or "nombre" in it:  # se saltan enchufes y cargas especiales, solo interesan las luminarias
                        continue  # no es luminaria, se salta
                    amb = str(it.get("amb", "")).strip() or "sin_amb"  # nombre del ambiente (pieza) de esta luminaria
                    key = amb.lower()  # normaliza a minúsculas para agrupar bien por ambiente
                    lum_by_amb_cond[key] = lum_by_amb_cond.get(key, 0) + 1  # cuenta las luminarias de este ambiente

                n_amb_cond = len(lum_by_amb_cond)  # cantidad de ambientes distintos que tiene el circuito
                n_puntos = 0  # total de chicotes (puntos de conexión) del circuito

                for i_amb, (amb_key, n_lum_amb) in enumerate(lum_by_amb_cond.items()):  # recorre cada ambiente sabiendo su posición, para saber si es el último
                    es_ultimo_amb = (i_amb == n_amb_cond - 1)  # ¿es el último ambiente del circuito?
                    n_conm_amb = int(min(n_lum_amb, ncon_map_local.get(amb_key, 0)))  # cuántas conmutadas tiene
                    n_rest_amb = int(n_lum_amb) - n_conm_amb  # el resto, no conmutadas

                    # Caja rectangular del ambiente
                    if es_ultimo_amb:  # el último ambiente no tiene salida hacia otro ambiente
                        n_puntos += 6   # 3 entrada + 1F hacia el interruptor + 1N hacia la luminaria + 1T hacia la luminaria
                    else:  # ambiente intermedio, el cable sigue hacia el siguiente
                        n_puntos += 9   # 6 + 3 salida hacia siguiente ambiente

                    # Cajas de interruptores del ambiente
                    if n_conm_amb > 0:  # conmutadas: van con un par de interruptores 9/24 (tipo escalera, se manda la luz desde 2 puntos)
                        # 9/24: siempre 1 PAR de interruptores por ambiente (no uno por luminaria)
                        # primer interruptor: 1 fase + 2 viajeros = 3 chicotes
                        # segundo interruptor: 2 viajeros + 1 retorno = 3 chicotes
                        n_puntos += 6  # 1 par × (3+3)
                    if n_rest_amb > 0:  # el resto de luminarias (no conmutadas) usa interruptores simples
                        c12, c15, c32 = descomponer_interruptores(n_rest_amb)  # separa la cantidad en combinaciones de interruptores 9/12, 9/15 y 9/32
                        n_puntos += c12 * 2   # 9/12: 1 fase + 1 retorno
                        n_puntos += c15 * 3   # 9/15: 1 fase + 2 retornos
                        n_puntos += c32 * 4   # 9/32: 1 fase + 3 retornos

                    # Cajas octogonales del ambiente
                    for i_lum in range(int(n_lum_amb)):  # recorre cada luminaria del ambiente
                        es_ultima_lum = (i_lum == int(n_lum_amb) - 1)  # ¿es la última luminaria del ambiente?
                        es_conm = (i_lum < n_conm_amb)  # ¿esta luminaria en particular está conmutada?
                        if es_ultima_lum:  # de la última luminaria no sale cable hacia otra caja
                            n_puntos += 3   # 1F + 1N + 1T (última siempre 3)
                        else:  # luminaria intermedia, el cable sigue a la siguiente octogonal
                            if es_conm:  # conmutada 9/24: pasan las 3 fases del par de interruptores
                                n_puntos += 9   # 3F + 3N + 3T (conmutada intermedia)
                            else:  # no conmutada: solo el retorno, más neutro y tierra
                                n_puntos += 7   # 1 retorno + 3N + 3T (no conmutada intermedia)

            elif es_especial_cond:  # circuito de carga especial (horno, climatización, agua caliente, etc.)
                # Recordatorio: n_puntos cuenta los chicotes del circuito.
                # Un chicote es el pedazo de cable (15cm) que se pierde cada
                # vez que el conductor se corta para conectarlo en una caja;
                # más adelante n_puntos*0.15 se suma a los metros de cable
                # que hay que comprar.
                # Especiales, clima, agua caliente: 1 chicote base (caja
                # junto al equipo) + 2 por cada caja de paso (entrada y
                # salida), solo si hay un tramo continuo >=20m
                _canal_circ_esp = str(r.get("Canalización", "")).strip().lower()  # tipo de canalización de este circuito, en minúsculas
                _es_emb_circ_esp = ("conduit" in _canal_circ_esp) or ("embutida" in _canal_circ_esp)  # ¿la canalización es embutida (conduit)?
                _tiene_20m_circ_esp = bool(r.get("_tiene_tramo_20m", True))  # ¿el circuito tiene confirmado un tramo continuo de 20m o más?
                n_cajas_paso_esp = int(L_total // 20) if (_es_emb_circ_esp and _tiene_20m_circ_esp) else 0  # 1 caja de paso cada 20m, solo si el circuito califica
                n_puntos = 1 + n_cajas_paso_esp * 2  # 1 chicote base junto al equipo, más 2 por cada caja de paso
            else:  # circuito de enchufes/tomas normales
                # Enchufes generales, por ambiente: intermedio=3 (entrada,
                # salida, al artefacto), último de CADA ambiente=1 (no solo
                # el último de todo el circuito). Caja adicional entre
                # ambientes=3 (entrada, salida 1, salida 2). Caja de paso=2
                # (entrada, salida), solo si hay un tramo continuo >=20m.
                n_cajas_adic_circ = int((cajas_adic_por_nombre or {}).get(circ, 0))  # cantidad de cajas adicionales entre ambientes de este circuito
                _canal_circ = str(r.get("Canalización", "")).strip().lower()  # tipo de canalización, en minúsculas
                _es_emb_circ = ("conduit" in _canal_circ) or ("embutida" in _canal_circ)  # ¿es canalización embutida?
                _tiene_20m_circ = bool(r.get("_tiene_tramo_20m", True))  # ¿tiene confirmado un tramo continuo de 20m o más?
                n_cajas_paso_circ = int(L_total // 20) if (_es_emb_circ and _tiene_20m_circ) else 0  # 1 caja de paso cada 20m, si el circuito califica
                _n_ench_por_amb = {}  # diccionario: ambiente -> cantidad de enchufes
                for it in items_circ:  # recorre los artefactos del circuito
                    if not isinstance(it, dict): continue  # descarta datos mal formados
                    if "n_ench" in it or "id_ench" in it or "modulos" in it:  # solo interesan los enchufes/módulos
                        _a = str(it.get("amb", "")).strip().lower() or "sin_amb"  # ambiente de este enchufe
                        _n_ench_por_amb[_a] = _n_ench_por_amb.get(_a, 0) + int(it.get("n_ench", 1) or 1)  # suma la cantidad de enchufes en ese ambiente
                if _n_ench_por_amb:  # si hay enchufes agrupados por ambiente...
                    n_puntos = sum((n_amb_ench - 1) * 3 + 1 for n_amb_ench in _n_ench_por_amb.values() if n_amb_ench > 0)  # por cada ambiente: los intermedios cuentan 3, y se suma 1 por el último de ese ambiente
                else:  # si no hay datos por ambiente, se calcula con el total del circuito
                    n = int(max(1, n_ench_cond))  # al menos 1 enchufe, para no restar de más
                    n_puntos = (n - 1) * 3 + 1  # todos intermedios cuentan 3, y se suma 1 por el último
                n_puntos += n_cajas_adic_circ * 3 + n_cajas_paso_circ * 2  # suma las cajas adicionales (3 c/u) y las cajas de paso (2 c/u)
            # 0,15 = 15cm de cable que se pierde en cada chicote (criterio propio
            # de obra, no es un valor del RIC). OJO: chicotes y L_con_chicotes
            # solo se usan en la rama "no iluminación" de acá abajo; en
            # iluminación el largo sale de chic_R / chic_N y estos dos se botan.
            chicotes = n_puntos * 0.15  # 15cm de cable por chicote
            L_con_chicotes = L_total + chicotes  # longitud real del circuito + lo que se pierde en los chicotes

            # --- NO iluminación: 3 conductores iguales ---
            if not es_ilum:  # enchufes y cargas especiales: los 3 conductores van del mismo largo
                # Los conductores se venden por metro entero, así que el largo se
                # redondea siempre hacia arriba (nunca hacia abajo)
                # +10% extra de holgura en los tres conductores (rojo, blanco, verde)
                Lr = math.ceil(L_con_chicotes * 1.10)  # metros de conductor rojo (fase)
                Lb = math.ceil(L_con_chicotes * 1.10)  # metros de conductor blanco (neutro)
                Lv = math.ceil(L_con_chicotes * 1.10)  # metros de conductor verde (tierra)
                total = Lr + Lb + Lv  # metros totales de cable a comprar en este circuito
                desc = f"Conductor {cond.replace('.', ',')} (Rojo = {Lr} m, Blanco = {Lb} m, Verde = {Lv} m)"  # texto de la partida, con el desglose por color
                # detecta si el circuito es de climatización o de agua caliente, por
                # palabras clave en su nombre, para usar una norma distinta
                _es_clima_cond = any(k in circ.lower() for k in ("climatiz", "aire", "split", "ac ", "a/c"))  # ¿el nombre del circuito dice que es climatización?
                _es_agua_cond  = any(k in circ.lower() for k in ("ducha", "termo", "calefon", "calefón", "calentador", "agua caliente"))  # ¿el nombre dice ducha, termo, calefón o parecido?
                _norma_cond = (  # clima y agua caliente citan más artículos del RIC que un circuito común
                    "RIC 7 (7.3.4, 7.5.1, 7.5.2) / RIC 3 (5.1.3) / RIC 4 (5.2, 5.4, 5.5, 5.34, 6.2.5)"  # norma ampliada, con los artículos que piden clima y agua caliente
                    if (_es_clima_cond or _es_agua_cond)  # condición: el circuito es de climatización o de agua caliente
                    else "RIC 4.1.1"  # circuito común: solo RIC 4.1.1
                )
                # Desglose "valor crudo (x1,10) = resultado", en una sola
                # línea porque en no-iluminación Rojo=Blanco=Verde siempre
                # son iguales (no hace falta repetir 3 veces lo mismo)
                _crudo_txt = f"{L_con_chicotes:.4f}".replace(".", ",")  # largo con chicotes antes del 10%, escrito con coma decimal
                _long_txt = f"Rojo=Blanco=Verde {_crudo_txt} (x1,10) = {Lr}m"  # desglose del cálculo, para poder revisar de dónde salió el número
                # agrega la fila de material para el conductor de este circuito
                add_row(
                    desc=desc,  # descripción armada arriba, con los metros de cada color
                    marcas_txt=marcas.get("Conductores", ""),  # marcas de conductores
                    norma=_norma_cond,  # norma según si es carga normal, clima o agua caliente
                    circuito=circ,  # circuito al que pertenece el conductor
                    unidad="m",  # el cable se compra por metro
                    k=1.10,  # solo deja constancia en la columna "K" de que ya se aplicó el 10% de holgura arriba
                    longitud_m=_long_txt,  # desglose visible en el informe
                    cantidad=total  # metros totales sumando los tres colores
                )
                continue  # este circuito ya quedó cargado, se pasa al siguiente

            #  Iluminación: calcular Lr, Lb, Lv con chicotes por color
            items = (items_por_nombre or {}).get(circ, [])  # artefactos del circuito de iluminación
            if not isinstance(items, list):  # por seguridad, si no vino una lista
                items = []  # queda vacío

            # Obtener ambientes del circuito
            ambs_circ = {}  # {amb_key: n_lum}
            for it in items:  # recorre los artefactos para agrupar luminarias por ambiente
                if not isinstance(it, dict): continue  # descarta datos mal formados
                if "n_ench" in it or "id_ench" in it or "modulos" in it or "nombre" in it: continue  # se saltan enchufes, módulos y cargas especiales
                a = str(it.get("amb", "")).strip().lower()  # ambiente de la luminaria, en minúsculas para agrupar bien
                if a: ambs_circ[a] = ambs_circ.get(a, 0) + 1  # cuenta las luminarias de este ambiente

            # ── Calcular chicotes por color ──────────────────────────────────
            # Lógica: recorre uno por uno los interruptores (= caja troncal) dentro de cada ambiente
            #
            # CAJA TRONCAL (1 por interruptor):
            #   n_ints==1, amb interm:  chic_R+=3  chic_N+=3
            #   n_ints==1, último amb:  chic_R+=2  chic_N+=2
            #   primera o intermedia:   chic_R+=3  chic_N+=3  (va hacia el siguiente tramo troncal)
            #   última, amb interm:     chic_R+=3  chic_N+=3  (va hacia el siguiente ambiente)
            #   última, último amb:     chic_R+=2  chic_N+=2
            #
            # CAJA INTERRUPTOR (solo rojo):
            #   9/24 suma 6 · 9/32 suma 4 · 9/15 suma 3 · 9/12 suma 2
            #
            # CAJA OCTOGONAL:
            #   Última: chic_R+=1 chic_N+=1
            #   Interm conm 9/24: chic_R+=3 chic_N+=3
            #   Interm no conm: chic_R+=1 chic_N+=3
            chic_R = 0  # total de chicotes de rojo (fase) del circuito
            chic_N = 0  # chicotes del neutro; la tierra usa el mismo número (por eso al final Lv = Lb)

            n_amb_circ = len(ambs_circ)  # cuántos ambientes distintos toca el circuito
            for i_amb, (amb_key, n_lum_amb) in enumerate(ambs_circ.items()):  # recorre ambiente por ambiente en orden, para saber cuál es el último
                # cuántas luminarias de este ambiente van con interruptor conmutado (9/24)
                n_conm_amb = int(min(n_lum_amb, ncon_map_local.get(amb_key, 0)))  # conmutadas del ambiente, nunca más que las luminarias que hay
                n_rest_amb = n_lum_amb - n_conm_amb  # el resto, con interruptor simple
                es_ultimo_amb = (i_amb == n_amb_circ - 1)  # ¿es el último ambiente del circuito?

                # Lista de interruptores del ambiente
                ints_amb = []  # interruptores del ambiente: (tipo, cuántas luminarias manda cada uno)
                if n_conm_amb > 0:  # si hay conmutadas, van todas con un par de 9/24
                    ints_amb.append(('9/24', n_conm_amb))  # un solo par de 9/24 se hace cargo de todas las conmutadas del ambiente
                if n_rest_amb > 0:  # las luminarias que sobran van con interruptores simples
                    c12, c15, c32 = descomponer_interruptores(n_rest_amb)  # reparte el resto en 9/12, 9/15 y 9/32
                    for _ in range(c32): ints_amb.append(('9/32', 3))  # cada 9/32 manda 3 luminarias
                    for _ in range(c15): ints_amb.append(('9/15', 2))  # cada 9/15 manda 2
                    for _ in range(c12): ints_amb.append(('9/12', 1))  # cada 9/12 manda 1
                n_ints = len(ints_amb)  # total de interruptores, o sea cuántas cajas troncales tiene el ambiente

                # Cajas troncales (1 por interruptor)
                for i_int, (tipo, n_lum_grupo) in enumerate(ints_amb):  # una caja troncal por cada interruptor del ambiente
                    es_primer = (i_int == 0)  # ¿es la primera troncal del ambiente?
                    es_ultimo = (i_int == n_ints - 1)  # ¿es la última troncal del ambiente?
                    if n_ints == 1:  # ambiente con un solo interruptor
                        # único interruptor del ambiente
                        if not es_ultimo_amb:  # todavía queda otro ambiente por delante
                            # entrada + salida hacia el interruptor + salida hacia el siguiente ambiente
                            chic_R += 3; chic_N += 3
                        else:  # no hay más ambientes después de este
                            # último ambiente: entrada + salida hacia el interruptor (no hay a dónde más salir)
                            chic_R += 2; chic_N += 2
                    elif es_primer:  # primera troncal cuando el ambiente tiene varios interruptores
                        # Primera troncal: entrada + salida hacia el interruptor + salida hacia el siguiente tramo troncal
                        chic_R += 3; chic_N += 3
                    elif not es_ultimo:  # troncales del medio
                        # Troncal intermedia: entrada + salida hacia el interruptor + salida hacia el siguiente tramo troncal
                        chic_R += 3; chic_N += 3
                    else:  # última troncal del ambiente
                        # Última troncal del ambiente
                        if not es_ultimo_amb:  # queda otro ambiente por delante
                            # Amb intermedio: entrada + salida hacia el interruptor + salida hacia el siguiente ambiente
                            chic_R += 3; chic_N += 3
                        else:  # acá termina el circuito
                            # Último ambiente: entrada + salida hacia el interruptor
                            chic_R += 2; chic_N += 2

                # Cajas interruptores (solo rojo, porque acá solo llega la fase)
                for tipo, n_lum_grupo in ints_amb:  # ahora las cajas de los interruptores propiamente tales
                    # 9/24 (conmutado): son 6 chicotes en total, 3 en el primer
                    # interruptor (1 fase + 2 viajeros) y 3 en el segundo
                    # (2 viajeros + 1 retorno)
                    if tipo == '9/24':   chic_R += 6
                    # 9/32: son 4 chicotes, 1 de entrada + 3 retornos (1 por cada luminaria del grupo)
                    elif tipo == '9/32': chic_R += 4
                    # 9/15: son 3 chicotes, 1 de entrada + 2 retornos
                    elif tipo == '9/15': chic_R += 3
                    # 9/12: son 2 chicotes, 1 de entrada + 1 retorno
                    elif tipo == '9/12': chic_R += 2

                # Cajas octogonales (donde va cada luminaria)
                for tipo, n_lum_grupo in ints_amb:  # y por último las octogonales de cada grupo de luminarias
                    for i_lum in range(n_lum_grupo):  # recorre las luminarias que cuelgan de ese interruptor
                        es_ultima_lum = (i_lum == n_lum_grupo - 1)  # ¿es la última del grupo?
                        es_conm = (tipo == '9/24')  # ¿el grupo va con par conmutado 9/24?
                        if es_ultima_lum:  # de la última no sale cable hacia otra octogonal
                            # última luminaria del grupo: solo llega 1 chicote (el retorno desde el interruptor)
                            chic_R += 1; chic_N += 1
                        else:  # luminaria intermedia
                            # luminaria intermedia:
                            # si es conmutada (9/24): son 3 chicotes, 1 que llega, 1 que va a la
                            #   luminaria y 1 que sigue a la siguiente caja octogonal
                            # si NO es conmutada: es 1 chicote, el que llega del retorno
                            chic_R += 3 if es_conm else 1
                            chic_N += 3  # el neutro/tierra siempre pone 3 chicotes en la octogonal intermedia, sea conmutada o no

            # Caja de paso (1 cada 20m del circuito, solo canalización
            # embutida, y solo si el circuito tiene un tramo continuo >=20m
            # confirmado en el input): 2 chicotes rojo + 2 neutro/tierra por
            # caja. Se cuenta por circuito completo (L_total), no por ambiente.
            _canal_circ_ilum = str(r.get("Canalización", "")).strip().lower()  # tipo de canalización del circuito de iluminación
            _es_emb_circ_ilum = ("conduit" in _canal_circ_ilum) or ("embutida" in _canal_circ_ilum)  # ¿va embutida en conduit?
            _tiene_20m_circ_ilum = bool(r.get("_tiene_tramo_20m", True))  # ¿se confirmó un tramo continuo de 20m o más?
            if _es_emb_circ_ilum and _tiene_20m_circ_ilum:  # solo ahí corresponde poner cajas de paso
                n_cajas_paso_ilum = int(L_total // 20)  # 1 caja de paso cada 20m de circuito
                chic_R += n_cajas_paso_ilum * 2  # 2 chicotes de rojo por caja, el que entra y el que sale
                chic_N += n_cajas_paso_ilum * 2  # lo mismo para neutro y tierra

            # ── Calcular extra longitud rojo vs neutro/tierra ────────────────
            # Rojo recorre tramos adicionales al interruptor y viajeros
            # Neutro/tierra no van al interruptor, así que hay que restar esos tramos
            # Si hay múltiples interruptores en el mismo ambiente, se multiplica
            # por n_interruptores (todos usan el mismo L_ida)
            # El rojo (fase) pasa por cada interruptor y vuelve (ida y
            # vuelta), por eso recorre más metros que la longitud medida
            # del circuito. El neutro y la tierra van directo a la
            # luminaria sin pasar por el interruptor, por eso recorren
            # menos. extra_R y extra_NT corrigen esa diferencia.
            extra_R = 0.0   # metros extra que recorre el rojo (a SUMAR)
            extra_NT = 0.0  # metros a RESTAR del neutro/tierra

            for amb_key in ambs_circ:  # recorre los ambientes para sacar las distancias al interruptor
                amb_row = _get_amb_row(ambientes_df, amb_key)  # busca los datos de distancias de este ambiente
                if amb_row is None: continue  # ambiente sin fila de distancias, no se puede corregir el largo
                ncon = int(pd.to_numeric(amb_row.get("N_conmutadas_924 (u)", 0), errors="coerce") or 0)  # cuántas luminarias del ambiente van conmutadas con el par 9/24
                n_lum_amb = ambs_circ[amb_key]  # total de luminarias de este ambiente dentro del circuito
                n_rest_amb = n_lum_amb - ncon  # las que quedan con interruptor normal, sin conmutar

                if ncon > 0:  # solo si el ambiente tiene conmutadas hay que corregir los largos
                    # 9/24: rojo suma L_via, resta L_troncal_primera_oct (ese
                    # tramo va directo de la caja a la luminaria, el rojo no
                    # lo recorre porque se desvía por los interruptores) · neutro/tierra
                    # resta L_fas + L_via + L_ret
                    # L_via: entre los 2 interruptores
                    L_via = float(pd.to_numeric(amb_row.get("L_viajeros_924 (m)", 0.0), errors="coerce") or 0.0)
                    # L_ret: del interruptor a la lámpara
                    L_ret = float(pd.to_numeric(amb_row.get("L_retorno_lampara (m)", 0.0), errors="coerce") or 0.0)
                    # L_fas: de la caja al primer interruptor
                    L_fas = float(pd.to_numeric(amb_row.get("L_fase_caja_primer_int (m)", 0.0), errors="coerce") or 0.0)
                    # L_tr_oct: de la caja troncal a la primera caja octogonal (1 vez por ambiente)
                    L_tr_oct = float(pd.to_numeric(amb_row.get("L_troncal_primera_oct_924 (m)", 0.0), errors="coerce") or 0.0)
                    extra_R  += L_via - L_tr_oct  # el rojo suma los viajeros y descuenta el tramo troncal-octogonal
                    extra_NT += L_fas + L_via + L_ret  # blanco y verde no pasan por los interruptores: estos tramos se acumulan para restarlos después

                if n_rest_amb > 0:  # ahora las luminarias que van con interruptor normal
                    # Un ambiente puede tener más de 1 interruptor (por
                    # ejemplo un 9/32 para las luminarias del centro y un
                    # 9/12 para un foco aparte), y cada interruptor puede
                    # estar a una distancia distinta de la caja troncal.
                    # "Grupos_interruptor_ilum" guarda esa lista: 1 grupo
                    # por cada interruptor del ambiente, con sus propias
                    # distancias (L_ida, L_tr_oct1, L_o1_o2).
                    grupos_amb = amb_row.get("Grupos_interruptor_ilum", None)  # lista de grupos de interruptor del ambiente, si viene el detalle
                    if isinstance(grupos_amb, list) and len(grupos_amb) > 0:  # solo sirve si la lista existe y trae al menos un grupo
                        # Se calcula interruptor por interruptor, cada uno
                        # con su propia distancia. Si en cambio se usara un
                        # solo promedio para todo el ambiente, el resultado
                        # queda inflado cuando las distancias son distintas
                        # entre sí — por eso se recorre grupo por grupo.
                        for _g in grupos_amb:  # un grupo = un interruptor de este ambiente
                            _tipo_g  = _g.get("tipo", "")       # 9/12, 9/15 o 9/32
                            _L_ida_g = float(_g.get("L_ida", 0.0) or 0.0)       # distancia de la caja troncal a este interruptor
                            _L_tr1_g = float(_g.get("L_tr_oct1", 0.0) or 0.0)   # distancia de la caja troncal a la primera luminaria de este grupo
                            _L_o12_g = float(_g.get("L_o1_o2", 0.0) or 0.0)     # entre la 1ra y 2da luminaria de este grupo
                            if _tipo_g == '9/32':  # interruptor triple
                                # 3 luminarias, entonces 3 retornos (una vuelta por cada una)
                                # + el tramo hasta la 1ra luminaria y hasta la 2da
                                extra_R += _L_ida_g * 3 + (_L_tr1_g + (_L_tr1_g + _L_o12_g))
                            elif _tipo_g == '9/15':  # interruptor doble
                                # 2 luminarias, entonces 2 retornos + el tramo hasta la 1ra
                                # (L_o1_o2 no aplica en este grupo, son solo 2 luminarias)
                                extra_R += _L_ida_g * 2 + _L_tr1_g
                            elif _tipo_g == '9/12':  # interruptor simple
                                # 1 sola luminaria, entonces 1 retorno, sin tramo extra
                                extra_R += _L_ida_g * 1
                            # el neutro y la tierra no pasan por el interruptor,
                            # así que solo se restan el tramo de ida hasta él
                            # (no van y vuelven como el rojo)
                            extra_NT += _L_ida_g
                    else:  # el ambiente no trajo el detalle por grupo de interruptor
                        # si no viene el detalle por grupo, se usan los totales
                        # agregados del ambiente. Exacto solo si el ambiente
                        # tiene 1 único interruptor.
                        L_ida     = float(pd.to_numeric(amb_row.get("L_caja_int_fase_ida (m)", 0.0), errors="coerce") or 0.0)  # distancia caja troncal - interruptor, promediada para todo el ambiente
                        L_tr_oct1 = float(pd.to_numeric(amb_row.get("L_troncal_oct1 (m)", 0.0), errors="coerce") or 0.0)  # de la caja troncal a la primera octogonal
                        L_o1_o2   = float(pd.to_numeric(amb_row.get("L_oct1_oct2 (m)", 0.0), errors="coerce") or 0.0)  # entre la primera y la segunda octogonal
                        c12, c15, c32 = descomponer_interruptores(n_rest_amb)  # reparte las luminarias en grupos 9/12, 9/15 y 9/32
                        n_retornos = c32 * 3 + c15 * 2 + c12 * 1  # un retorno por cada luminaria controlada
                        extra_R  += L_ida * n_retornos  # el rojo va y vuelve por cada retorno
                        extra_R  += c32 * (L_tr_oct1 + (L_tr_oct1 + L_o1_o2))  # los 9/32 además suman el tramo hasta la 1ra y hasta la 2da octogonal
                        extra_R  += c15 * L_tr_oct1  # el 9/15 solo suma el tramo hasta la primera octogonal
                        extra_NT += L_ida  # el blanco y el verde no llegan al interruptor: se acumula ese tramo para restarlo después

            # ── Longitudes finales por color ─────────────────────────────────
            # Lr = ceil((L_total + chicotes_R×0.15 + extra_R) × 1.10)
            # Lb = Lv = ceil((L_total + chicotes_N×0.15 - extra_NT) × 1.10)
            _crudo_r  = L_total + chic_R * 0.15 + extra_R  # largo crudo del rojo, antes de la holgura
            # si las distancias del ambiente vienen mal cargadas, extra_NT puede
            # superar al largo del circuito: el max(0.0, ...) evita que salga un
            # metraje negativo de blanco/verde
            _crudo_bv = max(0.0, L_total + chic_N * 0.15 - extra_NT)  # nunca negativo
            Lr = math.ceil(_crudo_r  * 1.10)  # redondea hacia arriba, se compra el metro completo
            Lb = math.ceil(_crudo_bv * 1.10)  # mismo redondeo hacia arriba para el blanco
            Lv = Lb  # blanco y verde siempre miden lo mismo
            total = Lr + Lb + Lv  # metros totales de este calibre para el circuito, sumando los 3 colores
            desc = f"Conductor {cond.replace('.', ',')} (Rojo = {Lr} m, Blanco = {Lb} m, Verde = {Lv} m)"  # texto de la columna descripción, con coma decimal como se usa acá
            _crudo_r_txt  = f"{_crudo_r:.4f}".replace(".", ",")  # el crudo con 4 decimales y coma
            _crudo_bv_txt = f"{_crudo_bv:.4f}".replace(".", ",")  # mismo formato para blanco y verde
            # detalle del cálculo que se muestra en la columna de longitud,
            # así queda a la vista de donde salió cada metro
            _long_txt = (
                f"Rojo {_crudo_r_txt} (x1,10) = {Lr}m / "
                f"Blanco {_crudo_bv_txt} (x1,10) = {Lb}m / "
                f"Verde {_crudo_bv_txt} (x1,10) = {Lv}m"
            )
            # agrega la fila de material para el conductor de este circuito de iluminación
            add_row(
                desc=desc,
                marcas_txt=marcas.get("Conductores", ""),  # marcas de conductor que se ofrecen en el informe
                norma="RIC 4.1.1",  # artículo del RIC que aplica a conductores
                circuito=circ,  # circuito al que se carga este material
                unidad="m",  # el conductor se vende por metro
                k=1.10,  # holgura del 10% que ya se aplico arriba
                longitud_m=_long_txt,
                cantidad=total  # metros a comprar
            )

    # =========================
    # 3) PROTECCIONES
    # =========================
    add_section("Protecciones")  # arranca la sección "Protecciones" en la hoja de materiales
    # Interruptor general omnipolar: la protección principal del tablero, siempre 1 unidad
    add_row(
        desc=f"Interruptor general omnipolar {texto_omni}",  # texto_omni viene armado antes, trae polos y amperaje
        marcas_txt=marcas.get("Protecciones", ""),
        norma="RIC 6.4.1",  # artículo del RIC de protecciones
        circuito="General",
        unidad="u",
        k=1,
        longitud_m="1 unid",
        cantidad=1  # siempre 1, es la protección de cabecera del tablero
    )
    # Supresor de transiente (SPD)
    add_row(
        desc="Supresor de transiente Tipo 2 (Clase II), 2P 230 V 20kA, Up ≤ 2.5kV",  # spd tipo 2 contra descargas, va en el tablero
        marcas_txt=marcas.get("Supresor de transiente", marcas.get("Protecciones", "")),
        norma="RIC 5 (8.7.7) / RIC 6 (6.2.2)",  # los artículos del RIC que lo exigen
        circuito="General",
        unidad="u",
        k=1,
        longitud_m="1 unid",
        cantidad=1  # 1 por tablero
    )
    # Protector sobrevoltaje y corriente ajustable
    add_row(
        desc="Protector sobrevoltaje y corriente ajustable 63A 220V",  # protege ante sobrevoltaje y sobrecorriente de la red
        marcas_txt=marcas.get("Protector sobrevoltaje", ""),
        norma="RIC 6 (6.6.2)",
        circuito="General",
        unidad="u",
        k=1,
        longitud_m="1 unid",
        cantidad=1  # 1 por tablero
    )
    # DIFERENCIALES: 1 por grupo
    if group_info and "Circuito" in circuitos_df.columns:  # solo si ya hay grupos armados y existe la columna Circuito
        for gid, meta in group_info.items():  # un diferencial por cada grupo de circuitos
            sens     = meta.get("sensibilidad_dif", "30mA")  # corriente de fuga que hace saltar el diferencial: 30mA salvo el equipo de agua caliente que quedó en Volumen 1 del baño, que va con 10mA
            dif_txt  = f"2X{meta['dif']} {sens} / Tipo A"  # texto del diferencial, ej "2X40 30mA / Tipo A" (el calibre va sin la "A")
            circuitos_grupo = []  # nombres de los circuitos que caen bajo este diferencial
            for idx in meta["indices"]:  # recorre los índices de circuito que pertenecen a este grupo
                if 0 <= idx < len(circuitos_df):  # se protege por si el índice quedó fuera de rango
                    circuitos_grupo.append(str(circuitos_df.loc[idx, "Circuito"]))  # nombres de los circuitos de este grupo
            lista_circuitos = "\n".join(circuitos_grupo)  # texto final con todos los circuitos del grupo, uno por línea
            # Marca diferente para 10mA
            marca_dif = (marcas.get("Diferencial 10mA agua caliente", marcas.get("Protecciones", ""))  # el de 10mA puede ser de otra marca que el común de 30mA
                         if sens == "10mA" else marcas.get("Protecciones", ""))  # si no es de 10mA queda la marca común de protecciones
            add_row(  # fila del diferencial en la hoja de materiales
                desc=f"Interruptor diferencial {dif_txt.replace(' / Tipo A', '')} Tipo A",  # quita " / Tipo A" del texto porque se agrega aparte al final
                marcas_txt=marca_dif,  # marca según la sensibilidad del diferencial
                norma="RIC N°11 (6.4.3)" if sens == "10mA" else "RIC 6.5.1",  # el artículo del RIC cambia si el diferencial es de 10mA (agua caliente) o el normal de 30mA
                circuito=lista_circuitos if lista_circuitos else "Varios",  # si no hay nombres de circuitos, se deja "Varios"
                unidad="u",
                k=1,
                longitud_m="1 unid",
                cantidad=1  # 1 diferencial por grupo
            )

    # Termomagnéticos: 1 por circuito
    if {"Disyuntor termomagnético", "Circuito"}.issubset(set(circuitos_df.columns)):  # sin las columnas de TM y de circuito no se puede armar la lista
        tms_df = circuitos_df[["Disyuntor termomagnético", "Circuito"]].copy()  # copia solo las 2 columnas que interesan, para no tocar el circuitos_df original
        tms_df["Disyuntor termomagnético"] = tms_df["Disyuntor termomagnético"].astype(str).str.strip()  # limpia espacios en el texto del TM
        tms_df["Circuito"] = tms_df["Circuito"].astype(str).str.strip()  # limpia espacios en el nombre del circuito
        tms_df = tms_df[tms_df["Disyuntor termomagnético"] != ""]  # se saltan los circuitos sin TM definido
        for _, rr in tms_df.iterrows():  # un TM por cada circuito
            tm_txt = rr["Disyuntor termomagnético"]  # texto del TM tal como quedó definido, ej 2X16A
            circ2 = rr["Circuito"]  # nombre del circuito al que pertenece este TM
            add_row(  # fila del termomagnético de este circuito
                desc=f"Disyuntor termomagnético {tm_txt}",
                marcas_txt=marcas.get("Protecciones", ""),
                norma="RIC 6.4.1",  # mismo artículo del RIC que el omnipolar
                circuito=circ2,
                unidad="u",
                k=1,
                longitud_m="1 unid",
                cantidad=1  # 1 TM por circuito
            )

    # =========================
    # 4) ACCESORIOS (cajas + tapas + portalámparas + interruptores + enchufes + especiales)
    # =========================
    add_section("Accesorios")  # arranca la sección Accesorios en la hoja de materiales
    # OJO: de acá para abajo es_embutida / es_sobrepuesta se refieren al tipo de
    # canalización de TODO el proyecto (parámetro tipo_canalizacion), no al de
    # cada circuito. Son los mismos nombres de variable que se usaron dentro del
    # bucle de Canalizaciones, pero allá salían de la columna "Canalización" fila
    # por fila; acá se pisan con el valor global y ya no vuelven a cambiar.
    # tipo de canalización en minúsculas, para poder buscar "embut" o "sobre" adentro
    tipo_can = (tipo_canalizacion or "").strip().lower()  # texto normalizado para comparar
    es_embutida = "embut" in tipo_can  # True si la instalación es embutida (conduit dentro de la muralla)
    es_sobrepuesta = "sobre" in tipo_can  # True si es sobrepuesta (canalización a la vista)
    total_luminarias = 0  # total de luminarias de toda la casa (se suma más abajo)
    total_luminarias_sobrepuestas = 0  # solo las que su MONTAJE individual es sobrepuesto (llevan tornillo), las embutidas no
    total_enchufes = 0  # total de enchufes de toda la casa
    total_interruptores = 0  # total de grupos de interruptor de toda la casa (9/12, 9/15, 9/32, 9/24)
    total_ampolletas = 0  # total de ampolletas (led/incandescentes), necesitan portalámpara

    # Salidas de caja octogonal encadenadas, por ambiente: en canalización
    # sobrepuesta TODA luminaria (sea el foco embutido o sobrepuesto) lleva su
    # propia caja octogonal embutida puntual para la conexión/derivación (no
    # "centro a centro"). Cada luminaria del ambiente lleva 2 salidas
    # (entrada + salida a la siguiente), excepto la última del ambiente que
    # lleva solo 1 (entrada). Parecido a lo de embutida (~2634-2647), pero
    # allá la cadena se corta por grupo de interruptor y acá por ambiente.
    # (esa referencia ~2634-2647 es la numeración del archivo completo)
    _salidas_oct_luminarias = 0  # cuenta las salidas de caja que se van encadenando de luminaria en luminaria
    # Recorre cada ambiente y suma luminarias, enchufes e interruptores del ambiente.
    # OJO: acá los totales salen de ambientes_df (una fila por ambiente), no de
    # items_por_nombre (una lista por circuito) como en los bloques de más arriba.
    # Además usa dos columnas distintas para lo mismo: total_luminarias cuenta
    # LÍNEAS de "Detalle iluminación", mientras que los interruptores se calculan
    # con "Cantidad luminarias (u)". Si esas dos columnas no coinciden, las cajas
    # y las tapas quedan descuadradas.
    if ambientes_df is not None:  # sin la tabla de ambientes no hay nada que contar
        # ---- Luminarias (contar líneas de Detalle iluminación) ----
        if "Detalle iluminación" in ambientes_df.columns:  # columna con el detalle de cada luminaria del ambiente, una por línea de texto
            for v in ambientes_df["Detalle iluminación"].fillna("").astype(str):  # recorre el detalle de iluminación de cada ambiente
                lineas = [x.strip() for x in v.split("\n") if x.strip() and x.strip().lower() != "ninguna"]  # separa el texto en líneas, ignorando vacias y la palabra "ninguna"
                n_lum_amb = len(lineas)  # una línea de detalle = una luminaria
                total_luminarias += n_lum_amb  # va acumulando las luminarias de todos los ambientes
                if n_lum_amb > 0:  # solo ambientes que tengan al menos una luminaria
                    _salidas_oct_luminarias += 2 * (n_lum_amb - 1) + 1  # todas x2, última x1
                for ln in lineas:  # revisa el detalle de cada luminaria del ambiente
                    lnl = ln.strip().lower()  # en minúsculas para poder buscar palabras clave
                    if "sobrepuest" in lnl:  # el montaje de esa luminaria es sobrepuesto
                        total_luminarias_sobrepuestas += 1  # esas son las que llevan tornillo y tarugo
                    if ("ampolleta led" in lnl) or ("ampolleta incand" in lnl) or ("incandescente" in lnl and "ampolleta" in lnl):  # las distintas formas en que se escribe la ampolleta en el detalle
                        total_ampolletas += 1  # cuenta cuántas son ampolletas (necesitan portalámpara)

        # Enchufes (cantidad enchufes por ambiente)
        if "Cantidad enchufes (u)" in ambientes_df.columns:  # columna con la cantidad de enchufes que pidió el usuario en ese ambiente
            total_enchufes = int(  # suma los enchufes de todos los ambientes, tratando los vacíos como 0
                pd.to_numeric(ambientes_df["Cantidad enchufes (u)"], errors="coerce").fillna(0).sum()  # lo que no sea número queda en 0
            )
        # Interruptores reales por grupo (normal y 9/24)
        total_interruptores_pared = 0  # 1 por grupo normal, 2 por grupo 9/24 (sus 2 mecanismos físicos)
        if ("Cantidad luminarias (u)" in ambientes_df.columns) and ("N_conmutadas_924 (u)" in ambientes_df.columns):  # necesita las 2 columnas: luminarias del ambiente y conmutadas
            for _, ar in ambientes_df.iterrows():  # recorre ambiente por ambiente
                n_lum = int(pd.to_numeric(ar.get("Cantidad luminarias (u)", 0), errors="coerce") or 0)  # cuántas luminarias tiene el ambiente
                n_con = int(pd.to_numeric(ar.get("N_conmutadas_924 (u)", 0), errors="coerce") or 0)  # cuántas de esas luminarias son conmutadas (par 9/24)
                n_con = max(0, min(n_con, n_lum))  # nunca más conmutadas que luminarias totales
                n_rest = max(0, n_lum - n_con)  # el resto de luminarias que se manejan con interruptores normales
                c12, c15, c32 = descomponer_interruptores(n_rest)  # reparte esas luminarias "normales" en grupos de interruptor 9/12, 9/15, 9/32
                grupo_924 = 1 if n_con > 0 else 0  # el par 9/24 es 1 solo grupo, sin importar cuántas luminarias controle
                total_interruptores += grupo_924 + c12 + c15 + c32  # grupos totales (para cónicos/curvas/caja troncal)
                total_interruptores_pared += (2 * grupo_924) + c12 + c15 + c32  # 9/24 pesa 2 (2 mecanismos)

    # Sumar cajas adicionales de enchufes (misma fuente que usan los cónicos)
    _cajas_adic_enchufes = sum(int(v) for v in (cajas_adic_por_nombre or {}).values())  # suma cajas extra de enchufes que se hayan definido aparte (no vienen del conteo normal)

    # Sumar cajas y tapas de climatización a los contadores
    # para que queden en la misma fila que las cajas normales
    _clima_cajas_emb  = sum(1 for _cl in _clima_items if _cl["es_emb"])  # cajas embutidas de climatización (split, etc.)
    _clima_cajas_sob  = sum(1 for _cl in _clima_items if not _cl["es_emb"])  # cajas sobrepuestas de climatización
    # tapa ciega solo si la caja del equipo NO queda cubierta por un enchufe.
    # con_enchufe es None cuando no se pudo leer el TM: ahí "not None" da True y
    # el equipo igual suma tapa (se prefiere cotizarla de más antes que faltar)
    _clima_tapas      = sum(1 for _cl in _clima_items if _cl["es_emb"] and not _cl["con_enchufe"])  # cajas de climatización embutidas que además necesitan tapa ciega (no llevan enchufe)
    _clima_circ_str   = ", ".join(_cl["circ"].split("(")[0].strip() for _cl in _clima_items) if _clima_items else ""  # texto con los nombres de los circuitos de climatización, para mostrar en el Excel

    # Sumar cajas y tapas de agua caliente (si la conexión es fija, siempre lleva tapa)
    _agua_cajas_emb  = sum(1 for _ac in _agua_items if _ac["es_emb"])  # cajas embutidas de agua caliente (termo, calefont, etc.)
    _agua_cajas_sob  = sum(1 for _ac in _agua_items if not _ac["es_emb"])  # cajas sobrepuestas de agua caliente
    # mismo criterio de fondo que clima (tapa solo si no hay enchufe que cubra la
    # caja, o sea TM > 16A), con una diferencia: si el TM no se pudo leer, agua
    # caliente NO suma tapa (in_tm es None), mientras que clima sí la suma
    _agua_tapas      = sum(1 for _ac in _agua_items if _ac["es_emb"] and _ac["in_tm"] is not None and _ac["in_tm"] > 16)  # tapa solo si va sin enchufe, o sea TM sobre 16A
    _agua_circ_str   = ", ".join(_ac["circ"].split("(")[0].strip() for _ac in _agua_items) if _agua_items else ""  # texto con los nombres de los circuitos de agua caliente

    # Circuitos especiales genéricos (no clima/agua): cualquiera sea su
    # amperaje, todos necesitan su propia caja (con enchufe dedicado si es
    # ≤16A, o con prensaestopa si es >16A).
    # Un circuito se toma como "especial genérico" cuando alguno de sus items
    # trae la clave "nombre" (horno, lavadora, encimera...); los enchufes traen
    # "id_ench"/"módulos"/"n_ench" y las luminarias no traen ninguna de las dos.
    _nombres_esp_genericos = []  # nombres de los circuitos especiales que no son clima ni agua
    if circuitos_df is not None and "Circuito" in circuitos_df.columns:  # necesita la tabla de circuitos con su columna de nombre
        for _, _rr in circuitos_df.iterrows():  # recorre todos los circuitos definidos
            _nombre_c = str(_rr.get("Circuito", "")).strip()  # nombre del circuito, sin espacios sobrantes
            if not _nombre_c:  # fila sin nombre de circuito, se salta
                continue  # fila sin nombre de circuito, no hay nada que revisar
            _es_clima_o_agua = any(k in _nombre_c.lower() for k in  # True si el nombre del circuito suena a climatización o agua caliente (esos ya se contaron aparte)
                                    ("clima", "split", "aire", "ducha", "termo", "calefon", "agua caliente"))  # palabras que delatan un circuito de clima o de agua caliente
            if _es_clima_o_agua:  # clima y agua caliente ya se contaron en sus propios bloques
                continue  # clima y agua ya se contaron en su bloque, no se cuentan de nuevo
            _items_c = (items_por_nombre or {}).get(_nombre_c, [])  # items especiales guardados para este circuito (si los hay)
            _es_generico = isinstance(_items_c, list) and any(  # es "genérico" si alguno de sus items tiene la clave "nombre" (formato de circuito especial)
                isinstance(_it, dict) and ("nombre" in _it) for _it in _items_c  # basta con que un item traiga la clave 'nombre' para tratarlo como especial
            )
            if _es_generico:  # el circuito califica como especial genérico: necesita su propia caja
                _nombres_esp_genericos.append(_nombre_c)  # se guarda el nombre para contarlo más abajo
    # cuántos circuitos especiales genéricos hay. El nombre engaña: NO son solo
    # los de conexión fija, son TODOS (con enchufe o sin él), porque todos
    # necesitan su caja. Las tapas ciegas se cuentan aparte, más abajo, en
    # _esp_conexion_fija_tapas (solo los >16A).
    _esp_conexion_fija = len(_nombres_esp_genericos)  # usado para CAJAS (todos)
    # nombres de esos circuitos, para mostrar en la columna "Circuito" del Excel
    _esp_circ_str = ", ".join(n.split("(")[0].strip() for n in _nombres_esp_genericos)
    # Solo los >16A van SIN enchufe (conexión directa con prensaestopa) — esos
    # sí necesitan tapa ciega, porque su caja no queda cubierta por ningún
    # enchufe. Los ≤16A ya quedan tapados por el enchufe dedicado, sin tapa.
    _esp_conexion_fija_tapas = 0  # cuenta las tapas ciegas de los especiales sobre 16A
    if circuitos_df is not None and "Circuito" in circuitos_df.columns:  # se recorren de nuevo los circuitos para ver a cuáles les corresponde tapa ciega
        for _, _rr in circuitos_df.iterrows():  # segunda pasada por los circuitos, ahora buscando a cuáles les toca tapa ciega
            _nombre_c = str(_rr.get("Circuito", "")).strip()  # nombre del circuito, sin espacios sobrantes
            if _nombre_c in _nombres_esp_genericos:  # solo se revisan los circuitos que quedaron marcados como especiales genéricos
                _tm_c = parse_in_tm(_rr.get("Disyuntor termomagnético", "")) or 0  # corriente nominal del TM de ese circuito
                if _tm_c > 16:  # si es mayor a 16A va sin enchufe (conexión directa), por eso necesita tapa ciega
                    _esp_conexion_fija_tapas += 1  # una tapa ciega más para ese circuito

    # ── CAJAS DE DERIVACIÓN ──────────────────────────────────────────────────
    # Acá se cuentan y se emiten TODAS las cajas del proyecto (octogonales de
    # luminaria + rectangulares/chuqui de enchufes, interruptores, uniones,
    # cajas de paso, clima, agua y especiales). Los subtotales que quedan en
    # el subtotal cajas_rect_troncal se reutiliza más abajo, en el bloque
    # TAPAS CIEGAS, para saber cuántas cajas quedan sin mecanismo a la vista.
    # Según el tipo de canalización se cuentan las cajas de un modo distinto:
    if es_embutida:  # caso canalización embutida
        cajas_octogonales = total_luminarias  # en embutida, 1 caja octogonal por cada luminaria
        # Por cada GRUPO de interruptor va 1 caja troncal (la derivación) más 1
        # caja por cada mecanismo físico. En 9/12, 9/15 y 9/32 eso son 2 cajas;
        # en el 9/24 son 3, porque el par conmutado tiene 2 mecanismos.
        cajas_rect_troncal      = total_interruptores  # 1 troncal por GRUPO de interruptor (9/24 cuenta 1 grupo, no 2)
        cajas_rect_interruptor  = total_interruptores_pared  # 1 caja por cada interruptor FÍSICO (el 9/24 aporta 2)
        cajas_rect_union        = cajas_rect_troncal + cajas_rect_interruptor  # rectangulares que se van por interruptores: troncal + mecanismo
        cajas_rect_enchufe      = total_enchufes + _cajas_adic_enchufes  # cajas rectangulares para los enchufes (más las adicionales)
        if cajas_octogonales > 0:  # solo si hay luminarias que alimentar
            # Tipo de caja octogonal según diámetro conduit iluminación
            if _mm_conduit_ilumin <= 20:  # si el conduit de iluminación es chico (<=20mm) se usa la caja octogonal chica
                desc_oct  = 'Caja de derivación embutida octogonal de PVC 100x41 mm 4" (12 salidas)'  # octogonal chica de 4 pulgadas
            else:  # conduit de iluminación sobre 20mm
                desc_oct  = "Caja de derivación embutida octogonal grande de PVC 109x70x45 mm (12 salidas)"  # con conduit más grueso se necesita la octogonal grande
            add_row(  # fila de las cajas octogonales de iluminación
                desc=desc_oct,
                marcas_txt=marcas.get("Cajas derivación embutidas", ""),
                norma="RIC 4.3.1",  # artículo del RIC de cajas de derivación
                circuito="Iluminación",
                unidad="u",
                k=1,
                longitud_m=f"{cajas_octogonales} unid",
                cantidad=cajas_octogonales  # una caja por luminaria
            )
        # Todo circuito especial genérico suma su caja de derivación en
        # embutida, tenga enchufe o no (la variable se llama
        # _esp_conexion_fija por historia, pero los cuenta a todos).
        # (en sobrepuesta ya se suma, ver cajas_chuqui más abajo)
        total_rectangulares = cajas_rect_union + cajas_rect_enchufe + _clima_cajas_emb + _agua_cajas_emb + _esp_conexion_fija  # cajas rectangulares totales de todo tipo (uniones + enchufes + climatización + agua + especiales)
        total_rectangulares += _cajas_paso_total  # y las cajas de paso de TODOS los circuitos embutidos (enchufes, iluminación, especiales, clima y agua caliente)
        # Las cajas de paso se suman a _cajas_total para que espuma/tornillos/tarugos
        # de Panel SIP (que usan _cajas_total) las contemplen correctamente.
        _cajas_total += int(cajas_octogonales) + int(total_rectangulares)
        if total_rectangulares > 0:  # solo si salió al menos una caja rectangular
            circ_cajas = "Enchufes / Interruptores / Uniones"  # texto base de la columna Circuito de esta fila
            if _cajas_paso_total > 0:  # se nombran también las cajas de paso si las hay
                circ_cajas += " / Cajas de paso"  # suma las cajas de paso al texto del circuito
            if _clima_cajas_emb > 0:  # y los circuitos de climatización
                circ_cajas += f" / {_clima_circ_str}"  # agrega los nombres de los circuitos de clima
            if _agua_cajas_emb > 0:  # y los de agua caliente
                circ_cajas += f" / {_agua_circ_str}"  # agrega los nombres de los circuitos de agua caliente
            if _esp_conexion_fija > 0:  # y los especiales genéricos
                circ_cajas += f" / Especiales ({_esp_circ_str})"  # agrega los especiales genéricos, con sus nombres entre paréntesis
            add_row(  # fila de las cajas rectangulares embutidas
                desc="Caja de derivación embutida de PVC para tabiques 110x110x67 mm (12 salidas)",
                marcas_txt=marcas.get("Cajas derivación embutidas", ""),
                norma="RIC 4.3.1",
                circuito=circ_cajas,
                unidad="u",
                k=1,
                longitud_m=f"{total_rectangulares} unid",
                cantidad=total_rectangulares  # total de rectangulares contadas más arriba
            )
        # NOTA: las salidas de caja de las cajas de paso interior (RIC N°4,
        # art. 7.16.1.13) ya se generan por circuito, más arriba, dentro de la
        # fila "Salida de caja conduit" de cada circuito embutido ("salidas +=
        # 2 * cajas_paso_circ", línea ~2684 del archivo completo) — no hace
        # falta (ni corresponde) generarlas de nuevo acá agregadas, se
        # duplicaría. Lo que sí se agrega en este bloque es la caja física.
    elif es_sobrepuesta:  # caso canalización sobrepuesta
        # Cajas chuqui por interruptor: 1 troncal por grupo + 1 mecanismo por
        # interruptor físico (el 9/24 pesa 2 mecanismos, igual que en embutida)
        cajas_chuqui = total_enchufes + _cajas_adic_enchufes + total_interruptores + total_interruptores_pared + _cajas_paso_total + _esp_conexion_fija + _clima_cajas_sob + _agua_cajas_sob  # suma todo lo que necesita una caja chuqui sobrepuesta: enchufes, interruptores, cajas de paso y circuitos especiales/clima/agua
        _cajas_total += int(total_luminarias) + int(cajas_chuqui)  # se acumulan al total general de cajas (luminarias + chuqui)
        if cajas_chuqui > 0:  # solo si salió al menos una chuqui
            circ_chuqui = "Enchufes / Interruptores / Uniones"  # texto base de la columna Circuito de esta fila
            if _clima_cajas_sob > 0:  # se nombran los circuitos de climatización
                circ_chuqui += f" / {_clima_circ_str}"  # agrega los circuitos de clima al texto
            if _agua_cajas_sob > 0:  # y los de agua caliente
                circ_chuqui += f" / {_agua_circ_str}"  # agrega los de agua caliente
            if _esp_conexion_fija > 0:  # y los especiales genéricos
                circ_chuqui += f" / Especiales ({_esp_circ_str})"  # agrega los especiales genéricos
            add_row(  # fila de las cajas chuqui sobrepuestas
                desc="Caja de derivación sobrepuesta chuqui de PVC 12x8,5 cm",
                marcas_txt=marcas.get("Cajas derivación sobrepuestas", ""),
                norma="RIC 4.3.1",
                circuito=circ_chuqui,
                unidad="u",
                k=1,
                longitud_m=f"{cajas_chuqui} unid",
                cantidad=cajas_chuqui  # total de chuqui contadas más arriba
            )
        # Focos: en canalización sobrepuesta, TODA luminaria (sea el foco
        # embutido o sobrepuesto) lleva su propia caja octogonal EMBUTIDA
        # puntual (no chuqui, porque el foco sobrepuesto delgado no calza
        # sobre una caja que sobresale) para poder hacer la conexión/
        # derivación dentro de la caja y no "centro a centro". Misma regla de
        # cadena que en embutida: 2 salidas por luminaria (entrada + salida a
        # la siguiente), 1 sola (solo entrada) para la última del ambiente.
        if total_luminarias > 0:  # solo si la casa tiene luminarias
            if _mm_conduit_ilumin <= 20:  # si el conduit de iluminación es chico (<=20mm) se usa la caja octogonal chica
                desc_oct_sob = 'Caja de derivación embutida octogonal de PVC 100x41 mm 4" (12 salidas)'  # octogonal chica de 4 pulgadas
            else:  # conduit de iluminación sobre 20mm
                desc_oct_sob = "Caja de derivación embutida octogonal grande de PVC 109x70x45 mm (12 salidas)"  # con conduit más grueso se necesita la octogonal grande
            add_row(  # fila de las octogonales embutidas de cada luminaria
                desc=desc_oct_sob,
                marcas_txt=marcas.get("Cajas derivación embutidas", ""),
                norma="RIC 4.3.1",
                circuito="Iluminación",
                unidad="u",
                k=1,
                longitud_m=f"{total_luminarias} unid",
                cantidad=total_luminarias  # una caja por luminaria
            )
            add_row(  # fila de las salidas de caja de esas octogonales
                desc=f"Salida de caja conduit de PVC de {_mm_conduit_ilumin}mm",
                marcas_txt=marcas.get("Salida de caja conduit", ""),
                norma="RIC 4.7.2",  # artículo del RIC de salidas de caja
                circuito="Iluminación",
                unidad="u",
                k=1,
                longitud_m=f"{_salidas_oct_luminarias} unid",
                cantidad=_salidas_oct_luminarias  # salidas encadenadas contadas arriba
            )
    else:  # no se definió el tipo de canalización
        add_row(  # se deja la fila marcada para que el usuario la complete después
            desc="Caja de derivación (definir embutida o sobrepuesta)",
            marcas_txt=f"{marcas.get('Cajas derivación embutidas','')} / {marcas.get('Cajas derivación sobrepuestas','')}",
            norma="RIC 4.3.1",
            circuito="Varios",
            unidad="u",
            k=1,
            longitud_m="Por definir",
            cantidad=""  # sin cantidad, no se puede calcular hasta saber el tipo de canalización
        )


    # TAPAS CIEGAS
    # Una tapa ciega es la placa que cubre una caja que NO tiene mecanismo
    # (interruptor o enchufe), solo cables empalmados adentro. Acá se cuentan
    # cuántas tapas van, según si la canalización es embutida o sobrepuesta.
    tipo_can_tap = (tipo_canalizacion or "").strip().lower()  # texto en minúsculas, para comparar fácil
    es_emb_tap = "embut" in tipo_can_tap  # True si la canalización es embutida
    es_sob_tap = "sobre" in tipo_can_tap  # True si la canalización es sobrepuesta
    # _cajas_paso_total solo se llenó con circuitos de canalización embutida
    # (así se contaron más arriba), aunque después se sume en las dos ramas.
    # NOTA: _cajas_paso_total ya quedó sumado a _cajas_total dentro del bloque
    # ACCESORIOS (vía total_rectangulares si el proyecto es embutido, o vía
    # cajas_chuqui si es sobrepuesto). Por eso acá NO se vuelve a sumar a
    # _cajas_total: si se sumara de nuevo, tornillos/tarugos/espuma quedarían
    # sobreestimados. Lo que sí se hace acá es contarlo como TAPA.

    if es_emb_tap:  # tapas ciegas para el caso de canalización embutida
        # Embutida: octogonal por luminaria (incluye foco) + rectangular por
        # caja TRONCAL (no por el mecanismo del interruptor, que ya lo cubre
        # la placa) + cajas de paso + cajas adicionales entre ambientes + clima + agua
        tapas_octogonales   = int(max(0, total_luminarias))  # 1 tapa octogonal por cada luminaria (la caja octogonal queda tapada aunque el foco vaya encima)
        # La caja de conexión fija de un circuito especial también necesita su
        # tapa ciega — solo los >16A (sin enchufe), los ≤16A ya quedan
        # tapados por su propio enchufe dedicado
        tapas_rectangulares = int(max(0, cajas_rect_troncal)) + int(_cajas_adic_enchufes) + _clima_tapas + _agua_tapas + _cajas_paso_total + _esp_conexion_fija_tapas  # total de tapas rectangulares: troncal + cajas extra + clima + agua + cajas de paso + especiales fijos
        total_tapas_ciegas  = tapas_octogonales + tapas_rectangulares  # todas las tapas ciegas juntas, octogonales más rectangulares
        _tapas_ciegas_total = int(total_tapas_ciegas)  # se guarda el total para el resumen de más abajo
        # Tapa ciega octogonal
        if tapas_octogonales > 0:  # solo emite la fila si de verdad quedaron tapas octogonales
            # El tamaño de la tapa depende del diámetro del conduit de iluminación
            if _mm_conduit_ilumin <= 20:  # conduit de iluminación de 20mm o menos: alcanza la tapa de 4 pulgadas
                desc_tapa_oct = 'Tapa ciega octogonal de PVC 4" (10,2 cm diámetro)'  # tapa octogonal chica
            else:  # conduit sobre 20mm: tapa octogonal grande
                desc_tapa_oct = "Tapa ciega octogonal grande de PVC (11,0 x 11,0 cm)"  # conduit más grueso: se va a la tapa grande de 11x11
            add_row(  # fila de la tapa octogonal en la lista de materiales
                desc=desc_tapa_oct,  # descripción que se eligió según el diámetro del conduit
                marcas_txt=marcas.get("Tapa ciega octogonal", ""),  # marcas sugeridas para tapa octogonal
                norma="RIC 4.3.1",  # la RIC pide tapar toda caja que quede sin uso
                circuito="Iluminación",  # estas tapas van al circuito de iluminación
                unidad="u",  # se cotiza por unidad
                k=1,  # sin holgura, es conteo directo
                longitud_m=f"{tapas_octogonales} unid",  # en la columna de largo se muestra el conteo
                cantidad=tapas_octogonales  # cantidad final que va al Excel
            )
        # Tapa ciega rectangular (troncal + cajas adicionales + cajas de paso + clima + agua)
        if tapas_rectangulares > 0:  # mismo criterio pero para las rectangulares
            circ_tapas = "Uniones"  # texto base del circuito, se le va sumando cada origen
            # Va agregando al texto del circuito cada origen que aportó tapas,
            # para que en el Excel se entienda de dónde salió la cantidad
            if _cajas_adic_enchufes > 0:  # hubo tapas de cajas extra entre ambientes
                circ_tapas += " / Cajas adicionales entre ambientes"  # queda anotado ese origen
            if _cajas_paso_total > 0:  # hubo tapas de cajas de paso
                circ_tapas += " / Cajas de paso"  # queda anotado
            if _clima_tapas > 0:  # hubo tapas de equipos de clima
                circ_tapas += f" / {_clima_circ_str}"  # agrega el nombre del circuito de clima
            if _agua_tapas > 0:  # hubo tapas de agua caliente
                circ_tapas += f" / {_agua_circ_str}"  # agrega el nombre del circuito de agua
            if _esp_conexion_fija_tapas > 0:  # hubo tapas de especiales con conexión fija
                circ_tapas += f" / Especiales ({_esp_circ_str})"  # agrega los circuitos especiales
            add_row(  # fila de las tapas rectangulares
                desc="Tapa ciega de PVC de 110 x 67 mm",  # tapa rectangular estándar de 110x67
                marcas_txt=marcas.get("Tapa ciega", ""),  # marcas sugeridas para tapa ciega
                norma="RIC 4.3.1",  # misma norma que la octogonal
                circuito=circ_tapas,  # acá va el texto con todos los origenes
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{tapas_rectangulares} unid",  # conteo en la columna de largo
                cantidad=tapas_rectangulares  # cantidad final
            )
    elif es_sob_tap:  # instalación sobrepuesta: cambia el tipo de tapa
        # Sobrepuesta: tapa chuqui en caja troncal + cajas_adic + cajas_paso
        # NO en caja mecanismo (interruptor) ni enchufe (tienen mecanismo
        # instalado) — los especiales solo suman tapa si son >16A (sin
        # enchufe), los ≤16A ya quedan tapados por su propio enchufe dedicado
        tapas_chuqui = int(total_interruptores) + int(_cajas_adic_enchufes) + int(_cajas_paso_total) + _clima_tapas + _agua_tapas + _esp_conexion_fija_tapas  # en sobrepuesta la tapa es chuqui: interruptores + cajas extra + cajas de paso + clima + agua + especiales fijos
        # Luminarias: en sobrepuesta, cada foco ahora sí lleva su propia caja
        # octogonal embutida puntual (ver bloque de ACCESORIOS más arriba), así
        # que también necesita su tapa ciega octogonal, igual que en embutida.
        tapas_octogonales_sob = int(max(0, total_luminarias))  # una tapa octogonal por cada luminaria
        total_tapas_ciegas  = tapas_chuqui + tapas_octogonales_sob  # total de tapas ciegas en sobrepuesta
        _tapas_ciegas_total = int(total_tapas_ciegas)  # se guarda para el resumen
        if tapas_octogonales_sob > 0:  # si no hay focos no hay tapa octogonal
            # Mismo criterio de tamaño que en embutida, según diámetro de conduit
            if _mm_conduit_ilumin <= 20:  # mismo corte de 20mm que en embutida
                desc_tapa_oct_sob = 'Tapa ciega octogonal de PVC 4" (10,2 cm diámetro)'  # tapa de 4 pulgadas
            else:  # conduit sobre 20mm: tapa octogonal grande
                desc_tapa_oct_sob = "Tapa ciega octogonal grande de PVC (11,0 x 11,0 cm)"  # tapa grande
            add_row(  # fila de tapas octogonales en sobrepuesta
                desc=desc_tapa_oct_sob,  # descripción según el diámetro
                marcas_txt=marcas.get("Tapa ciega octogonal", ""),  # marcas sugeridas
                norma="RIC 4.3.1",  # misma norma de cajas tapadas
                circuito="Iluminación",  # circuito de iluminación
                unidad="u",  # por unidad
                k=1,  # conteo directo
                longitud_m=f"{tapas_octogonales_sob} unid",  # conteo en la columna de largo
                cantidad=tapas_octogonales_sob  # cantidad final
            )
        if tapas_chuqui > 0:  # fila de las tapas chuqui
            circ_tapas = "Uniones"  # texto base del circuito
            # Igual que en embutida: se va anotando de dónde salió cada tapa
            if _clima_tapas > 0:  # aporte de clima
                circ_tapas += f" / {_clima_circ_str}"  # lo anota
            if _agua_tapas > 0:  # aporte de agua caliente
                circ_tapas += f" / {_agua_circ_str}"  # lo anota
            if _esp_conexion_fija_tapas > 0:  # aporte de especiales con conexión fija
                circ_tapas += f" / Especiales ({_esp_circ_str})"  # lo anota
            add_row(  # fila de la tapa chuqui
                desc="Tapa ciega chuqui de PVC 12x8,5 cm",  # medida típica de la tapa chuqui
                marcas_txt=marcas.get("Tapa ciega", ""),  # marcas sugeridas
                norma="RIC 4.3.1",  # misma norma
                circuito=circ_tapas,  # texto con todos los origenes
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{tapas_chuqui} unid",  # conteo en la columna de largo
                cantidad=tapas_chuqui  # cantidad final
            )

    # PORTALÁMPARAS SOLO SI HAY AMPOLLETAS
    # El portalámpara es la pieza donde se atornilla la ampolleta. Solo se
    # agrega si el proyecto de verdad tiene ampolletas contadas.
    if total_ampolletas > 0:  # sin ampolletas contadas no hay donde atornillar nada
        add_row(  # fila del portalámpara
            desc="Portalámpara plafón E27 redondo",  # plafon E27, el casquillo típico de casa
            marcas_txt=marcas.get("Portalamparas", ""),  # marcas sugeridas
            norma="RIC 10 (5.1.4.4, 5.1.4.5)",  # la RIC 10 pide portalámpara en cada punto de luz
            circuito="Iluminación",  # circuito de iluminación
            unidad="u",  # por unidad
            k=1,  # conteo directo
            longitud_m=f"{total_ampolletas} unid",  # conteo en la columna de largo
            cantidad=total_ampolletas  # uno por cada ampolleta
        )

    # INTERRUPTORES (9/12 - 9/15 - 9/24 - 9/32)
    # Recorre todos los circuitos de iluminación, agrupa sus luminarias por
    # ambiente, y decide cuántos interruptores (simples o conmutados) van en
    # cada ambiente, sumando todo por descripción al final.
    ncon_map = {}  # {nombre de ambiente en minúsculas: cuántas LUMINARIAS conmutadas pidió ahí (los interruptores 9/24 son siempre 2 por ambiente)}
    if ambientes_df is not None and "Ambiente" in ambientes_df.columns:  # solo si vino la tabla de ambientes con su columna de nombres
        amb_col = ambientes_df["Ambiente"].astype(str).str.strip()  # nombres de ambiente limpios, sin espacios de sobra
        if "N_conmutadas_924 (u)" in ambientes_df.columns:  # revisa si el usuario lleno la columna de conmutadas
            # Columna con la cantidad de LUMINARIAS conmutadas que el usuario pidió
            # por ambiente (no es la cantidad de interruptores: todas las conmutadas
            # de un ambiente se manejan con un solo par 9/24)
            ncon_col = pd.to_numeric(ambientes_df["N_conmutadas_924 (u)"], errors="coerce").fillna(0).astype(int)
            for a, nval in zip(amb_col, ncon_col):  # arma el mapa ambiente -> cuántos conmutados pidió
                ncon_map[a.lower()] = int(max(0, nval))  # nunca negativo
        else:  # la tabla de ambientes no trae la columna de conmutadas
            # Si no existe la columna, se asume que no hay conmutados en ningún ambiente
            for a in amb_col:
                ncon_map[a.lower()] = 0  # sin esa columna, ningún ambiente queda con conmutados
    if {"Circuito", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # necesita el nombre del circuito y su TM para poder trabajar
        _interruptores_acum = {}  # {desc: {"total": N, "detalle": {"circuito y ambiente": cant}}}

        # Suma "cant" unidades de un interruptor con descripción "desc",
        # guardando además el detalle de a qué circuito/ambiente pertenecen
        # qué recibe:
        #   desc ......... descripción del interruptor, ej: "Interruptor 9/12 10A 250V"
        #   detalle_txt .. de dónde viene, ej: "Circuito 1 - Iluminación -> Living"
        #   cant ......... cuántas unidades sumar
        # no devuelve nada: va llenando el diccionario _interruptores_acum.
        def _acum_interruptor(desc, detalle_txt, cant):
            if desc not in _interruptores_acum:  # primera vez que aparece este tipo de interruptor
                _interruptores_acum[desc] = {"total": 0, "detalle": {}}  # lo parte en cero
            _interruptores_acum[desc]["total"] += cant  # suma al total de ese interruptor
            _interruptores_acum[desc]["detalle"][detalle_txt] = _interruptores_acum[desc]["detalle"].get(detalle_txt, 0) + cant  # y guarda de que circuito y ambiente vino

        for _, rr in circuitos_df.iterrows():  # revisa circuito por circuito
            nombre_circ = str(rr.get("Circuito", "")).strip()  # nombre del circuito tal como viene en la tabla
            if "ilumin" not in nombre_circ.lower():  # si el nombre no dice ilumin, no lleva interruptor de pared
                continue  # solo interesa iluminación
            tm_txt = rr.get("Disyuntor termomagnético", "")  # texto del disyuntor termomagnético, ej 1x10A curva C
            in_tm = parse_in_tm(tm_txt)  # saca la corriente nominal del TM, en amperes
            amp_sw = "16A" if (in_tm == 16) else "10A"  # mecanismo de 16A solo si el TM es exactamente 16A; con cualquier otro valor (o TM ilegible, in_tm None) queda en 10A
            items = (items_por_nombre or {}).get(nombre_circ, [])  # luminarias y demás items que tiene este circuito
            if not isinstance(items, list) or not items:  # sin items no hay nada que contar
                continue  # circuito de iluminación sin items cargados, no aporta interruptores
            # agrupa las luminarias de este circuito por ambiente
            lum_by_amb = {}
            for it in items:  # revisa item por item del circuito
                if not isinstance(it, dict):  # ignora cualquier item que no sea diccionario
                    continue  # pasa al siguiente
                if ("modulos" in it) or ("id_ench" in it) or ("n_ench" in it) or ("nombre" in it):  # si trae módulos, id de enchufe o nombre, no es luminaria
                    continue  # descarta los items que no son luminarias (enchufes, especiales, etc.)
                amb = str(it.get("amb", "")).strip()  # ambiente donde esta la luminaria
                if not amb:  # si el item no trae ambiente
                    amb = "sin_amb"  # lo mete en un grupo genérico
                key = amb.lower()  # clave en minúsculas para no duplicar por mayúsculas
                lum_by_amb[key] = lum_by_amb.get(key, 0) + 1  # cuenta una luminaria más en ese ambiente
            # recorre cada ambiente y arma la fila de interruptores correspondiente
            for amb_key, n_lum in lum_by_amb.items():
                amb_show = amb_key  # nombre "bonito" del ambiente (con mayúsculas originales)
                try:  # el nombre bonito puede fallar si el ambiente no está en la tabla
                    # busca el nombre del ambiente tal como lo escribió el usuario, para mostrarlo bonito
                    match = ambientes_df[ambientes_df["Ambiente"].astype(str).str.strip().str.lower() == amb_key]
                    if len(match) > 0:  # encontró el ambiente en la tabla
                        amb_show = str(match.iloc[0]["Ambiente"]).strip()  # se queda con el nombre tal como lo escribió el usuario
                except:  # si algo falla se queda con el nombre en minúsculas
                    pass  # se queda con el nombre en minúsculas, no es grave
                n_conmutadas = int(max(0, ncon_map.get(amb_key, 0)))  # cuántos conmutados 9/24 pidió el usuario en ese ambiente
                n_conmutadas = min(n_conmutadas, int(n_lum))  # no puede pedir más conmutados que luminarias hay
                n_restantes = int(n_lum) - n_conmutadas  # luminarias que quedan con interruptor normal (no conmutado)
                if n_conmutadas > 0:  # si hay conmutadas, van de a par (9/24)
                    _acum_interruptor(  # suma el par de conmutados
                        f"Interruptor 9/24 (conmutado) {amp_sw} 250V",  # el amperaje del mecanismo sale del TM del circuito
                        f"{nombre_circ} -> {amb_show}",  # detalle de circuito y ambiente
                        2  # siempre 2 (1 par), sin importar cuántas luminarias controle
                    )
                if n_restantes > 0:  # el resto de las luminarias va con interruptor normal
                    c12, c15, c32 = descomponer_interruptores(int(n_restantes))  # cómo se agrupan las no conmutadas
                    if c12 > 0:  # los de 1 módulo
                        _acum_interruptor(f"Interruptor 9/12 {amp_sw} 250V", f"{nombre_circ} -> {amb_show}", c12)  # los suma al acumulado
                    if c15 > 0:  # los de 2 módulos
                        _acum_interruptor(f"Interruptor 9/15 {amp_sw} 250V", f"{nombre_circ} -> {amb_show}", c15)  # los suma
                    if c32 > 0:  # los de 3 módulos (controlan 3 luminarias)
                        _acum_interruptor(f"Interruptor 9/32 {amp_sw} 250V", f"{nombre_circ} -> {amb_show}", c32)  # los suma

        # emite 1 sola fila por descripción, sumando todos los circuitos/ambientes
        for desc, info in sorted(_interruptores_acum.items(), key=lambda x: x[0].lower()):
            detalle_txt = ", ".join(f"{k} ({v})" for k, v in info["detalle"].items())  # arma el texto con cada circuito/ambiente y su cantidad
            add_row(  # fila del interruptor en la lista
                desc=desc,  # descripción del interruptor
                marcas_txt=marcas.get("Interruptores", ""),  # marcas sugeridas
                norma="RIC 7.3.2",  # la RIC 7.3.2 cubre los mecanismos de mando
                circuito=detalle_txt,  # detalle de donde van
                unidad="u",  # por unidad
                k=1,  # conteo directo
                longitud_m=f"{info['total']} unid",  # conteo en la columna de largo
                cantidad=info["total"]  # total sumado de todos los circuitos
            )

    # ENCHUFES COMUNES: por circuito + ambiente + tipo
    # Recorre todos los circuitos: los de hasta 16A arman enchufes comunes
    # y los de más de 16A se van a la fila de conexión directa. Junta sus enchufes,
    # agrupándolos por circuito+ambiente+amperaje+cantidad de módulos
    # (simple/doble/triple), para después sumar todo por descripción.
    # Si el TM no se pudo leer (in_tm None), el circuito cae del lado de los
    # enchufes comunes y se cotiza como 10A.
    ench_rows = []  # acá se van juntando todos los enchufes antes de agruparlos
    _conexion_directa_acum = {"total": 0, "detalle": {}}  # circuitos >16A sin enchufe común
    if {"Circuito", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # necesita nombre de circuito y TM
        for _, rr in circuitos_df.iterrows():  # recorre los circuitos uno por uno
            nombre_circ = str(rr.get("Circuito", "")).strip()  # nombre del circuito
            tm_txt = rr.get("Disyuntor termomagnético", "")  # texto del TM del circuito
            in_tm = parse_in_tm(tm_txt)  # corriente nominal del TM, en amperes
            # Usar items_por_nombre en vez de _items del DataFrame
            items = (items_por_nombre or {}).get(nombre_circ, [])
            if not isinstance(items, list):  # si no vino como lista, la deja vacía
                items = []  # queda vacía y no aporta enchufes
            if not nombre_circ or len(items) == 0:  # sin nombre o sin items no hay nada que contar
                continue  # circuito sin nombre o sin items, no aporta enchufes
            items_ench = []  # acá quedan solo los items que son enchufes
            for it in items:  # revisa item por item
                if isinstance(it, dict) and (("modulos" in it) or ("id_ench" in it)):  # un enchufe se reconoce porque trae módulos o id_ench
                    items_ench.append(it)  # solo se queda con los items que son enchufes
            if not items_ench:  # este circuito no tenía enchufes
                continue  # este circuito no tenía ningún enchufe, pasa al siguiente
            if in_tm is not None and in_tm > 16:  # circuito de más de 16A: no lleva enchufe común
                # circuito de más de 16A: no lleva enchufe común, va a conexión directa
                _conexion_directa_acum["total"] += len(items_ench)  # esos puntos se suman como conexión directa
                _conexion_directa_acum["detalle"][nombre_circ] = _conexion_directa_acum["detalle"].get(nombre_circ, 0) + len(items_ench)  # y se guarda de que circuito salieron
                continue  # no sigue armando enchufes para este circuito

            # arma una fila por cada enchufe, para después agruparlas
            for it in items_ench:
                amb = str(it.get("amb", "")).strip()  # ambiente donde va el enchufe
                mod = int(it.get("modulos", 1)) if str(it.get("modulos", "")).strip() != "" else 1  # módulos del enchufe: 1 simple, 2 doble, 3 triple
                if in_tm == 16:  # TM de exactamente 16A
                    nominal = "10/16A"  # el enchufe tiene que ser 10/16A
                else:  # TM más chico, o TM que no se pudo leer
                    nominal = "10A"  # alcanza el enchufe de 10A
                ench_rows.append({  # guarda una fila por cada enchufe
                    "Circuito": nombre_circ,  # circuito al que pertenece
                    "Ambiente": amb if amb else "Sin ambiente",  # ambiente, o Sin ambiente si no lo escribieron
                    "Nominal": nominal,  # amperaje nominal del enchufe
                    "Modulos": mod  # cuántos módulos tiene
                })  # cierra la fila de este enchufe
    if _conexion_directa_acum["total"] > 0:  # hubo puntos sin enchufe, van como fila aparte
        # Estos son los puntos de circuitos >16A que no llevan enchufe (ej. horno,
        # cocina eléctrica): el cable queda conectado directo al artefacto.
        detalle_txt = ", ".join(f"{k} ({v})" for k, v in _conexion_directa_acum["detalle"].items())  # texto con cada circuito y cuántos puntos aporto
        add_row(  # fila del punto de conexión directa
            desc="Punto para conexión directa (>16A) - sin enchufe",  # queda claro en la descripción que no lleva enchufe
            marcas_txt="",  # no es un producto de marca, es un punto de conexión
            norma="RIC 13.5.1",  # la RIC 13.5.1 cubre los puntos de enchufe/conexión
            circuito=detalle_txt,  # detalle de circuitos
            unidad="u",  # por unidad
            k=1,  # conteo directo
            longitud_m=f"{_conexion_directa_acum['total']} unid",  # conteo en la columna de largo
            cantidad=_conexion_directa_acum["total"]  # cantidad final
        )
    if ench_rows:  # si se junto algún enchufe común
        # agrupa todos los enchufes por circuito+ambiente+amperaje+módulos, para
        # sacar la cantidad de cada combinación
        ench_df = pd.DataFrame(ench_rows)
        grp = ench_df.groupby(["Circuito", "Ambiente", "Nominal", "Modulos"]).size().reset_index(name="Cantidad")  # cuenta cuántos enchufes hay en cada combinación
        mod_order = {1: 0, 2: 1, 3: 2}  # orden para mostrar primero simple, luego doble, luego triple
        grp["__ord"] = grp["Modulos"].map(lambda x: mod_order.get(int(x), 9))  # columna auxiliar solo para ordenar por tipo
        grp = grp.sort_values(["Nominal", "__ord", "Circuito", "Ambiente"]).drop(columns="__ord")  # ordena y después bota la columna auxiliar
        # consolida en 1 fila por descripción (nominal+módulos), sumando todos
        # los circuitos/ambientes, con el detalle de cada uno en "Circuito"
        _enchufes_acum = {}  # {desc: {"total": N, "detalle": {"circuito y ambiente": cant}}}
        # recorre cada combinación de circuito+ambiente+amperaje+módulos y la
        # va sumando dentro de _enchufes_acum, según el nombre final que le
        # corresponda (simple/doble/triple)
        for _, g in grp.iterrows():
            circ = g["Circuito"]  # circuito de esta combinación
            amb = g["Ambiente"]  # ambiente
            nominal = g["Nominal"]  # amperaje nominal, 10A o 10/16A
            mod = int(g["Modulos"])  # módulos del enchufe
            cant = int(g["Cantidad"])  # cuántos van de esa combinación
            if mod == 1:  # 1 módulo
                tipo = "Enchufe simple"  # enchufe simple
            elif mod == 2:  # 2 módulos
                tipo = "Enchufe doble"  # enchufe doble
            else:  # 3 o más módulos
                tipo = "Enchufe triple"  # enchufe triple
            desc = f"{tipo} 2P+T {nominal} 250V"  # descripción final, siempre 2P+T (con tierra)
            detalle_txt = f"{circ} -> {amb}"  # de que circuito y ambiente viene
            if desc not in _enchufes_acum:  # primera vez que sale este tipo de enchufe
                _enchufes_acum[desc] = {"total": 0, "detalle": {}}  # lo parte en cero
            _enchufes_acum[desc]["total"] += cant  # suma la cantidad al total
            _enchufes_acum[desc]["detalle"][detalle_txt] = _enchufes_acum[desc]["detalle"].get(detalle_txt, 0) + cant  # y anota el detalle por circuito/ambiente
        # emite 1 fila por descripción de enchufe, ya sumada
        for desc, info in _enchufes_acum.items():
            detalle_txt = ", ".join(f"{k} ({v})" for k, v in info["detalle"].items())  # texto con todos los circuitos y ambientes
            add_row(  # fila del enchufe en la lista
                desc=desc,  # descripción del enchufe
                marcas_txt=marcas.get("Enchufes", ""),  # marcas sugeridas
                norma="RIC 13.5.1",  # misma norma de puntos de enchufe
                circuito=detalle_txt,  # detalle de donde van
                unidad="u",  # por unidad
                k=1,  # conteo directo
                longitud_m=f"{info['total']} unid",  # conteo en la columna de largo
                cantidad=info["total"] if info["total"] > 0 else ""  # si quedó en cero, deja la celda vacía
            )

    # Enchufe dedicado para climatización, siempre 2P+T 10/16A
    # Recorre cada equipo de climatización guardado antes en _clima_items y,
    # si ese equipo tiene enchufe (no es conexión fija), agrega la fila del
    # enchufe correspondiente a la lista de materiales.
    for _cl in _clima_items:
        # con_enchufe es None cuando no se pudo leer el TM: ahí no se emite
        # enchufe ni conexión fija, el aviso lo da el bloque de cónicos/ferrules
        if _cl["con_enchufe"]:  # solo los equipos que van con enchufe, los de conexión fija no
            add_row(  # fila del enchufe dedicado del aire
                desc="Enchufe dedicado para aire acondicionado simple 2P+T 10/16A 250V",  # el aire simple usa 2P+T 10/16A
                marcas_txt=marcas.get("Enchufes", ""),  # marcas sugeridas
                norma="RIC 7 (7.3.1, 7.3.2, 7.4.4)",  # la RIC 7 cubre el enchufe del circuito dedicado
                circuito=_cl["circ"],  # circuito de clima de ese equipo
                unidad="u",  # por unidad
                k=1,  # conteo directo
                longitud_m="1 unid",  # uno por equipo
                cantidad=1  # un enchufe por equipo
            )

    # Tabla prensaestopa según sección conductor (PG/Métrico)
    # Incluye calibres AWG equivalentes
    # Forma del diccionario: sección del cable en mm^2 -> (tamaño PG, rosca
    # métrica equivalente). Las claves "raras" (2.08, 3.31, 5.26, 8.37) son las
    # secciones exactas de los calibres AWG 14, 12, 10 y 8, para que un cable
    # americano caiga en el mismo prensaestopa que su equivalente métrico.
    # Es una tabla de catálogo comercial, no sale de un artículo del RIC.
    PRENSAESTOPA_POR_SECCION = {
        1.5:  ("PG11",   "M16"),   # 1,5mm^2
        2.08: ("PG11",   "M16"),   # AWG 14
        2.5:  ("PG11",   "M20"),   # 2,5mm^2
        3.31: ("PG11",   "M20"),   # AWG 12
        4.0:  ("PG13,5", "M20"),   # 4mm^2
        5.26: ("PG13,5", "M20"),   # AWG 10
        6.0:  ("PG16",   "M25"),   # 6mm^2
        8.37: ("PG16",   "M25"),   # AWG 8
        10.0: ("PG21",   "M25"),   # 10mm^2
    }

    # elige el prensaestopa que le corresponde al cable según su sección,
    # devuelve el texto listo para la lista de materiales
    def prensaestopa_para_seccion(sec):
        # sec = sección del cable en mm^2 (ej: 4.0), este es el dato de entrada
        # pg  = tamaño del prensaestopa (ej: "PG11")
        # mt  = la rosca métrica equivalente a ese pg (ej: "M20")
        # c   = cada una de las secciones que hay en la tabla PRENSAESTOPA_POR_SECCION
        # claves = todas las secciones "c" de la tabla, ordenadas de menor a mayor
        claves = sorted(PRENSAESTOPA_POR_SECCION.keys())
        pg, mt = PRENSAESTOPA_POR_SECCION[claves[-1]]  # parte con el prensaestopa más grande, por si acaso
        for c in claves:  # va probando de la sección más chica a la más grande
            # +0.01 es solo para evitar problemas de redondeo con decimales
            if sec <= c + 0.01:  # ¿el cable (sec) ya cabe en esta sección de tabla (c)?
                pg, mt = PRENSAESTOPA_POR_SECCION[c]  # sí cabe: se queda con el pg/mt de esta sección
                break  # ya encontró el tamaño, no sigue buscando
        sec_txt = str(round(sec, 2)).replace(".", ",")  # sección con coma decimal (2,5 en vez de 2.5)
        return f"Prensaestopa {pg} ({mt}) para cordón (cable 3x{sec_txt}mm^2)"  # texto final que se muestra en la lista de materiales

    # ─────────────────────────────────────────────────────────────────────────
    # ACCESORIOS AGUA CALIENTE
    # Con enchufe (TM≤16A) o conexión fija sin enchufe (TM>16A) — RIC N°07 7.2.8 / 7.3
    # Genera prensaestopa en equipo (solo sin enchufe) + tablero externo si corresponde
    # NOTA: este bloque va DESPUÉS de definir PRENSAESTOPA_POR_SECCION
    # ─────────────────────────────────────────────────────────────────────────
    if {"Circuito", "Conductor", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):
        # Solo entra a este bloque si el DataFrame de circuitos tiene todas
        # las columnas que se necesitan para trabajar (Circuito, Conductor y TM)
        for _, rr_ac in circuitos_df.iterrows():
            # recorre cada fila (cada circuito) de circuitos_df, una por una
            nombre_ac = str(rr_ac.get("Circuito", "")).strip()  # nombre del circuito de esta fila
            # ¿el nombre del circuito suena a que es para agua caliente? (ducha, termo, calefón, etc.)
            es_agua_mat = any(k in nombre_ac.lower() for k in
                              ("ducha", "termo", "calefon", "calefón",
                               "calentador", "agua caliente"))  # resto de palabras que indican agua caliente
            if not es_agua_mat:  # filtra: acá abajo solo siguen los circuitos de agua caliente
                continue  # no es un circuito de agua caliente, se salta

            sec_ac = extraer_seccion_mm2(str(rr_ac.get("Conductor", ""))) or 4.0  # sección del conductor en mm^2, si no se puede leer asume 4
            in_tm_ac = parse_in_tm(str(rr_ac.get("Disyuntor termomagnético", "")))  # corriente nominal del TM de ese circuito
            if in_tm_ac is None:  # no se pudo leer el TM
                # no se pudo leer el TM del circuito: se marca para que lo revisen a mano
                add_row(  # deja una fila de aviso para que la revisen a mano
                    desc=f"(definir) {nombre_ac} - TM no legible, revisar con/sin enchufe manualmente",  # parte con (definir) para que salte a la vista en el Excel
                    marcas_txt="",  # sin marca, es un aviso
                    norma="",  # sin norma, no es un material
                    circuito=nombre_ac,  # circuito afectado
                    unidad="u",  # por unidad
                    k=1,  # conteo directo
                    longitud_m="1 unid",  # una sola
                    cantidad=1  # cantidad 1
                )
                continue  # sin TM no se puede decidir si va con o sin enchufe

            # Buscar datos del equipo en circuitos_agua_caliente
            _datos_ac = {}
            for eq_ac in circuitos_agua_caliente:  # busca el equipo entre los que cargo el usuario
                if eq_ac.get("nombre_circ", "").lower() in nombre_ac.lower():  # calza si el nombre del equipo aparece dentro del nombre del circuito
                    _datos_ac = eq_ac  # se queda con esos datos
                    break  # ya lo encontró
            lleva_tab_ext = _datos_ac.get("lleva_tablero_externo", False)  # si el equipo pide tablero externo propio

            # Calcular prensaestopa según sección real del conductor
            _pg_ac, _mt_ac = ("PG11", "M16")  # valor por defecto
            for _sec_k, (_pg_k, _mt_k) in sorted(PRENSAESTOPA_POR_SECCION.items()):  # recorre la tabla de prensaestopas de menor a mayor sección
                if sec_ac <= _sec_k + 0.01:  # primera sección de la tabla donde cabe el cable
                    _pg_ac, _mt_ac = _pg_k, _mt_k  # primer tamaño de la tabla que alcanza
                    break  # se queda con ese tamaño
            sec_txt_ac = str(round(sec_ac, 1)).replace(".", ",")  # sección con coma decimal para mostrarla

            # Sección de la bornera PE: se toma la sección comercial más chica que
            # cubre al conductor real del circuito, con tope en 10mm^2. Es un
            # criterio de continuidad del PE (el borne no puede ser más chico que
            # el cable de tierra que entra), no una tabla del RIC.
            # Normalizar a sección comercial de bornera disponible
            if sec_ac <= 1.5:   sec_born_pe = 1.5  # cable chico: bornera de 1,5
            elif sec_ac <= 2.5: sec_born_pe = 2.5  # bornera de 2,5
            elif sec_ac <= 4.0: sec_born_pe = 4.0  # bornera de 4
            elif sec_ac <= 6.0: sec_born_pe = 6.0  # bornera de 6
            else:               sec_born_pe = 10.0  # de ahí para arriba, bornera de 10
            sec_born_pe_txt = str(sec_born_pe).replace(".", ",")  # texto de la sección de bornera con coma

            # 1) Punto conexión fija (sin enchufe) o enchufe (TM≤16A), igual criterio que climatización
            con_enchufe_ac = (in_tm_ac <= 16)  # hasta 16A el equipo va con enchufe, sobre eso es conexión fija
            if con_enchufe_ac:  # caso con enchufe
                # Los enchufes solo existen comercialmente en 10A o 16A — se
                # redondea el TM real al calibre comercial correcto
                _amp_ench_ac = 10 if in_tm_ac <= 10 else 16
                add_row(  # fila del enchufe del equipo de agua caliente
                    desc=f"Enchufe 2P+T {_amp_ench_ac}A / 250V para {_datos_ac.get('tipo_equipo', 'equipo agua caliente')}",  # descripción con el amperaje y el tipo de equipo
                    marcas_txt=marcas.get("Enchufes", ""),  # marcas sugeridas
                    norma="RIC N°07 (7.3.1, 7.3.2, 7.4.4)",  # la RIC N07 cubre los enchufes
                    circuito=nombre_ac,  # circuito de agua caliente
                    unidad="u",  # por unidad
                    k=1,  # conteo directo
                    longitud_m="1 unid",  # uno por equipo
                    cantidad=1  # cantidad 1
                )
            else:  # sin enchufe: el equipo queda conectado fijo
                add_row(  # fila del punto de conexión fija del equipo de agua caliente
                    desc=f"Punto para conexión fija — {_datos_ac.get('tipo_equipo', 'equipo agua caliente')} (sin enchufe)",  # descripción que sale en el Excel, con el tipo de equipo
                    marcas_txt="",  # sin marca sugerida, es un punto de instalación
                    norma="RIC N°07 (7.2.8) / RIC N°13 (5.1)",  # artículos RIC que respaldan el punto de conexión
                    circuito=nombre_ac,  # se carga al circuito de agua caliente
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra, no es cable
                    longitud_m="1 unid",  # en la columna de largo va como 1 unidad
                    cantidad=1  # 1 punto por equipo
                )

            # 2) Prensaestopa en equipo — solo para conexión fija (sin enchufe);
            # con enchufe no hay cordón entrando a una caja que sellar
            if not con_enchufe_ac:  # solo si el equipo va conectado fijo, sin enchufe
                add_row(  # prensaestopa que sujeta el cordón al entrar al equipo
                    desc=f"Prensaestopa {_pg_ac} ({_mt_ac}) para cordón — entrada equipo (cable 3×{sec_txt_ac}mm^2)",  # el PG y la métrica ya se eligieron antes según la sección del cable
                    marcas_txt=marcas.get("Prensaestopa", ""),  # marca de prensaestopa que eligió el usuario
                    norma="RIC N°04 (5.15, 5.24) / RIC N°07 (5.2.8)",  # RIC 4 y 7: entradas de cable selladas
                    circuito=nombre_ac,  # se carga al circuito de agua caliente
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra
                    longitud_m="1 unid",  # 1 pieza
                    cantidad=1  # 1 prensaestopa por equipo
                )

            if not lleva_tab_ext:  # el equipo no pide tablero externo de desconexión
                continue  # si no tiene tablero externo, este circuito termina aquí

            # ── TABLERO EXTERNO DE DESCONEXIÓN ──────────────────────────────
            # RIC N°07 art. 7.2.8: tablero de comando a la vista del equipo
            # RIC N°11 sección 6: fuera de Volúmenes 0, 1 y 2

            # 3) Tablero sobrepuesto 6 puestos IP41
            # TM bipolar (2 puestos) + la bornera PE en riel: con 6 puestos sobra espacio
            add_row(  # gabinete del tablero de desconexión, a la vista del equipo
                desc="Tablero sobrepuesto de PVC 6 puestos IP41 (desconexión agua caliente)",  # 6 puestos alcanzan para el TM bipolar y las borneras
                marcas_txt=marcas.get("Tablero externo agua caliente", marcas.get("Tablero sobrepuesto", "")),  # marca del tablero externo, si no hay usa la del sobrepuesto
                norma="RIC N°02 (5.2, 6.1) / RIC N°07 (7.2.8, 7.3.3, 7.4.1, 7.4.2) / RIC N°11 (6, Vol.3)",  # RIC 2, 7 y 11: tablero, comando a la vista y fuera de volúmenes
                circuito=nombre_ac,  # va cargado al circuito de agua caliente
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m="1 unid",  # 1 tablero
                cantidad=1  # un tablero por equipo
            )

            # 4) Riel DIN 35×7,5mm tira de 10cm
            add_row(  # riel donde se montan el TM y las borneras
                desc="Riel DIN 35×7,5mm tira de 10cm (tablero externo agua caliente)",  # tira corta, lo justo para un tablero de 6 puestos
                marcas_txt=marcas.get("Riel DIN", ""),  # marca de riel DIN
                norma="RIC N°02 (6.1.15, 6.1.23)",  # RIC 2: montaje sobre riel
                circuito=nombre_ac,  # se carga al circuito de agua caliente
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m="1 unid",  # 1 tira
                cantidad=1  # una tira por tablero
            )

            # 5) TM bipolar 1P+N — mismo calibre que TM del tablero principal
            add_row(  # TM local, para cortar el equipo sin ir al tablero general
                desc=f"Disyuntor termomagnético 1P+N {in_tm_ac}A / 6kA / Curva C (desconexión local agua caliente)",  # misma corriente que el TM del tablero principal, curva C y 6kA
                marcas_txt=marcas.get("TM bipolar agua caliente", marcas.get("Protecciones", "")),  # marca del TM, si no hay usa la de protecciones
                norma="RIC N°07 (7.2.8, 7.3.4, 7.4.2, 7.4.5)",  # RIC 7: protección y desconexión local del equipo
                circuito=nombre_ac,  # se carga al circuito de agua caliente
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m="1 unid",  # 1 disyuntor
                cantidad=1  # un TM por tablero externo
            )

            # 6) Bornera PE tablero externo — la sección sale del conductor real
            #    del circuito (sec_born_pe, calculada más arriba). La fila cita
            #    RIC N°02 6.2.7 y RIC N°06 5.11/5.14, que son los de continuidad
            #    del conductor de protección.
            add_row(  # bornera de tierra dentro del tablero externo
                desc=f"Bornera de conexión PE {sec_born_pe_txt}mm^2 (tierra tablero externo agua caliente)",  # la sección sale del PE real del circuito, no de un fijo
                marcas_txt=marcas.get("Bornera PE agua caliente", marcas.get("Barra unipolar verde", "")),  # marca de bornera PE, si no hay usa la de barra verde
                norma="RIC N°02 (6.2.7) / RIC N°06 (5.11, 5.14)",  # RIC 2 y 6: bornera de tierra en el tablero
                circuito=nombre_ac,  # se carga al circuito de agua caliente
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m="1 unid",  # 1 bornera
                cantidad=1  # una bornera PE por tablero
            )

            # (la numeración de los ítems salta del 6 al 9: no existen un 7 ni un 8)
            # 9) Fijación tablero externo — 6 puntos para el tablero a la pared
            #    + 2 puntos para el riel DIN = 8 puntos en total, SIEMPRE del
            #    mismo tipo de tornillo (no tiene sentido mezclar 2 tipos para
            #    fijar la misma pieza). En volcanita/fibrocemento, además,
            #    los 8 puntos necesitan tarugo (no solo un tornillo suelto),
            #    porque ahí van montados los equipos de protección y pesan —
            #    sin tarugo el tornillo se puede salir. En madera u otro
            #    forrado sólido, el riel DIN no necesita tarugo (el tornillo
            #    muerde bien directo), por lo que van 8 tornillos simples.
            _amb_row_ac = _get_amb_row(ambientes_df, _datos_ac.get("ambientes_str", nombre_ac))  # datos del ambiente donde va el tablero
            _mat_amb_ac = str(_amb_row_ac.get("Material forrado interior", "")).strip().lower() if _amb_row_ac is not None else ""  # material del muro, en minúsculas
            if ("volcan" in _mat_amb_ac) or ("vulcan" in _mat_amb_ac) or ("fibro" in _mat_amb_ac):  # volcanita o fibrocemento: el tornillo solo se sale, va con tarugo
                # 6 (tablero a la pared) + 2 (riel DIN) = 8 puntos, todos con tarugo
                add_row(  # los tarugos de los 8 puntos de fijación
                    desc="Tarugo paloma 6mm",  # tarugo paloma, el que se abre por detras de la plancha
                    marcas_txt=marcas.get("Tarugo paloma", ""),  # marca de tarugo paloma
                    norma="-",  # tornillería, no tiene artículo RIC
                    circuito=nombre_ac,  # se carga al circuito de agua caliente
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra
                    longitud_m="8 unid",  # 6 del tablero + 2 del riel DIN
                    cantidad=8  # 8 puntos de fijación
                )
                add_row(  # el tornillo que va dentro de cada tarugo paloma
                    desc='Tornillo volcanita punta fina 6x1 1/4"',  # tornillo de volcanita, punta fina
                    marcas_txt=marcas.get("Tornillo para tarugo paloma", ""),  # marca del tornillo que acompana al tarugo paloma
                    norma="-",  # tornillería, no tiene artículo RIC
                    circuito=nombre_ac,  # se carga al circuito de agua caliente
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra
                    longitud_m="8 unid",  # uno por cada tarugo
                    cantidad=8  # 8 tornillos, igual que los tarugos
                )
            else:  # madera u otro forrado sólido: el tornillo muerde directo, sin tarugo
                # Madera u otro forrado sólido: NO tiene sentido usar 2 tipos
                # de tornillo distintos (6 para el tablero + 2 "para tarugo"
                # para el riel DIN) si no hay tarugo de por medio en ningún
                # punto. Se unifica todo en 1 solo tipo de tornillo, 8 unidades
                # (6 tablero + 2 riel DIN), igual que en volcanita/fibrocemento
                # donde los 8 puntos también son del mismo tipo (ahí con tarugo).
                if "madera" in _mat_amb_ac:  # forrado de madera: tornillo punta fina cabeza lenteja
                    _desc_torn_ext = 'Tornillo punta fina para madera cabeza lenteja 6x1/2"'  # esa es la descripción que se cotiza
                else:  # no se pudo identificar el forrado del muro
                    _desc_torn_ext = 'Tornillo (definir según forrado)'  # no se pudo saber el material del muro
                add_row(  # los 8 tornillos, todos del mismo tipo
                    desc=_desc_torn_ext,  # la descripción depende del forrado
                    marcas_txt=marcas.get("Tornillos", ""),  # marca genérica de tornillos
                    norma="Instalación",  # es tema de montaje, no de norma
                    circuito=nombre_ac,  # se carga al circuito de agua caliente
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra
                    longitud_m="8 unid",  # 6 del tablero + 2 del riel DIN
                    cantidad=8  # 8 puntos de fijación
                )

            # 10) Nota normativa obligatoria en informe
            add_row(  # fila de aviso, no es material: queda escrita en el informe
                desc=(  # el texto es largo, se arma en varios pedazos
                    f"NOTA NORMATIVA — Tablero externo {nombre_ac}: "
                    f"Instalar FUERA de Volúmenes 0, 1 y 2 (RIC N°11 sección 6). "
                    f"Debe quedar a la vista directa del equipo (RIC N°07 art. 7.4.2). "
                    f"IP mínimo 41 en interior seco, IP44 en ambiente húmedo."
                ),
                marcas_txt="",  # es una nota, no lleva marca
                norma="RIC N°07 (7.2.8, 7.3.3, 7.4.1, 7.4.2) / RIC N°11 (6, Tabla Volúmenes)",  # artículos que respaldan la nota
                circuito=nombre_ac,  # se carga al circuito de agua caliente
                unidad="—",  # el guion marca que no se cotiza
                k=1,  # sin factor extra
                longitud_m="—",  # no tiene largo, es solo texto
                cantidad=""  # cantidad vacía para que no sume en los totales
            )
    # ─────────────────────────────────────────────────────────────────────────
    # PRENSAESTOPA CLIMATIZACIÓN — mismo criterio que agua caliente y especiales:
    # con enchufe (TM≤16A) no necesita, porque el cordón termina en un enchufe,
    # no entra a una caja. Sin enchufe (TM>16A) sí necesita, porque el cordón
    # del equipo entra a la caja de derivación y hay que sujetarlo/sellarlo ahí
    # (la conexión eléctrica en sí, dentro de la caja, ya se hace con cónico —
    # eso no cambia, esto es solo la entrada física del cordón).
    # ─────────────────────────────────────────────────────────────────────────
    for _cl_pe in _clima_items:  # recorre los equipos de climatización que se juntaron más arriba
        # solo interesa el equipo si va canalización embutida (es_emb) y además
        # se sabe con certeza que NO lleva enchufe (con_enchufe es False, no None)
        if _cl_pe["es_emb"] and _cl_pe["con_enchufe"] is False:  # con_enchufe es False solo si se sabe seguro que no lleva enchufe
            # busca en circuitos_df la fila cuyo nombre de circuito coincide con este equipo de clima
            _fila_cl_pe = circuitos_df[circuitos_df["Circuito"].astype(str).str.contains(  # busca en la tabla la fila del circuito de este equipo
                _cl_pe["circ"].split("(")[0].strip(), regex=False, na=False  # compara contra el nombre sin lo que va entre paréntesis
            )]  # cierra el filtro, queda el subconjunto de filas que calzan con el nombre
            sec_cl_pe = extraer_seccion_mm2(str(_fila_cl_pe["Conductor"].iloc[0])) if len(_fila_cl_pe) > 0 else None  # sección del conductor de ese circuito, en mm^2
            sec_cl_pe = sec_cl_pe or 2.5  # si no se encontró la sección, usa 2,5mm^2 por defecto
            _pg_cl, _mt_cl = ("PG11", "M16")  # valor por defecto
            for _sec_k, (_pg_k, _mt_k) in sorted(PRENSAESTOPA_POR_SECCION.items()):  # recorre la tabla de prensaestopas de menor a mayor
                if sec_cl_pe <= _sec_k + 0.01:  # el 0.01 es para que no falle por decimales
                    _pg_cl, _mt_cl = _pg_k, _mt_k  # primer tamaño de la tabla que alcanza
                    break  # ya encontró el tamaño que sirve, corta la búsqueda
            sec_txt_cl = str(round(sec_cl_pe, 1)).replace(".", ",")  # sección con coma decimal
            # agrega la fila del prensaestopa para la entrada del equipo de climatización
            add_row(  # fila del prensaestopa del equipo de climatización
                desc=f"Prensaestopa {_pg_cl} ({_mt_cl}) para cordón — entrada equipo (cable 3×{sec_txt_cl}mm^2)",  # la descripción muestra el PG, la métrica y el cable
                marcas_txt=marcas.get("Prensaestopa", ""),  # marca de prensaestopa
                norma="RIC N°04 (5.15, 5.24) / RIC N°07 (5.2.8)",  # RIC 4 y 7: entradas de cable selladas
                circuito=_cl_pe["circ"],  # se carga al circuito de ese equipo de clima
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m="1 unid",  # 1 pieza
                cantidad=1  # un prensaestopa por equipo
            )

    # CIRCUITOS ESPECIALES GENÉRICOS (horno, lavadora, encimera, etc.):
    # punto de conexión (enchufe o directo si supera 16A) para cada equipo especial
    if {"Circuito", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # solo entra a este bloque si el DataFrame tiene esas 2 columnas
        _esp_acum = {}  # {desc: {"total": N, "detalle": {"circuito y ambiente": cant}}}

        # Junta, para cada descripción de material especial, el total y el
        # detalle (que circuito/ambiente aporto cuántas unidades)
        # qué recibe:
        #   desc ......... descripción del material especial
        #   detalle_txt .. de qué circuito y ambiente viene
        #   cant ......... cuántas unidades sumar
        # no devuelve nada: va llenando el diccionario _esp_acum.
        def _acum_especial(desc, detalle_txt, cant):  # va sumando cantidades por descripción de material
            if desc not in _esp_acum:  # primera vez que aparece este material
                _esp_acum[desc] = {"total": 0, "detalle": {}}  # parte el total en 0 y el detalle vacío
            _esp_acum[desc]["total"] += cant  # suma al total general de ese material
            _esp_acum[desc]["detalle"][detalle_txt] = _esp_acum[desc]["detalle"].get(detalle_txt, 0) + cant  # y suma también en la línea de ese circuito y ambiente

        for _, rr in circuitos_df.iterrows():  # recorre circuito por circuito de la tabla
            nombre_circ = str(rr.get("Circuito", "")).strip()  # nombre del circuito, ej: 'Circuito 5 - Cocina'

            # Excluir circuitos de climatización
            es_clima_row = any(k in nombre_circ.lower() for k in ("climatiz", "aire acond", "split"))  # True si el nombre del circuito menciona climatización
            if es_clima_row:  # climatización se calcula en otra parte
                continue  # pasa al siguiente circuito
            # Excluir circuitos de agua caliente (tienen su propia lógica de materiales)
            # (climatización y agua caliente tienen su propio cálculo de
            # materiales en otra parte del programa, por eso se saltan aquí)
            es_agua_row = any(k in nombre_circ.lower() for k in  # busca palabras de agua caliente en el nombre del circuito
                              ("ducha", "termo", "calefon", "calefón", "calentador", "agua caliente"))  # así se suele escribir el equipo de agua caliente
            if es_agua_row:  # agua caliente ya se resolvió en el bloque de arriba
                continue  # pasa al siguiente circuito

            tm_txt = rr.get("Disyuntor termomagnético", "")  # texto del TM tal como está en la tabla, ej: '1x16A'
            in_tm = parse_in_tm(tm_txt)  # convierte el texto del TM a número (amperes), o None si no se pudo leer
            items = (items_por_nombre or {}).get(nombre_circ, [])  # lista de equipos/enchufes/luminarias que el usuario cargo para este circuito
            if not isinstance(items, list):  # si vino algo que no es lista
                items = []  # se deja vacía para no reventar más abajo
            if not nombre_circ or len(items) == 0:  # sin nombre o sin items no hay nada que cotizar
                continue  # pasa al siguiente circuito
            items_esp = [it for it in items if isinstance(it, dict) and ("nombre" in it)]  # solo los items "especiales"
            if not items_esp:  # el circuito no trae equipos especiales
                continue  # pasa al siguiente circuito
            # Revisa cada equipo especial del circuito: según el TM decide
            # si va con enchufe dedicado o con conexión fija + prensaestopa
            for it in items_esp:  # revisa equipo por equipo del circuito
                amb = str(it.get("amb", "")).strip() or "Sin ambiente"  # ambiente donde esta el equipo, ej: 'Cocina'
                equipo = str(it.get("nombre", "")).strip() or "equipo especial"  # nombre del equipo, ej: 'Horno empotrado'
                tm_val = in_tm if in_tm is not None else 16  # si no se pudo leer el TM, se asume 16A por defecto
                detalle_txt = f"{nombre_circ} -> {amb}"  # texto para identificar de donde sale este material
                if tm_val > 16:  # sobre 16A no hay enchufe domiciliario que aguante, va conexión fija
                    # más de 16A: no lleva enchufe, va a conexión directa + prensaestopa
                    desc = f"Punto para conexión directa de {equipo} (>16A) - sin enchufe"  # material 1: el punto de conexión fija en si
                    _acum_especial(desc, detalle_txt, 1)  # guarda el punto de conexión en el acumulador
                    # Prensaestopa según sección del conductor
                    sec_esp = extraer_seccion_mm2(str(rr.get("Conductor", "")))  # de que grosor es el cable para elegir el prensaestopa correcto
                    desc_pg = prensaestopa_para_seccion(sec_esp) if sec_esp else "Prensaestopa PG13,5 (M20) para cordón (cable 3x4,0mm^2)"  # material 2: el prensaestopa que sella la entrada del cable a la caja
                    _acum_especial(desc_pg, detalle_txt, 1)  # guarda el prensaestopa en el acumulador
                else:  # hasta 16A si se puede dejar enchufe dedicado
                    # 16A o menos: sí lleva enchufe dedicado — siempre simple,
                    # porque un enchufe "dedicado" es para 1 solo equipo
                    nominal = "10A" if tm_val <= 10 else "10/16A"  # TM de 10A o menos: enchufe de 10A; si no, enchufe 10/16A
                    desc = f"Enchufe dedicado para {equipo} simple 2P+T {nominal} 250V"  # el enchufe dedicado en si
                    _acum_especial(desc, detalle_txt, 1)  # guarda el enchufe en el acumulador

        # emite 1 sola fila por descripción, sumando todos los circuitos/ambientes
        # Recorre lo acumulado y escribe 1 fila por cada material distinto,
        # con el total sumado y el detalle de circuitos/ambientes en el texto
        for desc, info in sorted(_esp_acum.items(), key=lambda x: x[0].lower()):  # recorre lo acumulado, ordenado por nombre de material
            detalle_txt = ", ".join(f"{k} ({v})" for k, v in info["detalle"].items())  # arma el texto '(circuito y ambiente) (cantidad), ...' para la columna Circuito
            if "prensaestopa" in desc.lower():  # si el material es un prensaestopa usa su norma y marca; si no, la de enchufes
                marca_txt, norma_txt = marcas.get("Prensaestopa", ""), "RIC 4.5.1"  # si es prensaestopa usa marca y norma de prensaestopa
            else:  # cualquier otra cosa acumulada es un enchufe dedicado
                marca_txt, norma_txt = marcas.get("Enchufes", ""), "RIC 13.5.1"  # marca y norma de enchufes
            # agrega la fila del material al listado final
            add_row(  # una sola fila por material, con el total sumado
                desc=desc,  # la descripción tal cual quedó en el acumulador
                marcas_txt=marca_txt,  # marca elegida arriba
                norma=norma_txt,  # norma elegida arriba
                circuito=detalle_txt,  # en la columna circuito va el detalle de donde salió cada unidad
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m=f"{info['total']} unid",  # en la columna de largo va el total como unidades
                cantidad=info["total"]  # total sumado de todos los circuitos
            )

    # Función chica de formato: para mostrar la sección del conductor
    # con coma decimal (2,5 en vez de 2.5), como se escribe en Chile
    def texto_seccion(sec):  # recibe la sección en mm^2 y la devuelve como texto
        # convierte 2.5 en "2,5" para que se vea con coma como acá en Chile
        return str(sec).replace(".", ",")

    # Cuántas borneras (de una sección de conductor dada) alcanzan a
    # entrar en 1 puesto del riel DIN del tablero. A mayor sección,
    # el cable es más grueso y entran menos borneras por puesto.
    def borneras_por_puesto(sec):  # sec viene en mm^2
        """Cuántas borneras de una sección caben en 1 puesto del riel DIN."""
        if sec <= 1.5:  return 4  # conductores finos: caben 4 borneras por puesto
        if sec <= 2.5:  return 3  # 3 borneras por puesto
        if sec <= 4.0:  return 3  # 3 borneras por puesto
        if sec <= 6.0:  return 2  # 2 borneras por puesto
        return 1  # 10mm^2 y mayor

    # Nota: el cálculo de borneras de neutro/organización vive más abajo,
    # en _borneras_neutro_acum / _borneras_organizacion_acum (RIC 2, 6.2.12).

    # =========================
    # CONECTORES CÓNICOS (POR SECCIÓN)
    # Qué resuelve el bloque: cuántos conectores cónicos hay que comprar y de
    # qué color/número, recorriendo circuito por circuito de circuitos_df.
    # Reglas:
    #   - 3 por enchufe intermedio (el último de cada ambiente: 0)
    #   - 3 por CADA interruptor del ambiente (son los de su caja troncal)
    #   - 2 por octogonal intermedia (la última del grupo: 0; conmutadas: 3)
    #   - 3 por caja de derivación adicional entre ambientes
    # El color y el número salen de conico_por_seccion(), que mira la sección
    # del conductor del circuito y cuántos cables se unen en el punto (2 o 3).
    # Casos que NO generan cónico y se resuelven por otro lado:
    #   - conductor >6mm^2: no existe cónico de ese porte, va soldado con
    #     estaño; se anota en n_circ_estaño / circ_estaño_txt y se cobra en el
    #     bloque "ESTAÑO, PASTA Y CINTAS" de más abajo.
    #   - equipo con enchufe (TM<=16A): no lleva cónico, se conecta a los
    #     bornes del enchufe; eso se cobra como ferrules en los bloques C2
    #     (climatización), D2 (agua caliente) y D3 (especiales genéricos).
    #   - TM ilegible: se acumula una fila de aviso "(definir)" en vez de un
    #     conector real, para que alguien lo revise a mano.
    # =========================
    add_section("Conectores cónicos")  # título de la sección en el listado de materiales del Excel
    conicos_por_color = {}  # key=(color, num, rango) -> {"cantidad":int, "circuitos":set()}; acá también caen las filas de aviso "(definir)"
    # Suma cuántos conectores cónicos de cada tipo (color/número/rango)
    # se necesitan en total, y va anotando de que circuitos vienen.
    def _accum_conico(color, num, rango, cantidad, circuito_nombre):
        # va sumando cuántos conectores cónicos de cada color/tamaño se
        # necesitan, y anota de qué circuito son
        key = (color, num, rango)  # identifica el tipo de conector cónico
        if key not in conicos_por_color:  # todavía no existe este tipo de conector: hay que crear su entrada
            conicos_por_color[key] = {"cantidad": 0, "circuitos": set()}  # primera vez que aparece
        conicos_por_color[key]["cantidad"] += int(max(0, cantidad))  # suma la cantidad
        if circuito_nombre:  # si vino el nombre del circuito, lo agrega al detalle
            conicos_por_color[key]["circuitos"].add(str(circuito_nombre))  # anota de qué circuito viene
    # Contadores que después alimentan el bloque "ESTAÑO, PASTA Y CINTAS"
    n_circ_estaño = 0  # cuántos circuitos van con estaño en vez de cónico (conductor >6mm^2, incluidos los equipos de conexión fija)
    circ_estaño_txt = []  # nombres de esos circuitos, para el detalle en el Excel
    _n_focos_led_estaño = 0  # cuántas luminarias con cable (LED, panel, tubo) se estañan, sumando todos los circuitos
    _circ_focos_led = []  # nombres de los circuitos con luminarias LED que se estañan
    if {"Circuito", "Conductor"}.issubset(set(circuitos_df.columns)):  # solo se puede calcular si el DataFrame tiene estas 2 columnas
        # Recorre cada circuito y calcula cuántos conectores cónicos
        # necesita, según si es enchufe, luminaria o equipo especial
        for _, rr in circuitos_df.iterrows():  # fila por fila de la tabla de circuitos
            circ_name = str(rr.get("Circuito", "")).strip()  # nombre del circuito
            cond_txt = str(rr.get("Conductor", "")).strip()  # texto del conductor, ej: 'THHN 2.5mm^2'
            items = (items_por_nombre or {}).get(circ_name, [])  # items (enchufes/luminarias/equipos) que el usuario cargo para este circuito
            if not isinstance(items, list):  # si no vino lista
                items = []  # se deja vacía y sigue
            secc = extraer_seccion_mm2(cond_txt)  # sección del conductor en mm^2, sacada del texto

            # Climatización, agua caliente y especiales (horno, lavadora, etc.) tienen
            # su propia lógica de conexión — se resuelven acá, antes de que el código
            # más abajo salte este circuito por no tener sus datos en "items"
            # (esos casos guardan sus datos aparte, en _clima_items, no en items_por_nombre)
            es_clima_row    = any(k in circ_name.lower() for k in ("climatiz","aire","split","ac ","a/c"))  # el nombre del circuito menciona climatización
            es_agua_row     = any(k in circ_name.lower() for k in ("ducha","termo","calefon","calefón","calentador","agua caliente"))  # el nombre del circuito menciona agua caliente
            es_especial_row = any(k in circ_name.lower() for k in ("especial","horno","encimera","lavadora","lavaplatos","secadora","jacuzzi","piscina"))  # el nombre del circuito menciona un equipo especial (horno, lavadora, etc.)
            _especial_generico = es_especial_row and not es_agua_row and not es_clima_row  # especial 'puro': no es agua caliente ni climatización (esos ya tienen su propio bloque arriba)
            # Agua caliente: si no se puede leer el TM se avisa para revisar
            # a mano; si tiene enchufe (TM<=16A) no lleva cónico — la conexión
            # son los bornes del enchufe, y eso se cobra como ferrules en D2
            if es_agua_row:  # circuito de agua caliente
                _in_tm_ac_row = parse_in_tm(str(rr.get("Disyuntor termomagnético", "")))  # TM del circuito de agua caliente
                if _in_tm_ac_row is None:  # no se pudo leer el TM: no se adivina, se deja aviso para revisar a mano
                    _accum_conico("(definir)", None, "", 1, f"{circ_name} - TM no legible, revisar con/sin enchufe manualmente")  # fila de aviso '(definir)' en vez de un conector real
                    continue  # no se cuenta nada hasta que alguien lo revise
                if _in_tm_ac_row <= 16:  # con enchufe: no lleva cónico, se conecta a los bornes del enchufe
                    continue  # los 3 ferrules F/N/PE de ese enchufe se agregan en el bloque D2, más abajo
                es_especial_row = True  # sin enchufe (TM>16A) = conexión fija, mismo caso que horno/lavadora
            # Climatización: mismo criterio que agua caliente, pero busca
            # el equipo en _clima_items para saber si tiene enchufe o no
            if es_clima_row:  # circuito de climatización
                _circ_base_cl = circ_name.split("(")[0].strip()  # nombre del circuito sin el texto entre paréntesis, para buscarlo
                _cl_match_cl = next((_cl for _cl in _clima_items if _cl["circ"].split("(")[0].strip() == _circ_base_cl), None)  # busca el equipo de climatización que corresponde a este circuito
                if _cl_match_cl is not None and _cl_match_cl["con_enchufe"] is None:  # no se sabe si tiene enchufe: aviso para revisar a mano
                    _accum_conico("(definir)", None, "", 1, f"{circ_name} - TM no legible, revisar con/sin enchufe manualmente")  # fila de aviso '(definir)'
                    continue  # no se cuenta nada hasta que alguien lo revise
                if _cl_match_cl is not None and _cl_match_cl["con_enchufe"]:  # el equipo va enchufado
                    continue  # ya se cuenta en la sección de ferrules del enchufe (C2)
                es_especial_row = True  # climatización sin enchufe = conexión fija, mismo caso que horno/lavadora
            if _especial_generico:  # horno, lavadora, encimera y esos
                # Mismo criterio que agua caliente y climatización (arriba): TM
                # decide con/sin enchufe, no la potencia. Si tiene enchufe (TM≤16A),
                # se salta acá, los ferrules del enchufe van en su propio bloque
                # (D3, más abajo en el archivo). Si no tiene enchufe (TM>16A), sigue
                # como es_especial_row=True y cae al tratamiento de conexión fija.
                # Si el TM no se puede leer, no se adivina — se deja un aviso
                # "(definir)" para que se revise a mano, igual que ya se hace
                # cuando falla la columna Conductor.
                _in_tm_esp_row = parse_in_tm(str(rr.get("Disyuntor termomagnético", "")))  # TM del equipo especial
                if _in_tm_esp_row is None:  # no se pudo leer el TM: aviso para revisar a mano
                    _accum_conico("(definir)", None, "", 1, f"{circ_name} - TM no legible, revisar con/sin enchufe manualmente")  # fila de aviso '(definir)'
                    continue  # no se cuenta nada hasta que alguien revise el TM
                if _in_tm_esp_row <= 16:  # con enchufe (TM<=16A): no lleva cónico, se conecta a los bornes del enchufe
                    continue  # los 3 ferrules F/N/PE de ese enchufe se agregan en el bloque D3, más abajo
            if es_especial_row:  # equipo de conexión fija: se empalma directo al cable
                # Conexión fija = cola de rata + soldadura/cónico (NO bornera con
                # ferrule — el equipo no tiene bornes, se empalma directo al cable).
                # Se suma acá mismo (antes de que se cierren los totales de cónico/
                # estaño/cintas más abajo) para no generar un rollo aparte por poco.
                # El corte "con/sin enchufe" ya se resolvió arriba por TM, igual
                # criterio que agua caliente y climatización — este bloque solo
                # se alcanza para circuitos sin enchufe (TM>16A), por lo que el
                # cónico/estaño se genera siempre, sin límite de potencia.
                if secc is not None and secc > 6.0:  # conductor grueso: se suelda con estaño en vez de cónico
                    n_circ_estaño += 1  # suma 1 circuito más a la lista de los que se sueldan con estaño
                    circ_estaño_txt.append(f"{circ_name} - conexión fija equipo")  # guarda el nombre para el detalle en el Excel
                else:  # conductor de 6mm^2 o menos (o sección ilegible: ahí conico_por_seccion devuelve "(definir)")
                    color_esp, num_esp, rango_esp = conico_por_seccion(secc, n_cables=3)  # conductor normal: 3 conectores cónicos (fase, neutro, tierra) para la conexión fija
                    _accum_conico(color_esp, num_esp, rango_esp, 3, f"{circ_name} - conexión fija equipo")  # guarda los 3 cónicos en el acumulador
                continue  # ya se resolvió este circuito (especial), pasa al siguiente

            # Circuitos con conductor >6mm^2 se conectan con estaño, no con conector cónico
            if secc is not None and secc > 6.0:  # mismo criterio que arriba: conductor grueso va con estaño, no con cónico
                n_circ_estaño += 1  # suma 1 circuito más a la lista de los que se sueldan con estaño
                circ_estaño_txt.append(circ_name)  # guarda el nombre para el detalle en el Excel
                continue  # ese circuito no lleva cónicos, va todo con estaño

            if not isinstance(items, list) or len(items) == 0:  # circuito sin items cargados: no hay nada que calcular, pasa al siguiente
                continue  # pasa al siguiente circuito

            # Contar enchufes y luminarias por ambiente
            n_ench_circ = 0  # cuenta cuántos enchufes tiene el circuito
            n_lum_estaño_circ = 0  # cuenta cuántas luminarias del circuito necesitan estaño
            lum_by_amb_con = {}  # {amb_key: [items_lum]}  # agrupa las luminarias por ambiente, para calcular los cónicos por grupo

            # Recorre los items del circuito y separa: enchufes, equipos
            # especiales (se saltan, ya se resolvieron arriba) y luminarias
            for it in items:  # revisa item por item del circuito
                if not isinstance(it, dict):  # no es un dict válido: se salta
                    continue  # item malo, se salta
                if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it):  # es un enchufe (tiene alguna de estas claves)
                    n_ench_circ += int(it.get("n_ench", 1) or 1)  # suma la cantidad de enchufes de este item
                    continue  # ya contado como enchufe, sigue con el siguiente
                if "nombre" in it:  # es un equipo especial: ya se contó arriba, se salta acá
                    continue  # sigue con el siguiente item
                # Es luminaria
                amb = str(it.get("amb", "")).strip().lower() or "sin_amb"  # ambiente de la luminaria, en minúsculas para agrupar bien
                if amb not in lum_by_amb_con:  # primera luminaria de ese ambiente
                    lum_by_amb_con[amb] = []  # primera luminaria de este ambiente: crea la lista
                lum_by_amb_con[amb].append(it)  # agrega la luminaria al grupo de su ambiente
                if _lum_necesita_estaño(it.get("tipo_lum", ""), it.get("desc_lum", "")):  # marca si esta luminaria necesita estaño en vez de cónico
                    n_lum_estaño_circ += 1  # esta luminaria se conecta con estaño, no cónico

            # ── ILUMINACIÓN: cónicos ──────────────────────────────────────────
            # Caja troncal:            +3 cónicos POR CADA interruptor del ambiente
            # Caja octogonal intermedia (no última del grupo): +2 cónicos
            # Caja octogonal última del grupo: 0 cónicos
            # Caja interruptor (9/12/15/32/24): 0 cónicos
            #
            # De TODOS los cónicos de iluminación del circuito, los últimos 3
            # se unen con 2 cables (van con la tabla de 2 cables), el resto
            # (total - 3) se une con 3 cables (tabla de 3 cables).
            if lum_by_amb_con:  # hay luminarias agrupadas por ambiente que resolver
                con_ilum = 0  # total de conectores cónicos de iluminación para este circuito
                for amb_key, lums_amb in lum_by_amb_con.items():  # recorre cada ambiente con sus luminarias
                    n_lum_amb = len(lums_amb)  # cuántas luminarias tiene este ambiente
                    n_conm_amb = int(min(n_lum_amb, ncon_map_local.get(amb_key, 0)))  # conmutadas del ambiente: viene de la columna "N_conmutadas_924 (u)" de ambientes_df (ncon_map_local, armado más arriba); se topa al total de luminarias de ese ambiente
                    n_rest_amb = n_lum_amb - n_conm_amb  # el resto: luminarias no conmutadas
                    _c32 = n_rest_amb // 3  # grupos completos de 3 luminarias no conmutadas (grupo tipo 9/32)
                    _rem = n_rest_amb % 3  # luminarias no conmutadas que sobran después de armar grupos de 3
                    _c15 = 1 if _rem == 2 else 0  # sobran 2: forman un grupo tipo 9/15
                    _c12 = 1 if _rem == 1 else 0  # sobra 1: forma un grupo tipo 9/12

                    # Caja troncal: 3 por CADA interruptor NO conmutado del
                    # ambiente (grupos 9/12/15/32), más 3 si el ambiente tiene
                    # conmutadas — siempre 1 sola caja troncal para el par
                    # conmutado, sin importar cuántas luminarias conmutadas
                    # haya (mismo criterio de "1 par por ambiente" que ya se
                    # usa en el cálculo de chicotes para cinta aislante).
                    n_interruptores_amb = (1 if n_conm_amb > 0 else 0) + _c12 + _c15 + _c32  # cuenta cuántos interruptores tiene el ambiente (conmutado + grupos 9/12/15/32)
                    con_ilum += 3 * n_interruptores_amb  # 3 cónicos por cada interruptor, para la caja troncal

                    # Cajas octogonales intermedias: última del grupo: 0
                    # Grupo conmutadas 9/24: 3 cónicos c/u (lleva 1 cable extra
                    # por el conductor rojo que sigue hacia la siguiente
                    # luminaria conmutada), entonces son n_conm_amb-1 intermedias
                    # Grupos no conmutadas: 2 cónicos c/u, descomponer en c12/c15/c32
                    #   9/32 (3 lum): 2 intermedias · 9/15 (2 lum): 1 intermedia · 9/12 (1 lum): 0 intermedias
                    if n_conm_amb > 1:  # con 2 o más conmutadas hay cajas octogonales en el medio
                        con_ilum += 3 * (n_conm_amb - 1)  # 3 cónicos por cada caja octogonal intermedia del par conmutado
                    if n_rest_amb > 0:  # hay luminarias no conmutadas, van con sus propias cajas intermedias
                        con_ilum += 2 * (_c32 * 2)  # 9/32: 2 intermedias por grupo
                        con_ilum += 2 * (_c15 * 1)  # 9/15: 1 intermedia por grupo
                        # 9/12: 0 intermedias

                if con_ilum > 0:  # el circuito quedó con cónicos de iluminación por comprar
                    con_ilum_2cables = min(3, con_ilum)  # de todos los cónicos de iluminación, los últimos 3 se unen con 2 cables
                    con_ilum_3cables = max(0, con_ilum - 3)  # el resto (si sobra algo) se une con 3 cables
                    if con_ilum_3cables > 0:  # hay cónicos de 3 cables: busca el color/número según la sección
                        color_i3, num_i3, rango_i3 = conico_por_seccion(secc, n_cables=3)  # cónico de iluminación, versión de 3 cables
                        _accum_conico(color_i3, num_i3, rango_i3, con_ilum_3cables, f"{circ_name} (iluminación, 3 cables)")  # guarda los cónicos de iluminación (3 cables) en el acumulador
                    if con_ilum_2cables > 0:  # hay cónicos de 2 cables: busca el color/número según la sección
                        color_i2, num_i2, rango_i2 = conico_por_seccion(secc, n_cables=2)  # cónico de iluminación, versión de 2 cables
                        _accum_conico(color_i2, num_i2, rango_i2, con_ilum_2cables, f"{circ_name} (iluminación, 2 cables)")  # guarda los cónicos de iluminación (2 cables) en el acumulador

            # ── ENCHUFES: cónicos ─────────────────────────────────────────────
            # Enchufe intermedio: +3 cónicos
            # Último enchufe de CADA ambiente (no solo el último de todo el
            # circuito): 0 cónicos
            # Caja adicional entre ambientes: +3 cónicos
            # Enchufes siempre se unen con 3 cables (F/N/T)
            if n_ench_circ > 0:  # este circuito tiene enchufes, se le calculan sus cónicos
                _n_ench_por_amb_conico = {}  # cuenta enchufes por ambiente, para saber cual es el 'último' de cada uno
                for it in items:  # revisa cada item cargado del circuito buscando los enchufes
                    if not isinstance(it, dict): continue  # no es un dict: se salta
                    if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it):  # es un enchufe
                        _a = str(it.get("amb", "")).strip().lower() or "sin_amb"  # ambiente de este enchufe, en minúsculas
                        _n_ench_por_amb_conico[_a] = _n_ench_por_amb_conico.get(_a, 0) + int(it.get("n_ench", 1) or 1)  # suma la cantidad de enchufes de este ambiente
                _cajas_adic_circ = int((cajas_adic_por_nombre or {}).get(circ_name, 0))  # cuántas cajas de derivación adicionales tiene este circuito
                con_ench = sum(3 * max(0, n_amb_ench - 1) for n_amb_ench in _n_ench_por_amb_conico.values())  # 3 cónicos por cada enchufe intermedio (todos menos el último de cada ambiente)
                con_ench += 3 * _cajas_adic_circ  # más 3 cónicos por cada caja adicional entre ambientes
                if con_ench > 0:  # quedaron cónicos de enchufe que anotar
                    color, num, rango = conico_por_seccion(secc, n_cables=3)  # cónico de enchufes: siempre 3 cables (fase, neutro, tierra)
                    _accum_conico(color, num, rango, con_ench, f"{circ_name} (enchufes)")  # guarda los cónicos de enchufes en el acumulador


            # Acumular focos LED para estaño
            if n_lum_estaño_circ > 0:  # si este circuito tiene focos LED que necesitan estañarse
                _n_focos_led_estaño += n_lum_estaño_circ  # suma al contador general de focos LED con estaño
                _circ_focos_led.append(circ_name)  # guarda el nombre del circuito para mostrarlo después en la fila de material
    # si por algún motivo no se generó ningún cónico (columnas faltantes, etc.),
    # se hace una estimación aproximada en base a los totales generales
    # (sin dato de sección real disponible acá, se asume la típica: 2.5mm^2
    # para enchufes y 1.5mm^2 para iluminación)
    if not conicos_por_color:  # no se generó ningún cónico circuito por circuito: se hace una estimación general
        # total_enchufes y total_luminarias son los conteos por ambiente que se
        # hicieron mucho más arriba en esta misma función (bloque ACCESORIOS)
        if total_enchufes > 0:  # hay enchufes en la instalación
            _total_cajas_adic = sum(int(v) for v in (cajas_adic_por_nombre or {}).values())  # cuenta cuántas cajas de derivación adicionales hay en total
            # Enchufes: 3*(n_ench-1) + 3*cajas_adic, siempre 3 cables
            _con_ench_fb = 3 * max(0, total_enchufes - 1) + 3 * _total_cajas_adic  # estimado: 3 cónicos por enchufe intermedio más 3 por cada caja adicional
            if _con_ench_fb > 0:  # solo si de verdad hay conexiones de enchufe que resolver
                _color_efb, _num_efb, _rango_efb = conico_por_seccion(2.5, n_cables=3)  # cónico que corresponde a sección 2.5mm^2 (la típica de enchufes)
                _accum_conico(_color_efb, _num_efb, _rango_efb, _con_ench_fb, "Enchufes (estimado)")  # suma esos cónicos estimados al total
        if total_luminarias > 0:  # hay luminarias en la instalación
            # Sin info de ambientes, se aproxima con 1 ambiente por circuito:
            # troncal (3×n_interruptores) + octogonales intermedias
            # (no se conoce cuántas son conmutadas, se asume 0)
            _c32_fb = total_luminarias // 3  # grupos de 3 luminarias (una caja tipo 32 por grupo)
            _rem_fb = total_luminarias % 3  # luminarias que sobran fuera de los grupos de 3
            _c15_fb = 1 if _rem_fb == 2 else 0  # sobran 2: se agrega 1 caja tipo 15
            _c12_fb = 1 if _rem_fb == 1 else 0  # sobra 1: se agrega 1 caja tipo 12
            _n_int_fb = _c12_fb + _c15_fb + _c32_fb  # total de interruptores estimados
            _con_ilum_fb = 3 * _n_int_fb + 2 * (_c32_fb * 2 + _c15_fb * 1)  # troncal + oct intermedias
            if _con_ilum_fb > 0:  # solo si hay conexiones de iluminación que resolver
                _con_ilum_fb_2 = min(3, _con_ilum_fb)  # como máximo 3 conexiones se unen con 2 cables
                _con_ilum_fb_3 = max(0, _con_ilum_fb - 3)  # el resto se une con 3 cables
                if _con_ilum_fb_3 > 0:  # hay conexiones de 3 cables
                    _color_i3fb, _num_i3fb, _rango_i3fb = conico_por_seccion(1.5, n_cables=3)  # cónico para 1.5mm^2 (iluminación) con 3 cables
                    _accum_conico(_color_i3fb, _num_i3fb, _rango_i3fb, _con_ilum_fb_3, "Iluminación (estimado, 3 cables)")  # suma esos cónicos estimados
                if _con_ilum_fb_2 > 0:  # hay conexiones de 2 cables
                    _color_i2fb, _num_i2fb, _rango_i2fb = conico_por_seccion(1.5, n_cables=2)  # cónico para 1.5mm^2 (iluminación) con 2 cables
                    _accum_conico(_color_i2fb, _num_i2fb, _rango_i2fb, _con_ilum_fb_2, "Iluminación (estimado, 2 cables)")  # suma esos cónicos estimados
    # arma una fila por cada color/tamaño de conector cónico acumulado
    for (color, num, rango), info in sorted(conicos_por_color.items(), key=lambda x: (str(x[0][0]), str(x[0][1]))):  # recorre cada tipo de cónico ya acumulado, ordenado por color y número
        cant = int(info["cantidad"])  # cantidad total de ese tipo de cónico
        if cant <= 0:  # si no se necesita ninguno, se salta esta fila
            continue  # sin cantidad no se escribe la fila
        # el amarillo N°44 tiene su propia rama, pero el texto que sale es
        # igual al genérico; la única diferencia real aparece cuando el
        # cónico no trae número
        if color == "Amarillo" and num == 44:  # caso especial: cónico amarillo N°44
            desc = f"Conector cónico Amarillo (N° {num}) para {rango}"  # descripción del amarillo N°44, con su rango de secciones
        elif num is not None:  # cónico con número de catálogo conocido
            desc = f"Conector cónico {color} (N° {num}) para {rango}"  # descripción genérica del cónico, con su número de catálogo
        else:  # cónico sin número definido (ej. "(definir)")
            desc = f"Conector cónico {color} para {rango}"  # descripción cuando el cónico no tiene número definido
        circuitos_txt = "\n".join(sorted(info["circuitos"])) if info["circuitos"] else "Varios"  # lista de circuitos que usan este cónico, uno por línea
        # agrega la fila de material para este tipo de cónico
        add_row(
            desc=desc,  # descripción ya armada arriba
            marcas_txt=marcas.get("Conectores cónicos", ""),  # marcas sugeridas para conectores cónicos
            norma="RIC 4.4.1",  # texto que se guarda; en el Excel la celda queda como "Ver normativa" con link a la fila "Conector cónico" de BLOQUES_NORMATIVA, que cita RIC 4 (5.11.3)
            circuito=circuitos_txt,  # circuitos donde se usan estos cónicos
            unidad="u",  # se compran por unidad
            k=1,  # sin holgura extra, la cantidad ya viene contada
            longitud_m=f"{cant} unid",  # texto que se muestra en la columna de cantidad
            cantidad=cant  # cónicos a comprar de este tipo
        )

    # ================================================================
    # ESTAÑO, PASTA Y CINTAS — TODO UNIFICADO
    # Fuentes:
    #   A) Circuitos >6mm^2  dan conexiones_circ (conductor grueso)
    #   B) Puesta a tierra  da conexiones_pt1 + conexiones_pt2 + conexiones_desnudo (conductor grueso, 16mm^2, va con A al grupo grueso ÷4)
    #   C) Focos LED        dan _n_focos_led_estaño × 3 (conductor fino)
    # ================================================================
    conexiones_circ    = n_circ_estaño * 3  # 3 conexiones por cada circuito grueso (fase, neutro y tierra)
    conexiones_pt1     = n_barras_pt1 * 2  # 2 conexiones por cada barra de puesta a tierra N°1
    conexiones_pt2     = n_barras_pt2 * 1  # 1 conexión por cada barra de puesta a tierra N°2
    conexiones_desnudo = ((n_barras_pt1 - 1) * 2) + ((n_barras_pt2 - 1) * 2)  # uniones entre barras de tierra hechas con conductor desnudo
    _conexiones_grueso = conexiones_circ + conexiones_pt1 + conexiones_pt2 + conexiones_desnudo  # total de conexiones con conductor grueso (>6mm^2)
    _conexiones_focos  = _n_focos_led_estaño * 3  # conductor fino

    _hay_estaño = (_conexiones_grueso > 0) or (_conexiones_focos > 0)  # True si hay algo que estañar, sea conductor grueso o focos LED

    if _hay_estaño:  # solo genera materiales de estaño si de verdad se necesita
        # Tubos: conductor grueso (>6mm^2, incluye circuitos y PT) usa 1 cada 4 conexiones
        #        conductor fino (solo focos LED) usa 1 cada 15 conexiones
        _conexiones_fino = _conexiones_focos  # mismo valor, con otro nombre para el cálculo de tubos
        _conexiones_grueso_solo = _conexiones_grueso  # circuitos >6mm^2 + PT (pt1+pt2+desnudo)
        _tubos_total = math.ceil(_conexiones_grueso_solo / 4 + _conexiones_fino / 15) if (_conexiones_grueso_solo + _conexiones_fino) > 0 else 0  # tubos de estaño necesarios, redondeando siempre hacia arriba
        _tubos_total = max(1, _tubos_total)  # siempre se pide al menos 1 tubo
        # Pasta: 1 frasco cada 4 tubos
        _pasta_total = max(1, math.ceil(_tubos_total / 4))  # frascos de pasta para soldar necesarios

        # Armar texto de circuitos
        _circ_estaño_all = []  # lista de circuitos que van con estaño, para mostrar en la fila
        if circ_estaño_txt:  # si hay circuitos con conductor grueso que llevan estaño
            _circ_estaño_all += circ_estaño_txt  # los suma a la lista que se muestra en la fila
        if conexiones_pt1 > 0 or conexiones_pt2 > 0:  # si hubo conexiones de puesta a tierra
            _circ_estaño_all.append("Puesta a tierra N°1 y N°2")  # agrega la puesta a tierra a la lista de circuitos a mostrar
        if _circ_focos_led:  # si hubo focos LED estañados
            _circ_estaño_all += _circ_focos_led  # también agrega los circuitos con focos LED
        _circ_estaño_str_unif = " / ".join(sorted(set(_circ_estaño_all))) if _circ_estaño_all else "General"  # une todos los circuitos en un solo texto para la columna Circuito

        # fila de material: tubos de estaño necesarios
        add_row(
            desc="Tubo de estaño 1m / 17gr",  # nombre del material tal como va en el listado
            marcas_txt=marcas.get("Estaño", ""),  # marcas sugeridas de estaño
            norma="RIC 4 (5.11.1)",  # artículo del RIC que pide soldar las conexiones
            circuito=_circ_estaño_str_unif,  # circuitos donde se usa el estaño
            unidad="u",  # se compra por unidad (tubos)
            k=1,  # sin holgura, los tubos ya salen redondeados hacia arriba
            longitud_m=f"{_tubos_total} unid",  # texto de la columna cantidad
            cantidad=_tubos_total  # tubos a comprar
        )
        # fila de material: pasta para soldar necesaria
        add_row(
            desc="Pasta para soldar 50gr",  # nombre del material
            marcas_txt=marcas.get("Pasta para soldar", ""),  # marcas sugeridas de pasta para soldar
            norma="-",  # no tiene artículo del RIC asociado
            circuito=_circ_estaño_str_unif,  # mismos circuitos que el estaño
            unidad="u",  # se compra por frasco
            k=1,  # sin holgura
            longitud_m=f"{_pasta_total} unid",  # texto que se muestra
            cantidad=_pasta_total  # frascos a comprar
        )

        # Cinta autofundente de goma — sobre estaño (focos LED y >6mm^2)
        # 20cm por conexión, rollo 3m = 300cm, redondeando siempre hacia arriba (conexiones × 20 / 300)
        _todas_conexiones_goma = _conexiones_focos + (n_circ_estaño * 3)  # total de conexiones que llevan cinta de goma
        if _todas_conexiones_goma > 0:  # solo si hay conexiones que aislar con cinta de goma
            _rollos_goma = max(1, math.ceil(_todas_conexiones_goma * 20 / 300))  # rollos de cinta de goma necesarios
            _circ_goma_all = []  # lista de circuitos para mostrar en la fila
            if circ_estaño_txt:  # agrega los circuitos con conductor grueso
                _circ_goma_all += circ_estaño_txt  # agrega esos circuitos a la lista
            if _circ_focos_led:  # agrega los circuitos con focos LED
                _circ_goma_all += _circ_focos_led  # y también los de focos LED
            _circ_goma_str = " / ".join(sorted(set(_circ_goma_all))) if _circ_goma_all else "General"  # texto final de circuitos para la fila
            # fila de material: cinta autofundente de goma
            add_row(
                desc="Cinta autofundente de goma 3m",  # nombre del material, rollo de 3m
                marcas_txt=marcas.get("Cinta autofundente goma", ""),  # marcas sugeridas de cinta de goma
                norma="-",  # no tiene artículo del RIC asociado
                circuito=_circ_goma_str,  # circuitos que llevan cinta de goma
                unidad="u",  # se compra por rollo
                k=1,  # sin holgura
                longitud_m=f"{_rollos_goma} unid",  # texto que se muestra
                cantidad=_rollos_goma  # rollos a comprar
            )

    # Cinta aislante PVC — unificada: cónicos (30cm c/u) + focos LED (20cm c/u) + circuitos >6mm^2 (35cm c/u)
    # No incluye puesta a tierra ni conductor desnudo (mismo criterio que la cinta de goma)
    # Rollo 20m = 2000cm
    total_conicos = sum(info["cantidad"] for info in conicos_por_color.values())  # total de cónicos de toda la instalación; incluye también las filas de aviso "(definir)", así que la cinta puede quedar levemente sobrestimada
    _cm_conicos = total_conicos * 30  # 30cm de cinta PVC por cada cónico
    _cm_focos   = _conexiones_focos * 20  # 20cm de cinta PVC por cada conexión de foco LED
    _cm_circ_grueso = conexiones_circ * 35  # 35cm de cinta PVC por cada conexión de circuito grueso
    _cm_total_pvc = _cm_conicos + _cm_focos + _cm_circ_grueso  # total de centímetros de cinta PVC necesarios
    if _cm_total_pvc > 0:  # solo si se necesita algo de cinta
        _rollos_pvc = max(1, math.ceil(_cm_total_pvc / 2000))  # rollos de cinta PVC necesarios (rollo de 20m)
        _circs_pvc = set()  # junta, sin repetir, todos los circuitos que usan cinta PVC
        for info in conicos_por_color.values():  # recorre cada tipo de cónico acumulado
            _circs_pvc.update(info["circuitos"])  # agrega los circuitos de cada tipo de cónico
        if _circ_focos_led:  # hubo focos LED estañados
            _circs_pvc.update(_circ_focos_led)  # agrega los circuitos con focos LED
        if circ_estaño_txt:  # hubo circuitos gruesos estañados
            _circs_pvc.update(circ_estaño_txt)  # agrega los circuitos con conductor grueso (estaño)
        _circs_pvc_txt = " / ".join(sorted(_circs_pvc)) if _circs_pvc else "Varios circuitos"  # texto final de circuitos para la fila
        # fila de material: cinta aislante PVC
        add_row(
            desc="Cinta aislante PVC 20m",  # nombre del material, rollo de 20m
            marcas_txt=marcas.get("Cinta aislante PVC", ""),  # marcas sugeridas de cinta PVC
            norma="-",  # no tiene artículo del RIC asociado
            circuito=_circs_pvc_txt,  # circuitos que llevan cinta PVC
            unidad="u",  # se compra por rollo
            k=1,  # sin holgura
            longitud_m=f"{_rollos_pvc} unid",  # texto que se muestra
            cantidad=_rollos_pvc  # rollos a comprar
        )

    # =========================
    # LUMINARIAS: 1 fila por tipo de luminaria (todos los ambientes juntos)
    # De dónde sale la cantidad: cada LÍNEA de la columna "Detalle iluminación"
    # de ambientes_df es UNA luminaria (ese texto lo arma desc_luminaria_auto()
    # cuando el usuario carga el ambiente). Se agrupa por el texto tal cual, así
    # que dos luminarias descritas igual caen en la misma fila.
    # Ojo: los portalámparas NO van acá — se cuentan aparte, más arriba en el
    # archivo (línea ~3729), solo para las luminarias tipo ampolleta.
    # =========================
    if ambientes_df is not None and "Detalle iluminación" in ambientes_df.columns and "Ambiente" in ambientes_df.columns:  # solo si hay datos de ambientes con detalle de iluminación cargado
        add_section("Iluminarias")  # título de sección en el Excel
        conteo = {}  # {desc_bonita: {"total": N, "por_ambiente": {amb: cant}}}
        # recorre cada ambiente y cuenta cuántas veces se repite cada descripción de luminaria
        for _, ar in ambientes_df.iterrows():  # recorre ambiente por ambiente
            amb_show = str(ar.get("Ambiente", "")).strip() or "Sin ambiente"  # nombre del ambiente a mostrar
            det = str(ar.get("Detalle iluminación", "") or "").strip()  # texto con el detalle de luminarias de ese ambiente
            if not det or det.lower() == "ninguna":  # ambiente sin luminarias cargadas, se salta
                continue  # ambiente sin luminarias, pasa al siguiente
            lineas = [x.strip() for x in det.split("\n") if x.strip()]  # separa el detalle en una luminaria por línea
            for ln in lineas:  # cada línea del detalle es una luminaria
                if ln not in conteo:  # primera vez que aparece este tipo de luminaria
                    conteo[ln] = {"total": 0, "por_ambiente": {}}  # primera aparicion: crea el registro de esa luminaria
                conteo[ln]["total"] += 1  # suma 1 al total de esta luminaria
                conteo[ln]["por_ambiente"][amb_show] = conteo[ln]["por_ambiente"].get(amb_show, 0) + 1  # suma 1 al conteo de este ambiente en particular
        # arma 1 fila por cada tipo de luminaria distinta, sumando todos los ambientes
        for desc_bonita, info in sorted(conteo.items(), key=lambda x: x[0].lower()):  # una fila por cada descripción distinta, en orden alfabetico
            detalle_amb = ", ".join(  # arma texto tipo "Living (2), Cocina (1)" con el detalle por ambiente
                f"{amb} ({cant})" for amb, cant in sorted(info["por_ambiente"].items(), key=lambda x: x[0].lower())  # una entrada por ambiente, ordenadas por nombre
            )
            # fila de material: esta luminaria (todos los ambientes juntos)
            add_row(
                desc=desc_bonita,  # descripción de la luminaria tal como la escribió el usuario
                marcas_txt=marcas.get("Iluminarias", ""),  # marcas sugeridas de iluminarias
                norma="RIC 10 (5.1.4.4, 5.1.4.5)",  # artículos del RIC 10 sobre iluminación
                circuito=f"Iluminación -> {detalle_amb}",  # muestra en que ambientes va y cuántas en cada uno
                unidad="u",  # se compra por unidad
                k=1,  # sin holgura
                longitud_m=f"{info['total']} unid",  # texto que se muestra
                cantidad=info["total"]  # total de luminarias de este tipo en toda la instalación
            )

    # =========================================================
    # Barra unipolar verde según cantidad de circuitos
    # Regla usuario:
    #  1    circuito:   4 polos
    #  2-3  circuitos:  6 polos
    #  4-5  circuitos:  8 polos
    #  6-7  circuitos:  10 polos
    #  8-9  circuitos:  12 polos
    #  10+  circuitos:  15 polos
    # Si n_circ = 0 no cae en ningún rango: polos_barra_verde queda en None y
    # no se agrega ni la fila de la barra verde ni su puesto en el tablero.
    # =========================================================
    n_circ = 0  # cantidad de circuitos, se calcula abajo
    try:  # la tabla de circuitos puede venir rara, por eso el try
        # cuenta circuitos reales (filas con "Circuito" no vacío)
        if circuitos_df is not None and "Circuito" in circuitos_df.columns:  # si existe la columna "Circuito" en la tabla
            n_circ = int(
                circuitos_df["Circuito"].astype(str).str.strip().replace("", np.nan).dropna().shape[0]  # cuenta las filas con nombre de circuito no vacío
            )
        else:  # no viene la columna Circuito
            n_circ = int(len(circuitos_df)) if circuitos_df is not None else 0  # si no hay columna Circuito, cuenta todas las filas de la tabla
    except:  # si algo falla contando, se usa el respaldo de abajo
        n_circ = int(len(circuitos_df)) if circuitos_df is not None else 0  # mismo respaldo si algo falla al leer la tabla
    polos_barra_verde = None  # todavía no se sabe cuántos polos lleva la barra
    # busca cuántos polos corresponde según el rango de circuitos (ver tabla arriba)
    if n_circ == 1:  # 1 circuito -> 4 polos
        polos_barra_verde = 4  # se guarda el número de polos para armar la fila de la barra más abajo
    elif 2 <= n_circ <= 3:  # 2 o 3 circuitos -> 6 polos
        polos_barra_verde = 6  # dos polos más que el caso anterior
    elif 4 <= n_circ <= 5:  # 4 o 5 circuitos -> 8 polos
        polos_barra_verde = 8  # sube a 8 polos
    elif 6 <= n_circ <= 7:  # 6 o 7 circuitos -> 10 polos
        polos_barra_verde = 10  # 10 polos
    elif 8 <= n_circ <= 9:  # 8 o 9 circuitos -> 12 polos
        polos_barra_verde = 12  # 12 polos
    elif n_circ >= 10:  # 10 o más circuitos -> 15 polos
        polos_barra_verde = 15  # tope de la tabla: 15 polos, aunque haya muchos más circuitos
    if polos_barra_verde is not None:  # solo se agrega barra verde si algún rango de arriba definió polos
        # fila de material: barra unipolar verde (solo si hay circuitos)
        add_row(
            desc=f"Barra unipolar verde de {polos_barra_verde} polos 63A",  # nombre de la barra con los polos ya calculados
            marcas_txt=marcas.get("Barra unipolar verde", ""),  # marcas sugeridas de barra unipolar verde
            norma="RIC 9.2.1",  # ojo: en BLOQUES_NORMATIVA las barras están en RIC 2, así que el link lleva a la fila del RIC 2
            circuito="General",  # es material del tablero, no de un circuito puntual
            unidad="u",  # se compra por unidad
            k=1,  # sin holgura
            longitud_m=f"{polos_barra_verde} polos",  # en vez de metros, muestra los polos
            cantidad=1  # siempre va 1 sola barra verde
        )

    # =========================================================
    # BARRA REPARTIDORA TETRAPOLAR PRINCIPAL — siempre 4 polos fija
    # =========================================================
    n_dif = 0  # cantidad de diferenciales, se calcula abajo
    try:  # group_info puede venir vacío o mal formado
        n_dif = int(len(group_info)) if group_info else 0  # cantidad de diferenciales (1 por grupo)
    except:  # group_info venía vacío o mal formado
        n_dif = 0  # si falla la lectura, se asume 0 diferenciales
    # fila de material: barra repartidora tetrapolar principal (siempre se agrega)
    add_row(
        desc="Barra repartidora tetrapolar de 4 polos de 125[A]",  # la barra principal es siempre de 4 polos
        marcas_txt=marcas.get("Barra repartidora", ""),  # marcas sugeridas de barra repartidora
        norma="RIC 9.2.1",  # ojo: en BLOQUES_NORMATIVA las barras están en RIC 2, así que el link lleva a la fila del RIC 2
        circuito="General",  # material del tablero, no de un circuito
        unidad="u",  # se compra por unidad
        k=1,  # sin holgura
        longitud_m="4 polos",  # muestra los polos en la columna
        cantidad=1  # siempre 1 barra principal
    )


    # =========================================================
    # BARRA REPARTIDORA BIPOLAR según nº de diferenciales
    #  1-3 diferenciales:  4 polos
    #  4-6 diferenciales:  7 polos
    #  7-10 diferenciales: 11 polos
    # =========================================================
    polos_bipolar = None  # de a poco: si no cae en ningún rango de abajo, no se agrega barra bipolar
    # revisa en que rango cae la cantidad de diferenciales para saber cuántos polos necesita la barra
    if 1 <= n_dif <= 3:  # hasta 3 diferenciales alcanza con 4 polos
        polos_bipolar = 4  # se guarda para armar la descripción de la barra bipolar
    elif 4 <= n_dif <= 6:  # entre 4 y 6 diferenciales, barra de 7 polos
        polos_bipolar = 7  # 7 polos
    elif 7 <= n_dif <= 10:  # entre 7 y 10 diferenciales, barra de 11 polos
        polos_bipolar = 11  # tope: 11 polos, sobre 10 diferenciales queda sin barra bipolar
    if polos_bipolar is not None:  # solo agrega el material si algún rango de arriba definió una cantidad de polos
        # agrega al listado de materiales la barra repartidora bipolar con los polos calculados
        add_row(
            desc=f"Barra repartidora bipolar de {polos_bipolar} polos de 125[A]",  # barra bipolar con los polos según cuántos diferenciales hay
            marcas_txt=marcas.get("Barra repartidora", ""),  # marcas sugeridas de barra repartidora
            norma="RIC 2 (6.2.1, 6.2.4, 6.2.7)",  # artículos del RIC 2 sobre protecciones y tablero
            circuito="General",  # material general del tablero
            unidad="u",  # se compra por unidad
            k=1,  # sin holgura
            longitud_m=f"{polos_bipolar} polos",  # muestra los polos en la columna
            cantidad=1  # 1 sola barra bipolar
        )
    # =========================================================
    # CONDICIÓN 2:
    # Barras repartidoras tetrapolares EXTRA de 4 polos 125A
    # Regla (2 grupos por barra, no 2 TM por barra):
    #  - Tomar SOLO diferenciales que tengan 2 o 3 TM
    #  - Contar cuántos GRUPOS válidos hay: eso da total_grupos_validos
    #  - Cada barra 4P alimenta 2 grupos
    #  - barras_extra_4p = ceil(total_grupos_validos / 2)
    # Ejemplos:
    #   1 grupo (2 o 3 TM):  1 barra
    #   2 grupos (2 o 3 TM): 1 barra
    #   3 grupos (2 o 3 TM): 2 barras
    #   4 grupos (2 o 3 TM): 2 barras
    # =========================================================
    total_grupos_validos = 0  # contador de grupos (diferenciales) que tienen 2 o 3 TM asociados
    # recorre cada diferencial y cuenta los que tienen 2 o 3 TM (los que califican para una barra extra)
    try:  # group_info puede no venir cargado
        if group_info:  # solo si hay diferenciales armados
            for meta in group_info.values():  # revisa el detalle de cada diferencial
                n_tm = int(len(meta.get("indices", [])))  # cantidad de TM que cuelgan de este diferencial
                if n_tm in (2, 3):  # solo califican los diferenciales con 2 o 3 TM colgando
                    total_grupos_validos += 1  # este grupo califica: suma 1 al total
    except:  # si group_info venía mal, no se cuenta ningún grupo
        total_grupos_validos = 0  # si algo fallo al leer group_info, se asume que no hay grupos válidos
    barras_extra_4p = int(math.ceil(total_grupos_validos / 2.0)) if total_grupos_validos > 0 else 0  # cada barra 4P alcanza para 2 grupos, por eso se divide por 2 y se redondea hacia arriba
    # si se necesita al menos 1 barra extra, se agrega al listado de materiales
    if barras_extra_4p > 0:
        # agrega la(s) barra(s) tetrapolar(es) extra con la cantidad ya calculada
        add_row(
            desc="Barra repartidora tetrapolar de 4 polos de 125[A]",  # misma barra tetrapolar, pero estas son las extra
            marcas_txt=marcas.get("Barra repartidora", ""),  # marcas sugeridas de barra repartidora
            norma="RIC 9.2.1",  # ojo: en BLOQUES_NORMATIVA las barras están en RIC 2, así que el link lleva a la fila del RIC 2
            circuito="General",  # material general del tablero
            unidad="u",  # se compra por unidad
            k=1,  # sin holgura
            longitud_m=f"{barras_extra_4p} barra(s) de 4 polos",  # muestra cuántas barras y de cuántos polos
            cantidad=barras_extra_4p  # barras extra calculadas arriba
        )
    # =========================================================
    # TABLERO (PVC IP41) – CONDICIONES 1 + 2 + 3 + REDONDEO
    # (DEBE IR DESPUÉS DE calcular polos_barra_verde, polos_bipolar y barras_extra_4p)
    # =========================================================
    # contar TMs reales (1 puesto c/u)
    n_tm = 0  # valor por defecto si no se puede contar
    # cuenta cuántas filas de la tabla de circuitos tienen un disyuntor termomagnético asignado
    try:  # la tabla puede no traer la columna del disyuntor
        if {"Disyuntor termomagnético", "Circuito"}.issubset(set(circuitos_df.columns)):  # solo cuenta si la tabla de circuitos tiene esas columnas
            tms_df = circuitos_df[["Disyuntor termomagnético", "Circuito"]].copy()  # copia solo las columnas que interesan, para no tocar el dataframe original
            tms_df["Disyuntor termomagnético"] = tms_df["Disyuntor termomagnético"].astype(str).str.strip()  # saca espacios en blanco del texto del disyuntor
            tms_df = tms_df[tms_df["Disyuntor termomagnético"] != ""]  # se queda solo con las filas que SI tienen disyuntor termomagnético
            n_tm = int(len(tms_df))  # cantidad final de TM contados
    except:  # faltaba alguna columna o fallo la lectura de la tabla
        n_tm = 0  # si algo fallo (columnas no existen, etc), se asume 0 TM
    # contar diferenciales (2 puestos c/u)
    n_dif = int(len(group_info)) if group_info else 0  # recalcula n_dif acá porque esta función lo necesita en este punto del cálculo

    # CONDICIÓN 1: puestos por protecciones — se van sumando los puestos que
    # ocupa cada elemento del tablero (cada uno tiene un ancho fijo en puestos)
    puestos_omni = 2  # el interruptor general (omnipolar) siempre ocupa 2 puestos
    puestos_tm = n_tm * 1  # cada TM (disyuntor termomagnético) ocupa 1 puesto
    puestos_dif = n_dif * 2  # cada diferencial ocupa 2 puestos
    puestos_luz_piloto = 1 if n_circ >= 3 else 0  # la luz piloto solo se agrega si el tablero tiene 3 o más circuitos
    puestos_portafusible = 1 if n_circ >= 3 else 0  # el portafusible también depende de tener 3 o más circuitos
    puestos_spd = 2  # el protector contra sobretensiones (SPD) ocupa 2 puestos fijos
    puestos_protector_sobrevoltaje = 2  # el protector de sobrevoltaje ocupa 2 puestos fijos (es un elemento aparte del SPD)
    puestos_barra_verde = 1 if polos_barra_verde is not None else 0  # la barra verde (tierra) ocupa 1 puesto, solo si se cálculo antes
    puestos_barra_principal = 2  # siempre 4 polos = 2 puestos
    puestos_barra_bipolar = 0  # se va sumando según los polos que resulten de la barra bipolar
    # convierte los polos de la barra bipolar a puestos de riel que ocupa
    if polos_bipolar == 4:  # barra bipolar de 4 polos
        puestos_barra_bipolar = 2  # ocupa 2 puestos del riel
    elif polos_bipolar == 7:  # barra bipolar de 7 polos
        puestos_barra_bipolar = 4  # ocupa 4 puestos
    elif polos_bipolar == 11:  # barra bipolar de 11 polos
        puestos_barra_bipolar = 6  # ocupa 6 puestos
    puestos_barras_extra = int(barras_extra_4p) * 2 if barras_extra_4p > 0 else 0  # cada barra tetrapolar extra ocupa 2 puestos
    # Puestos de bornera — se agrupan por sección, ya que varias borneras
    # chicas comparten el mismo puesto del riel (no es 1 puesto fijo por
    # circuito). Hay 2 tipos, con condiciones distintas, y NO se reemplazan
    # entre sí porque son para conductores distintos:
    #  - Bornera de neutro: existe siempre que haya un diferencial exclusivo
    #    (1:1), sin importar la cantidad total de circuitos del proyecto.
    #    Es solo para el conductor neutro.
    #  - Bornera de organización: solo si el proyecto tiene más de 8
    #    circuitos. Es para el conductor fase, y aplica a TODOS los
    #    circuitos en ese caso, tengan o no bornera de neutro.
    # Ojo: este conteo es SOLO para dimensionar el tablero. Las filas de
    # material de bornera se arman aparte, más abajo, con
    # _borneras_neutro_acum / _borneras_organizacion_acum. El criterio es el
    # mismo, pero aquellas saltan las filas incompletas de circuitos_df (sin
    # nombre, conductor o TM), así que la cantidad de acá puede quedar un poco
    # por encima de la que finalmente se compra.
    puestos_borneras = 0  # acá se van acumulando los puestos que ocupan las borneras

    # función auxiliar (copia local) para elegir la sección mínima del conductor de una corriente dada
    def _seccion_A1_local(corriente_A):  # versión chica de seccion_A1, para poder usarla antes de que se defina la otra
        # busca la sección de conductor (mm^2) más chica que aguante esta
        # corriente, usando la tabla de ampacidad método A1 (embutido/ducto)
        secciones_A1 = [  # misma tabla que seccion_A1 (definida más abajo): método A1, RIC 2 punto 6.2.2, conductor H07Z1-K
            (1.5, 14), (2.5, 18), (4.0, 24), (6.0, 31),  # (sección mm^2, corriente máxima admisible en A); secciones chicas, hasta 31A
            (10.0, 42), (16.0, 56), (25.0, 73), (35.0, 89),  # secciones gruesas, de 42A para arriba
        ]
        for sec, iz in secciones_A1:  # recorre de la sección más chica a la más grande
            if iz >= corriente_A:  # la primera que aguante la corriente sirve
                return sec  # devuelve la sección en mm^2
        return 35.0  # si ninguna sección de la lista alcanzo, se usa la más grande disponible

    _borneras_por_seccion = {}  # {sección mm^2 (float): cuántas borneras de esa sección}; solo para contar puestos del riel

    # Bornera de neutro (siempre, sin condición de n_circ)
    if group_info:  # solo si ya están armados los grupos de diferencial
        for _gid, _meta in group_info.items():  # revisa grupo por grupo
            _indices_grupo = _meta.get("indices", [])  # circuitos que cuelgan de ese diferencial
            if len(_indices_grupo) == 1:  # diferencial exclusivo: protege un solo circuito
                # si el diferencial es exclusivo, lleva bornera de neutro, con la sección del diferencial
                _dif_a = _meta.get("dif", None)  # corriente del diferencial del grupo, en amperes
                _sec_dif_local = _seccion_A1_local(_dif_a) if _dif_a else None  # sección del conductor que sale de ese diferencial
                if _sec_dif_local:  # si se pudo sacar la sección, cuenta la bornera
                    _borneras_por_seccion[_sec_dif_local] = _borneras_por_seccion.get(_sec_dif_local, 0) + 1  # suma una bornera de neutro de esa sección

    # Bornera de organización (solo si n_circ>8) — aplica a TODOS los
    # circuitos, incluyendo los que ya tienen bornera de neutro, porque
    # es para el conductor fase, no el neutro.
    if n_circ > 8 and {"Conductor", "Circuito"}.issubset(set(circuitos_df.columns)):  # bornera de organización solo con más de 8 circuitos, y si están las columnas
        for _idx, _r in circuitos_df.iterrows():  # una bornera de fase por cada circuito del proyecto
            _cond_b = str(_r.get("Conductor", "")).strip()  # texto del conductor del circuito, ej 3x2.5mm^2
            _sec_b = extraer_seccion_mm2(_cond_b)  # saca la sección en mm^2 de ese texto
            if _sec_b:  # si el conductor traía una sección legible
                _borneras_por_seccion[_sec_b] = _borneras_por_seccion.get(_sec_b, 0) + 1  # suma una bornera de organización de esa sección

    # suma los puestos que ocupan todas las borneras, agrupadas por sección
    for _sec_b, _cantidad_b in _borneras_por_seccion.items():  # recorre sección por sección del acumulado de borneras
        puestos_borneras += int(math.ceil(_cantidad_b / borneras_por_puesto(_sec_b)))  # cada tipo de bornera comparte puesto entre varias unidades, por eso se divide por borneras_por_puesto
    # suma todos los puestos que se van a necesitar por protecciones, barras y borneras
    puestos_base = int(puestos_omni + puestos_tm + puestos_dif + puestos_luz_piloto + puestos_portafusible + puestos_spd  # protecciones: omnipolar, TM, diferenciales, luz piloto, portafusible y SPD
        + puestos_protector_sobrevoltaje + puestos_barra_principal + puestos_barra_bipolar  # más el protector de sobrevoltaje y las barras principal y bipolar
        + puestos_barras_extra + puestos_barra_verde + puestos_borneras)  # más las barras extra, la barra verde y las borneras

    # CONDICIÓN 2: reserva por circuitos — deja espacio libre para futuras ampliaciones
    puestos_reserva = int(math.ceil(n_circ * 0.25)) * 3 if n_circ > 0 else 0  # reserva para ampliaciones: 25% de los circuitos, contando 3 puestos cada uno
    # CONDICIÓN 3: suma total
    puestos_total = int(puestos_base + puestos_reserva)  # puestos definitivos que tiene que tener el tablero
    # REDONDEO A TABLEROS COMERCIALES: busca el tamaño de tablero comercial
    # más chico que alcance a cubrir los puestos necesarios
    tamanos_tablero = [2, 4, 6, 8, 12, 16, 18, 24, 36, 42, 48, 54, 56, 72]  # tamaños de tablero que se consiguen comercialmente, en cantidad de puestos
    puestos_tablero = None  # acá se va a guardar el tamaño de tablero elegido
    # recorre los tamaños de menor a mayor y elige el primero que alcance a cubrir lo necesario
    for t in tamanos_tablero:  # prueba los tamaños comerciales de menor a mayor
        if t >= puestos_total:  # el primero que cubra lo que se necesita
            puestos_tablero = t  # se queda con ese tamaño de tablero
            break  # ya lo encontró, corta la búsqueda
    if puestos_tablero is None:  # no alcanzo ninguno de la lista
        puestos_tablero = tamanos_tablero[-1]  # se pasó de todos los tamaños, usa el más grande
    _puestos_tablero = int(puestos_tablero)  # copia auxiliar del tamaño de tablero elegido, como número entero
    # descripción según canalización (embutida/sobrepuesta)
    tipo_can = (tipo_canalizacion or "").strip().lower()  # normaliza el texto de canalización para comparar en minúsculas
    es_embutida_tab = "embut" in tipo_can  # revisa si la canalización es embutida (va dentro de la muralla) o sobrepuesta
    # arma la descripción y la marca del tablero según el tipo de canalización
    if es_embutida_tab:  # el tablero embutido va dentro de la muralla
        desc_tab = f"Tablero embutido de PVC de {puestos_tablero} puestos IP41"  # descripción del tablero embutido con su cantidad de puestos
        marca_tab = marcas.get("Tablero embutido", "")  # marca sugerida para tablero embutido
    else:  # si no, el tablero va montado sobre la muralla
        desc_tab = f"Tablero sobrepuesto de PVC de {puestos_tablero} puestos IP41"  # descripción del tablero sobrepuesto
        marca_tab = marcas.get("Tablero sobrepuesto", "")  # marca sugerida para tablero sobrepuesto
    # agrega el tablero al listado de materiales, con el tamaño y la cantidad de puestos requeridos
    # agrega el tablero como material, con su tamaño comercial
    add_row(
        desc=desc_tab,  # texto que sale en el informe
        marcas_txt=marca_tab,  # marcas sugeridas para el tablero
        norma="RIC 9.1.1",  # ojo: en BLOQUES_NORMATIVA el tablero está en RIC 2, y el link lleva a esa fila
        circuito=f"General ({puestos_total} puestos requeridos)",  # en la columna circuito se deja anotado cuántos puestos se pidieron
        unidad="u",  # se cotiza por unidad
        k=1,  # multiplicador de cantidad, acá 1 porque es un solo tablero
        longitud_m=f"{puestos_tablero} puestos",  # en vez de metros se informa el tamaño en puestos
        cantidad=1  # un tablero por proyecto
    )
    # =========================================================
    # RIEL DIN según tamaño de tablero
    # - 1m para cualquier tablero que no sea de 56 o 72 puestos
    # - 2m para: 56 o 72 puestos
    # =========================================================
    # el riel viene en tiras de 1 o 2 metros; para tableros grandes (56 o 72 puestos) se ocupan 2m
    if puestos_tablero in [56, 72]:
        desc_riel = "Riel DIN de 35x7,5 mm tira de 2m"  # descripción del riel de 2 metros
        largo_riel = 2  # largo del riel en metros
    else:  # tableros más chicos con 1 metro alcanzan
        desc_riel = "Riel DIN de 35x7,5 mm tira de 1m"  # descripción del riel de 1 metro
        largo_riel = 1  # largo del riel en metros
    _riel_m = largo_riel  # 1 o 2 (metros); lo usa el bloque de Tornillería, más abajo, para pedir 7 o 14 tornillos de riel
    # agrega el riel DIN al listado de materiales
    # el riel DIN es donde se montan las protecciones adentro del tablero
    add_row(
        desc=desc_riel,  # descripción del riel que se eligió arriba
        marcas_txt=marcas.get("Riel DIN", ""),  # marcas sugeridas de riel DIN
        norma="RIC 9.1.1",  # mismo texto que el tablero; el link del Excel lleva a la fila "Riel DIN" de BLOQUES_NORMATIVA = RIC 2 (6.1.15, 6.1.23)
        circuito="General",  # no pertenece a un circuito en particular
        unidad="u",  # se compra por tira
        k=1,  # sin multiplicador
        longitud_m=f"{largo_riel} m",  # largo de la tira, en metros
        cantidad=1  # una sola tira
    )
    # =========================================================
    # LUZ PILOTO + PORTAFUSIBLE + FUSIBLE (solo si hay >=3 circuitos)
    # =========================================================
    # estos 3 elementos solo se agregan si el tablero tiene 3 o más circuitos
    if n_circ >= 3:  # con menos de 3 circuitos no se pone senalizacion en el tablero
        # Luz piloto tablero
        # la luz piloto avisa que el tablero esta energizado
        add_row(
            desc="Luz piloto LED Riel Din 220 VAC IP44",  # luz piloto que va montada sobre el riel DIN
            marcas_txt=marcas.get("Luz Piloto", ""),  # marcas sugeridas de luz piloto
            norma="RIC 9.1.1",  # mismo texto que el tablero; el link lleva a la fila "Luz piloto" = RIC 2 (6.2.14, 6.2.15)
            circuito="Tablero de alumbrado",  # se carga al tablero de alumbrado
            unidad="u",  # se cotiza por unidad
            k=1,  # sin multiplicador
            longitud_m="1 Unid",  # se informa como 1 unidad
            cantidad=1  # una sola luz piloto
        )
        # Portafusible tablero
        # el portafusible protege el circuito de la luz piloto, no el tablero
        add_row(
            desc="Portafusible 1P 32A 10x38mm 500V",  # portafusible que protege el circuito de la luz piloto
            marcas_txt=marcas.get("Portafusible", ""),  # marcas sugeridas de portafusible
            norma="RIC 9.1.1",  # mismo texto que el tablero; el link lleva a la fila "Portafusible tablero" = RIC 2 (6.2.15, 6.3.6)
            circuito="Tablero de alumbrado",  # también va al tablero de alumbrado
            unidad="u",  # por unidad
            k=1,  # sin multiplicador
            longitud_m="1 Unid",  # 1 unidad
            cantidad=1  # uno solo
        )
        # Fusible tablero
        # el fusible chico que va adentro del portafusible
        add_row(
            desc="Fusible cilíndrico 2A 10x38mm 500V",  # fusible de 2A que va dentro del portafusible
            marcas_txt=marcas.get("Fusible", ""),  # marcas sugeridas de fusible
            norma="RIC 9.1.1",  # mismo texto que el tablero; el link lleva a la fila "Fusible cilíndrico tablero" = RIC 2 (6.2.15, 6.3.6)
            circuito="Tablero de alumbrado",  # se carga al tablero de alumbrado
            unidad="u",  # por unidad
            k=1,  # sin multiplicador
            longitud_m="1 Unid",  # 1 unidad
            cantidad=1  # uno solo
        )

        # =========================
    # TERMINALES FERRUL INTERIORES
    # =========================
    # a partir de acá se arman las funciones para calcular los terminales ferrul (terminales de los conductores)
    add_section("Terminales ferrul interiores")  # abre la sección de ferrules en el informe

    def normalizar_seccion_ferrul(seccion_mm2):  # deja la sección en un valor de ferrul que se venda
        # redondea la sección hacia arriba a la sección comercial de ferrul
        # más cercana (1.5, 2.5, 4.0, 6.0, 10.0, 16.0, 25.0 o 35.0mm^2)
        try:  # la sección puede venir como texto
            s = float(seccion_mm2)  # intenta convertir a número
        except:  # no se pudo convertir a número
            return None  # no era un número válido

        secciones = [1.5, 2.5, 4.0, 6.0, 10.0, 16.0, 25.0, 35.0]  # secciones comerciales de ferrul

        for sec in secciones:  # recorre de la sección más chica a la más grande
            if s <= sec:  # el ferrul tiene que ser igual o más grande que el conductor
                return sec  # primera sección comercial que alcanza a cubrirla

        return secciones[-1]  # era más grande que todas, se usa la máxima

    def color_ferrul_por_seccion(sec):  # entrega el color del ferrul según la sección
        # color del ferrul según la sección del conductor (código de colores
        # típico: 1.5mm^2=rojo, 2.5mm^2=azul, etc.)
        colores = {  # tabla sección mm^2 a color de ferrul
            1.5: "rojo",  # 1.5mm^2 va rojo
            2.5: "azul",  # 2.5mm^2 va azul
            4.0: "naranjo",  # 4mm^2 va naranjo
            6.0: "amarillo",  # 6mm^2 va amarillo
            10.0: "rojo",  # en 10mm^2 se repite el rojo
            16.0: "azul",  # 16mm^2 azul
            25.0: "amarillo",  # 25mm^2 amarillo
            35.0: "gris"  # 35mm^2 gris
        }  # tabla de colores por sección

        try:  # la sección puede venir como texto
            return colores.get(float(sec), "por definir")  # busca el color, si no está lo marca "por definir"
        except:  # no era un número, queda sin color definido
            return "por definir"  # la sección no era un número válido


    # Acumulador de ferrules del tablero, con forma:
    #   {sección_mm^2 (float ya normalizada a comercial): {texto_detalle: cantidad}}
    # Lo llena agregar_ferrules() desde los bloques A) a E) de más abajo, y el
    # bloque F) arma 1 fila de Excel por sección, sumando todos sus detalles y
    # poniendo cada detalle en una línea de la columna Circuito.
    ferrules = {}  # diccionario que junta cuántos ferrules se necesitan, agrupados por sección

    def agregar_ferrules(seccion_mm2, cantidad, detalle):  # va acumulando los ferrules en el diccionario de más arriba
        # suma ferrules al conteo total, según la sección que corresponda.
        # Si la sección no es válida o la cantidad es 0, no hace nada
        # qué recibe:
        #   seccion_mm2 .. sección del conductor al que va el ferrul, en mm^2
        #   cantidad ..... cuántos ferrules sumar
        #   detalle ...... para qué son, ej: "Circuito 2 - fase en bornera"
        # no devuelve nada: va llenando el diccionario "ferrules".
        sec = normalizar_seccion_ferrul(seccion_mm2)  # sección comercial de ferrul

        if sec is None:  # la sección no servia, no cuenta ferrules
            return  # sección inválida, no hace nada

        cantidad = int(cantidad)  # asegura que cantidad sea un número entero

        if cantidad <= 0:  # no tiene sentido sumar 0 o negativo
            return  # no hay nada que sumar

        if sec not in ferrules:  # primera vez que aparece esta sección
            ferrules[sec] = {}  # primera vez que se usa esta sección

        ferrules[sec][detalle] = ferrules[sec].get(detalle, 0) + cantidad  # suma al total de esa sección/detalle

    def seccion_circuito(cond_txt):  # atajo para sacar la sección del texto del conductor
        # saca la sección en mm^2 del texto del conductor (ej: "3x2.5mm^2" da 2.5)
        return extraer_seccion_mm2(str(cond_txt))  # usa la función general de más arriba

    def seccion_A1(corriente_A):  # misma idea que _seccion_A1_local, pero para el resto de la función
        """Selecciona la sección mínima del conductor interior del tablero
        usando método A1 (RIC 2 punto 6.2.2), secciones comerciales H07Z1-K."""
        secciones_A1 = [  # tabla de ampacidad método A1 para conductor H07Z1-K
            (1.5, 14), (2.5, 18), (4.0, 24), (6.0, 31),  # secciones chicas, hasta 31A
            (10.0, 42), (16.0, 56), (25.0, 73), (35.0, 89),  # secciones gruesas, de 42A para arriba
        ]  # (sección mm^2, capacidad de corriente A)
        for sec, iz in secciones_A1:  # recorre la tabla de menor a mayor sección
            if iz >= corriente_A:  # busca la primera capacidad que alcance la corriente pedida
                return sec  # la primera que alcanza a soportar la corriente
        return 35.0  # máximo disponible

    def contar_enchufes_items(items):  # cuenta cuántos enchufes trae el circuito
        # suma los enchufes del circuito (misma idea que _contar_enchufes_en_items,
        # pero definida acá porque se usa en otra parte de la función)
        total = 0  # contador de enchufes encontrados

        if isinstance(items, list):  # solo procesa si items es una lista válida
            for it in items:  # recorre cada item del circuito
                if not isinstance(it, dict):  # si el item no es diccionario no se puede leer
                    continue  # item raro, se salta

                if ("id_ench" in it) or ("modulos" in it) or ("n_ench" in it):  # detecta si el item es un enchufe (por sus claves)
                    total += int(it.get("n_ench", 1) or 1)  # suma la cantidad de este item

        return int(total)  # entrega el total de enchufes como entero

    def buscar_grupo_por_indice(idx_circuito):  # dice a que diferencial esta conectado el circuito
        # busca a qué grupo de diferencial pertenece este circuito
        if not group_info:  # sin grupos armados no hay nada que buscar
            return None, None  # no hay agrupación definida

        for gid, meta in group_info.items():  # recorre cada grupo de diferencial definido
            indices = meta.get("indices", [])  # circuitos (por índice) que pertenecen a este grupo

            if idx_circuito in indices:  # el circuito está dentro de este grupo
                return gid, meta  # encontró el grupo al que pertenece

        return None, None  # no pertenece a ningún grupo

    try:  # el alimentador puede no venir calculado
        sec_alimentador = float(res_alim["S"])  # sección del alimentador ya calculada, en mm^2 (viene de seleccionar_alimentador())
        # Sección del conductor de protección (PE) del tablero, con la regla
        # clásica del RIC 6 (sección 7, conductor de protección):
        #   S <= 16mm^2      -> PE = S
        #   16 < S <= 35mm^2 -> PE = 16mm^2 fijo
        #   S > 35mm^2       -> PE = S/2
        if sec_alimentador <= 16:  # hasta 16mm^2 el PE va del mismo calibre que la fase
            sec_pt_tablero = sec_alimentador  # PE igual a la fase
        elif sec_alimentador <= 35:  # entre 16 y 35mm^2 el PE queda fijo en 16
            sec_pt_tablero = 16.0  # PE fijo en 16mm^2
        else:  # sobre 35mm^2 el PE va a la mitad de la fase
            sec_pt_tablero = sec_alimentador / 2.0  # PE a la mitad de la fase
    except:  # si res_alim no traía una sección válida
        sec_alimentador = None  # no se pudo calcular; ojo: con None, agregar_ferrules() descarta en silencio los ferrules del alimentador
        sec_pt_tablero = 4.0  # respaldo: 4mm^2 es la sección mínima de PE que usa el resto del programa

    # el cableado de control siempre va en 1.5mm^2, no depende de la carga
    sec_control = 1.5  # sección fija en mm^2 para el cableado de control (luz piloto, etc.)

    try:  # cuenta los circuitos con nombre, para repartir los ferrules
        n_circ_ferrul = int(  # el total queda como número entero
            circuitos_df["Circuito"]  # toma la columna de circuitos
            .astype(str)  # la pasa a texto para poder limpiarla
            .str.strip()  # saca los espacios de los lados
            .replace("", np.nan)  # las celdas vacias quedan como nulas
            .dropna()  # y las nulas se botan
            .shape[0]  # lo que queda es la cantidad de circuitos reales
        )
    except:  # si la columna no existe o el dataframe viene raro
        n_circ_ferrul = int(len(circuitos_df)) if circuitos_df is not None else 0  # si algo falla, respaldo: cuenta todas las filas del dataframe

    usa_luz_piloto = n_circ_ferrul >= 3  # con 3 o más circuitos, el tablero lleva luz piloto de senalizacion

    # -------------------------
    # A) TABLERO FIJO
    # -------------------------
    # Todo lo que cuelga de la barra principal, protegido por el mismo
    # interruptor general omnipolar (entre la barra y el omnipolar, el SPD,
    # el protector de sobrevoltaje) se dimensiona según Método A1 (RIC 2,
    # 6.2.2) con la corriente nominal del omnipolar — NO con la sección del
    # alimentador.
    # Ej: alimentador 4mm^2 (RV-K) pero omnipolar 2x25A: el método A1 exige 6mm^2
    # (4mm^2 solo aguanta hasta 24A, no alcanza para un interruptor de 25A).
    # Solo el tramo entre el alimentador y la barra principal sigue usando la
    # sección real del alimentador (es el mismo cable que ya viene puesto).
    # ojo: se usa la corriente del omnipolar, no la sección del alimentador
    sec_omni_A1 = seccion_A1(interruptor_empalme)  # sección del conductor según la corriente del interruptor general (método A1)
    agregar_ferrules(sec_alimentador, 2, "Alimentador F/N a barra principal")  # 2 ferrules: fase y neutro del alimentador llegando a la barra principal
    agregar_ferrules(sec_alimentador, 1, "Alimentador PE a barra tierra")  # 1 ferrul para el PE del alimentador a la barra de tierra
    agregar_ferrules(sec_pt_tablero, 1, "Barra verde PE tablero")  # ferrul del PE que sale de la barra verde
    agregar_ferrules(sec_omni_A1, 4, "Barra principal a omnipolar")  # 4 ferrules entre la barra principal y la entrada del omnipolar
    agregar_ferrules(sec_omni_A1, 4, "Omnipolar salida a barra principal")  # 4 más en la salida del omnipolar de vuelta a la barra
    agregar_ferrules(sec_omni_A1, 4, "Barra principal a SPD F/N")  # 4 para conectar el SPD a fase y neutro
    agregar_ferrules(sec_omni_A1, 2, "SPD a barra PE")  # 2 para la tierra del SPD
    agregar_ferrules(sec_omni_A1, 4, "Barra principal a protector sobrevoltaje")  # 4 hacia la entrada del protector de sobrevoltaje
    agregar_ferrules(sec_omni_A1, 4, "Protector sobrevoltaje a barra bipolar")  # 4 de la salida del protector a la barra bipolar

    if usa_luz_piloto:  # solo agrega estos ferrules si el tablero lleva luz piloto
        agregar_ferrules(sec_control, 6, "Luz piloto + portafusible")  # 6 ferrules del cableado de control de la luz piloto y su portafusible

    # -------------------------
    # B) DIFERENCIALES + TM + SALIDAS
    # -------------------------
    if {"Circuito", "Conductor", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # revisa que existan las columnas necesarias para procesar diferenciales
        grupos_ya_contados = set()  # para no contar 2 veces los ferrules de entrada al mismo diferencial
        grupos_barra_extra_contados = set()  # para no contar 2 veces la barra extra cuando el grupo tiene 2 o 3 circuitos
        _borneras_neutro_acum = {}  # {sec_dif: {"cantidad": N, "circuitos": [...]}} — se fusionan en 1 fila por sección
        _borneras_organizacion_acum = {}  # {sec_circ: {"cantidad": N, "circuitos": [...]}} — 1 por circuito, cuando n_circ>8

        for idx, r in circuitos_df.iterrows():  # recorre cada circuito del tablero, uno por uno
            circ = str(r.get("Circuito", "")).strip()  # nombre del circuito
            cond = str(r.get("Conductor", "")).strip()  # texto del conductor, ej: "3x2.5mm^2"
            tm = str(r.get("Disyuntor termomagnético", "")).strip()  # texto del disyuntor termomagnético asignado a este circuito

            if not circ or not cond or not tm:  # fila vacía o incompleta, se salta
                continue  # fila incompleta, pasa al siguiente circuito

            sec_circ = seccion_circuito(cond)  # sección en mm^2 del conductor de este circuito
            gid, meta = buscar_grupo_por_indice(idx)  # busca el grupo de diferencial al que pertenece (si tiene)

            if meta is None:  # no se encontró grupo de diferencial para este circuito
                # este circuito no pertenece a ningún grupo de diferencial (caso raro/individual)
                agregar_ferrules(sec_circ, 2, f"{circ} - barra a TM")  # 2 ferrules: de la barra a la entrada del TM
                if n_circ > 8:  # con más de 8 circuitos se usan borneras de organización para ordenar el cableado
                    agregar_ferrules(sec_circ, 2, f"{circ} - salida TM a entrada bornera")  # 2 ferrules entre la salida del TM y la bornera
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida bornera al circuito")  # 3 ferrules de la bornera al circuito: fase, neutro y tierra
                    if sec_circ not in _borneras_organizacion_acum:  # primera vez que aparece esta sección en el acumulado
                        _borneras_organizacion_acum[sec_circ] = {"cantidad": 0, "circuitos": []}  # crea el registro para esta sección
                    _borneras_organizacion_acum[sec_circ]["cantidad"] += 1  # suma 1 al conteo de esta sección
                    _borneras_organizacion_acum[sec_circ]["circuitos"].append(circ)  # deja anotado que circuito uso esa bornera
                else:  # con 8 circuitos o menos la salida va directa del TM, sin bornera
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida F/N/PE")  # 3 ferrules de salida: fase, neutro y tierra
                continue  # este circuito no pertenece a ningún grupo, ya se procesaron sus ferrules

            indices_grupo = meta.get("indices", [])  # lista de circuitos que comparten este mismo diferencial
            n_circuitos_grupo = int(len(indices_grupo))  # cuántos circuitos comparten este diferencial
            dif_a = meta.get("dif", None)  # corriente nominal (A) del diferencial de este grupo
            sec_dif = seccion_A1(dif_a)  # tabla métrica limpia (1,5-2,5-4-6-10-16-25-35), igual que el omnipolar

            if gid not in grupos_ya_contados:  # grupo nuevo, todavía no se le cuenta la entrada
                # ferrules de entrada al diferencial, solo se cuentan 1 vez por grupo
                agregar_ferrules(sec_dif, 4, f"Diferencial G{gid} entrada desde barra bipolar")  # 4 ferrules de la barra bipolar a la entrada del diferencial
                grupos_ya_contados.add(gid)  # marca este grupo para no repetir el conteo

            if n_circuitos_grupo == 1:  # caso: el diferencial protege un solo circuito
                # el diferencial tiene 2 salidas: fase y neutro. La FASE va
                # del diferencial a la entrada del TM — este tramo usa la
                # sección del DIFERENCIAL (sec_dif), no la del circuito,
                # porque todavía es cable "grueso" antes de llegar al TM.
                agregar_ferrules(sec_dif, 2, f"{circ} - diferencial a TM")  # 2 ferrules en el tramo diferencial a TM, con calibre del diferencial

                # El NEUTRO del diferencial no pasa por el TM (el TM solo
                # censura la fase) — va directo a una bornera propia, con
                # el mismo criterio: sección del diferencial, 1 ferrule en
                # la salida del diferencial + 1 en la entrada de la bornera.
                agregar_ferrules(sec_dif, 2, f"{circ} - diferencial neutro a bornera")  # 2 ferrules: salida de neutro del diferencial y entrada de la bornera
                if sec_dif not in _borneras_neutro_acum:  # primera vez que aparece esta sección en el acumulado de neutros
                    _borneras_neutro_acum[sec_dif] = {"cantidad": 0, "circuitos": []}  # crea el registro para esta sección
                _borneras_neutro_acum[sec_dif]["cantidad"] += 1  # suma 1 al conteo de esta sección
                _borneras_neutro_acum[sec_dif]["circuitos"].append(circ)  # guarda a que circuito pertenece esa bornera de neutro

                # Salida de la FASE: si el proyecto tiene más de 8 circuitos,
                # también necesita su propia bornera de organización — es un
                # conductor distinto al neutro, así que ambas borneras
                # coexisten para este circuito, no se reemplazan entre sí.
                # Mismo criterio que el caso de circuitos compartidos: entrada
                # a la bornera y salida al circuito, las dos con la sección del circuito (calibre
                # del circuito).
                if n_circ > 8:  # igual que antes: con más de 8 circuitos se usan borneras de organización
                    agregar_ferrules(sec_circ, 2, f"{circ} - salida TM a entrada bornera organización (fase)")  # 2 ferrules entre la salida del TM y la bornera de fase
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida al circuito (fase+neutro+tierra, cada una desde su propia bornera)")  # 3 ferrules de salida al circuito, uno por conductor
                    if sec_circ not in _borneras_organizacion_acum:  # primera vez que aparece esta sección en el acumulado
                        _borneras_organizacion_acum[sec_circ] = {"cantidad": 0, "circuitos": []}  # crea el registro para esta sección
                    _borneras_organizacion_acum[sec_circ]["cantidad"] += 1  # suma 1 al conteo de esta sección
                    _borneras_organizacion_acum[sec_circ]["circuitos"].append(circ)  # anota el circuito en la bornera de organización
                else:  # con pocos circuitos la salida va directa, sin bornera
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida F/N/PE")  # 3 ferrules de salida: fase, neutro y tierra

            elif n_circuitos_grupo in (2, 3):  # el diferencial lo comparten 2 o 3 circuitos
                # el diferencial comparte 2 o 3 circuitos: necesita una barra extra para repartir
                if gid not in grupos_barra_extra_contados:  # solo se agrega la barra extra una vez por grupo
                    agregar_ferrules(sec_dif, 4, f"Diferencial G{gid} a barra extra")  # 4 ferrules del diferencial a la barra extra que reparte
                    grupos_barra_extra_contados.add(gid)  # marca este grupo para no repetir

                # "barra extra a TM" usa sec_dif (cable grueso), no sec_circ
                # — todavía es cable del calibre del diferencial, igual que
                # "diferencial a TM" en el caso de 1 circuito por diferencial
                agregar_ferrules(sec_dif, 2, f"{circ} - barra extra a TM")  # 2 ferrules de la barra extra a la entrada del TM
                if n_circ > 8:  # igual que antes: más de 8 circuitos usa bornera de organización
                    agregar_ferrules(sec_circ, 2, f"{circ} - salida TM a entrada bornera")  # 2 ferrules entre la salida del TM y la bornera
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida bornera al circuito")  # 3 ferrules de la bornera al circuito
                    if sec_circ not in _borneras_organizacion_acum:  # primera vez que aparece esta sección en el acumulado
                        _borneras_organizacion_acum[sec_circ] = {"cantidad": 0, "circuitos": []}  # crea el registro para esta sección
                    _borneras_organizacion_acum[sec_circ]["cantidad"] += 1  # suma 1 al conteo de esta sección
                    _borneras_organizacion_acum[sec_circ]["circuitos"].append(circ)  # anota el circuito en esa bornera de organización
                else:  # con 8 circuitos o menos no se usa bornera de organización
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida F/N/PE")  # 3 ferrules a la salida del TM: fase, neutro y tierra

            else:  # grupo con 4 o más circuitos colgando del mismo diferencial
                # cualquier otro caso (no debería pasar, pero por seguridad se cubre igual)
                agregar_ferrules(sec_dif, 2, f"{circ} - diferencial a TM")  # 2 ferrules del tramo diferencial->TM (salida del dif y entrada del TM), con el calibre del diferencial
                if n_circ > 8:  # igual que antes: más de 8 circuitos usa bornera de organización
                    agregar_ferrules(sec_circ, 2, f"{circ} - salida TM a entrada bornera")  # 2 ferrules del TM a la entrada de la bornera
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida bornera al circuito")  # 3 ferrules de la bornera al circuito (F/N/PE)
                    if sec_circ not in _borneras_organizacion_acum:  # primera vez que aparece esta sección en el acumulado
                        _borneras_organizacion_acum[sec_circ] = {"cantidad": 0, "circuitos": []}  # crea el registro para esta sección
                    _borneras_organizacion_acum[sec_circ]["cantidad"] += 1  # suma 1 al conteo de esta sección
                    _borneras_organizacion_acum[sec_circ]["circuitos"].append(circ)  # deja anotado que circuito uso esta bornera
                else:  # con 8 circuitos o menos no se usa bornera, va directo
                    agregar_ferrules(sec_circ, 3, f"{circ} - salida F/N/PE")  # salida directa del TM al circuito: fase, neutro y tierra

        # Borneras de neutro (modo "1 diferencial por circuito"): 1 fila por
        # sección, con la cantidad total y el detalle de cada circuito en
        # la columna Circuito (en vez de 1 fila repetida por cada circuito).
        for _sec_b, _info_b in sorted(_borneras_neutro_acum.items()):  # recorre las secciones de bornera de neutro que se juntaron
            _sec_b_txt = texto_seccion(_sec_b)  # pasa la sección a texto con coma decimal, ej 2.5 queda "2,5"
            _circ_txt = ", ".join(_info_b["circuitos"])  # une los nombres de los circuitos en un solo texto
            add_row(  # arma la fila de bornera de neutro para el Excel
                desc=f"Bornera de conexión {_sec_b_txt}mm^2",  # descripción con la sección en mm^2
                marcas_txt=marcas.get("Bornera de conexion", ""),  # marca sugerida para la bornera
                norma="RIC 2 (6.2.12)",  # artículo del RIC que pide la bornera
                circuito=f"Neutro diferencial (1 por circuito): {_circ_txt}",  # detalle de a que circuitos corresponde
                unidad="u",  # las borneras se cotizan por unidad
                k=1,  # sin holgura, se cuentan exactas
                longitud_m=f"{_info_b['cantidad']} unid",  # en vez de metros va la cantidad de unidades
                cantidad=_info_b["cantidad"]  # total de borneras de esta sección
            )

        # Borneras de organización (n_circ>8): 1 fila por sección, mismo
        # criterio que la bornera de neutro — 1 bornera por cada circuito
        # que use esta rama.
        for _sec_o, _info_o in sorted(_borneras_organizacion_acum.items()):  # recorre las secciones de bornera de organización juntadas
            _sec_o_txt = texto_seccion(_sec_o)  # convierte la sección numerica a texto legible
            _circ_o_txt = ", ".join(_info_o["circuitos"])  # une los nombres de los circuitos en un solo texto
            add_row(  # arma la fila de bornera de organización para el Excel
                desc=f"Bornera de conexión {_sec_o_txt}mm^2",  # descripción con la sección en mm^2
                marcas_txt=marcas.get("Bornera de conexion", ""),  # marca sugerida para la bornera
                norma="RIC 2 (6.2.12)",  # artículo del RIC que pide la bornera
                circuito=f"Organización salida circuitos (>8 circuitos): {_circ_o_txt}",  # deja claro que son las del caso de más de 8 circuitos
                unidad="u",  # se cotizan por unidad
                k=1,  # sin holgura
                longitud_m=f"{_info_o['cantidad']} unid",  # en vez de metros va la cantidad de unidades
                cantidad=_info_o["cantidad"]  # total de borneras de organización de esta sección
            )

    # -------------------------
    # C) ENCHUFES COMUNES
    # -------------------------
    if {"Circuito", "Conductor"}.issubset(set(circuitos_df.columns)):  # revisa que existan las columnas necesarias para procesar enchufes
        fallback_enchufes = False  # marca si hubo que estimar enchufes en vez de contarlos por item

        for _, r in circuitos_df.iterrows():  # recorre cada circuito del tablero
            circ = str(r.get("Circuito", "")).strip()  # nombre del circuito

            if "enchufe" not in circ.lower():  # si el nombre del circuito no dice enchufe, no es de este bloque
                continue  # solo interesan los circuitos de enchufes

            sec_ench = seccion_circuito(r.get("Conductor", ""))  # sección en mm^2 del conductor de este circuito de enchufes
            n_ench_circ = contar_enchufes_items((items_por_nombre or {}).get(circ, []))  # cuenta cuántos enchufes tiene este circuito, según sus items


            if n_ench_circ > 0:  # si el circuito tiene enchufes contados, se suman sus ferrules
                agregar_ferrules(sec_ench, 3 * n_ench_circ, f"{circ} - enchufes ({n_ench_circ})")  # 3 ferrules por enchufe: fase, neutro y tierra
            else:  # sin items detallados no se puede contar enchufe por enchufe
                fallback_enchufes = True  # no se pudo contar por items, hay que estimar más abajo

        # si algún circuito de enchufes no tenía items detallados, se estima
        # con el total general de enchufes por ambiente
        if fallback_enchufes and ambientes_df is not None and "Cantidad enchufes (u)" in ambientes_df.columns:  # solo entra si falto contar y la hoja de ambientes trae esa columna
            # suma la columna 'Cantidad enchufes (u)' de todos los ambientes
            # (total estimado del proyecto)
            total_enchufes = int(  # total de enchufes estimados de todo el proyecto
                pd.to_numeric(
                    ambientes_df["Cantidad enchufes (u)"],  # columna con los enchufes de cada ambiente
                    errors="coerce"  # lo que no sea número queda como NaN
                ).fillna(0).sum()  # los NaN pasan a 0 y se suma toda la columna
            )

            # acá se van juntando las secciones de conductor usadas
            # en los circuitos de enchufe
            secciones_ench = []

            # recorre todos los circuitos buscando cuáles son de enchufe
            for _, r in circuitos_df.iterrows():
                if "enchufe" in str(r.get("Circuito", "")).lower():  # filtra solo los circuitos de enchufes
                    sec = seccion_circuito(r.get("Conductor", ""))  # sección del conductor de este circuito

                    # guarda la sección para después usar la más grande
                    if sec is not None:
                        secciones_ench.append(float(sec))  # se guarda como float para poder compararlas después

            sec_enchufe = max(secciones_ench) if secciones_ench else None  # usa la sección más grande, para no quedar corto

            if total_enchufes > 0:  # si la estimación dio 0 no hay nada que agregar
                agregar_ferrules(sec_enchufe, 3 * total_enchufes, f"Enchufes comunes ({total_enchufes})")  # 3 ferrules por cada enchufe estimado (F/N/T)

    # -------------------------
    # C2) ENCHUFE CLIMATIZACIÓN (solo cuando TM <= 16A)
    # -------------------------
    for _cl in _clima_items:  # recorre cada equipo de climatización detectado
        if _cl["con_enchufe"]:  # solo si el equipo tiene enchufe (TM<=16A)
            # busca la sección del conductor de este circuito de climatización
            sec_cl = seccion_circuito(
                circuitos_df[circuitos_df["Circuito"].astype(str).str.contains(  # filtra las filas cuyo nombre de circuito contenga el del equipo
                    _cl["circ"].split("(")[0].strip(), regex=False, na=False  # usa solo el nombre antes del paréntesis, sin la potencia
                )]["Conductor"].iloc[0]  # toma el conductor de la primera fila que calce
                if len(circuitos_df[circuitos_df["Circuito"].astype(str).str.contains(  # misma búsqueda de nuevo, pero solo para ver si encontró algo
                    _cl["circ"].split("(")[0].strip(), regex=False, na=False  # mismo nombre recortado que en la búsqueda de arriba
                )]) > 0 else ""  # si no calza ninguna fila, entrega texto vacío y cae al 2,5mm^2 por defecto
            ) or 2.5  # si no encuentra el circuito, usa 2.5mm^2 por defecto
            agregar_ferrules(sec_cl, 3, f"{_cl['circ']} - enchufe climatización F/N/T")  # 3 ferrules: fase, neutro y tierra del enchufe de climatización

    # NOTA: los circuitos especiales sin enchufe (horno, lavadora,
    # climatización sin enchufe van con conector cónico o estaño + cintas)
    # se calculan arriba, en el bloque "CONECTORES CÓNICOS (POR SECCIÓN)",
    # rama es_especial_row (línea ~4477 del archivo completo), para que
    # se sumen a los totales generales de cónico/estaño/cinta del proyecto,
    # en vez de generar un rollo aparte por cada circuito especial.

    # -------------------------
    # D2) AGUA CALIENTE — ferrules tablero externo + equipo con enchufe
    # Los conductores van directo a bornes del TM/bornera y, si tiene enchufe,
    # a los bornes propios del enchufe (TM≤16A, igual criterio que climatización):
    #   - 3 ferrules entrada equipo (F+N+PE) — SOLO si tiene enchufe (TM≤16A)
    #     Si no tiene enchufe (TM>16A), la conexión es cola de rata: cónico/
    #     estaño, ya contado arriba en el bloque "CONECTORES CÓNICOS (POR
    #     SECCIÓN)", rama es_especial_row (línea ~4477 del archivo completo)
    #   - Si además lleva tablero externo:
    #       2 ferrules entrada TM (F+N)
    #       2 ferrules salida TM (F+N)
    #       2 ferrules bornera PE (entrada+salida tierra)
    # -------------------------
    if {"Circuito", "Conductor"}.issubset(set(circuitos_df.columns)):  # sigue solo si el DataFrame tiene estas columnas
        for _, r_ac in circuitos_df.iterrows():  # recorre cada circuito buscando los de agua caliente
            circ_ac = str(r_ac.get("Circuito", "")).strip()  # nombre del circuito
            # detecta si este circuito es de agua caliente, por palabras
            # clave en el nombre
            _es_agua_ferrul = any(k in circ_ac.lower() for k in
                                  ("ducha", "termo", "calefon", "calefón",
                                   "calentador", "agua caliente"))  # resto de palabras que indican agua caliente
            if not _es_agua_ferrul:  # el circuito no es de agua caliente
                continue  # no es agua caliente, se salta

            sec_ac_f = seccion_circuito(r_ac.get("Conductor", "")) or 4.0  # sección del conductor, 4.0mm^2 si no se pudo leer
            in_tm_ac_row2 = parse_in_tm(str(r_ac.get("Disyuntor termomagnético", "")))  # corriente del TM de este circuito

            # Ferrules en los bornes del enchufe — solo si tiene enchufe (TM≤16A).
            # Si el TM no se pudo leer, no se genera nada acá: ya se avisó una
            # vez para este circuito con la fila "(definir)" del bloque de
            # cónicos (línea ~4446 del archivo completo).
            if in_tm_ac_row2 is not None and in_tm_ac_row2 <= 16:
                agregar_ferrules(sec_ac_f, 3, f"{circ_ac} - enchufe equipo F/N/PE")  # 3 ferrules: fase, neutro y tierra en el enchufe del equipo

            # Buscar si lleva tablero externo
            _datos_ac_f = {}  # acá se guardan los datos de este equipo si se encuentra
            # busca en la lista de equipos de agua caliente el que
            # coincide con este circuito
            for eq_ac_f in circuitos_agua_caliente:
                if eq_ac_f.get("nombre_circ", "").lower() in circ_ac.lower():  # calza si el nombre del equipo aparece dentro del nombre del circuito
                    _datos_ac_f = eq_ac_f  # se queda con los datos de ese equipo de agua caliente
                    break  # ya lo encontró, corta la búsqueda

            if _datos_ac_f.get("lleva_tablero_externo", False) and in_tm_ac_row2 is not None:  # solo si el equipo tiene tablero externo y se pudo leer el TM
                sec_int_tab = seccion_A1(in_tm_ac_row2)  # sección del conductor interior del tablero externo
                # 6 ferrules en el tablero externo (el TM ya trae sus propios
                # bornes, no se usa bornera F+N):
                # 2 entrada TM (F+N), 2 salida TM (F+N), 2 bornera PE (tierra)
                agregar_ferrules(sec_int_tab, 2, f"{circ_ac} - tablero externo entrada TM (F+N)")  # 2 ferrules a la entrada del TM del tablero externo
                agregar_ferrules(sec_int_tab, 2, f"{circ_ac} - tablero externo salida TM (F+N)")  # 2 ferrules a la salida del TM
                agregar_ferrules(sec_ac_f,    2, f"{circ_ac} - tablero externo bornera PE (tierra)")  # 2 ferrules en la bornera de tierra, con la sección del circuito

    # -------------------------
    # D3) ESPECIALES GENÉRICOS (horno, lavadora, encimera, etc.) — ferrules
    # equipo con enchufe. Mismo criterio y mismo formato que climatización (C2)
    # y agua caliente (D2): si TM≤16A, el equipo tiene enchufe, entonces van 3 ferrules
    # F/N/PE en los bornes de ese enchufe.
    # -------------------------
    if {"Circuito", "Conductor"}.issubset(set(circuitos_df.columns)):  # sigue solo si existen estas columnas
        for _, r_esp in circuitos_df.iterrows():  # recorre cada circuito buscando los especiales genéricos
            circ_esp = str(r_esp.get("Circuito", "")).strip()  # nombre del circuito
            # detecta si el circuito es un equipo especial genérico
            # (horno, lavadora, etc.) por palabras clave
            _es_esp_ferrul = any(k in circ_esp.lower() for k in
                                  ("especial", "horno", "encimera", "lavadora",
                                   "lavaplatos", "secadora", "jacuzzi", "piscina"))  # resto de equipos que cuentan como especiales
            # pero excluye los que ya se cuentan en agua caliente o
            # climatización, para no duplicar ferrules
            _es_agua_o_clima = any(k in circ_esp.lower() for k in
                                    ("ducha", "termo", "calefon", "calefón", "calentador",
                                     "agua caliente", "climatiz", "aire", "split", "ac ", "a/c"))  # palabras de agua caliente y clima, esas se descartan acá
            if not _es_esp_ferrul or _es_agua_o_clima:  # no es especial genérico, o ya se contó en agua/clima
                continue  # no es un especial genérico nuevo (o ya se contó en agua/clima), se salta

            sec_esp = seccion_circuito(r_esp.get("Conductor", "")) or 2.5  # sección del conductor, 2.5mm^2 si no se pudo leer
            in_tm_esp = parse_in_tm(str(r_esp.get("Disyuntor termomagnético", "")))  # corriente del TM de este circuito

            # Si el TM no se pudo leer, no se asume "con enchufe" — el aviso
            # "(definir)" ya se generó una vez para este mismo circuito en el
            # bloque de cónicos (línea ~4473 del archivo completo), no hace
            # falta duplicarlo ni adivinar acá.
            if in_tm_esp is not None and in_tm_esp <= 16:
                agregar_ferrules(sec_esp, 3, f"{circ_esp} - enchufe equipo F/N/PE")  # 3 ferrules: fase, neutro y tierra en el enchufe del equipo

    # -------------------------
    # E) INTERRUPTORES
    # -------------------------
    if ambientes_df is not None:  # solo si hay datos de ambientes
        secciones_ilum = []  # acá se juntan las secciones de los circuitos de iluminación

        # busca la sección más grande usada en los circuitos de iluminación
        if {"Circuito", "Conductor"}.issubset(set(circuitos_df.columns)):  # sigue solo si existen estas columnas
            for _, r in circuitos_df.iterrows():  # recorre los circuitos buscando los de iluminación
                if "ilumin" in str(r.get("Circuito", "")).lower():  # filtra solo los circuitos de iluminación
                    sec = seccion_circuito(r.get("Conductor", ""))  # sección del conductor de este circuito

                    # guarda la sección para usar la más grande al final
                    if sec is not None:
                        secciones_ilum.append(float(sec))  # se guarda como float para poder compararlas después

        sec_iluminacion = max(secciones_ilum) if secciones_ilum else None  # usa la sección más grande encontrada (o None si no hay ninguna)

        # recorre cada ambiente y calcula cuántos ferrules necesitan sus interruptores
        for _, ar in ambientes_df.iterrows():  # recorre cada ambiente (pieza, baño, cocina, etc.)
            amb = str(ar.get("Ambiente", "")).strip()  # nombre del ambiente
            n_lum = int(pd.to_numeric(ar.get("Cantidad luminarias (u)", 0), errors="coerce") or 0)  # cantidad de luminarias de este ambiente
            n_con = int(pd.to_numeric(ar.get("N_conmutadas_924 (u)", 0), errors="coerce") or 0)  # cuántas de esas luminarias son conmutadas (interruptor 9/24)

            n_con = max(0, min(n_con, n_lum))  # no puede haber más conmutadas que luminarias totales
            n_rest = max(0, n_lum - n_con)  # el resto de luminarias que van con interruptor normal

            c12, c15, c32 = descomponer_interruptores(n_rest)  # reparte esas luminarias en interruptores 9/12, 9/15 y 9/32 según cuántas controla cada uno

            # Ferrules por interruptor: 9/12 lleva 2 · 9/15 lleva 5 · 9/32 lleva 8 · conmutado 9/24 lleva 6 (fijo por grupo, sin importar cuántas luminarias controle)
            grupo_924_ferrul = 1 if n_con > 0 else 0  # el grupo de conmutados cuenta 1 sola vez, aunque tenga varias luminarias
            total_int = (c12 * 2) + (c15 * 5) + (c32 * 8) + (grupo_924_ferrul * 6)  # suma los ferrules de todos los tipos de interruptor de este ambiente

            if total_int > 0:  # si el ambiente no lleva interruptores no se agrega nada
                agregar_ferrules(sec_iluminacion, total_int, f"Interruptores {amb}")  # agrega los ferrules de los interruptores de este ambiente

    # -------------------------
    # F) AGREGAR FILAS
    # -------------------------
    # arma una fila del Excel por cada sección de ferrul acumulada, con el
    # detalle de a qué corresponde cada cantidad (para que se pueda revisar)
    for sec, detalles in sorted(ferrules.items()):  # recorre cada sección de ferrul juntada, en orden
        cantidad_total = int(sum(detalles.values()))  # suma todas las cantidades de esa sección
        color = color_ferrul_por_seccion(sec)  # color de ferrul que corresponde a esta sección

        # arma el texto con el detalle de donde se uso cada cantidad,
        # una línea por item
        detalle_txt = "\n".join(
            [f"{nombre} ({cant}u)" for nombre, cant in detalles.items()]  # una línea por item: donde se usaron los ferrules y cuántos
        )

        # agrega la fila de ferrules de esta sección al listado de materiales
        add_row(
            desc=f"Terminal ferrul color {color} {texto_seccion(sec)}mm^2",  # nombre que va al Excel, con color y sección del ferrul
            marcas_txt=marcas.get("Ferrule", ""),  # marcas sugeridas de ferrul
            norma="RIC 4.1.1",  # artículo del RIC
            circuito=detalle_txt,  # el detalle armado arriba, un item por línea
            unidad="u",  # los ferrules se compran por unidad
            k=1,  # sin holgura, se cuentan exactos
            longitud_m=f"{cantidad_total} unid",  # en vez de metros va la cantidad
            cantidad=cantidad_total  # total de ferrules de esta sección
        )

        # =========================
    # CABLEADO INTERIOR DEL TABLERO
    # =========================
    # Cálculo basado en dimensiones reales del tablero y posición en riel DIN
    # Dimensiones tablero (alto x ancho en cm):
    TAB_DIMS = {  # clave = cantidad de puestos del tablero, valor = (alto, ancho) en cm
        2:  (17, 11),  4:  (20, 12),  6:  (20, 16),  8:  (20, 20),  # tableros chicos, de 2 a 8 puestos
        12: (25, 28), 16: (26, 36), 18: (26, 39), 24: (36, 30),  # tableros medianos
        36: (48, 30), 42: (54, 36), 48: (65, 35), 54: (54, 44),  # tableros grandes
        56: (75, 36), 72: (75, 46),  # los más grandes, de 56 y 72 puestos
    }
    tab_h_cm, tab_w_cm = TAB_DIMS.get(_puestos_tablero, (36, 30))  # medidas del tablero de este proyecto (36x30 si no está en la tabla)

    # ============================================================
    # REGLA de estimación de largo de cableado interior del tablero:
    #
    #  - Por cada circuito: 1m Rojo + 1m Blanco + 1m Verde, fijo, en la
    #    sección de ESE circuito (sec_circ) — sin importar diferencial,
    #    barra extra, ni posición en el riel.
    #  - Luz piloto + portafusible: 1 sola vez para todo el tablero,
    #    +1m Rojo +1m Blanco (sin verde, no necesita tierra).
    #  - Tramo grueso del tablero (entre la barra principal y el omnipolar / SPD /
    #    protector sobrevoltaje / diferenciales — el tramo grueso se estima
    #    todo con sec_omni_A1, aunque el calibre del diferencial se calcule
    #    aparte con su propia corriente):
    #      Verde: 1m fijo (salida del SPD a tierra)
    #      Rojo = Blanco = (alto + ancho + alto) del tablero, en metros,
    #             × cantidad total de circuitos del proyecto
    #  - El alimentador (F/N/PE hasta la barra principal) NO se cuenta acá
    #    — ya está incluido en el metraje propio del alimentador.
    # ============================================================
    # el alto se cuenta 2 veces y el ancho 1 vez, en metros
    _factor_dim_m = (tab_h_cm + tab_w_cm + tab_h_cm) / 100.0
    HOLGURA_TABLERO = 1.10  # +10% de holgura sobre el total final

    add_section("Cableado interior tablero")  # nueva sección en la hoja de materiales

    cableado_tablero = {}  # acá se va acumulando el cable interior del tablero, por sección y color

    # suma metros de cable interior del tablero, agrupados por sección y
    # color. seccion_mm2: sección del conductor; colores: uno o una lista;
    # longitud_base: metros a sumar; detalle: de donde viene ese cable
    def agregar_cable_tablero(seccion_mm2, colores, longitud_base, detalle):
        # va sumando metros de cable interior del tablero por sección y color
        sec = normalizar_seccion_ferrul(seccion_mm2)  # redondea a sección comercial
        if sec is None:  # la sección no era un número válido (si es más grande que todas, devuelve 35)
            return  # sección inválida, no se puede agregar
        if isinstance(colores, str):  # vino un solo color como texto y no una lista
            colores = [colores]  # si mandaron un solo color como texto, lo mete en una lista
        if sec not in cableado_tablero:  # sección que aparece por primera vez en el tablero
            cableado_tablero[sec] = {"base_por_color": {}, "detalles": []}  # primera vez que aparece esta sección
        if detalle not in cableado_tablero[sec]["detalles"]:  # evita repetir el mismo detalle dos veces
            cableado_tablero[sec]["detalles"].append(detalle)  # anota de dónde viene este cable
        for color in colores:  # reparte los mismos metros a cada color pedido
            color = str(color).strip().capitalize()  # ej: "rojo" queda como "Rojo"
            cableado_tablero[sec]["base_por_color"][color] = (  # acumula los metros de este color en esta sección
                cableado_tablero[sec]["base_por_color"].get(color, 0.0)  # parte de lo que ya había acumulado antes
                + float(longitud_base)  # suma los metros de este color
            )

    # n_circ ya viene calculado más arriba (cuenta solo filas con "Circuito"
    # no vacío) — se reutiliza acá para no contar de nuevo con un método
    # menos preciso.

    # -------------------------
    # A) TABLERO FIJO — tramo grueso (todo comparte sec_omni_A1)
    # -------------------------
    agregar_cable_tablero(sec_omni_A1, ["Rojo", "Blanco"],
                           _factor_dim_m * n_circ,  # largo del tramo grueso multiplicado por la cantidad de circuitos
                           "Barra principal a omnipolar, SPD, protector sobrevoltaje y diferenciales (tramo grueso)")  # rojo y blanco, largo = tramo grueso x cantidad de circuitos
    agregar_cable_tablero(sec_omni_A1, "Verde", 1.0, "SPD a barra PE (fijo)")  # 1 metro fijo de verde (tierra) para el tramo grueso

    # Luz piloto + portafusible: 1 sola vez para todo el tablero
    if usa_luz_piloto:  # solo si el tablero lleva luz piloto
        agregar_cable_tablero(sec_control, "Rojo",   1.0, "Luz piloto + portafusible")  # 1m de rojo (fase) para la luz piloto
        agregar_cable_tablero(sec_control, "Blanco", 1.0, "Luz piloto + portafusible")  # 1m de blanco (neutro) para la luz piloto

    # -------------------------
    # B) POR CADA CIRCUITO — 1m Rojo + 1m Blanco + 1m Verde, fijo
    # Usa el conductor que ya quedó definido para ese circuito (sec_circ),
    # no uno recalculado de nuevo con otro método.
    # -------------------------
    if {"Circuito", "Conductor", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # sigue solo si existen estas columnas
        for idx, r in circuitos_df.iterrows():  # recorre cada circuito del proyecto
            circ = str(r.get("Circuito", "")).strip()  # nombre del circuito
            cond = str(r.get("Conductor", "")).strip()  # texto del conductor (ej: 3x2.5mm^2)
            tm   = str(r.get("Disyuntor termomagnético", "")).strip()  # texto del TM
            if not circ or not cond or not tm:  # al circuito le falta nombre, conductor o TM
                continue  # falta algún dato, no se puede calcular, se salta

            sec_circ = seccion_circuito(cond)  # sección del conductor de este circuito

            # 1 metro de cada color (rojo, blanco, verde) para este circuito
            agregar_cable_tablero(sec_circ, "Rojo",   1.0, f"{circ} - cableado interior")  # fase
            agregar_cable_tablero(sec_circ, "Blanco", 1.0, f"{circ} - cableado interior")  # neutro
            agregar_cable_tablero(sec_circ, "Verde",  1.0, f"{circ} - cableado interior")  # tierra

        # -------------------------
    # C) AGREGAR FILAS
    # -------------------------
    orden_colores = ["Rojo", "Blanco", "Verde"]  # orden en que se muestran los colores en el detalle

    # arma una fila del Excel por cada sección de cable acumulada
    for sec, data in sorted(cableado_tablero.items()):  # recorre cada sección de cable acumulada (ordenadas), con sus metros por color
        base_por_color = data["base_por_color"]  # metros acumulados por color, para esta sección de cable

        metros_base_total = round(sum(base_por_color.values()), 2)  # suma de todos los colores, sin holgura
        metros_final_total = round(metros_base_total * HOLGURA_TABLERO, 2)  # aplica la holgura de tablero: cantidad final a comprar

        resumen_colores = []  # aquí se arma el texto tipo 'Rojo: 12,5 = 13m' por color

        for color in orden_colores:  # recorre Rojo, Blanco y Verde en ese orden fijo
            if color in base_por_color:  # solo si este color se uso en esta sección
                # arma el texto legible de este color, ej: 'Rojo: 12,5 = 13m'
                base_color = round(base_por_color[color], 2)  # metros base de este color, redondeado a 2 decimales
                _txt = f"{base_color}"  # pasa el número a texto
                if "." in _txt:  # si tiene parte decimal...
                    _txt = _txt.rstrip("0").rstrip(".")  # ...saca los ceros y el punto sobrantes, ej 12.50 -> 12.5
                _base_color_txt = _txt.replace(".", ",")  # usa coma decimal (formato chileno)
                # guarda 'Color: base = techoM' para mostrarlo en el Excel
                resumen_colores.append(
                    f"{color}: {_base_color_txt} = {math.ceil(base_color)}m"  # queda como 'Rojo: 12,5 = 13m': el crudo y el redondeo hacia arriba
                )

        detalle_longitud = "\n".join(resumen_colores)  # une los textos de cada color con salto de línea, para la celda del Excel

        detalles_txt = "\n".join(data["detalles"]).lower()  # todo el texto de detalles juntos, en minúsculas, para buscar palabras clave

        componentes = []  # lista de que elementos del tablero usan este calibre de cable

        # busca palabras clave en el texto de detalles para saber que
        # partes del tablero usan este calibre (así el Excel muestra algo
        # entendible, en vez de solo el nombre tecnico de la sección)
        if "barra principal" in detalles_txt:  # el detalle nombra la barra principal
            componentes.append("Barras repartidoras")  # el tramo grueso sale de las barras repartidoras

        if "omnipolar" in detalles_txt:  # el detalle nombra el omnipolar
            componentes.append("Interruptor general")  # el omnipolar es el interruptor general del tablero

        if "spd" in detalles_txt:  # el detalle nombra el SPD
            componentes.append("SPD")  # protector contra sobrevoltajes transitorios

        if "portafusible" in detalles_txt:  # el detalle nombra el portafusible
            componentes.append("Portafusible")  # portafusible de la luz piloto

        if "luz piloto" in detalles_txt:  # el detalle nombra la luz piloto
            componentes.append("Luz piloto")  # luz piloto del tablero

        if "diferencial" in detalles_txt:  # el detalle nombra los diferenciales
            componentes.append("Interruptores diferenciales")  # cable que llega a los diferenciales

        if ("tm" in detalles_txt) or ("circuito" in detalles_txt):  # si el detalle menciona TM o circuitos, es cable de salida
            componentes.append("Interruptores termomagnéticos y salidas de circuitos")  # cable de los TM y las salidas a cada circuito

        if not componentes:  # si no calzo ninguna palabra clave, deja una etiqueta genérica
            componentes.append("Cableado interior de tablero")  # etiqueta genérica cuando no calzo ninguna palabra clave

        detalle_circuito = "Tablero: " + ", ".join(componentes)  # texto final que va en la columna 'Circuito' del Excel

        # agrega la fila de este conductor de tablero a la lista de materiales
        add_row(
            desc=f"Conductor flexible libre de halógenos {tipo_cable_default} {texto_seccion(sec)}mm^2",  # descripción del cable, libre de halógenos y con su sección
            marcas_txt=marcas.get("Conductores", ""),  # marcas sugeridas de conductor
            norma="RIC 4.1.1",  # artículo del RIC
            circuito=detalle_circuito,  # que partes del tablero usan este calibre
            unidad="m",  # el cable se compra por metro
            k=HOLGURA_TABLERO,  # el 1.10 de holgura queda anotado en la fila
            longitud_m=detalle_longitud,  # detalle de metros por color
            cantidad=metros_final_total  # metros finales a comprar, con holgura incluida
        )

    # ── CABLEADO INTERIOR TABLERO EXTERNO AGUA CALIENTE ─────────────────────
    # Conductor método A1 dentro del tablero externo (mismo tipo de cable que
    # el resto de la instalación: THWN-2 en zona húmeda, H07Z1-K en zona seca —
    # ver tipo_cable_default, definido según la zona ingresada por el usuario):
    #   Tramo de entrada del prensaestopa al TM bipolar F (Rojo): 0.25m
    #   Tramo de entrada del prensaestopa al TM bipolar N (Blanco): 0.25m
    #   Tramo del TM bipolar a la salida del prensaestopa F (Rojo): 0.25m
    #   Tramo del TM bipolar a la salida del prensaestopa N (Blanco): 0.25m
    #   Tramo de la bornera PE de entrada a la bornera PE de salida (Verde): 0.25m
    # Total por color antes de la holgura: Rojo=0.50m, Blanco=0.50m, Verde=0.25m
    # (con el 1,15 de holgura que se aplica más abajo quedan 0,58 / 0,58 / 0,29m)
    # Sección F+N: seccion_A1(In_tm); Sección tierra: sección real conductor circuito
    # Se genera como fila SEPARADA del tablero principal (no se mezcla)
    if {"Circuito", "Conductor", "Disyuntor termomagnético"}.issubset(set(circuitos_df.columns)):  # solo sigue si el DataFrame de circuitos tiene esas columnas
        for _, r_cab_ac in circuitos_df.iterrows():  # recorre cada circuito ingresado por el usuario
            circ_cab = str(r_cab_ac.get("Circuito", "")).strip()  # nombre del circuito, ej 'Circuito 5 - Ducha'
            # revisa si el nombre del circuito corresponde a un equipo de
            # agua caliente (ducha, termo, calefont, calentador, etc.)
            _es_agua_cab = any(k in circ_cab.lower() for k in
                               ("ducha", "termo", "calefon", "calefón",
                                "calentador", "agua caliente"))  # resto de palabras que indican agua caliente
            if not _es_agua_cab:  # no es un circuito de agua caliente: pasa al siguiente
                continue  # circuito común, no le corresponde este cableado
            _datos_cab = {}  # aquí se guardan los datos ingresados de este equipo de agua caliente
            # busca en la lista de equipos de agua caliente el que coincide con este circuito
            for eq_cab in circuitos_agua_caliente:  # recorre los equipos de agua caliente que ingreso el usuario
                if eq_cab.get("nombre_circ", "").lower() in circ_cab.lower():  # calza el nombre del equipo con el nombre del circuito
                    _datos_cab = eq_cab  # guarda los datos de ese equipo
                    break  # ya lo encontró, no sigue buscando
            if not _datos_cab.get("lleva_tablero_externo", False):  # si este equipo no tiene tablero externo, no le corresponde este cableado
                continue  # equipo de agua caliente sin tablero externo, no lleva cableado aparte

            in_tm_cab = parse_in_tm(str(r_cab_ac.get("Disyuntor termomagnético", "")))  # corriente del TM bipolar del tablero externo (en Amperes)
            if in_tm_cab is None:  # no se pudo leer la corriente del TM desde el texto
                continue  # el aviso "(definir)" por TM ilegible ya se generó en el bloque de cónicos (línea ~4446 del archivo completo)
            sec_int_cab  = seccion_A1(in_tm_cab)   # F+N: método A1
            sec_ext_cab  = extraer_seccion_mm2(str(r_cab_ac.get("Conductor", ""))) or 4.0  # tierra: sección real
            L_fn  = round(0.25 * 2 * 1.15, 2)  # 2 tramos × 0.25m × holgura 1.15
            L_t   = round(0.25 * 1 * 1.15, 2)  # 1 tramo × 0.25m × holgura 1.15
            sec_fn_txt = str(sec_int_cab).replace(".", ",")  # sección F+N en texto, con coma decimal
            sec_t_txt  = str(sec_ext_cab).replace(".", ",")  # sección de tierra en texto, con coma decimal
            # fila del conductor F+N (fase y neutro) del tablero externo
            add_row(  # fila del listado para el cable de fase y neutro
                desc=f"Conductor flexible libre de halógenos {tipo_cable_default} {sec_fn_txt}mm^2 (tablero externo agua caliente F+N)",  # descripción que sale en el Excel, con tipo de cable y sección
                marcas_txt=marcas.get("Conductores", ""),  # marcas sugeridas para conductores
                norma="RIC 4.1.1",  # artículo del RIC que aplica a conductores
                circuito=f"{circ_cab} - tablero externo",  # queda identificado como cableado del tablero externo
                unidad="m",  # el cable se mide en metros
                k=1,  # sin multiplicador
                longitud_m=f"Rojo {L_fn}m / Blanco {L_fn}m",  # detalle de colores: rojo fase, blanco neutro
                cantidad=math.ceil(L_fn * 2)  # los dos colores juntos, redondeado hacia arriba
            )
            # fila del conductor de tierra (PE) del tablero externo
            add_row(  # fila del listado para el conductor de protección
                desc=f"Conductor flexible libre de halógenos {tipo_cable_default} {sec_t_txt}mm^2 (tablero externo agua caliente tierra)",  # misma descripción pero para el cable de tierra
                marcas_txt=marcas.get("Conductores", ""),  # marcas sugeridas para conductores
                norma="RIC 4.1.1",  # mismo artículo del RIC
                circuito=f"{circ_cab} - tablero externo",  # también va al tablero externo del equipo
                unidad="m",  # se mide en metros
                k=1,  # sin multiplicador
                longitud_m=f"Verde {L_t}m",  # el PE siempre va en verde
                cantidad=L_t  # es un solo tramo de tierra
            )
    add_section("Tornillería")  # nueva sección del Excel: tornillos y fijaciones
    tipo_can2 = (tipo_canalizacion or "").strip().lower()  # tipo de canalización (embutida o sobrepuesta), en minúsculas
    es_emb2 = "embut" in tipo_can2  # True si la canalización es embutida
    es_sob2 = "sobre" in tipo_can2  # True si la canalización es sobrepuesta
    # tornillos totales ----------
    tornillos_total = 0  # contador de tornillos de todo el proyecto
    if es_emb2:  # reglas de tornillos para canalización embutida
        tornillos_total += 2 * int(_abrazaderas_total)        # 2 por abrazadera
        tornillos_total += 4 * int(_cajas_total)              # 4 por caja
        tornillos_total += 2 * int(_tapas_ciegas_total)       # 2 por tapa ciega
        tornillos_total += 2 * int(total_luminarias_sobrepuestas)  # 2 por luminaria, SOLO las de montaje sobrepuesto (las embutidas no llevan)
        tornillos_total += (14 if int(_riel_m) == 2 else 7)   # riel DIN
        # Tornillos de fijación del tablero según su cantidad de puestos
        # más puestos en el tablero = gabinete más grande = más tornillos de fijación
        if _puestos_tablero in [2, 4, 6, 8, 12, 16]:  # tableros chicos, hasta 16 puestos
            tornillos_total += 6  # 6 tornillos de fijación
        elif _puestos_tablero in [18, 24, 36, 42]:  # tableros medianos
            tornillos_total += 8  # 8 tornillos de fijación
        elif _puestos_tablero in [48, 54, 56, 72]:  # tableros grandes
            tornillos_total += 10  # 10 tornillos de fijación
    elif es_sob2:  # reglas de tornillos para canalización sobrepuesta
        tornillos_total += int(_long_sobrepuesta_total)       # 1 por metro de canaleta
        tornillos_total += 4 * int(_cajas_total)              # 4 por caja
        tornillos_total += 2 * int(_tapas_ciegas_total)       # 2 por tapa ciega
        tornillos_total += 2 * int(total_luminarias_sobrepuestas)  # 2 por luminaria, SOLO las de montaje sobrepuesto (las embutidas no llevan)
        tornillos_total += (14 if int(_riel_m) == 2 else 7)   # riel DIN
        # Tornillos de fijación del tablero según su cantidad de puestos
        if _puestos_tablero in [2, 4, 6, 8, 12, 16]:  # tableros chicos, hasta 16 puestos
            tornillos_total += 6  # 6 tornillos de fijación
        elif _puestos_tablero in [18, 24, 36, 42]:  # tableros medianos
            tornillos_total += 8  # 8 tornillos de fijación
        elif _puestos_tablero in [48, 54, 56, 72]:  # tableros grandes
            tornillos_total += 10  # 10 tornillos de fijación
    # Según tipo de material de construcción
    col_mat = "Material tabique" if es_emb2 else "Material forrado interior"  # columna del DataFrame de ambientes que indica el material de la superficie
    # El reparto de tornillos se pondera por la cantidad real de cajas de
    # cada ambiente (luminarias + enchufes), así un ambiente con más cajas
    # aporta más al reparto de tornillos que uno con pocas.
    mats_peso = {}  # {material: peso acumulado (n° de cajas reales)}
    if ambientes_df is not None and col_mat in ambientes_df.columns:  # solo si esa columna existe con datos
        for _, _ar_mat in ambientes_df.iterrows():  # recorre cada ambiente (living, cocina, dormitorio, etc.)
            _m = str(_ar_mat.get(col_mat, "")).strip().lower()  # material de este ambiente, en minúsculas
            if not _m:  # ambiente sin material definido: se salta
                continue  # ambiente sin material declarado, no entra al reparto de tornillos
            _n_lum_amb = int(pd.to_numeric(_ar_mat.get("Cantidad luminarias (u)", 0), errors="coerce") or 0)  # cantidad de luminarias de este ambiente
            _n_ench_amb = int(pd.to_numeric(_ar_mat.get("Cantidad enchufes (u)", 0), errors="coerce") or 0)  # cantidad de enchufes de este ambiente
            _peso_amb = _n_lum_amb + _n_ench_amb  # cajas reales de este ambiente
            if _peso_amb <= 0:  # ambiente que no declaro luminarias ni enchufes
                _peso_amb = 1  # ambiente sin cajas registradas: igual cuenta como mínimo 1, para no perderlo del reparto
            mats_peso[_m] = mats_peso.get(_m, 0) + _peso_amb  # acumula el peso (cajas reales) de este material
    # relaciona cada material con su tipo de tornillo
    def tornillo_por_material(m: str) -> str:  # recibe el material y devuelve el nombre exacto del tornillo a usar
        # tornillo adecuado según el material de la superficie (metalcón,
        # volcanita, madera, etc.) y si es embutida o no
        mm = (m or "").lower()  # material en minúsculas
        if es_emb2:  # canalización embutida
            if "metalcon" in mm:  # tabique de metalcon: perfil metálico
                return 'Tornillo autoperforante punta broca 6x1 1/2"'  # para estructura metálica
            return 'Tornillo punta fina para madera cabeza trompeta 6x1"'  # para madera, por defecto
        else:  # canalización sobrepuesta
            if ("volcan" in mm) or ("vulcan" in mm) or ("fibro" in mm):  # volcanita, vulcanita o fibrocemento
                return 'Tornillo volcanita punta fina 1 1/4"'  # para volcanita/fibrocemento
            if "madera" in mm:  # forrado de madera
                return 'Tornillo punta fina para madera cabeza lenteja 6x1/2"'  # cabeza lenteja, queda más plano sobre la superficie
            return 'Tornillo (definir según forrado)'  # material no identificado

    # salida al listado (UNA fila por tipo de tornillo)
    if tornillos_total > 0:  # solo genera filas si hay al menos un tornillo que reportar
        if len(mats_peso) == 0:  # no se pudo determinar el material de ningún ambiente
            # no hay datos de material de construcción: se deja una fila genérica
            add_row(  # fila genérica de tornillos, sin detalle de material
                desc="Tornillo definir según construcción",  # el instalador define el tornillo en terreno
                marcas_txt=marcas.get("Tornillos", ""),  # marcas sugeridas de tornillería
                norma="-",  # la tornillería no tiene artículo del RIC
                circuito="General",  # no pertenece a un circuito puntual
                unidad="u",  # se cuentan por unidad
                k=1,  # sin multiplicador
                longitud_m=f"{tornillos_total} unid",  # detalle en texto para el Excel
                cantidad=tornillos_total  # el total de tornillos calculado arriba
            )
        else:  # si hay datos de material: reparte los tornillos por tipo
            # reparte el total de tornillos proporcionalmente según cuántas
            # cajas reales tiene cada material (no según cantidad de piezas)
            total_peso = sum(mats_peso.values())  # total de cajas reales de todo el proyecto
            # Estas 2 porciones nunca necesitan tarugo:
            #  - Tapas ciegas (2 c/u): se atornillan a la caja, no a la pared.
            #  - Cajas octogonales en sobrepuesta (4 c/u): van fijadas
            #    directo al tabique, no a la superficie exterior con tarugo.
            # Se separan del resto antes de repartir, y se vuelven a sumar
            # al final como tornillo simple.
            _tornillos_tapas = 2 * int(_tapas_ciegas_total)  # tornillos de tapas ciegas (nunca llevan tarugo)
            _tornillos_oct_sob = 4 * int(total_luminarias) if es_sob2 else 0  # tornillos de cajas octogonales en sobrepuesta (tampoco llevan tarugo)
            _tornillos_sin_tarugo_total = _tornillos_tapas + _tornillos_oct_sob  # total de tornillos que nunca necesitan tarugo
            tornillos_total_resto = max(0, tornillos_total - _tornillos_sin_tarugo_total)  # el resto si se reparte con tarugo, según el material
            # agrupar por tipo de tornillo
            tornillos_por_tipo = {}  # acumula cantidades por cada tipo de tornillo
            for mat, peso in mats_peso.items():  # recorre cada material encontrado en los ambientes
                desc_tipo = tornillo_por_material(mat)  # nombre del tornillo que corresponde a este material
                # frac_material: qué fracción del total de cajas del proyecto
                # corresponde a este material puntual (ej. si "volcanita" tiene
                # 30 de 50 cajas totales, frac_material = 0.6 = 60%)
                frac_material = (peso / total_peso) if total_peso > 0 else 1.0  # queda entre 0 y 1; si no hay peso, se le asigna todo
                # torn_estim: cuántos tornillos de ESTE material le tocan del
                # total elegible para tarugo, según su fracción del proyecto
                torn_estim = int(math.ceil(tornillos_total_resto * frac_material))  # redondea hacia arriba para no quedar corto
                # sin_tarugo_estim: mismo reparto proporcional, pero para la
                # porción que nunca lleva tarugo (tapas + octogonales sobrepuesta)
                sin_tarugo_estim = int(math.ceil(_tornillos_sin_tarugo_total * frac_material))  # también redondeado hacia arriba
                if desc_tipo not in tornillos_por_tipo:  # primer material que usa este tipo de tornillo
                    tornillos_por_tipo[desc_tipo] = {"tornillos": 0, "tornillos_sin_tarugo": 0, "mats": set()}  # primera vez que aparece este tipo de tornillo: se inicializa en 0
                tornillos_por_tipo[desc_tipo]["tornillos"] += torn_estim  # suma los tornillos con tarugo de este material
                tornillos_por_tipo[desc_tipo]["tornillos_sin_tarugo"] += sin_tarugo_estim  # suma los tornillos sin tarugo de este material
                tornillos_por_tipo[desc_tipo]["mats"].add(mat)  # guarda que materiales usan este tipo de tornillo (para mostrarlo en el Excel)
            # escribir UNA fila por tipo (excepto volcanita/fibrocemento en
            # sobrepuesta, que usa tarugo+tornillo en vez de tornillo directo;
            # ese caso se calcula acá mismo, no se repite en otra sección)
            for desc_tipo, info in tornillos_por_tipo.items():  # recorre cada tipo de tornillo ya agrupado
                cantidad_tornillos = int(info["tornillos"])  # tornillos con tarugo de este tipo
                cantidad_sin_tarugo = int(info["tornillos_sin_tarugo"])  # tornillos sin tarugo de este tipo
                if cantidad_tornillos <= 0 and cantidad_sin_tarugo <= 0:  # si no hay ninguno de este tipo, no genera fila
                    continue  # no salió ningún tornillo de este tipo, no se genera fila
                mats_txt = " + ".join(sorted(info["mats"]))  # texto con los materiales que usan este tornillo, ej 'madera + metalcon'
                if es_sob2 and desc_tipo == 'Tornillo volcanita punta fina 1 1/4"':  # caso especial: volcanita en sobrepuesta necesita tarugo además del tornillo
                    if cantidad_tornillos > 0:  # solo si hay tornillos con tarugo de este material
                        add_row(  # fila del tarugo para volcanita en sobrepuesta
                            desc="Tarugo paloma 6mm",  # el tarugo que agarra en la volcanita
                            marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas sugeridas de tarugo paloma
                            norma="-",  # sin norma asociada
                            circuito=f"General ({col_mat}: {mats_txt})",  # deja anotado que material obligo el tarugo
                            unidad="u",  # se cuentan por unidad
                            k=1,  # sin multiplicador
                            longitud_m=f"{cantidad_tornillos} unid",  # detalle para el Excel
                            cantidad=cantidad_tornillos  # un tarugo por cada tornillo
                        )
                    # tornillo para tarugo (parte elegible) + tornillo simple
                    # de la porción sin tarugo (tapas ciegas + octogonales
                    # sobrepuesta), mismo tipo, se suman en la misma fila
                    cantidad_tornillo_total = cantidad_tornillos + cantidad_sin_tarugo  # tornillo para el tarugo + el que no necesita tarugo, mismo tipo de tornillo
                    if cantidad_tornillo_total > 0:  # solo genera la fila si el total es mayor a 0
                        add_row(  # fila del tornillo que va dentro del tarugo
                            desc='Tornillo volcanita punta fina 6x1 1/4"',  # el tornillo que entra en el tarugo paloma
                            marcas_txt=marcas.get("Tornillo para tarugo paloma", ""),  # marcas del tornillo para tarugo
                            norma="-",  # sin norma asociada
                            circuito=f"General ({col_mat}: {mats_txt})",  # mismo detalle de materiales
                            unidad="u",  # se cuentan por unidad
                            k=1,  # sin multiplicador
                            longitud_m=f"{cantidad_tornillo_total} unid",  # detalle para el Excel
                            cantidad=cantidad_tornillo_total  # los con tarugo más los que no lo necesitan
                        )
                    continue  # ya se generaron las filas de este caso especial: sigue con el próximo tipo
                cantidad_total_tipo = cantidad_tornillos + cantidad_sin_tarugo  # casos normales: junta ambos conteos en una sola fila
                # agrega la fila de este tipo de tornillo al listado de materiales
                add_row(
                    desc=desc_tipo,  # el tornillo que corresponde a ese material
                    marcas_txt=marcas.get("Tornillos", ""),  # marcas sugeridas de tornillería
                    norma="-",  # sin norma asociada
                    circuito=f"General ({col_mat}: {mats_txt})",  # deja anotado el material de este grupo
                    unidad="u",  # se cuentan por unidad
                    k=1,  # sin multiplicador
                    longitud_m=f"{cantidad_total_tipo} unid",  # detalle para el Excel
                    cantidad=cantidad_total_tipo  # total de tornillos de este tipo
                )

    # NOTA: los tarugos paloma 6mm para volcanita/fibrocemento (canalización
    # sobrepuesta) ya se generan arriba, dentro de la sección "Tornillería".


    # =========================
    # SELLOS DEL PANEL SIP (tablero embutido en panel SIP)
    # Esto es distinto a los tornillos: acá se agrega espuma de poliuretano
    # para sellar el hueco que queda alrededor de las cajas de derivación
    # embutidas en el núcleo del panel SIP.
    # SOLO para Panel SIP + canalización embutida
    # Criterio
    #  - SOLO por cajas de derivación
    # Regla:
    #  - 1 tubo cada 10 cajas de derivación
    #  - mínimo 1 tubo si aplica a material SIP
    # El tablero se fija con tornillos, no con espuma.
    # =========================
    tipo_can_esp = (tipo_canalizacion or "").strip().lower()  # tipo de canalización, en minúscula
    es_emb_esp = "embut" in tipo_can_esp  # True si la canalización es embutida
    if es_emb_esp:  # la espuma solo se usa en canalización embutida
        # detectar Panel SIP en "Material tabique"
        mats_tab = []  # acá se juntan los materiales de tabique de cada ambiente
        if ambientes_df is not None and "Material tabique" in ambientes_df.columns:  # solo si el usuario lleno esa columna
            # lista con el material de tabique de cada ambiente, en minúscula
            mats_tab = (
                ambientes_df["Material tabique"].fillna("").astype(str).str.strip().str.lower().tolist())  # queda una lista con el material de tabique de cada ambiente
        hay_sip = any("sip" in m for m in mats_tab)  # True si algún tabique es Panel SIP
        if hay_sip:  # hay al menos un tabique de Panel SIP
            # (opcional) sección nueva para que quede ordenado en el Excel
            add_section("Sellos / Aislación (Panel SIP)")  # sección aparte en el Excel para que no se mezcle
            # 1) SOLO cajas de derivación
            n_cajas = int(_cajas_total) if _cajas_total is not None else 0  # cajas de derivación embutidas del proyecto
            # 3) cálculo de tubos — la espuma es solo para sellar/rellenar el
            # hueco de las cajas embutidas en el núcleo del panel SIP, el
            # tablero no la necesita.
            tubos = int(math.ceil(n_cajas / 10.0)) if n_cajas > 0 else 0  # 1 tubo cada 10 cajas, redondeando hacia arriba
            tubos = max(1, tubos)  # mínimo 1 si aplica SIP
            # 4) agregar al listado
            add_row(  # fila de la espuma en el listado
                desc="Espuma expansiva de poliuretano 750 ml",  # espuma para rellenar el hueco de la caja en el panel
                marcas_txt=marcas.get("Espuma expansiva PU", ""),  # marcas sugeridas de espuma PU
                norma="Sellado de cajas de derivación embutidas en Panel SIP",  # no es artículo del RIC, es el criterio de uso
                circuito="Alumbrado (iluminación y enchufes)",  # se carga al circuito de alumbrado
                unidad="u",  # se compran por tubo
                k=1,  # sin multiplicador
                longitud_m=f"{tubos} tubo(s) 750 ml",  # detalle para el Excel
                cantidad=tubos  # los tubos calculados arriba
            )

# =========================
# CLIMATIZACIÓN (RIC N°07 SEC. 7)
# Materiales exclusivos por equipo:
#  - Enchufe 2P+T 16A (si el equipo llega con enchufe)
#  - Caja de derivación cerca del equipo
#  - Canalización exclusiva (conduit o canaleta)
#  - Abrazaderas / accesorios de canalización
#  - Salidas de caja conduit (si embutida)
#  - Conector cónico 2.5 mm^2 (unión en caja)
#  - Boquilla bordes redondeados
#  - Prensaestopa (si conexión directa sin enchufe)
# Materiales que se actualizan automáticamente
# por ser circuitos más en circuitos_df:
#  - Conductor (3×L ya generado en sección Conductores)
#  - TM y diferencial (ya generados en sección Protecciones)
#  - Ferrules (ya generados en sección Ferrules)
#  - Tablero, riel, barra verde, barra repartidora
#    (n_circ ya incluye circuitos de climatización)
# =========================
# EMPALME
# Acá empieza la parte del listado de materiales del empalme: lo que
# conecta la red pública con la instalación (medidor, caja, cable de
# acometida, tubo de protección, terminales, etc.).
# =========================
    add_section("Empalme")  # nueva sección del Excel: materiales del empalme
    # Empalme en fachada con o sin mástil
    # Se activa solo cuando: la instalación del empalme es "fachada", el
    # alimentador va en ducto, y la acometida es aérea o subterránea.
    if (str(tipo_instalacion_empalme).strip().lower() == "fachada"  # el empalme va montado en la fachada
        and "duct" in str(tipo_alimentador).strip().lower()  # el alimentador baja por ducto
        and ("aer" in str(tipo_acometida).strip().lower() or "sub" in str(tipo_acometida).strip().lower())  # la acometida puede ser aérea o subterránea
        ):
        # Unidad de medida monofásica (medidor)
        add_row(  # fila del medidor
            desc="Unidad de medida monofásica 220V 50Hz 50A (medidor)",  # medidor monofásico del empalme
            marcas_txt=marcas.get("Medidor empalme", ""),  # marcas sugeridas de medidor
            norma="SEC",  # marca Sello SEC; en la columna Norma termina saliendo "RIC 1" (empalme)
            circuito="Empalme",  # va en la sección empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin multiplicador
            longitud_m="1 unid",  # detalle para el Excel
            cantidad=1  # uno solo por instalación
        )
        #Caja metálica del empalme
        add_row(  # fila de la caja del empalme
            desc="Caja de empalme metálica para medidor monofásico 405x200x137 mm IP54",  # caja metálica IP54 donde se monta el medidor
            marcas_txt=marcas.get("Caja empalme", ""),  # marcas sugeridas de caja de empalme
            norma="SEC",  # lo define la SEC
            circuito="Empalme",  # va en la sección empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin multiplicador
            longitud_m="1 unid",  # detalle para el Excel
            cantidad=1  # una sola caja
        )
        #Disyuntor termomagnético del empalme
        add_row(  # fila del TM del empalme
            desc=f"Disyuntor termomagnético {interruptor_texto}",  # el texto del TM viene del cálculo de la corriente del empalme
            marcas_txt=marcas.get("Disyuntor empalme", ""),  # marcas sugeridas de disyuntor de empalme
            norma="SEC",  # lo define la SEC
            circuito="Empalme",  # va en la sección empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin multiplicador
            longitud_m="1 unid",  # detalle para el Excel
            cantidad=1  # uno solo
        )
        # Cable de la acometida
        add_row(  # fila del cable de acometida
            desc=f"Cable {acometida_txt}",  # el texto de la acometida ya trae la sección calculada
            marcas_txt=marcas.get("Cable acometida", ""),  # marcas sugeridas de cable de acometida
            norma="SEC",  # lo define la SEC
            circuito="Acometida",  # se carga a la acometida
            unidad="m",  # el cable se mide en metros
            k=1,  # sin multiplicador
            longitud_m=f"{math.ceil(float(longitud_transformador_empalme))} m",  # largo desde el transformador hasta el empalme
            cantidad=math.ceil(float(longitud_transformador_empalme))  # mismos metros, redondeados hacia arriba
        )
        # Tubo conduit galvanizado según sección de acometida
        sec_acom = float(acometida_txt.split("2x")[1].replace("mm^2", "").strip())  # extrae el número de mm^2 del texto
        # El diámetro del tubo galvanizado depende de la sección del cable de
        # la acometida: mientras más gruesa la sección, más grande el tubo.
        if sec_acom <= 6:  # acometida de hasta 6 mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 25mm, 3mts"  # le basta tubo de 25mm
        elif sec_acom <= 16:  # acometida de hasta 16 mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 32mm, 3mts"  # necesita tubo de 32mm
        elif sec_acom <= 35:  # acometida de hasta 35 mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 40mm, 3mts"  # necesita tubo de 40mm
        else:  # sección fuera de la tabla
            desc_tubo_acom = ""  # sección fuera de rango, no hay tubo definido
        if desc_tubo_acom:  # solo sigue si quedó un tubo definido
            # SIN MÁSTIL
            if str(requiere_mastil).strip().lower() == "no":  # el empalme no lleva mástil
                # Si la acometida es subterránea,
                # calcular cantidad según longitud hasta el medidor
                if "sub" in str(tipo_acometida).strip().lower():  # acometida subterránea: el tubo va enterrado hasta el medidor
                    largo_tubo_acom = float(longitud_subterraneo_medidor)  # largo del tramo subterráneo
                    cantidad_tubo_acom = int(math.ceil(largo_tubo_acom / 3))  # los tubos vienen en tramos de 3 m
                else:  # acometida aérea sin mástil
                    largo_tubo_acom = float(dist_vertical_acometida)  # el tubo sube por la fachada
                    cantidad_tubo_acom = int(math.ceil(largo_tubo_acom / 3))  # los tubos vienen en tramos de 3 m
            # CON MÁSTIL
            elif str(requiere_mastil).strip().lower() == "si":  # el empalme lleva mástil
                largo_tubo_acom = float(longitud_mastil)  # el tubo cubre todo el largo del mástil
                cantidad_tubo_acom = int(math.ceil(largo_tubo_acom / 3))  # los tubos vienen en tramos de 3 m
            # Puestas a tierra
            # el tubo galvanizado también cubre el tramo hasta las 2 puestas a tierra
            largo_total_pt = float(dist_empalme_pt1) + float(dist_tda_pt2)  # tramo del empalme a la PT1 más el del TDA a la PT2
            cantidad_tubo_pt = int(math.ceil(largo_total_pt / 3))  # tramos de 3 m
            # Totales
            largo_total_tubos = largo_tubo_acom + largo_total_pt  # metros totales entre acometida y puesta a tierra
            cantidad_total_tubos = cantidad_tubo_acom + cantidad_tubo_pt  # tubos de la acometida más los de las puestas a tierra
            add_row(  # fila del tubo galvanizado
                desc=desc_tubo_acom,  # el tubo que se eligió según la sección
                marcas_txt=marcas.get("Tubo conduit galvanizado acometida", ""),  # marcas sugeridas de tubo galvanizado
                norma="SEC",  # lo define la SEC
                circuito="Acometida / Puesta a tierra",  # el mismo tubo cubre los dos tramos
                unidad="u",  # se compran por tira
                k=1,  # sin multiplicador
                longitud_m=f"{round(largo_total_tubos, 2)} m",  # metros totales, como referencia
                cantidad=cantidad_total_tubos  # tiras de 3m que hay que comprar
            )

        # Cabeza de servicio según tubo galvanizado (solo acometida aérea)
        # La cabeza de servicio va en la punta del tubo, para que no entre
        # agua; su tamaño depende del diámetro del tubo galvanizado.
        if "aer" in str(tipo_acometida).strip().lower():  # la cabeza de servicio solo va si la acometida es aérea
            if "25mm" in desc_tubo_acom:  # tubo de 25mm
                desc_cabeza = 'Cabeza de servicio 3/4"'  # le corresponde cabeza de 3/4
            elif "32mm" in desc_tubo_acom:  # tubo de 32mm
                desc_cabeza = 'Cabeza de servicio 1"'  # le corresponde cabeza de 1
            elif "40mm" in desc_tubo_acom:  # tubo de 40mm
                desc_cabeza = 'Cabeza de servicio 1 1/4"'  # le corresponde cabeza de 1 1/4
            elif "50mm" in desc_tubo_acom:  # tubo de 50mm
                desc_cabeza = 'Cabeza de servicio 2"'  # le corresponde cabeza de 2
            else:  # tubo que no está en la tabla
                desc_cabeza = ""  # diámetro fuera de tabla: no va cabeza
            if desc_cabeza:  # solo si quedó una cabeza definida
                add_row(  # fila de la cabeza de servicio
                    desc=desc_cabeza,  # la cabeza que calzo con el diámetro del tubo
                    marcas_txt=marcas.get("Cabeza de servicio", ""),  # marcas sugeridas de cabeza de servicio
                    norma="SEC",  # lo define la SEC
                    circuito="Acometida",  # se carga a la acometida
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin multiplicador
                    longitud_m="1 unid",  # detalle para el Excel
                    cantidad=1  # una sola
                )
        # Cáncamo de acero galvanizado
        # Sirve para sujetar el cable de la acometida aérea al poste/muro.
        if "aer" in str(tipo_acometida).strip().lower():  # el cáncamo solo aplica en acometida aérea
            add_row(  # fila del cáncamo
                desc="Cáncamo abierto de acero galvanizado 7,8 x 110 mm",  # de acá se amarra el cable de la acometida
                marcas_txt=marcas.get("Cancamo abierto", ""),  # marcas sugeridas de cáncamo
                norma="-",  # sin norma asociada
                circuito="Acometida",  # se carga a la acometida
                unidad="u",  # se cuenta por unidad
                k=1,  # sin multiplicador
                longitud_m="1 unid",  # detalle para el Excel
                cantidad=1  # uno solo
            )
        #Granpa de retención — según sección de la acometida
        # La granpa (grapa) de retención sujeta el cable de la acometida
        # aérea sin cortarlo; el modelo depende del grosor del cable.
        if "aer" in str(tipo_acometida).strip().lower():  # solo aplica si la acometida es aérea
            if sec_acom <= 10:  # cable delgado, hasta 10mm^2 de sección
                desc_granpa = "Granpa de retención tipo cuña 2,5-25mm^2 (6-10mm)"  # granpa chica, para cable de 6 a 10mm de diámetro
            else:  # cable más grueso
                desc_granpa = "Granpa de retención tipo cuña 16-25mm^2 (10-15mm)"  # granpa grande, para cable de 10 a 15mm
            add_row(  # ítem: granpa de retención de la acometida aérea
                desc=desc_granpa,  # nombre tal cual sale en el listado de materiales
                marcas_txt=marcas.get("Granpa de retención", ""),  # marcas del catálogo para este ítem
                norma="-",  # no la exige un artículo, va por instalación
                circuito="Acometida",  # pertenece al tramo de acometida
                unidad="u",  # se compra por unidad
                k=1,  # sin holgura extra
                longitud_m="1 unid",  # texto que se muestra en la columna de largo
                cantidad=1  # 1 granpa por acometida
            )
        # Conector HUB según tubo galvanizado
        # El conector HUB une el tubo galvanizado con la caja metálica;
        # el tamaño sigue el mismo diámetro que el tubo calculado antes.
        if "25mm" in desc_tubo_acom:  # el hub sigue el mismo diámetro del tubo galvanizado ya elegido
            desc_hub = "Conector HUB 25mm de acero galvanizado"  # hub de 25mm
        elif "32mm" in desc_tubo_acom:  # tubo de 32mm
            desc_hub = "Conector HUB 32mm de acero galvanizado"  # hub de 32mm
        elif "40mm" in desc_tubo_acom:  # tubo de 40mm
            desc_hub = "Conector HUB 40mm de acero galvanizado"  # hub de 40mm
        elif "50mm" in desc_tubo_acom:  # tubo de 50mm
            desc_hub = "Conector HUB 50mm de acero galvanizado"  # hub de 50mm
        else:  # no se reconoció el diámetro del tubo
            desc_hub = ""  # queda vacío y más abajo no se agrega la fila
        # cantidad cambia si acometida es subterránea
        if "sub" in str(tipo_acometida).strip().lower():  # acometida subterránea
            cantidad_hub = 6  # una entrada y salida más, por la caja de derivación extra
        else:  # acometida aérea
            cantidad_hub = 5  # 5 hubs: acometida más las dos puestas a tierra
        if desc_hub:  # solo agrega la fila si se pudo determinar el diámetro
            add_row(  # ítem: conectores HUB de las cajas metálicas
                desc=desc_hub,  # descripción armada arriba según el diámetro
                marcas_txt=marcas.get("Conector HUB", ""),  # marcas del catálogo
                norma="SEC",  # exigencia SEC
                circuito="Acometida, puesta a tierra 1, puesta a tierra 2",  # se usan en acometida y en las dos puestas a tierra
                unidad="u",  # por unidad
                k=1,  # sin factor de pérdida
                longitud_m=f"{cantidad_hub} unid",  # texto de la columna de largo
                cantidad=cantidad_hub  # 5 o 6 según si la acometida es subterránea
            )

        # Terminal PVC conduit con 2 tuercas — igual criterio que en poste:
        # se separa acometida (solo si subterránea, Tabla N°4.29 con 2
        # conductores) y alimentador (siempre en ducto en fachada, Tabla
        # N°4.29 con 3 conductores). Si coinciden en diámetro, 1 fila;
        # si no, 2 filas separadas.
        _es_acom_sub_term_fach = "sub" in str(tipo_acometida).strip().lower()  # True si la acometida es subterránea
        qty_term_acom_fach = 1 if _es_acom_sub_term_fach else 0  # terminal de la acometida (solo si es subterránea)
        qty_term_alim_fach = 4  # el alimentador en fachada siempre va en ducto

        # diámetro del ducto para la acometida (tabla de subterráneo, 2 conductores)
        diam_term_acom_fach = ducto_nominal_tablas(sec_acom, 2, "subterraneo") if qty_term_acom_fach else None

        import re  # se usa para sacar el número de mm desde los textos
        # función chica para no repetir el mismo re.search en cada tramo
        def _extraer_diam_num_fach(txt):
            # saca el número de milímetros de un texto tipo "conduit 25mm"
            m = re.search(r'(\d+)\s*mm', str(txt))  # busca algo tipo "25mm" dentro del texto
            return int(m.group(1)) if m else None  # si el texto no trae mm, devuelve None
        diam_term_alim_fach = _extraer_diam_num_fach(canalizacion_txt)  # diámetro del ducto del alimentador

        if diam_term_acom_fach and diam_term_alim_fach and diam_term_acom_fach == diam_term_alim_fach:  # los dos existen y además coinciden en diámetro
            # mismo diámetro para acometida y alimentador: se suman en una sola fila
            add_row(  # ítem: terminales conduit de acometida y alimentador en una sola fila
                desc=f"Terminal PVC conduit con 2 tuercas {diam_term_acom_fach} mm",  # el diámetro es el mismo para los dos tramos
                marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas del catálogo
                norma="RIC 4.7.2",  # RIC N°4, artículo 7.2
                circuito="Acometida y alimentador en caja empalme, entrada y salida de caja metálica, entrada al TDA",  # todos los puntos donde entra o sale el conduit
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{qty_term_acom_fach + qty_term_alim_fach} unid",  # suma de los dos tramos, en texto
                cantidad=qty_term_acom_fach + qty_term_alim_fach  # acometida más alimentador
            )
        else:  # diámetros distintos, o falta el dato de alguno
            # diámetros distintos (o no aplica alguno): se agregan filas separadas
            if diam_term_acom_fach and qty_term_acom_fach:  # solo si la acometida es subterránea y tiene diámetro
                add_row(  # ítem: terminal del tramo PVC de la acometida
                    desc=f"Terminal PVC conduit con 2 tuercas {diam_term_acom_fach} mm",  # diámetro que salió de la tabla de subterráneo
                    marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas del catálogo
                    norma="RIC 4.7.2",  # RIC N°4, artículo 7.2
                    circuito="Acometida (tramo PVC subterráneo)",  # esta fila es solo del tramo enterrado
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{qty_term_acom_fach} unid",  # texto de la columna de largo
                    cantidad=qty_term_acom_fach  # 1 si la acometida es subterránea, 0 si no
                )
            if diam_term_alim_fach and qty_term_alim_fach:  # solo si el alimentador trajo diámetro
                add_row(  # ítem: terminal conduit del alimentador en fachada
                    desc=f"Terminal PVC conduit con 2 tuercas {diam_term_alim_fach} mm",  # diámetro sacado del texto de canalización
                    marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas del catálogo
                    norma="RIC 4.7.2",  # RIC N°4, artículo 7.2
                    circuito="Alimentador en caja empalme, entrada y salida de caja metálica, entrada al TDA",  # entrada y salida de la caja metálica, y entrada al TDA
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{qty_term_alim_fach} unid",  # texto de la columna de largo
                    cantidad=qty_term_alim_fach  # 4 terminales, fijo para el alimentador en fachada (siempre va en ducto)
                )

        # =========================
        # PUESTA A TIERRA
        # =========================
        #Conductor puesta a tierra
        sec_pt = "4" # valor por defecto: 4mm^2 es la sección mínima de puesta a tierra
        # intenta calcular la sección real de PT a partir del texto del alimentador (ej. "3x4mm^2");
        # si algo falla (formato inesperado), se usa el mínimo definido arriba (sec_pt = "4")
        try:  # si el texto no viene con el formato esperado, cae al except
            sec_base = float(alim_txt.split("3x")[1].replace("mm^2", "").strip())  # sección del alimentador
            # la sección de PT sigue a la del alimentador, redondeada a la comercial más cercana
            if sec_base <= 4:  # hasta 4mm^2 el PT también queda en 4
                sec_pt = "4"  # sección comercial más cercana
            elif sec_base <= 6:  # hasta 6mm^2
                sec_pt = "6"  # PT de 6mm^2
            elif sec_base <= 10:  # hasta 10mm^2
                sec_pt = "10"  # PT de 10mm^2
            elif sec_base <= 16:  # hasta 16mm^2
                sec_pt = "16"  # PT de 16mm^2
            elif sec_base <= 25:  # hasta 25mm^2
                sec_pt = "25"  # PT de 25mm^2
            else:  # alimentador más grueso que 25mm^2
                sec_pt = "25"  # se tope en 25mm^2, es el máximo de la tabla usada acá
        except:  # el texto del alimentador vino raro
            sec_pt = "4"  # no se pudo leer el alimentador, usa el mínimo
        # +2m: 1 chicote en la camarilla N°1 + 1 chicote en la caja metálica del empalme
        metros_pt = math.ceil(float(dist_empalme_pt1)) + 2  # total de metros de conductor PT1 (ya incluye los 2m extra de chicotes)
        # ítem: conductor blanco (neutro) de la puesta a tierra N°1
        add_row(  # agrega la fila del conductor blanco de PT1
            desc=f"Conductor THWN-2 {sec_pt}mm^2 blanco",  # la sección sale del cálculo de arriba
            marcas_txt=marcas.get("Conductor THWN-2 blanco", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # tramo entre la caja del empalme y la camarilla N°1
            unidad="m",  # se vende por metro
            k=1,  # sin holgura, los chicotes ya están sumados
            longitud_m=f"{metros_pt} m",  # metros que se muestran en el listado
            cantidad=metros_pt  # metros de conductor blanco
        )
        # ítem: conductor verde (tierra), mismo largo que el blanco de arriba
        add_row(  # agrega la fila del conductor verde
            desc=f"Conductor THWN-2 {sec_pt}mm^2 verde",  # misma sección que el blanco
            marcas_txt=marcas.get("Conductor THWN-2 verde", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # mismo tramo que el blanco
            unidad="m",  # por metro
            k=1,  # sin holgura
            longitud_m=f"{metros_pt} m",  # mismos metros que el blanco
            cantidad=metros_pt  # metros de conductor verde
        )
        # Barra copperweld PT1 (empalme - camarilla N°1)
        add_row(  # ítem: barras copperweld de la puesta a tierra N°1
            desc=desc_barra_pt,  # descripción de la barra, ya calculada antes
            marcas_txt=marcas.get("Barra copperweld", ""),  # marcas del catálogo
            norma="RIC 6 (8.3.2, 8.5, 8.6, Tabla 6.1)",  # artículos de puesta a tierra del RIC N°6
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # tierra de servicio, la del empalme
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{n_barras_pt1} unid",  # cuántas barras van, en texto
            cantidad=n_barras_pt1  # número de barras que pide la resistividad del terreno
        )
        # Camarilla PT1 según tubo conduit galvanizado
        diam_cam = ""  # se va llenando abajo, según el diámetro del tubo de acometida
        if "25mm" in desc_tubo_acom:  # la camarilla se elige con el mismo diámetro del tubo de acometida
            diam_cam = "25mm"  # camarilla de 25mm
        elif "32mm" in desc_tubo_acom:  # tubo de 32mm
            diam_cam = "32mm"  # camarilla de 32mm
        elif "40mm" in desc_tubo_acom:  # tubo de 40mm
            diam_cam = "40mm"  # camarilla de 40mm
        elif "50mm" in desc_tubo_acom:  # tubo de 50mm
            diam_cam = "50mm"  # camarilla de 50mm
        # RIC N°6 art. 8.3.2: si hay más de 1 barra, hay que unirlas con
        # un conductor desnudo de cobre, mínimo 16mm^2 de sección (fijo,
        # ese mínimo no depende de la sección del alimentador)
        if n_barras_pt1 > 1:  # con una sola barra no hay nada que unir
            add_row(  # ítem: conductor desnudo entre barras de PT1
                desc=f"Conductor desnudo Cu 16mm^2 (unión entre barras PT1)",  # sección fija de 16mm^2, no depende del alimentador
                marcas_txt=marcas.get("Conductor desnudo Cu", ""),  # marcas del catálogo
                norma="RIC 6 (8.3.2, 8.7, 8.9)",  # artículos de puesta a tierra del RIC N°6
                circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # queda dentro de la tierra de servicio
                unidad="m",  # por metro
                k=1,  # sin holgura
                longitud_m=f"{long_cond_desnudo_pt1:.1f} m",  # largo con un decimal
                cantidad=long_cond_desnudo_pt1  # metros de conductor desnudo entre las barras de PT1
            )
        # Barra copperweld PT2 (TDA - camarilla N°2)
        add_row(  # ítem: barras copperweld de la puesta a tierra N°2
            desc=desc_barra_pt,  # misma descripción de barra que en PT1
            marcas_txt=marcas.get("Barra copperweld", ""),  # marcas del catálogo
            norma="RIC 6 (8.3.2, 8.5, 8.6, Tabla 6.1)",  # artículos de puesta a tierra del RIC N°6
            circuito="Puesta a tierra N°2 (TDA - camarilla N°2)",  # tierra de protección, la del tablero
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{n_barras_pt2} unid",  # cuántas barras van, en texto
            cantidad=n_barras_pt2  # número de barras de la camarilla N°2
        )
        # Camarilla PT1 + PT2: mismo diámetro siempre (diam_cam es único,
        # compartido entre ambas), así que se fusionan en 1 sola fila en
        # vez de 2 filas idénticas repetidas.
        if diam_cam:  # solo si arriba se logró definir el diámetro
            add_row(  # ítem: camarillas de registro de las dos puestas a tierra
                desc=f"Camarilla de registro con tapa PVC naranjo 160 x {diam_cam}",  # la camarilla es de 160 y la entrada va según el tubo
                marcas_txt=marcas.get("Camarilla PVC naranjo", ""),  # marcas del catálogo
                norma="RIC 6 (5.15)",  # RIC N°6, artículo 5.15
                circuito="Puesta a tierra N°1 y N°2 (empalme - camarilla N°1 / TDA - camarilla N°2)",  # cubre las dos tierras en una sola fila
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{n_barras_pt1 + n_barras_pt2} unid",  # total de camarillas, en texto
                cantidad=n_barras_pt1 + n_barras_pt2  # una camarilla por cada barra de PT1 y PT2
            )
        # RIC N°6 art. 8.3.2: mismo criterio que PT1, mínimo 16mm^2 fijo
        # si hay más de 1 barra (esta es la puesta a tierra de PROTECCIÓN,
        # la del tablero, distinta de la de SERVICIO que es PT1)
        if n_barras_pt2 > 1:  # igual que en PT1, con una sola barra no se une nada
            add_row(  # ítem: conductor desnudo entre barras de PT2
                desc=f"Conductor desnudo Cu 16mm^2 (unión entre barras PT2)",  # también 16mm^2 fijo
                marcas_txt=marcas.get("Conductor desnudo Cu", ""),  # marcas del catálogo
                norma="RIC 6 (8.3.2, 8.7, 8.9)",  # artículos de puesta a tierra del RIC N°6
                circuito="Puesta a tierra N°2 (TDA - camarilla N°2)",  # queda dentro de la tierra de protección
                unidad="m",  # por metro
                k=1,  # sin holgura
                longitud_m=f"{long_cond_desnudo_pt2:.1f} m",  # largo con un decimal
                cantidad=long_cond_desnudo_pt2  # metros de conductor desnudo entre las barras de PT2
            )
        # ---- abrazaderas tipo caddy (empalme en fachada) ----
        # criterio de conteo: 3 abrazaderas por cada tubo galvanizado de 3m.
        # cantidad_total_tubos viene del bloque del tubo galvanizado de más
        # arriba (línea ~6253 del archivo completo) y ya suma los tubos de la
        # acometida MÁS los de las dos puestas a tierra (PT1 y PT2). el
        # alimentador no entra acá: en fachada siempre va en conduit PVC y
        # lleva sus propias abrazaderas de PVC (bloque de la Tabla N°4.24).
        # ojo: cantidad_total_tubos solo se define si desc_tubo_acom quedó con
        # diámetro; si la acometida pasa de 35mm^2 ese texto queda vacío y esta
        # línea se cae con NameError.
        # función interna: arma el nombre de la abrazadera caddy según el diámetro de texto que reciba
        # recibe el texto del tubo como parámetro; en este bloque siempre se
        # le pasa el de la acometida (desc_tubo_acom), que en fachada es el
        # único tubo galvanizado que existe
        def _desc_caddy_de_fach(diam_txt):
            if "25mm" in diam_txt:  # el texto trae 25mm
                return "Abrazadera tipo caddy 25mm"  # caddy de 25mm
            elif "32mm" in diam_txt:  # trae 32mm
                return "Abrazadera tipo caddy 32mm"  # caddy de 32mm
            elif "40mm" in diam_txt:  # trae 40mm
                return "Abrazadera tipo caddy 40mm"  # caddy de 40mm
            elif "50mm" in diam_txt:  # trae 50mm
                return "Abrazadera tipo caddy 50mm"  # caddy de 50mm
            return "Abrazadera tipo caddy"  # si el texto no calzó con ningún diámetro, va sin medida
        # 3 abrazaderas por cada tubo de 3m: criterio de montaje del autor,
        # no sale de una tabla del RIC. cantidad_caddy se reusa más abajo en
        # el grupo B de tornillos (1 tornillo por abrazadera).
        cantidad_caddy = int(cantidad_total_tubos * 3)  # 3 abrazaderas por cada tubo galvanizado
        add_row(  # ítem: abrazaderas caddy que fijan los tubos a la fachada
            desc=_desc_caddy_de_fach(desc_tubo_acom),  # el nombre se arma con el diámetro del tubo de acometida
            marcas_txt=marcas.get("Abrazadera tipo caddy", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra 1 y 2, acometida",  # sujetan los tubos de las tierras y de la acometida
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{cantidad_caddy} unid",  # total de abrazaderas, en texto
            cantidad=cantidad_caddy  # 3 por cada tubo
        )
        #Sellador de roscas: 1 frasco de 50ml alcanza para 6 conectores HUB
        # en fachada cantidad_hub vale 5 (acometida aérea) o 6 (acometida
        # subterránea), así que este ceil siempre termina dando 1 frasco.
        cantidad_sellador = int(math.ceil(cantidad_hub / 6))  # frascos necesarios, redondeando hacia arriba
        add_row(  # ítem: sellador de roscas para los hub
            desc="Sellador de roscas con teflón 50ml",  # frasco de 50ml
            marcas_txt=marcas.get("Sellador de roscas", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Conectores Hub(cajas metálicas empalme y derivación)",  # se usa en las roscas de todos los hub
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{cantidad_sellador} unid",  # cuántos frascos, en texto
            cantidad=cantidad_sellador  # 1 frasco cada 6 hub
        )
        # Terminal ferrul acometida: mismo color/sección que el tablero (color_ferrul_por_seccion)
        # intenta leer el número de sección desde el texto de la acometida (ej. "2x4mm^2")
        try:  # el texto de acometida viene tipo "2x4mm^2", fase y neutro
            sec_acom = float(acometida_txt.split("2x")[1].replace("mm^2", "").strip())  # sección de la acometida en mm^2
        except:  # el texto vino con otro formato
            sec_acom = 4.0  # no se pudo leer, usa 4mm^2 por defecto
        color_ferrul = color_ferrul_por_seccion(sec_acom)  # color de ferrul según tabla, para que combine con el del tablero
        sec_ferrul = str(int(sec_acom)) if sec_acom == int(sec_acom) else texto_seccion(sec_acom)  # texto de la sección: sin decimales si es un número entero
        add_row(  # ítem: ferrules de la acometida
            desc=f"Terminal ferrul color {color_ferrul} {sec_ferrul}mm (acometida)",  # el color depende de la sección, así no se confunden en obra
            marcas_txt=marcas.get("Terminal ferrul acometida", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Acometida (Fase y Neutro)",  # van en las puntas de fase y neutro
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="2 unid",  # 2 unidades, en texto
            cantidad=2  # uno para la fase y uno para el neutro
        )
        # Terminal ferrul alimentador: mismo color/sección que el tablero
        # mismo cálculo que arriba, pero para el alimentador
        try:  # el alimentador viene tipo "3x6mm^2", son 3 conductores
            sec_alim = float(alim_txt.split("3x")[1].replace("mm^2", "").strip())  # sección del alimentador en mm^2
        except:  # el texto vino con otro formato
            sec_alim = 4.0  # no se pudo leer, usa 4mm^2 por defecto
        color_ferrul_alim = color_ferrul_por_seccion(sec_alim)  # color de ferrul del alimentador
        sec_ferrul_alim = str(int(sec_alim)) if sec_alim == int(sec_alim) else texto_seccion(sec_alim)  # texto de la sección del alimentador
        add_row(  # ítem: ferrules del alimentador
            desc=f"Terminal ferrul color {color_ferrul_alim} {sec_ferrul_alim}mm (alimentador)",  # color y sección del alimentador
            marcas_txt=marcas.get("Terminal ferrul alimentador", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Alimentador (Salida de medidor fase, entrada y salida termomagnético empalme)",  # salida del medidor y ambos lados del TM del empalme
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="3 unid",  # 3 unidades, en texto
            cantidad=3  # tres puntas a terminar
        )
        #terminal ferrul doble alimentador
        add_row(  # ítem: ferrul doble del alimentador
            desc=f"Terminal ferrul doble color {color_ferrul_alim} {sec_ferrul_alim}mm (alimentador)",  # el doble entra dos conductores en un mismo terminal
            marcas_txt=marcas.get("Terminal ferrul doble alimentador", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Salida de medidor (Neutro del alimentador y neutro aterrizado)",  # ahí se juntan el neutro del alimentador y el neutro aterrizado
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="1 unid",  # 1 unidad, en texto
            cantidad=1  # un solo ferrul doble
        )
        # Cantidad de cajas de derivación metálica del empalme en fachada.
        # criterio: 1 caja siempre (la del tramo de fachada) y 1 más si la
        # acometida es subterránea, porque el cambio de conduit PVC enterrado
        # a tubo galvanizado necesita su propia caja.
        # cantidad_caja_derivacion se vuelve a usar más abajo en: el terminal
        # de compresión tipo ojo, la fila de la caja de derivación, el tornillo
        # de tierra de las cajas metálicas y el grupo B de tornillos (4 c/u).
        if "sub" in str(tipo_acometida).strip().lower():  # acometida subterránea
            cantidad_caja_derivacion = 2  # subterránea necesita una caja extra
        else:  # acometida aérea
            cantidad_caja_derivacion = 1  # solo la caja de derivación del tramo de fachada
        #Terminal de compresión tipo ojo: 1 por caja de empalme + 1 por cada caja de derivación
        cantidad_terminal_ojo = 1 + cantidad_caja_derivacion  # 1 por la caja del empalme más 1 por cada caja de derivación
        add_row(  # ítem: terminales de compresión tipo ojo para aterrizar las cajas
            desc=f"Terminal de compresión tipo ojo {sec_pt}mm",  # del mismo grosor que el conductor de puesta a tierra
            marcas_txt=marcas.get("Terminal compresion tipo ojo", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Tramos metálicos (caja metálica empalme, caja de derivación, acometida, puesta a tierra 1 y 2)",  # todas las partes metálicas que hay que aterrizar
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{cantidad_terminal_ojo} unid",  # total de terminales, en texto
            cantidad=cantidad_terminal_ojo  # uno por cada caja metálica
        )
        # Nota: el sistema de tornillos/tarugos de fijación (caja del empalme,
        # abrazaderas caddy, cajas de derivación y abrazaderas PVC) se calcula
        # más abajo, una vez que ya se conoce abrazaderas_fachada. Estas
        # variables se calculan acá porque el bloque de "caja de paso del
        # alimentador" (más abajo, pero antes del sistema de tornillos) también las necesita.
        hay_madera = "madera" in material_forrado_exterior  # True si el forrado exterior es de madera
        hay_pvc = "pvc" in material_forrado_exterior  # ej. "siding pvc" — usa el mismo tornillo de madera, sin tarugo
        hay_siding_metalico = ("metál" in material_forrado_exterior) or ("metal" in material_forrado_exterior)  # con o sin tilde
        hay_fibro = "fibro" in material_forrado_exterior  # True si el forrado exterior es fibrocemento
        hay_forrado_valido = any(x in material_forrado_exterior for x in ("madera", "fibro", "siding"))  # True si el forrado es alguno de los tipos reconocidos (si no, no se calculan tornillos)
        # material_forrado_exterior es una variable GLOBAL del script (se
        # pregunta en el input de la línea ~9310), no un parámetro de esta
        # función; llega en minúsculas y sin espacios a los lados.
        # ojo con la combinación "metal" sin "siding" (ej. "metalcon"):
        # hay_siding_metalico queda True pero hay_forrado_valido queda False, y
        # entonces ninguna rama del bloque de tornillos de más abajo se cumple
        # (ni siquiera la genérica), así que no se emite fila de tornillos.
        #Alimentador (ya viene calculado en res_alim, de seleccionar_alimentador)
        add_row(  # ítem: alimentador que va del empalme al TDA
            desc=f'Alimentador RV-K Cu 3x{round(float(res_alim["S"]),2)}mm^2',  # sección que entregó el cálculo por ampacidad y caída de tensión
            marcas_txt=marcas.get("Alimentador RV-K", ""),  # marcas del catálogo
            norma="SEC",  # exigencia SEC
            circuito="Alimentador",  # tramo de alimentador
            unidad="m",  # se vende por metro
            k=1,  # sin holgura
            longitud_m=f"{math.ceil(float(longitud_alimentador))} m",  # metros redondeados hacia arriba, en texto
            cantidad=math.ceil(float(longitud_alimentador))  # metros de cable a comprar
        )
        # Conduit PVC para tramo subterráneo de acometida. Es un material
        # de PVC, no de acero — usa Tabla N°4.29 (2 conductores, F+N), NO
        # el diámetro de desc_tubo_acom (tubo de acero, tabla simple aparte).
        diam_pvc_sub = None  # inicializar siempre antes de usar
        if "sub" in str(tipo_acometida).strip().lower():  # solo si la acometida es subterránea
            # el 3er parámetro de ducto_nominal_tablas se llama tipo_alimentador,
            # pero acá se le pasa el literal "subterráneo" a propósito para
            # forzar la Tabla N°4.29 con 2 conductores (fase + neutro de la
            # acometida). esa rama de la tabla devuelve el diámetro tal cual,
            # sin el piso de 32mm que si aplica la rama de ducto/embutido.
            diam_pvc_sub = ducto_nominal_tablas(sec_acom, 2, "subterraneo")  # diámetro comercial del conduit PVC para el tramo subterráneo
            if diam_pvc_sub:  # solo si la tabla entregó un diámetro
                # longitud_subterraneo_medidor solo se pregunta cuando el
                # empalme es en fachada, sin mástil y con acometida subterránea
                # (input de la línea ~10765, que además exige alimentador en ducto); en cualquier otro caso vale 0.
                # ojo: esta resta NO se limita a 0 (a diferencia de
                # long_pvc_acom del bloque de cámaras, que si usa max(0.0,...)),
                # así que si el usuario ingresa un tramo al medidor mayor que el
                # total, la cantidad de tubos queda negativa.
                longitud_pvc_sub = (float(longitud_transformador_empalme) - float(longitud_subterraneo_medidor))  # metros de PVC enterrado = tramo total menos el tramo que va del suelo al medidor
                cantidad_pvc_sub = int(math.ceil(longitud_pvc_sub / 3))  # cada tubo de conduit PVC viene en tramos de 3 metros
                add_row(desc=f"Conduit PVC {diam_pvc_sub}mm, 3mts",  # ítem: tubos de conduit PVC del tramo enterrado
                    marcas_txt=marcas.get("Conduit PVC", ""),  # marcas del catálogo
                    norma="SEC",  # exigencia SEC
                    circuito="Acometida subterránea",  # tramo enterrado de la acometida
                    unidad="u",  # se compra por tubo
                    k=1,  # sin holgura
                    longitud_m=f"{round(longitud_pvc_sub, 2)} m",  # metros reales del tramo, en texto
                    cantidad=cantidad_pvc_sub)  # cuántos tubos de 3m hay que comprar

        # Abrazadera adicional para PVC subterráneo
        if "sub" in str(tipo_acometida).strip().lower() and diam_pvc_sub:  # la abrazadera solo aplica si hay tramo PVC subterráneo
            add_row(desc=f"Abrazadera conduit PVC {diam_pvc_sub}mm",  # ítem: abrazadera del conduit PVC
                marcas_txt=marcas.get("Abrazadera conduit PVC alimentador", ""),  # marcas del catálogo
                norma="SEC",  # exigencia SEC
                circuito="Acometida subterránea",  # tramo enterrado de la acometida
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m="1 Unid",  # 1 unidad, en texto
                cantidad=1)  # basta con una en la subida

        # Conduit PVC según canalización alimentador (mínimo 32 mm)
        # Se define ANTES del bloque cámara para que diam_conduit esté disponible
        # canalizacion_txt se arma FUERA de esta función (línea ~12043) como
        # "Ø PVC Conduit NN mm", con ducto_nominal_tablas(S_alim, 3, tipo_alimentador).
        # en fachada el alimentador siempre es "en ducto", así que ese NN sale
        # de la Tabla N°4.19 (o de la N°4.20 por AWG equivalente) y ya viene
        # con el piso de 32mm aplicado; el max(32, ...) de acá es redundante
        # pero deja el piso explícito y cubre el caso de que el regex no calce.
        _m_diam_conduit = re.search(r'(\d+)\s*mm', str(canalizacion_txt))  # busca el número de mm dentro del texto de la canalización
        diam_conduit = max(32, int(_m_diam_conduit.group(1))) if _m_diam_conduit else 32  # diámetro final del conduit: nunca menor a 32mm
        metros_cond = math.ceil(float(longitud_alimentador))  # metros de conduit para el alimentador, redondeados hacia arriba

        # =========================================================
        # CÁMARA TIPO C — CANALIZACIÓN SUBTERRÁNEA RESIDENCIAL
        # RIC N°04, sección 7.9, 7.9.5, 7.9.8.4.3 y Anexo 4.5
        # Regla RIC N°4 art. 7.9.7.8 / 7.9.7.9 / 7.9.7.10:
        #   - Si L es 20 m o menos: 0 cámaras (forma U, RIC N°4 art. 7.9.7.10)
        #   - Si L es mayor a 20 m: ceil(L / 90) cámaras
        #   Acá solo se calculan las cámaras del tramo de acometida
        #   subterránea; el alimentador no genera cámaras en este bloque.
        # Dimensiones mínimas (Anexo 4.5 Lámina 2):
        #   tapa 440mm, marco 440x440mm, cámara 400x450mm, drenaje ø10mm
        # =========================================================
        if "sub" in str(tipo_acometida).strip().lower():  # todo el bloque de cámaras es solo para acometida subterránea
            # PVC enterrado = total acometida menos el tramo galvanizado (del suelo al medidor)
            long_pvc_acom = max(0.0,  # largo del tramo enterrado, nunca negativo
                float(longitud_transformador_empalme) - float(longitud_subterraneo_medidor))  # al total se le descuenta el tramo galvanizado que sube al medidor


            # Cámaras según RIC N°4 art. 7.9.7.8 / 7.9.7.10
            # 20 m o menos: 0 cámaras (forma U)
            # más de 20 m: ceil(L / 90) cámaras
            if long_pvc_acom <= 20.0:  # 20m o menos se resuelve con la forma U
                camaras_tipo_c = 0  # sin cámaras
            else:  # tramo largo
                camaras_tipo_c = int(math.ceil(long_pvc_acom / 90.0))  # una cámara cada 90m como máximo

            # Diámetro boquilla = diámetro del PVC subterráneo de la acometida
            _diam_pvc_sub_safe_f = None  # parte en None por si no quedó definido el diámetro del PVC
            try:  # puede que la variable ni exista
                _diam_pvc_sub_safe_f = diam_pvc_sub  # siempre existe (parte en None más arriba); el try queda solo por seguridad
            except NameError:  # nunca se creó diam_pvc_sub
                _diam_pvc_sub_safe_f = None  # se queda en None y abajo cae al valor por defecto
            diam_cam_c = _diam_pvc_sub_safe_f if _diam_pvc_sub_safe_f else 25  # si no hay dato, usa 25mm por defecto

            # Si va a haber al menos 1 cámara, agrega la cámara, su marco y sus boquillas al listado
            if camaras_tipo_c > 0:  # sin cámaras no hay nada que agregar
                add_row(  # ítem: cámaras tipo C del tramo subterráneo
                    desc="Cámara tipo C de hormigón prefabricado con tapa de acero diamantado 440x440mm",  # medidas minimás del Anexo 4.5
                    marcas_txt=marcas.get("Camara tipo C", ""),  # marcas del catálogo
                    norma="RIC 4 (7.9, 7.9.5, 7.9.7.8, 7.9.7.10, 7.9.8.4.3, Anexo 4.5)",  # artículos de canalización subterránea del RIC N°4
                    circuito="Acometida subterránea",  # tramo enterrado de la acometida
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{camaras_tipo_c} unid (L={round(long_pvc_acom,1)} m PVC sub.)",  # cuántas cámaras y con qué largo salieron
                    cantidad=camaras_tipo_c  # cantidad de cámaras
                )
                add_row(  # ítem: marco metálico que sujeta la tapa de la cámara
                    desc="Marco metálico galvanizado para cámara tipo C 440x440mm",  # mismas medidas que la cámara
                    marcas_txt=marcas.get("Marco metalico camara C", ""),  # marcas del catálogo
                    norma="RIC 4 (7.9.8, Anexo 4.5)",  # RIC N°4, artículo 7.9.8 y Anexo 4.5
                    circuito="Acometida subterránea",  # tramo enterrado de la acometida
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{camaras_tipo_c} unid",  # cuántos marcos, en texto
                    cantidad=camaras_tipo_c  # un marco por cada cámara
                )
                boquillas_camara = camaras_tipo_c * 2  # 2 boquillas por cámara: entrada y salida del conduit
                # boquillas para la entrada y la salida del conduit en cada cámara tipo C
                add_row(
                    desc=(
                        f"Boquilla de PVC ø{diam_cam_c}mm con borde redondeado "
                        f"para entrada/salida conduit en cámara tipo C"
                    ),
                    marcas_txt=marcas.get("Boquilla camara tipo C", ""),  # marcas sugeridas para la boquilla de cámara
                    norma="RIC 4 (7.9.8.9, 5.14)",  # artículos del RIC que piden la boquilla
                    circuito="Acometida subterránea",  # se carga a la acometida subterránea
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin holgura, la cantidad va tal cual
                    longitud_m=f"{boquillas_camara} unid",  # lo que se muestra en la columna de largo
                    cantidad=boquillas_camara  # total de boquillas
                )
        # Conduit para el alimentador: tramos de 3 metros, se redondea hacia arriba
        cantidad_conduit_alim = int(math.ceil(metros_cond / 3))  # tiras de conduit que hay que comprar para el alimentador
        # fila del conduit de PVC del alimentador
        add_row(
            desc=f'Conduit de PVC de {diam_conduit}mm, 3mts',  # el conduit se vende en tiras de 3 metros
            marcas_txt=marcas.get("Conduit PVC", ""),  # marcas sugeridas para conduit PVC
            norma="RIC 4.7.2",  # el 4.7.2 es el artículo de canalizaciones
            circuito="Alimentador",  # se carga al alimentador
            unidad="u",  # se compra por tira, no por metro
            k=1,  # sin factor extra
            longitud_m=f"{metros_cond} m",  # metros reales de canalización, van de referencia
            cantidad=cantidad_conduit_alim  # tiras de 3m a comprar
        )
        # Abrazaderas según Tabla N°4.24
        # separación entre abrazaderas (Tabla N°4.24): como el conduit tiene piso de 32mm, en la práctica siempre queda 1.50m
        sep_abraz_alim = 1.20 if diam_conduit <= 25 else 1.50  # separación entre abrazaderas, en metros
        cantidad_abrazaderas_alim = int(math.ceil(metros_cond / sep_abraz_alim))  # abrazaderas para todo el tramo del alimentador
        # fila de las abrazaderas del conduit del alimentador
        add_row(
            desc=f"Abrazadera conduit de PVC {diam_conduit}mm",  # la abrazadera va del mismo diámetro del conduit
            marcas_txt=marcas.get("Abrazadera conduit PVC alimentador", ""),  # marcas sugeridas para la abrazadera
            norma="RIC 4 (Tabla N°4.24)",  # la Tabla N°4.24 es la que fija la separación máxima
            circuito="Alimentador",  # también se carga al alimentador
            unidad="u",  # se cuentan de a una
            k=1,  # sin holgura
            longitud_m=f"{cantidad_abrazaderas_alim} unid",  # cuántas abrazaderas son
            cantidad=cantidad_abrazaderas_alim  # total de abrazaderas
        )
        # Cajas de paso alimentador en ducto (RIC 7.16.1.13)
        # Caja de paso solo si el alimentador va en ducto: 1 cada 20 metros.
        # Acá el ducto va por fachada, por eso la caja que se pide es estanca.
        # en fachada el input obliga a que el alimentador sea "en ducto"
        # (validación de la línea ~10731), así que la condición siempre se
        # cumple. el // 20 es división entera: 1 caja por cada 20 m COMPLETOS
        # (con 39 m va 1 caja, con 40 m van 2) y 0 cajas bajo los 20 m.
        cajas_paso_alim = int(metros_cond // 20) if "duct" in str(tipo_alimentador).strip().lower() else 0  # si el alimentador no va en ducto queda en 0 y no se agrega nada
        if cajas_paso_alim > 0:  # solo se arman las filas si de verdad hace falta una caja
            # Caja de paso ESTANCA (el alimentador va por fachada/exterior, es
            # distinta a la caja de PVC de interior que se usa en los circuitos).
            # No lleva tapa aparte: la caja estanca viene con su tapa integrada
            # de fábrica (se atornilla directo a la caja, no es una pieza separada
            # como la tapa ciega de las cajas de interior).
            # El tamaño depende del diámetro real del conduit del alimentador:
            # con tubos de 32/40mm entra una caja compacta 150x110x70mm; con
            # 50mm hace falta una más grande 190x140x90mm.
            _dc_alim = 0  # diámetro del conduit del alimentador, en mm
            try:  # el diámetro puede venir como texto, hay que forzarlo a número
                _dc_alim = int(diam_conduit)  # diámetro real del conduit
            except Exception:  # si no se puede leer el diámetro
                _dc_alim = 0  # queda en 0 y cae en la caja chica
            if _dc_alim >= 50:  # con conduit de 50mm no cabe la caja compacta
                _medida_caja_paso_alim = "190x140x90 mm"  # caja grande
            else:  # para conduit de 32 o 40mm
                _medida_caja_paso_alim = "150x110x70 mm"  # caja compacta
            # fila de la caja de paso estanca del alimentador
            add_row(
                desc=f"Caja de paso estanca IP65 de PVC/policarbonato para exterior {_medida_caja_paso_alim} (incluye tapa)",  # IP65 porque va a la intemperie
                marcas_txt=marcas.get("Cajas de paso estancas", ""),  # marcas sugeridas para cajas de paso estancas
                norma="RIC 4 (7.16.1.13)",  # el 7.16.1.13 es el que pide caja de paso cada 20m
                circuito="Alimentador - Caja de paso (tramo > 20m)",  # queda identificada como caja de paso del alimentador
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cajas_paso_alim} unid",  # cuántas cajas de paso son
                cantidad=cajas_paso_alim  # total de cajas de paso
            )
            # Salidas de caja PVC: 2 por caja de paso (entrada + salida conduit)
            salidas_paso_alim = 2 * cajas_paso_alim  # una para el conduit que entra y otra para el que sale
            # fila de las salidas de caja de las cajas de paso
            add_row(
                desc=f"Salida de caja conduit de PVC de {diam_conduit}mm",  # la salida de caja va del mismo diámetro del conduit
                marcas_txt=marcas.get("Salida de caja conduit", ""),  # marcas sugeridas para salida de caja
                norma="RIC 4.7.2",  # mismo artículo de canalizaciones
                circuito="Alimentador - Caja de paso (tramo > 20m)",  # se carga a la caja de paso del alimentador
                unidad="u",  # se cuentan de a una
                k=1,  # sin factor extra
                longitud_m=f"{salidas_paso_alim} unid",  # cuántas salidas son
                cantidad=salidas_paso_alim  # total de salidas de caja
            )
            # Fijación de la caja de paso (fila propia, porque el acumulador
            # general de tornillos ya se calculó y escribió antes de llegar
            # acá — sumar solo a _cajas_total no alcanza a reflejarse en la
            # fila de "Tornillo" del Excel). Usa el mismo criterio por
            # material de forrado exterior que el resto de fijaciones del
            # alimentador (hay_madera/hay_pvc/hay_siding_metalico/hay_fibro,
            # calculadas más arriba) — con fibrocemento hace falta tarugo
            # antes del tornillo, no sirve un autoperforante directo.
            _tornillos_paso_alim = 4 * cajas_paso_alim  # 4 puntos de fijación por caja
            if (hay_madera or hay_pvc) and hay_forrado_valido:  # primero se revisa el caso más simple de fijación
                # forrado de madera o siding PVC: tirafondo directo, sin tarugo
                add_row(
                    desc='Tirafondo hexagonal para madera de 1/4" x 1 1/2" (caja de paso alimentador)',  # tirafondo que agarra directo en la madera
                    marcas_txt=marcas.get("Tirafondo hexagonal madera", ""),  # marcas sugeridas para tirafondo de madera
                    norma="-",  # es fijación, no la pide una norma en particular
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # se carga a la caja de paso del alimentador
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{_tornillos_paso_alim} unid",  # cuántos tornillos son
                    cantidad=_tornillos_paso_alim  # 4 por cada caja de paso
                )
            elif hay_siding_metalico and hay_forrado_valido:  # caso del siding metálico
                # forrado de siding metálico: tornillo autoperforante especial
                add_row(
                    desc='Tornillo autoperforante hexagonal 10 x 1-1/2" (caja de paso alimentador)',  # autoperforante, no hay que hacer el agujero antes
                    marcas_txt=marcas.get("Tornillo autoperforante hexagonal", ""),  # marcas sugeridas para autoperforante hexagonal
                    norma="-",  # sin norma asociada
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # misma caja de paso del alimentador
                    unidad="u",  # se cuentan de a uno
                    k=1,  # sin factor extra
                    longitud_m=f"{_tornillos_paso_alim} unid",  # cuántos tornillos son
                    cantidad=_tornillos_paso_alim  # 4 puntos de fijación por caja
                )
            elif hay_fibro:  # último caso con material conocido
                # fibrocemento: necesita tarugo primero y después el tirafondo, van los 2 juntos
                add_row(
                    desc="Tarugo paloma 8mm (caja de paso alimentador)",  # el tarugo paloma es el que sirve en plancha delgada
                    marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas sugeridas para tarugo paloma
                    norma="-",  # sin norma, es fijación
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # caja de paso del alimentador
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{_tornillos_paso_alim} unid",  # cuántos tarugos son
                    cantidad=_tornillos_paso_alim  # un tarugo por cada punto de fijación
                )
                # y el tirafondo que va dentro del tarugo
                add_row(
                    desc="Tirafondo zincado punta fina 4,5 x 30 mm, rosca gruesa (caja de paso alimentador)",  # punta fina para no reventar la plancha
                    marcas_txt=marcas.get("Tirafondo hexagonal madera", ""),  # se reutiliza la marca del tirafondo de madera
                    norma="-",  # sin norma
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # caja de paso del alimentador
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin factor extra
                    longitud_m=f"{_tornillos_paso_alim} unid",  # misma cantidad que los tarugos
                    cantidad=_tornillos_paso_alim  # van de a uno con el tarugo
                )
            else:  # no se reconoció el material del forrado
                # no se sabe el material del forrado: fila genérica para que se defina después
                add_row(
                    desc="Tornillo (definir según material de forrado exterior) — caja de paso alimentador",  # queda pendiente elegir el tornillo en terreno
                    marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas de tornillos
                    norma="-",  # sin norma
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # caja de paso del alimentador
                    unidad="u",  # se cuenta por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{_tornillos_paso_alim} unid",  # la cantidad igual queda estimada
                    cantidad=_tornillos_paso_alim  # 4 por caja de paso
                )
            # se suma al acumulador general de cajas por consistencia, pero a
            # esta altura ya no lo lee nadie: la fila general de tornillos
            # (líneas ~5940 y ~5954) y la espuma del panel SIP (línea ~6133)
            # se calcularon mucho antes que este bloque. por eso la fijación
            # de estas cajas se emitio recién como fila propia, arriba.
            _cajas_total += cajas_paso_alim  # suma estas cajas al total de cajas de la instalación
        #caja de derivación metálica con tapa (cantidad_caja_derivacion ya calculada más arriba)
        add_row(
            desc="Caja de derivación metálica pregalvanizada 100x65x65mm con tapa",  # caja de derivación del empalme, viene con tapa
            marcas_txt=marcas.get("Caja derivacion metalica", ""),  # marcas sugeridas para caja de derivación metálica
            norma="SEC",  # formato que exige la SEC en el empalme
            circuito="Empalme",  # se carga al empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin holgura
            longitud_m=f"{cantidad_caja_derivacion} unid",  # cuántas cajas de derivación son
            cantidad=cantidad_caja_derivacion  # la cantidad ya venía calculada de más arriba
        )
        # tornillo con golilla para aterrizar las cajas metálicas
        add_row(
            desc="Tornillo autoperforante punta broca 8 x 1/2\" cabeza lenteja + golilla (conexión tierra caja metálica)",  # la golilla es la que hace contacto con la caja
            marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas de tornillos
            norma="RIC 4 (5.13)",  # el 5.13 pide que las cajas metálicas queden a tierra
            circuito="Empalme (caja de empalme + caja(s) de derivación)",  # cubre la caja de empalme y las de derivación
            unidad="u",  # se cuenta por unidad
            k=1,  # sin factor extra
            longitud_m=f"{1 + cantidad_caja_derivacion} unid",  # 1 de la caja de empalme más 1 por cada derivación
            cantidad=1 + cantidad_caja_derivacion  # uno por cada caja metálica
        )

        # Sistema completo de tornillos/tarugos de fijación (empalme en
        # fachada), según el material del forrado exterior de la casa.
        # criterios de conteo (son de montaje, no hay artículo del RIC detras):
        #   Grupo A: caja metálica del empalme, fijo en 6 tornillos (4 esquinas
        #            + 2 al medio), siempre con golilla.
        #   Grupo B: 1 tornillo por abrazadera caddy (cantidad_caddy, que ya
        #            son 3 por tubo galvanizado) + 4 por cada caja de
        #            derivación metálica + 2 por cada abrazadera de conduit PVC
        #            del alimentador en fachada (abrazaderas_fachada).
        # se separan en 2 grupos porque llevan tornillo de distinto largo:
        # el A carga el peso del gabinete del medidor, el B solo tubos y cajas
        # chicas. en fibrocemento no se separan (ver rama de más abajo).
        if "sub" in str(tipo_acometida).strip().lower():  # se mira si la acometida es subterránea
            # acometida subterránea: hay 2 cajas en fachada y una abrazadera extra
            cajas_fachada = 2  # caja de empalme más caja de derivación. ojo: en este bloque la variable no se usa después, los tornillos salen de cantidad_caja_derivacion
            # abrazaderas_fachada = las del conduit PVC del alimentador
            # (cantidad_abrazaderas_alim, contadas por la Tabla N°4.24) más la
            # abrazadera del conduit PVC subterráneo de la acometida, esa que
            # se agrego con cantidad=1 en el bloque de más arriba.
            abrazaderas_fachada = cantidad_abrazaderas_alim + 1  # una abrazadera más por el tramo que sube desde el suelo
        else:  # caso aéreo
            # acometida aérea: solo 1 caja en fachada
            cajas_fachada = 1  # solo la caja del empalme
            abrazaderas_fachada = cantidad_abrazaderas_alim  # las mismas abrazaderas del alimentador

        cant_grupo_A = 6  # 6 tornillos fijos para la caja metálica del empalme
        cant_grupo_B = (1 * cantidad_caddy) + (4 * cantidad_caja_derivacion) + (2 * abrazaderas_fachada)  # 1 por caddy, 4 por caja de derivación y 2 por abrazadera PVC

        # Según el material del forrado exterior, elige el tipo de tornillo/tarugo correcto
        if (hay_madera or hay_pvc) and hay_forrado_valido:  # madera o siding PVC: tirafondo directo, sin tarugo
            # tornillos de la caja metálica del empalme
            add_row(
                desc='Tirafondo hexagonal 1/4" x 1 1/2" + golilla 1/4"',  # tirafondo con golilla para colgar la caja
                marcas_txt=marcas.get("Tirafondo hexagonal madera", ""),  # marcas sugeridas para tirafondo de madera
                norma="-",  # sin norma, es fijación
                circuito="Caja metálica del empalme",  # se carga a la caja del empalme
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_grupo_A} unid",  # los 6 del grupo A
                cantidad=cant_grupo_A  # grupo A
            )
            # tornillos del resto de las fijaciones de la fachada
            add_row(
                desc='Tornillo 8x1" cabeza lenteja punta fina',  # tornillo más chico para abrazaderas y cajas de derivación
                marcas_txt=marcas.get("Tornillo punta fina madera", ""),  # marcas sugeridas para tornillo punta fina
                norma="-",  # sin norma
                circuito="Abrazaderas caddy, caja(s) de derivación metálica y abrazaderas PVC",  # junta caddy, derivaciones y abrazaderas PVC
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m=f"{cant_grupo_B} unid",  # total del grupo B
                cantidad=cant_grupo_B  # grupo B
            )
        elif hay_fibro and hay_forrado_valido:  # fachada de fibrocemento
            # fibrocemento: necesita tarugo + tornillo juntos, en 1 solo grupo combinado
            cant_fibro_total = cant_grupo_A + cant_grupo_B  # acá no se separan los grupos, van todos juntos
            # el tornillo del combo tarugo más tornillo
            add_row(
                desc='Tornillo 8x1 1/2" punta fina cabeza lenteja',  # punta fina para plancha de fibrocemento
                marcas_txt=marcas.get("Tornillo punta fina madera", ""),  # marcas sugeridas para tornillo punta fina
                norma="-",  # sin norma
                circuito="Caja metálica del empalme, abrazaderas caddy, caja(s) de derivación y abrazaderas PVC",  # cubre todas las fijaciones del empalme en fachada
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_fibro_total} unid",  # grupo A más grupo B
                cantidad=cant_fibro_total  # total combinado
            )
            # y el tarugo que va antes del tornillo
            add_row(
                desc="Tarugo paloma N°8",  # tarugo paloma para plancha delgada
                marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas sugeridas para tarugo paloma
                norma="-",  # sin norma
                circuito="Caja metálica del empalme, abrazaderas caddy, caja(s) de derivación y abrazaderas PVC",  # mismas fijaciones que el tornillo de arriba
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m=f"{cant_fibro_total} unid",  # uno por cada tornillo
                cantidad=cant_fibro_total  # misma cantidad que los tornillos
            )
        elif hay_siding_metalico and hay_forrado_valido:  # siding metálico: acá todo es autoperforante
            # tornillos de la caja metálica del empalme
            add_row(
                desc='Tornillo autoperforante hexagonal 1/4" x 1 1/2" + golilla',  # hexagonal con golilla, perfora la plancha
                marcas_txt=marcas.get("Tornillo autoperforante hexagonal", ""),  # marcas sugeridas para autoperforante hexagonal
                norma="-",  # sin norma
                circuito="Caja metálica del empalme",  # se carga a la caja del empalme
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_grupo_A} unid",  # los 6 del grupo A
                cantidad=cant_grupo_A  # grupo A
            )
            # tornillos de abrazaderas y cajas de derivación
            add_row(
                desc='Tornillo cabeza lenteja 8x1 1/4" punta broca',  # punta broca, más chico que el de la caja del empalme
                marcas_txt=marcas.get("Tornillo autoperforante broca", ""),  # marcas sugeridas para autoperforante punta broca
                norma="-",  # sin norma
                circuito="Abrazaderas caddy, caja(s) de derivación metálica y abrazaderas PVC",  # caddy, derivaciones y abrazaderas PVC
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m=f"{cant_grupo_B} unid",  # total del grupo B
                cantidad=cant_grupo_B  # grupo B
            )
        elif not (hay_madera or hay_pvc or hay_siding_metalico or hay_fibro):  # no se identificó ningún material de forrado
            # no se sabe el material del forrado: 1 sola fila genérica con el total, para definir después
            cant_total_indef = cant_grupo_A + cant_grupo_B  # se junta todo en un solo total
            # fila genérica, el tipo de tornillo se define después en terreno
            add_row(
                desc='Tornillo (definir según material de tabique)',  # queda anotado que falta definirlo
                marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas de tornillos
                norma="-",  # sin norma
                circuito="Caja metálica del empalme, abrazaderas caddy, caja(s) de derivación y abrazaderas PVC",  # cubre todas las fijaciones del empalme
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_total_indef} unid",  # el total sin separar por grupo
                cantidad=cant_total_indef  # grupo A más grupo B
            )

        # Conductor puesta a tierra camarilla N°2 (TDA - puesta a tierra 2)
        # +2m: 1 chicote en la camarilla N°2 + 1 chicote en el TDA
        metros_pt2 = math.ceil(float(dist_tda_pt2)) + 2  # largo del conductor de tierra entre el TDA y la camarilla N°2, redondeado hacia arriba y sumando 2m de chicotes
        # fila del conductor verde de la puesta a tierra N°2
        add_row(
            desc=f"Conductor THWN-2 {sec_pt}mm^2 verde",  # verde porque es tierra, de la misma sección sec_pt
            marcas_txt=marcas.get("Conductor THWN-2 verde", ""),  # marcas sugeridas para THWN-2 verde
            norma="SEC",  # conductor de tierra exigido por la SEC
            circuito="Puesta a tierra N°2 (Tda - Camarilla N°2)",  # tramo entre el TDA y la camarilla N°2
            unidad="m",  # este va por metro, es cable
            k=1,  # sin holgura
            longitud_m=f"{metros_pt2} m",  # metros del tramo
            cantidad=metros_pt2  # los mismos metros
        )

        # Tarugo + tornillo chico (6mm) para las abrazaderas de conduit PVC
        # que van pegadas a la fachada. Solo se agregan con forrado de
        # fibrocemento, que es el único material que obliga a tarugo.
        # criterio: 2 tarugos por cada abrazadera de fachada, más 1 extra si
        # además hay conduit PVC subterráneo de acometida (por la abrazadera
        # de la subida desde el suelo).
        # ojo, doble conteo: en el caso fibrocemento el grupo B de más arriba
        # YA cobro 2 tornillos de 8x1 1/2 y 2 tarugos paloma N°8 por cada una
        # de estas mismas abrazaderas. estas dos filas de 6mm se suman además
        # de aquellas, o sea los mismos puntos de fijación quedan en el Excel
        # dos veces y con dos medidas distintas.
        # La caja de derivación metálica no entra en este conteo: sus tornillos
        # ya salieron en el grupo B (4 por caja).
        # hay_fibro ya se definió más arriba
        cantidad_tarugos = 0  # si el forrado no es fibrocemento queda en 0 y no se agrega nada
        if hay_fibro:  # solo el fibrocemento necesita tarugo
            cantidad_tarugos = int(abrazaderas_fachada * 2)  # 2 tarugos por cada abrazadera PVC en fachada
            if "sub" in str(tipo_acometida).strip().lower() and diam_pvc_sub:  # si además hay conduit PVC subterráneo hay un punto de fijación extra
                cantidad_tarugos += 1  # acometida subterránea con conduit PVC: 1 tarugo extra
            # fila de los tarugos de las abrazaderas en fachada
            add_row(
                desc="Tarugo paloma 6mm",  # tarugo de 6mm, más chico que el de las cajas
                marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas sugeridas para tarugo paloma
                norma="-",  # sin norma
                circuito="Abrazaderas PVC en fachada",  # abrazaderas del conduit en la fachada
                unidad="u",  # se cuenta por unidad
                k=1,  # sin factor extra
                longitud_m=f"{cantidad_tarugos} unid",  # cuántos tarugos son
                cantidad=cantidad_tarugos  # total de tarugos
            )
            # y el tornillo que va dentro de cada tarugo
            add_row(
                desc='Tornillo punta plana 6x1 1/4" para tarugo',  # punta plana, del diámetro que acepta el tarugo
                marcas_txt=marcas.get("Tornillo para tarugo paloma", ""),  # marcas sugeridas para tornillo de tarugo paloma
                norma="-",  # sin norma
                circuito="Abrazaderas PVC en fachada",  # mismas abrazaderas de la fachada
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m=f"{cantidad_tarugos} unid",  # uno por cada tarugo
                cantidad=cantidad_tarugos  # misma cantidad que los tarugos
            )

        # Portafusible aéreo loza según interruptor termomagnético del empalme
        # interruptor_texto se arma fuera de esta función (línea ~11366) como
        # "1x{N}A / 6kA / Curva D", y ese N puede ser 25, 32, 40, 50 o 63.
        # la tabla de abajo solo cubre 25, 32 y 40: con 50A o 63A fusible_A se
        # queda en None y NO se agrega ninguna fila de portafusible.
        tm_empalme_A = parse_in_tm(interruptor_texto)  # amperaje del TM del empalme, sacado del texto
        fusible_A = None  # si el TM no cae en ninguno de los casos de abajo no se agrega portafusible
        if tm_empalme_A == 25:  # empalme de 25A
            fusible_A = 30  # TM de 25A usa fusible de 30A
        elif tm_empalme_A in [32, 40]:  # empalmes más grandes
            fusible_A = 60  # TM de 32A o 40A usa fusible de 60A
        if fusible_A:  # solo se agrega la fila si quedó definido el fusible
            # fila del portafusible aéreo de loza del empalme
            add_row(
                desc=f"Portafusible de loza con fusibles cartucho {fusible_A}A",  # el fusible cartucho va acorde al TM del empalme
                marcas_txt=marcas.get("Portafusible de loza", ""),  # marcas sugeridas para portafusible de loza
                norma="SEC",  # lo pide la SEC en el empalme
                circuito="Empalme",  # se carga al empalme
                unidad="u",  # se cuenta por unidad
                k=1,  # sin holgura
                longitud_m="1 unid",  # siempre uno
                cantidad=1  # un portafusible
            )

#============================================
#Empalme independiente poste madera/metálico
    elif (str(tipo_instalacion_empalme).strip().lower() == "independiente"  # empalme parado en su propio poste, no pegado a la fachada
        and ("aer" in str(tipo_alimentador).strip().lower() or "sub" in str(tipo_alimentador).strip().lower())  # el alimentador puede ser aéreo o subterráneo, da igual
        and ("aer" in str(tipo_acometida).strip().lower() or "sub" in str(tipo_acometida).strip().lower())):  # y la acometida también, cualquiera de las dos sirve
        # ------------------------------------------------------------
        # Caso: empalme independiente en poste de madera o metálico
        # (no está pegado a la fachada, está en un poste aparte).
        # Aplica tanto si el alimentador/acometida son aéreos como
        # subterráneos.
        # ------------------------------------------------------------
        # Unidad de medida monofásico / medidor
        add_row(
            desc="Unidad de medida monofásica 220V 50Hz 50A (medidor)",  # el medidor propiamente tal
            marcas_txt=marcas.get("Medidor empalme", ""),  # marcas sugeridas para el medidor
            norma="SEC",  # el empalme lo aprueba la SEC
            circuito="Empalme",  # se carga al empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin holgura
            longitud_m="1 unid",  # siempre uno
            cantidad=1  # un medidor por vivienda
        )
        # Caja metálica empalme
        add_row(
            desc="Caja de empalme metálica para medidor monofásico 405x200x137 mm IP54",  # gabinete donde va el medidor, IP54 porque queda a la intemperie
            marcas_txt=marcas.get("Caja empalme", ""),  # marcas sugeridas para caja de empalme
            norma="SEC",  # formato que exige la SEC
            circuito="Empalme",  # se carga al empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin factor extra
            longitud_m="1 unid",  # una sola caja
            cantidad=1  # una
        )
        # Disyuntor termomagnético empalme
        add_row(
            desc=f"Disyuntor termomagnético {interruptor_texto}",  # el TM del empalme, con el amperaje que se definió más arriba
            marcas_txt=marcas.get("Disyuntor empalme", ""),  # marcas sugeridas para el disyuntor del empalme
            norma="SEC",  # lo fija la SEC según la potencia contratada
            circuito="Empalme",  # se carga al empalme
            unidad="u",  # se cuenta por unidad
            k=1,  # sin holgura
            longitud_m="1 unid",  # uno solo
            cantidad=1  # un TM
        )
        # Conductor acometida
        add_row(
            desc=f"Cable {acometida_txt}",  # cable de la acometida, el texto se armó más arriba
            marcas_txt=marcas.get("Cable acometida", ""),  # marcas sugeridas para el cable de acometida
            norma="SEC",  # sección y largo los revisa la SEC
            circuito="Acometida",  # se carga a la acometida
            unidad="m",  # este va por metro, es cable
            k=1,  # sin holgura
            longitud_m=f"{math.ceil(float(longitud_transformador_empalme))} m",  # distancia del transformador al empalme, redondeada hacia arriba
            cantidad=math.ceil(float(longitud_transformador_empalme))  # los mismos metros
        )
        # ============================================================
        # Tubo conduit galvanizado — dos diámetros posibles y distintos:
        #   - Tramo ACOMETIDA (+ puesta a tierra): según sección del cable
        #     de la acometida (sec_acom), sin piso mínimo (25/32/40mm).
        #   - Tramo ALIMENTADOR (aéreo O subterráneo, da igual): según
        #     sección del cable del alimentador (sec_alim), con PISO
        #     MÍNIMO de 32mm (nunca puede quedar en 25mm, aunque el cable
        #     sea delgado) — el alimentador SIEMPRE calcula su propio
        #     diámetro, no se fusiona en silencio con el de la acometida
        #     solo porque sea subterráneo.
        # Si ambos diámetros coinciden, se fusionan en 1 sola fila; si son
        # distintos, se generan 2 filas separadas.
        # ============================================================
        # acometida_txt se arma fuera de esta función (línea ~11950) como
        # "Concéntrico Cu 2x N mm^2", por eso el split por "2x". acá va SIN
        # try/except: si el texto no trae ese patrón, el programa se cae.
        # más abajo, en los ferrules, se vuelve a parsear lo mismo pero con
        # try/except y con 4.0mm^2 por defecto.
        sec_acom = float(acometida_txt.split("2x")[1].replace("mm^2", "").strip())  # sección del cable de acometida, sacada del texto "2x...mm^2"
        _acom_es_aerea = "aer" in str(tipo_acometida).strip().lower()  # True si la acometida es aérea
        _alim_es_aereo = "aer" in str(tipo_alimentador).strip().lower()  # True si el alimentador es aéreo

        # Tubo de ACERO galvanizado: siempre tabla simple de umbrales, sea
        # aérea o subterránea — la Tabla N°4.29 no aplica acá, es solo para
        # el Terminal PVC / Conduit PVC (materiales de PVC, no de acero).
        if sec_acom <= 6:  # cables delgados, hasta 6mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 25mm, 3mts"  # tubo de 25mm
        elif sec_acom <= 16:  # hasta 16mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 32mm, 3mts"  # tubo de 32mm
        elif sec_acom <= 35:  # hasta 35mm^2
            desc_tubo_acom = "Tubo conduit galvanizado 40mm, 3mts"  # tubo de 40mm
        else:  # sobre 35mm^2 no está tabulado acá
            desc_tubo_acom = ""  # sección fuera de rango: no hay diámetro definido

        # El alimentador SIEMPRE calcula su propio diámetro (con piso
        # mínimo 32mm), sea aéreo o subterráneo.
        try:  # la sección del alimentador también viene metida en el texto
            sec_alim = float(alim_txt.split("3x")[1].replace("mm^2", "").strip())  # sección del cable del alimentador, sacada del texto "3x...mm^2"
        except Exception:  # si el texto no trae el formato esperado
            sec_alim = 0.0  # queda en 0 y cae en el diámetro mínimo
        # Piso mínimo 32mm para el alimentador (no hay opción de 25mm)
        if sec_alim <= 16:  # hasta 16mm^2 se usa el mínimo permitido
            desc_tubo_alim = "Tubo conduit galvanizado 32mm, 3mts"  # el piso son 32mm, no baja de ahí
        elif sec_alim <= 35:  # hasta 35mm^2
            desc_tubo_alim = "Tubo conduit galvanizado 40mm, 3mts"  # de 16 a 35mm^2 va tubo de 40mm
        else:  # sobre 35mm^2
            desc_tubo_alim = ""  # sección fuera de rango: no hay diámetro definido

        cantidad_total_tubos = 0  # se usa después para abrazaderas caddy
        largo_total_acom_pt = 0.0  # metros de tubo que suman acometida + puesta a tierra
        cantidad_total_acom_pt = 0  # tubos de 3m de ese mismo grupo
        largo_tramo_alim = 0.0  # metros de tubo del tramo del alimentador
        cantidad_tramo_alim = 0  # tubos de 3m del alimentador
        cantidad_caddy_pt2 = 0  # abrazaderas caddy del tramo PT2, que se fijan en la fachada y no en el poste

        if desc_tubo_acom:  # solo si la acometida quedó con un diámetro de tubo definido
            # Largo del tramo de acometida (+PT). El alimentador (aéreo o
            # subterráneo) no se suma acá: tiene su propio diámetro y se
            # calcula aparte más abajo.
            if _acom_es_aerea:  # el largo del tramo depende de si la acometida es aérea o subterránea
                largo_tramo_acom = float(altura_acometida_aerea)  # acometida aérea: usa la altura del poste hasta el empalme
            else:  # acometida subterránea
                largo_tramo_acom = float(longitud_subterraneo_medidor2)  # acometida subterránea: usa el largo del tramo enterrado
            cantidad_tramo_acom = int(math.ceil(largo_tramo_acom / 3))  # cantidad de tubos de 3m necesarios, redondeado hacia arriba
            # Puesta a tierra: mismo diámetro que la acometida
            largo_total_pt = float(dist_empalme_pt1) + float(dist_tda_pt2)  # suma los 2 tramos de puesta a tierra (empalme-PT1 y TDA-PT2)
            cantidad_tubo_pt = int(math.ceil(largo_total_pt / 3))  # tubos de 3m para los dos tramos de tierra
            largo_total_acom_pt = largo_tramo_acom + largo_total_pt  # metros totales del grupo acometida + PT
            cantidad_total_acom_pt = cantidad_tramo_acom + cantidad_tubo_pt  # tubos totales del mismo grupo
            # PT2 (del TDA a la camarilla N°2) físicamente está en la fachada de la
            # casa, no en el poste — se guarda aparte para reasignar sus
            # tornillos de fijación de la abrazadera caddy hacia la fachada
            # más abajo (el material "Abrazadera tipo caddy" en sí no cambia,
            # solo QUIÉN paga el tornillo de fijación).
            # ceil(dist_tda_pt2 / 3) = tubos de 3m del tramo TDA - camarilla N°2,
            # y por cada tubo van 3 abrazaderas caddy (mismo criterio que el
            # resto). este número se RESTA más abajo del grupo B del poste y se
            # SUMA al de fachada: son las únicas caddy que no van en el poste.
            cantidad_caddy_pt2 = int(math.ceil(float(dist_tda_pt2) / 3)) * 3  # 3 abrazaderas caddy por cada tubo de 3m del tramo TDA - PT2

        if desc_tubo_alim:  # solo si el alimentador quedó con diámetro definido
            # Largo del tramo del alimentador: aéreo usa su propio poste,
            # subterráneo usa su propio "cuello".
            if _alim_es_aereo:  # igual que la acometida, el largo cambia según aéreo o subterráneo
                largo_tramo_alim = float(longitud_poste_alimentador_aereo)  # alimentador aéreo: largo del poste propio
            else:  # alimentador subterráneo
                largo_tramo_alim = float(longitud_subterraneo_medidor2)  # alimentador subterráneo: largo del tramo enterrado
            cantidad_tramo_alim = int(math.ceil(largo_tramo_alim / 3))  # cantidad de tubos de 3m necesarios

        # ¿Se fusionan en 1 fila (mismo diámetro) o van 2 filas separadas?
        if desc_tubo_acom and desc_tubo_alim and desc_tubo_acom == desc_tubo_alim:
            # mismo diámetro en ambos tramos: se suman y va todo en 1 sola fila
            largo_total_tubos = largo_total_acom_pt + largo_tramo_alim  # metros de tubo de todo junto
            cantidad_total_tubos = cantidad_total_acom_pt + cantidad_tramo_alim  # tubos de 3m de todo junto
            # fila única de tubo conduit: acometida, alimentador y las dos tierras
            add_row(
                desc=desc_tubo_acom,  # el diámetro es el mismo para los dos tramos, va uno solo
                marcas_txt=marcas.get("Tubo conduit galvanizado acometida", ""),  # marcas aprobadas para el tubo conduit
                norma="SEC",  # queda como exigencia SEC, sin artículo RIC puntual
                circuito="Acometida / Alimentador / Puesta a tierra",  # esta fila cubre los tres tramos a la vez
                unidad="u",  # se compra por tubo, no por metro
                k=1,  # sin factor de mayoracion
                longitud_m=f"{round(largo_total_tubos, 2)} m",  # metros reales sumados, solo de referencia
                cantidad=cantidad_total_tubos  # tubos de 3m que hay que comprar
            )
        else:  # no se pudo fusionar en una sola fila
            # diámetros distintos (o falta uno de los dos): van en filas separadas
            if desc_tubo_acom and cantidad_total_acom_pt:  # hay tubo de acometida y quedó con cantidad, va su propia fila
                cantidad_total_tubos += cantidad_total_acom_pt  # acumula estos tubos en el total (sirve para las abrazaderas caddy)
                # fila del tubo de acometida + puesta a tierra
                add_row(
                    desc=desc_tubo_acom,  # diámetro del tubo de la acometida
                    marcas_txt=marcas.get("Tubo conduit galvanizado acometida", ""),  # mismas marcas del tubo de acometida
                    norma="SEC",  # exigencia SEC
                    circuito="Acometida / Puesta a tierra",  # esta fila es acometida + las dos tierras
                    unidad="u",  # por tubo
                    k=1,  # sin mayoracion
                    longitud_m=f"{round(largo_total_acom_pt, 2)} m",  # metros del tramo acometida + PT
                    cantidad=cantidad_total_acom_pt  # tubos de 3m de ese grupo
                )
            if desc_tubo_alim and cantidad_tramo_alim:  # hay tubo de alimentador con cantidad, va aparte
                cantidad_total_tubos += cantidad_tramo_alim  # también se acumula al total de tubos
                # fila del tubo del alimentador, con su diámetro
                add_row(
                    desc=desc_tubo_alim,  # diámetro propio del alimentador
                    marcas_txt=marcas.get("Tubo conduit galvanizado acometida", ""),  # reusa las marcas del tubo de acometida
                    norma="SEC",  # exigencia SEC
                    circuito="Alimentador",  # esta fila es solo del tramo del alimentador
                    unidad="u",  # por tubo
                    k=1,  # sin mayoracion
                    longitud_m=f"{round(largo_tramo_alim, 2)} m",  # metros del tramo del alimentador
                    cantidad=cantidad_tramo_alim  # tubos de 3m del alimentador
                )

        # Cabeza de servicio — 1 por cada tramo aéreo presente, con el
        # TAMAÑO DE SU PROPIO diámetro (no siempre el mismo tamaño para
        # los dos tramos, ya que acometida y alimentador pueden diferir).
        def _desc_cabeza_de(diam_txt):  # arma el nombre de la cabeza de servicio según el diámetro del tubo
            # recibe el texto del tubo conduit (con su diámetro) y devuelve
            # la cabeza de servicio del mismo tamaño (o "" si no calza)
            if "25mm" in diam_txt:  # tubo de 25mm
                return 'Cabeza de servicio 3/4"'  # cabeza de 3/4"
            elif "32mm" in diam_txt:  # tubo de 32mm
                return 'Cabeza de servicio 1"'  # cabeza de 1"
            elif "40mm" in diam_txt:  # tubo de 40mm
                return 'Cabeza de servicio 1 1/4"'  # cabeza de 1 1/4"
            elif "50mm" in diam_txt:  # tubo de 50mm
                return 'Cabeza de servicio 2"'  # cabeza de 2"
            return ""  # diámetro raro o vacío: no se agrega cabeza

        _cabezas = {}  # relaciona cada descripción con su cantidad (se fusionan si coinciden)
        # Si la acometida es aérea, le corresponde 1 cabeza de servicio (con el
        # tamaño según el diámetro de SU tubo)
        if _acom_es_aerea and desc_tubo_acom:
            _d = _desc_cabeza_de(desc_tubo_acom)  # tamaño de cabeza que le toca al tubo de la acometida
            if _d:  # solo si el diámetro calzo con alguna medida
                _cabezas[_d] = _cabezas.get(_d, 0) + 1  # suma 1 cabeza de ese tamaño
        # Si el alimentador es aéreo, también le corresponde 1 cabeza de
        # servicio (puede ser de otro tamaño distinto a la de la acometida)
        if _alim_es_aereo and desc_tubo_alim:
            _d = _desc_cabeza_de(desc_tubo_alim)  # tamaño de cabeza del tubo del alimentador
            if _d:  # igual, solo si calzo
                _cabezas[_d] = _cabezas.get(_d, 0) + 1  # si coincide con la de la acometida quedan 2 en la misma fila
        # Agrega una fila de material por cada tamaño de cabeza de servicio
        # que haya quedado en el diccionario (1 o 2 filas, según si acometida
        # y alimentador usan el mismo diámetro o no)
        for _desc_cab, _cant_cab in _cabezas.items():
            add_row(
                desc=_desc_cab,  # cabeza de servicio del tamaño que corresponda
                marcas_txt=marcas.get("Cabeza de servicio", ""),  # marcas de cabeza de servicio
                norma="SEC",  # exigida por SEC en las bajadas aereas
                circuito="Acometida y/o alimentador",  # puede venir de la acometida, del alimentador o de los dos
                unidad="u",  # se cuenta por unidad
                k=1,  # sin mayoracion
                longitud_m=f"{_cant_cab} unid",  # acá no hay metros, se muestra la cantidad
                cantidad=_cant_cab  # 1 o 2 según cuántos tramos aéreos usan este diámetro
            )
        # Cáncamo de acero galvanizado
        # 1 por cada tramo aéreo presente (acometida y/o alimentador).
        # No tiene variantes de tamaño, por lo que no se ve afectado por
        # el split de diámetros de arriba.
        if _acom_es_aerea or _alim_es_aereo:
            cantidad_cancamo = (1 if _acom_es_aerea else 0) + (1 if _alim_es_aereo else 0)  # 1 o 2, según cuántos tramos son aéreos
            # fila del cáncamo, 1 o 2 unidades según los tramos aéreos
            add_row(
                desc="Cáncamo abierto de acero galvanizado 7,8 x 110 mm",  # cáncamo donde se amarra el cable aéreo al poste
                marcas_txt=marcas.get("Cancamo abierto", ""),  # marcas de cáncamo
                norma="-",  # ferreteria, no tiene norma asociada
                circuito="Acometida y/o alimentador",  # sirve para acometida y/o alimentador aéreo
                unidad="u",  # se cuenta por unidad
                k=1,  # sin mayoracion
                longitud_m=f"{cantidad_cancamo} unid",  # se muestra la cantidad, no metros
                cantidad=cantidad_cancamo  # 1 por cada tramo aéreo
                )
        #Granpa de retención — según sección de la acometida
        if "aer" in str(tipo_acometida).strip().lower():  # solo aplica si la acometida es aérea
            if sec_acom <= 10:  # hasta 10mm^2 va la cuna chica
                desc_granpa = "Granpa de retención tipo cuña 2,5-25mm^2 (6-10mm)"  # cuna para cable delgado
            else:  # sobre 10mm^2
                desc_granpa = "Granpa de retención tipo cuña 16-25mm^2 (10-15mm)"  # secciones mayores: cuna para cable más grueso
            add_row(
                desc=desc_granpa,  # cuna elegida según la sección de la acometida
                marcas_txt=marcas.get("Granpa de retención", ""),  # marcas de granpa de retención
                norma="-",  # ferreteria, sin norma
                circuito="Acometida y/o alimentador",  # va en el tramo aéreo
                unidad="u",  # por unidad
                k=1,  # sin mayoracion
                longitud_m="1 unid",  # siempre 1
                cantidad=1  # una sola granpa por acometida aérea
            )
        # Mordaza alimentador
        if "aer" in str(tipo_alimentador).strip().lower():  # solo si el alimentador es aéreo
            add_row(
                desc="Mordaza para alimentador aéreo",  # mordaza para tensar el alimentador aéreo
                marcas_txt=marcas.get("Mordaza acometida", ""),  # reusa las marcas de mordaza de acometida
                norma="SEC",  # exigencia SEC
                circuito="Alimentador",  # solo del alimentador
                unidad="u",  # por unidad
                k=1,  # sin mayoracion
                longitud_m="1 unid",  # siempre 1
                cantidad=1  # una sola mordaza
            )
        # Conector HUB — total según combinación acometida/alimentador (tabla
        # fija), y de ese total, cuántos corresponden específicamente al tramo
        # del ALIMENTADOR (el resto es acometida + PT1 + PT2). Si el diámetro
        # del alimentador difiere del de la acometida (aéreo o subterráneo,
        # da igual), van en 2 filas separadas (una por diámetro); si no, se
        # fusiona en 1 sola fila.
        _acom_aer_hub = "aer" in str(tipo_acometida).strip().lower()  # true si la acometida es aérea (se vuelve a chequear acá para el HUB)
        _alim_aer_hub = "aer" in str(tipo_alimentador).strip().lower()  # true si el alimentador es aéreo
        # Tabla fija: según si acometida/alimentador son aéreos o subterráneos,
        # cuántos conectores HUB van en total y cuántos de esos son del alimentador.
        # los números están escritos a mano (no salen de ninguna tabla del RIC):
        # cuentan las entradas y salidas de tubo galvanizado en las cajas
        # metálicas del empalme, la de derivación y las dos puestas a tierra.
        # cada tramo que pasa de aéreo a subterráneo agrega una caja más, y por
        # eso agrega hub. cantidad_hub_total también manda en el sellador de
        # roscas de más abajo (1 frasco cada 6 hub).
        if _acom_aer_hub and _alim_aer_hub:  # los dos tramos aéreos
            cantidad_hub_total, cantidad_hub_alim = 6, 1  # los dos aéreos: 6 hub en total, 1 es del alimentador
        elif _acom_aer_hub and not _alim_aer_hub:  # acometida aérea y alimentador subterráneo
            cantidad_hub_total, cantidad_hub_alim = 7, 2  # acometida aérea y alimentador subterráneo: 7 en total, 2 del alimentador
        elif (not _acom_aer_hub) and _alim_aer_hub:  # acometida subterránea y alimentador aéreo
            cantidad_hub_total, cantidad_hub_alim = 7, 1  # acometida subterránea y alimentador aéreo: 7 en total, 1 del alimentador
        else:  # los dos tramos subterráneos
            cantidad_hub_total, cantidad_hub_alim = 8, 2  # los dos subterráneos: 8 en total, 2 del alimentador
        cantidad_hub_acom = cantidad_hub_total - cantidad_hub_alim  # el resto (acometida + PT1 + PT2)

        # Da la descripción del conector HUB según el diámetro que aparece
        # en el texto del tubo (25/32/40/50mm)
        def _desc_hub_de(diam_txt):  # nombre del conector HUB según el diámetro del tubo
            if "25mm" in diam_txt:  # tubo de 25mm
                return "Conector HUB 25mm de acero galvanizado"  # hub de 25mm
            elif "32mm" in diam_txt:  # tubo de 32mm
                return "Conector HUB 32mm de acero galvanizado"  # hub de 32mm
            elif "40mm" in diam_txt:  # tubo de 40mm
                return "Conector HUB 40mm de acero galvanizado"  # hub de 40mm
            elif "50mm" in diam_txt:  # tubo de 50mm
                return "Conector HUB 50mm de acero galvanizado"  # hub de 50mm
            return ""  # diámetro que no calza: no se agrega hub

        if desc_tubo_alim and desc_tubo_alim != desc_tubo_acom:  # hay tubo de alimentador y con distinto diámetro que el de la acometida
            # Diámetros distintos (alimentador aéreo o subterráneo, da
            # igual): 2 filas separadas
            _hub_acom_desc = _desc_hub_de(desc_tubo_acom)  # hub del lado de la acometida, con su propio diámetro
            if _hub_acom_desc and cantidad_hub_acom:  # solo si el diámetro calzo y quedaron hub que cobrar
                add_row(
                    desc=_hub_acom_desc,  # hub del diámetro de la acometida
                    marcas_txt=marcas.get("Conector HUB", ""),  # marcas de conector HUB
                    norma="SEC",  # exigencia SEC
                    circuito="Acometida, puesta a tierra 1, puesta a tierra 2",  # estos hub son de acometida y de las dos tierras
                    unidad="u",  # por unidad
                    k=1,  # sin mayoracion
                    longitud_m=f"{cantidad_hub_acom} unid",  # se muestra la cantidad
                    cantidad=cantidad_hub_acom  # total menos los que se van al alimentador
                )
            _hub_alim_desc = _desc_hub_de(desc_tubo_alim)  # hub del lado del alimentador, con el otro diámetro
            if _hub_alim_desc and cantidad_hub_alim:  # solo si calzo y hay cantidad
                add_row(
                    desc=_hub_alim_desc,  # hub del diámetro del alimentador
                    marcas_txt=marcas.get("Conector HUB", ""),  # marcas de conector HUB
                    norma="SEC",  # exigencia SEC
                    circuito="Alimentador",  # solo del tramo alimentador
                    unidad="u",  # por unidad
                    k=1,  # sin mayoracion
                    longitud_m=f"{cantidad_hub_alim} unid",  # se muestra la cantidad
                    cantidad=cantidad_hub_alim  # los que le tocan al alimentador según la tabla de arriba
                )
        else:  # mismo diámetro en los dos tramos
            # Mismo diámetro (acometida y alimentador coinciden): se
            # mantiene fusionado en 1 sola fila
            _hub_desc = _desc_hub_de(desc_tubo_acom)  # un solo diámetro para todos los hub
            if _hub_desc and cantidad_hub_total:  # solo si calzo el diámetro y hay hub que cobrar
                add_row(
                    desc=_hub_desc,  # hub del diámetro común
                    marcas_txt=marcas.get("Conector HUB", ""),  # marcas de conector HUB
                    norma="SEC",  # exigencia SEC
                    circuito="Acometida, puesta a tierra 1, puesta a tierra 2",  # van juntos acometida y las dos tierras
                    unidad="u",  # por unidad
                    k=1,  # sin mayoracion
                    longitud_m=f"{cantidad_hub_total} unid",  # se muestra la cantidad
                    cantidad=cantidad_hub_total  # el total completo de la tabla de arriba
                )

        # Terminal PVC conduit con 2 tuercas — se separa en acometida y
        # alimentador, cada uno con su propia cantidad y diámetro (son 2
        # tramos de PVC enterrado distintos, no comparten medida):
        #  - Acometida subterránea: 1 terminal, con la Tabla N°4.29 (2
        #    conductores, F+N sin tierra) — este es el diámetro del ducto
        #    PVC real, DISTINTO al tubo de acero galvanizado (que usa la
        #    tabla simple) — si es aérea, no aplica (0).
        #  - Alimentador: 4 terminales si es subterráneo (diámetro real del
        #    ducto, Tabla N°4.29 con 3 conductores) o 3 si es aéreo (mismo
        #    piso mínimo de 32mm que ya usa el tubo galvanizado del
        #    alimentador, sin ser un valor fijo).
        # Si ambos diámetros coinciden, se fusionan en 1 fila; si no, 2.
        import re  # se usa para sacar el número de mm desde el texto del diámetro
        # Saca el número de milímetros de un texto tipo "32mm" (sirve para
        # comparar diámetros y armar la descripción del terminal)
        def _extraer_diam_num(txt):
            m = re.search(r'(\d+)\s*mm', str(txt))  # busca el primer NNmm que aparezca en el texto
            return int(m.group(1)) if m else None  # devuelve el número de mm, o None si no encontró nada

        _es_acom_sub_term = "sub" in str(tipo_acometida).strip().lower()  # acometida subterránea: lleva terminal PVC
        _es_alim_aer_term = "aer" in str(tipo_alimentador).strip().lower()  # alimentador aéreo: cambia la cantidad de terminales

        qty_term_acom = 1 if _es_acom_sub_term else 0  # 1 terminal solo si la acometida es subterránea
        qty_term_alim = 3 if _es_alim_aer_term else 4  # 3 si el alimentador es aéreo, 4 si es subterráneo

        # Diámetro del terminal de la acometida (Tabla N°4.29, 2 conductores)
        diam_term_acom = ducto_nominal_tablas(sec_acom, 2, "subterraneo") if qty_term_acom else None  # None si la acometida es aérea, ahí no va terminal
        # el diámetro del terminal del alimentador se saca distinto
        # según sea aéreo o subterráneo
        if qty_term_alim:  # si hay terminales de alimentador hay que definir su diámetro
            # Aéreo: usa el mismo diámetro (piso 32mm) que el tubo galvanizado
            # Subterráneo: saca el diámetro real de la canalización
            diam_term_alim = _extraer_diam_num(desc_tubo_alim) if _es_alim_aer_term else _extraer_diam_num(canalizacion_txt)
        else:  # sin terminales de alimentador no hay nada que medir
            diam_term_alim = None  # sin terminales de alimentador no hay diámetro que poner

        if diam_term_acom and diam_term_alim and diam_term_acom == diam_term_alim:  # los dos tramos con el mismo diámetro de ducto
            # Mismo diámetro: se fusiona todo en 1 sola fila
            add_row(  # una sola fila con los terminales de los dos tramos
                desc=f"Terminal PVC conduit con 2 tuercas {diam_term_acom} mm",  # diámetro común a los dos
                marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas de terminal PVC conduit
                norma="RIC 4.7.2",  # RIC 4.7.2: terminacion de ductos
                circuito="Acometida y alimentador en caja empalme, entrada y salida de caja metálica, entrada al TDA",  # puntos donde entra y sale el ducto
                unidad="u",  # por unidad
                k=1,  # sin mayoracion
                longitud_m=f"{qty_term_acom + qty_term_alim} unid",  # se muestra la cantidad
                cantidad=qty_term_acom + qty_term_alim  # los de acometida más los del alimentador
            )
        else:  # los diámetros no coinciden
            # Diámetros distintos: 1 fila para la acometida y otra para el alimentador
            if diam_term_acom and qty_term_acom:  # fila propia de la acometida, si es subterránea
                add_row(
                    desc=f"Terminal PVC conduit con 2 tuercas {diam_term_acom} mm",  # diámetro del ducto PVC de la acometida
                    marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas de terminal PVC conduit
                    norma="RIC 4.7.2",  # RIC 4.7.2
                    circuito="Acometida (tramo PVC subterráneo)",  # solo el tramo PVC enterrado de la acometida
                    unidad="u",  # por unidad
                    k=1,  # sin mayoracion
                    longitud_m=f"{qty_term_acom} unid",  # se muestra la cantidad
                    cantidad=qty_term_acom  # 1 terminal
                )
            if diam_term_alim and qty_term_alim:  # fila propia del alimentador
                add_row(
                    desc=f"Terminal PVC conduit con 2 tuercas {diam_term_alim} mm",  # diámetro del ducto del alimentador
                    marcas_txt=marcas.get("Terminal PVC conduit con 2 tuercas", ""),  # marcas de terminal PVC conduit
                    norma="RIC 4.7.2",  # RIC 4.7.2
                    circuito="Alimentador en caja empalme, entrada y salida de caja metálica, entrada al TDA",  # entrada y salida de caja metálica y entrada al TDA
                    unidad="u",  # por unidad
                    k=1,  # sin mayoracion
                    longitud_m=f"{qty_term_alim} unid",  # se muestra la cantidad
                    cantidad=qty_term_alim  # 3 si el alimentador es aéreo, 4 si es subterráneo
                )

        # =========================
        # PUESTA A TIERRA
        # =========================
        # A partir de acá se calculan los materiales del circuito de puesta
        # a tierra: conductores, barras copperweld, camarillas, etc.
        #Conductor thwn-2 puesta a tierra empalme
        sec_pt = "4" # valor por defecto: 4mm^2 es la sección mínima de puesta a tierra
        try:  # si el texto del alimentador no viene como 3xNmm2, se cae al except
            # La sección del conductor de puesta a tierra se elige según la
            # sección del alimentador (a mayor sección del alimentador, mayor
            # sección de tierra requerida)
            sec_base = float(alim_txt.split("3x")[1].replace("mm^2", "").strip())  # saca la sección del alimentador desde el texto tipo 3x10mm^2
            if sec_base <= 4:  # alimentador hasta 4mm^2
                sec_pt = "4"  # tierra de 4mm^2
            elif sec_base <= 6:  # alimentador hasta 6mm^2
                sec_pt = "6"  # tierra de 6mm^2
            elif sec_base <= 10:  # alimentador hasta 10mm^2
                sec_pt = "10"  # tierra de 10mm^2
            elif sec_base <= 16:  # alimentador hasta 16mm^2
                sec_pt = "16"  # tierra de 16mm^2
            elif sec_base <= 25:  # alimentador hasta 25mm^2
                sec_pt = "25"  # tierra de 25mm^2
            else:  # sobre 25mm^2
                sec_pt = "25"  # sobre 25mm^2 se deja en 25, no sube más
        except:  # texto raro o vacío
            sec_pt = "4"  # si no se pudo leer la sección, usa el mínimo
        # +2m: 1 chicote en la camarilla N°1 + 1 chicote en la caja metálica del empalme
        metros_pt = math.ceil(float(dist_empalme_pt1)) + 2  # metros de conductor para el tramo empalme - camarilla N°1
        # Conductor blanco (neutro) de puesta a tierra N°1, entre el
        # empalme y la camarilla N°1
        add_row(
            desc=f"Conductor THWN-2 {sec_pt}mm^2 blanco",  # blanco por norma para el neutro del empalme
            marcas_txt=marcas.get("Conductor THWN-2 blanco", ""),  # marcas de THWN-2 blanco
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # tramo empalme - camarilla N°1
            unidad="m",  # se compra por metro
            k=1,  # sin mayoracion
            longitud_m=f"{metros_pt} m",  # metros con los 2 de chicote incluidos
            cantidad=metros_pt  # misma cantidad en metros
        )
        #Conductor thwn-2 puesta a tierra empalme
        # Conductor verde (tierra) del mismo tramo, misma cantidad de metros
        add_row(
            desc=f"Conductor THWN-2 {sec_pt}mm^2 verde",  # verde por norma para la tierra
            marcas_txt=marcas.get("Conductor THWN-2 verde", ""),  # marcas de THWN-2 verde
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # mismo tramo que el blanco
            unidad="m",  # por metro
            k=1,  # sin mayoracion
            longitud_m=f"{metros_pt} m",  # mismos metros que el blanco
            cantidad=metros_pt  # misma cantidad
        )

        # Barra copperweld PT1 (empalme - camarilla N°1)
        add_row(
            desc=desc_barra_pt,  # largo de barra ya elegido más arriba según la resistividad
            marcas_txt=marcas.get("Barra copperweld", ""),  # marcas de barra copperweld
            norma="RIC 6 (8.3.2, 8.5, 8.6, Tabla 6.1)",  # RIC 6: puesta a tierra de servicio
            circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # esta es la tierra del empalme
            unidad="u",  # por unidad
            k=1,  # sin mayoracion
            longitud_m=f"{n_barras_pt1} unid",  # se muestra la cantidad de barras
            cantidad=n_barras_pt1  # barras que necesita PT1
        )
        # Camarilla PT1 según tubo conduit galvanizado
        diam_cam = ""  # queda vacío si el diámetro del tubo no calza con ninguna medida estándar
        if "25mm" in desc_tubo_acom:  # tubo de 25mm
            diam_cam = "25mm"  # camarilla de 25mm
        elif "32mm" in desc_tubo_acom:  # tubo de 32mm
            diam_cam = "32mm"  # camarilla de 32mm
        elif "40mm" in desc_tubo_acom:  # tubo de 40mm
            diam_cam = "40mm"  # camarilla de 40mm
        elif "50mm" in desc_tubo_acom:  # tubo de 50mm
            diam_cam = "50mm"  # camarilla de 50mm
        # RIC N°6 art. 8.3.2: si hay más de 1 barra, hay que unirlas con
        # un conductor desnudo de cobre, mínimo 16mm^2 de sección (fijo,
        # ese mínimo no depende de la sección del alimentador)
        if n_barras_pt1 > 1:
            add_row(
                desc=f"Conductor desnudo Cu 16mm^2 (unión entre barras PT1)",  # 16mm^2 fijo, es el mínimo del RIC para unir barras
                marcas_txt=marcas.get("Conductor desnudo Cu", ""),  # marcas de conductor desnudo de cobre
                norma="RIC 6 (8.3.2, 8.7, 8.9)",  # RIC 6: unión entre barras
                circuito="Puesta a tierra N°1 (empalme - camarilla N°1)",  # unión de las barras de PT1
                unidad="m",  # por metro
                k=1,  # sin mayoracion
                longitud_m=f"{long_cond_desnudo_pt1:.1f} m",  # metros del cable de unión, con 1 decimal
                cantidad=long_cond_desnudo_pt1  # largo calculado según la separación entre barras
            )
        # Barra copperweld PT2 (TDA - camarilla N°2)
        add_row(
            desc=desc_barra_pt,  # misma barra que en PT1
            marcas_txt=marcas.get("Barra copperweld", ""),  # marcas de barra copperweld
            norma="RIC 6 (8.3.2, 8.5, 8.6, Tabla 6.1)",  # RIC 6: puesta a tierra de protección
            circuito="Puesta a tierra N°2 (TDA - camarilla N°2)",  # esta es la tierra del tablero
            unidad="u",  # por unidad
            k=1,  # sin mayoracion
            longitud_m=f"{n_barras_pt2} unid",  # se muestra la cantidad de barras
            cantidad=n_barras_pt2  # barras que necesita PT2
        )
        # Camarilla PT1 + PT2: mismo diámetro siempre (diam_cam es único,
        # compartido entre ambas), así que se fusionan en 1 sola fila en
        # vez de 2 filas idénticas repetidas.
        if diam_cam:  # si el diámetro no calzo con ninguna medida, no se agrega camarilla
            add_row(
                desc=f"Camarilla de registro con tapa PVC naranjo 160 x {diam_cam}",  # camarilla del mismo diámetro del tubo conduit
                marcas_txt=marcas.get("Camarilla PVC naranjo", ""),  # marcas de camarilla PVC naranjo
                norma="RIC 6 (5.15)",  # RIC 6 (5.15): registro de la puesta a tierra
                circuito="Puesta a tierra N°1 y N°2 (empalme - camarilla N°1 / TDA - camarilla N°2)",  # una fila que cubre las camarillas de PT1 y PT2
                unidad="u",  # por unidad
                k=1,  # sin mayoracion
                longitud_m=f"{n_barras_pt1 + n_barras_pt2} unid",  # se muestra la cantidad total
                cantidad=n_barras_pt1 + n_barras_pt2  # una camarilla por cada barra de las dos tierras
            )
        # RIC N°6 art. 8.3.2: mismo criterio que PT1, mínimo 16mm^2 fijo
        # si hay más de 1 barra (esta es la puesta a tierra de PROTECCIÓN,
        # la del tablero, distinta de la de SERVICIO que es PT1)
        if n_barras_pt2 > 1:
            add_row(  # fila de materiales: cable desnudo que une las dos barras de tierra
                desc=f"Conductor desnudo Cu 16mm^2 (unión entre barras PT2)",  # texto que aparece en el informe
                marcas_txt=marcas.get("Conductor desnudo Cu", ""),  # marcas sugeridas para ese material
                norma="RIC 6 (8.3.2, 8.7, 8.9)",  # artículo del RIC que respalda la unión entre barras
                circuito="Puesta a tierra N°2 (TDA - camarilla N°2)",  # tramo al que pertenece el material
                unidad="m",  # se compra por metro
                k=1,  # sin holgura extra
                longitud_m=f"{long_cond_desnudo_pt2:.1f} m",  # largo del tramo con un decimal
                cantidad=long_cond_desnudo_pt2  # metros que van a la lista de compra
            )

        #Abrazaderas tipo caddy — 3 por cada tubo, con el tamaño de SU
        # PROPIO diámetro (acometida+PT vs. alimentador pueden diferir).
        # Se fusiona en 1 fila si ambos diámetros coinciden.
        # Función interna: recibe el texto del diámetro del tubo (ej. "32mm")
        # y devuelve el nombre del producto abrazadera caddy de ese tamaño.
        def _desc_caddy_de(diam_txt):
            if "25mm" in diam_txt:  # tubo de 25mm
                return "Abrazadera tipo caddy 25mm"  # abrazadera caddy de esa medida
            elif "32mm" in diam_txt:  # tubo de 32mm
                return "Abrazadera tipo caddy 32mm"  # abrazadera caddy de esa medida
            elif "40mm" in diam_txt:  # tubo de 40mm
                return "Abrazadera tipo caddy 40mm"  # abrazadera caddy de esa medida
            elif "50mm" in diam_txt:  # tubo de 50mm
                return "Abrazadera tipo caddy 50mm"  # abrazadera caddy de esa medida
            return ""  # diámetro que no está en la lista: no se agrega abrazadera

        cantidad_caddy_acom = cantidad_total_acom_pt * 3  # 3 abrazaderas por cada tubo de 3m del grupo acometida + PT
        cantidad_caddy_alim = cantidad_tramo_alim * 3  # 3 abrazaderas por cada tramo del alimentador
        cantidad_caddy = cantidad_caddy_acom + cantidad_caddy_alim  # total, usado más abajo en tornillos de fijación

        if desc_tubo_acom and desc_tubo_alim and desc_tubo_acom == desc_tubo_alim:  # si acometida+PT y alimentador usan el mismo diámetro de tubo
            # Mismo diámetro: 1 sola fila con el total
            _desc_caddy = _desc_caddy_de(desc_tubo_acom)  # nombre de la abrazadera de ese diámetro
            if _desc_caddy and cantidad_caddy:  # solo si el diámetro se reconoció y hay abrazaderas que contar
                add_row(  # fila de materiales: todas las abrazaderas caddy juntas
                    desc=_desc_caddy,  # nombre de la abrazadera según el diámetro
                    marcas_txt=marcas.get("Abrazadera tipo caddy", ""),  # marcas sugeridas
                    norma="SEC",  # no viene de una tabla del RIC, es criterio SEC
                    circuito="Puesta a tierra 1 y 2, acometida, alimentador",  # cubre los dos tramos de tierra más acometida y alimentador
                    unidad="u",  # se compran por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cantidad_caddy} unid",  # texto con el total de abrazaderas
                    cantidad=cantidad_caddy  # cantidad final
                )
        else:  # los tubos no son del mismo diámetro
            # Diámetros distintos: 2 filas separadas, una por cada tramo
            # (acometida y alimentador siempre tienen su propio diámetro)
            _desc_caddy_acom = _desc_caddy_de(desc_tubo_acom)  # abrazadera del tubo de acometida + PT
            if _desc_caddy_acom and cantidad_caddy_acom:  # solo si el diámetro se reconoció y hay tramos que fijar
                add_row(  # fila de materiales: abrazaderas del tramo acometida
                    desc=_desc_caddy_acom,  # nombre según el diámetro del tubo de acometida
                    marcas_txt=marcas.get("Abrazadera tipo caddy", ""),  # marcas sugeridas
                    norma="SEC",  # criterio SEC
                    circuito="Puesta a tierra 1 y 2, acometida",  # van en las tierras y en la acometida
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cantidad_caddy_acom} unid",  # texto con la cantidad
                    cantidad=cantidad_caddy_acom  # cantidad de abrazaderas de este tramo
                )
            _desc_caddy_alim = _desc_caddy_de(desc_tubo_alim)  # abrazadera del tubo del alimentador
            if _desc_caddy_alim and cantidad_caddy_alim:  # solo si hay diámetro reconocido y tramos que fijar
                add_row(  # fila de materiales: abrazaderas del tramo alimentador
                    desc=_desc_caddy_alim,  # nombre según el diámetro del tubo del alimentador
                    marcas_txt=marcas.get("Abrazadera tipo caddy", ""),  # marcas sugeridas
                    norma="SEC",  # criterio SEC
                    circuito="Alimentador",  # estas van solo en el alimentador
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cantidad_caddy_alim} unid",  # texto con la cantidad
                    cantidad=cantidad_caddy_alim  # cantidad de abrazaderas del alimentador
                )

        #Sellador de roscas: 1 frasco de 50ml alcanza para 6 conectores HUB
        cantidad_sellador = int(math.ceil(cantidad_hub_total / 6))  # frascos necesarios, redondeando hacia arriba
        add_row(  # fila de materiales: sellador de roscas
            desc="Sellador de roscas con teflón 50ml",  # el frasco de 50ml
            marcas_txt=marcas.get("Sellador de roscas", ""),  # marcas sugeridas
            norma="SEC",  # criterio SEC
            circuito="Conectores Hub(cajas metálicas empalme y derivación)",  # se usa al enroscar los HUB en las cajas metálicas
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m=f"{cantidad_sellador} unid",  # texto con los frascos
            cantidad=cantidad_sellador  # frascos que van a la compra
        )

        # Terminal ferrul acometida
        # Terminal ferrul acometida: mismo color/sección que el tablero (color_ferrul_por_seccion)
        # separa el texto de la acometida (ej. "2x10mm^2") y saca el número de sección;
        # si no logra leerlo, usa 4.0 mm^2 como valor por defecto
        try:  # puede fallar si el texto no viene en el formato esperado
            sec_acom = float(acometida_txt.split("2x")[1].replace("mm^2", "").strip())  # saca la sección del texto de la acometida (viene como 2x10mm^2)
        except:  # no se pudo leer la sección
            sec_acom = 4.0  # sección por defecto para la acometida
        color_ferrul = color_ferrul_por_seccion(sec_acom)  # color de ferrul según la sección (norma de colores)
        sec_ferrul = str(int(sec_acom)) if sec_acom == int(sec_acom) else texto_seccion(sec_acom)  # sección como texto, sin decimales si es entera
        add_row(  # fila de materiales: ferrules de la acometida
            desc=f"Terminal ferrul color {color_ferrul} {sec_ferrul}mm (acometida)",  # el color depende de la sección
            marcas_txt=marcas.get("Terminal ferrul acometida", ""),  # marcas sugeridas
            norma="SEC",  # criterio SEC
            circuito="Acometida (Fase y Neutro)",  # van uno en la fase y otro en el neutro
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="2 unid",  # siempre 2: fase y neutro
            cantidad=2  # cantidad fija
        )

        # Terminal ferrul alimentador: mismo color/sección que el tablero
        # mismo criterio de arriba, pero leyendo el texto del alimentador
        # (ej. "3x10mm^2"); si falla, también usa 4.0 mm^2 por defecto
        try:  # puede fallar si el texto del alimentador viene distinto
            sec_alim = float(alim_txt.split("3x")[1].replace("mm^2", "").strip())  # saca la sección del texto del alimentador (viene como 3x10mm^2)
        except:  # no se pudo leer la sección
            sec_alim = 4.0  # sección por defecto para el alimentador
        color_ferrul_alim = color_ferrul_por_seccion(sec_alim)  # color de ferrul según la sección del alimentador
        sec_ferrul_alim = str(int(sec_alim)) if sec_alim == int(sec_alim) else texto_seccion(sec_alim)  # sección como texto
        add_row(  # fila de materiales: ferrules del alimentador
            desc=f"Terminal ferrul color {color_ferrul_alim} {sec_ferrul_alim}mm (alimentador)",  # color y sección del alimentador
            marcas_txt=marcas.get("Terminal ferrul alimentador", ""),  # marcas sugeridas
            norma="SEC",  # criterio SEC
            circuito="Alimentador (Salida de medidor fase, entrada y salida termomagnético empalme)",  # salida del medidor más entrada y salida del TM del empalme
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="3 unid",  # 3 puntos donde se pone ferrul
            cantidad=3  # cantidad fija
        )

        #terminal ferrul doble alimentador
        add_row(  # fila de materiales: ferrul doble del neutro
            desc=f"Terminal ferrul doble color {color_ferrul_alim} {sec_ferrul_alim}mm (alimentador)",  # ferrul doble, entran dos conductores en el mismo terminal
            marcas_txt=marcas.get("Terminal ferrul doble alimentador", ""),  # marcas sugeridas
            norma="SEC",  # criterio SEC
            circuito="Salida de medidor (Neutro del alimentador y neutro aterrizado)",  # en la salida del medidor se juntan neutro del alimentador y neutro aterrizado
            unidad="u",  # por unidad
            k=1,  # sin holgura
            longitud_m="1 unid",  # basta con uno
            cantidad=1  # cantidad fija
        )

        # Cajas metálicas de derivación por ubicación (se usa acá y en la fijación del caddy del poste)
        # esto es específico del empalme "independiente": puede haber caja en
        # el poste (si algún tramo es subterráneo) además de la de la fachada
        if "aer" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower():  # todo aéreo: la única caja de derivación queda en la fachada
            cajas_poste = 0  # no se monta nada en el poste
            cajas_fachada = 1  # la caja de la fachada siempre va
        elif "aer" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower():  # acometida aérea y alimentador subterráneo
            cajas_poste = 1  # el cambio de aéreo a subterráneo pide caja en el poste
            cajas_fachada = 1  # más la de la fachada
        elif "sub" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower():  # acometida subterránea y alimentador aéreo
            cajas_poste = 1  # caja en el poste por el cambio de canalización
            cajas_fachada = 1  # más la de la fachada
        elif "sub" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower():  # los dos tramos subterráneos
            cajas_poste = 2  # una caja por cada bajada del poste
            cajas_fachada = 1  # más la de la fachada
        else:  # tipo de canalización no reconocido
            cajas_poste = 0  # sin cajas, para no inventar material
            cajas_fachada = 0  # sin cajas, para no inventar material

        # Tornillos según material y ubicación (cajas_poste/cajas_fachada ya calculadas más arriba)
        # - Fachada/casa: abrazaderas PVC del alimentador + cajas metálicas en fachada
        # - Poste: cajas metálicas en poste + abrazaderas PVC en poste
        # Diámetro real del ducto del alimentador, leido del texto
        # canalizacion_txt que se arma FUERA de esta función (línea ~12043)
        # como "Ø PVC Conduit NN mm", con ducto_nominal_tablas(S_alim, 3
        # conductores, tipo_alimentador). en empalme independiente el
        # alimentador solo puede ser aéreo o subterráneo (el input prohíbe
        # "en ducto", línea ~10735), así que:
        #   - subterráneo -> NN sale de la Tabla N°4.29
        #   - aéreo       -> canalizacion_txt queda en "-", el regex no calza
        #                    y se cae al piso de 32mm
        # La separación entre abrazaderas sigue la Tabla N°4.24 según ese
        # diámetro. Con el piso de 32mm, en la práctica siempre queda 1,50m.
        _m_diam_conduit_tornillos = re.search(r'(\d+)\s*mm', str(canalizacion_txt))  # busca el diámetro en mm dentro del texto de la canalización
        _diam_conduit_tornillos = max(32, int(_m_diam_conduit_tornillos.group(1))) if _m_diam_conduit_tornillos else 32  # mínimo 32mm si no se pudo leer
        _sep_abraz_tornillos = 1.20 if _diam_conduit_tornillos <= 25 else 1.50  # separación entre abrazaderas según Tabla N 4.24
        # Abrazaderas PVC por ubicación
        if "aer" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower():  # los dos tramos aéreos
            abrazaderas_poste = 0  # no se fija nada al poste
            abrazaderas_fachada = int(math.ceil(float(longitud_llegada_aerea_tda) / _sep_abraz_tornillos))  # abrazaderas repartidas en la llegada aérea al TDA
        elif "aer" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower():  # acometida aérea, alimentador subterráneo
            abrazaderas_poste = 1  # una abrazadera en la bajada del poste
            abrazaderas_fachada = int(math.ceil(float(longitud_abrazaderas_alimentador) / _sep_abraz_tornillos))  # abrazaderas del tramo visible del alimentador
        elif "sub" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower():  # acometida subterránea, alimentador aéreo
            abrazaderas_poste = 1  # una abrazadera en la bajada del poste
            abrazaderas_fachada = int(math.ceil(float(longitud_llegada_aerea_tda) / _sep_abraz_tornillos))  # abrazaderas de la llegada aérea al TDA
        elif "sub" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower():  # los dos tramos subterráneos
            abrazaderas_poste = 2  # dos bajadas por el poste, una abrazadera cada una
            abrazaderas_fachada = int(math.ceil(float(longitud_abrazaderas_alimentador) / _sep_abraz_tornillos))  # abrazaderas del tramo visible del alimentador
        else:  # canalización no reconocida
            abrazaderas_poste = 0  # sin abrazaderas para no inventar material
            abrazaderas_fachada = 0  # sin abrazaderas para no inventar material
        # Tornillos fachada/casa: cuenta SOLO lo que se atornilla contra la
        # casa, con el material del forrado exterior. Criterio de conteo:
        #   2 tornillos por abrazadera de conduit PVC en fachada
        # + 4 tornillos por cada caja de derivación metálica en fachada
        # + 1 tornillo por cada abrazadera caddy del tramo PT2 (del TDA a la
        #   camarilla N°2), que físicamente va pegada a la casa y no al poste.
        # el total se usa recién al final del bloque, en cant_grupo_B_total,
        # que es donde se elige el tipo de tornillo según material_forrado_exterior.
        cant_tornillos_fachada = int((abrazaderas_fachada * 2) + (cajas_fachada * 4) + cantidad_caddy_pt2)  # 2 tornillos por abrazadera, 4 por caja y 1 por caddy de PT2

        # Grupo A: caja metálica del empalme (fijo, siempre 6), según el
        # material del POSTE (madera/metálico) — con golilla solo acá.
        cant_grupo_A_poste = 6  # la caja del empalme siempre se fija con 6 tornillos
        if "madera" in str(tipo_poste).strip().lower():  # poste de madera: tirafondo
            add_row(  # fila de materiales: tirafondos de la caja del empalme
                desc='Tirafondo hexagonal 1/4" x 1 1/2" + golilla 1/4"',  # tirafondo con golilla, para madera
                marcas_txt=marcas.get("Tirafondo hexagonal madera", ""),  # marcas sugeridas
                norma="-",  # no aplica norma, es criterio de montaje
                circuito="Caja metálica del empalme",  # solo para fijar la caja del empalme
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_grupo_A_poste} unid",  # texto con los 6 tornillos
                cantidad=cant_grupo_A_poste  # cantidad fija
            )
        elif "metal" in str(tipo_poste).strip().lower():  # poste metálico: tornillo con broca
            add_row(  # fila de materiales: autoperforantes de la caja del empalme
                desc='Tornillo autoperforante 1/4" x 1 1/2" + golilla 1/4"',  # autoperforante con golilla, atraviesa el metal
                marcas_txt=marcas.get("Tornillo autoperforante broca", ""),  # marcas sugeridas
                norma="-",  # no aplica norma
                circuito="Caja metálica del empalme",  # solo para fijar la caja del empalme
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{cant_grupo_A_poste} unid",  # texto con los 6 tornillos
                cantidad=cant_grupo_A_poste  # cantidad fija
            )

        # Grupo B (poste): abrazaderas caddy del poste + caja(s) de derivación
        # del poste + abrazaderas PVC del poste — usa el material del POSTE
        # (tipo_poste), NO el de la fachada, ya que físicamente van montadas
        # sobre el poste (mismo criterio de textos que fachada, pero
        # seleccionado por tipo_poste en vez de material_forrado_exterior).
        # Criterio de conteo del grupo B del poste, espejo del de fachada:
        #   1 tornillo por abrazadera caddy que quede en el poste, o sea todas
        #     (cantidad_caddy) menos las del tramo PT2 que ya se cobraron en
        #     cant_tornillos_fachada
        # + 4 tornillos por cada caja de derivación metálica montada en el poste
        # + 2 tornillos por cada abrazadera de conduit PVC del poste
        # así cada punto de fijación queda cobrado una sola vez, y con el
        # tornillo que corresponde al material donde se atornilla.
        cant_grupo_B_poste = (1 * (cantidad_caddy - cantidad_caddy_pt2)) + (4 * cajas_poste) + (2 * abrazaderas_poste)  # cuenta lo que falta fijar en el poste: caddys que no son del tramo PT2, cajas de derivación (4 c/u) y abrazaderas (2 c/u)
        # si hay algo que fijar en el poste, agrega los tornillos correspondientes
        if cant_grupo_B_poste > 0:
            # poste de madera: tornillo punta fina
            if "madera" in str(tipo_poste).strip().lower():
            # fila de materiales: tornillos para poste de madera
                add_row(
                    desc='Tornillo 8x1" cabeza lenteja punta fina',  # tornillo punta fina, agarra en madera
                    marcas_txt=marcas.get("Tornillo punta fina madera", ""),  # marcas sugeridas
                    norma="-",  # no aplica norma
                    circuito="Poste: abrazaderas caddy, caja(s) de derivación metálica y abrazaderas PVC",  # todo lo que se atornilla al poste
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cant_grupo_B_poste} unid",  # texto con la cantidad del grupo B
                    cantidad=cant_grupo_B_poste  # tornillos que van a la compra
                )
            # poste metálico: tornillo autoperforante (con broca, para atravesar el metal)
            elif "metal" in str(tipo_poste).strip().lower():
            # fila de materiales: tornillos para poste metálico
                add_row(
                    desc='Tornillo cabeza lenteja 8x1 1/4" punta broca',  # punta broca para perforar el poste metálico
                    marcas_txt=marcas.get("Tornillo autoperforante broca", ""),  # marcas sugeridas
                    norma="-",  # no aplica norma
                    circuito="Poste: abrazaderas caddy, caja(s) de derivación metálica y abrazaderas PVC",  # todo lo que se atornilla al poste
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{cant_grupo_B_poste} unid",  # texto con la cantidad del grupo B
                    cantidad=cant_grupo_B_poste  # tornillos que van a la compra
                )

        # cable que va desde el empalme hasta el tablero (TDA): 3 conductores de cobre (fase, neutro y tierra),
        # la sección S sale del cálculo del alimentador (res_alim) hecho antes
        #Alimentador
        # fila de materiales: metros de cable alimentador (se redondea hacia arriba)
        add_row(
            desc=f'Alimentador RV-K Cu 3x{round(float(res_alim["S"]),2)}mm^2',  # RV-K de 3 conductores con la sección calculada
            marcas_txt=marcas.get("Alimentador RV-K", ""),  # marcas sugeridas
            norma="SEC",  # criterio SEC
            circuito="Alimentador",  # es el tramo empalme - TDA
            unidad="m",  # se compra por metro
            k=1,  # sin holgura
            longitud_m=f"{math.ceil(float(longitud_alimentador))} m",  # metros redondeados hacia arriba
            cantidad=math.ceil(float(longitud_alimentador))  # metros que van a la compra
        )

        # Conduit PVC para tramo subterráneo solo de acometida en empalme.
        # Es un material de PVC, no de acero — usa Tabla N°4.29 (2
        # conductores, F+N), NO el diámetro de desc_tubo_acom (que es del
        # tubo de acero galvanizado y va con tabla simple aparte).
        diam_pvc_sub = None  # inicializar siempre antes de usar
        # solo corre este bloque si la acometida es subterránea
        if "sub" in str(tipo_acometida).strip().lower():
            diam_pvc_sub = ducto_nominal_tablas(sec_acom, 2, "subterraneo")  # diámetro del conduit según sección del conductor de acometida, 2 conductores (F+N), tramo subterráneo
            # si la tabla no entrega diámetro, no se agrega material
            if diam_pvc_sub:
                longitud_pvc_sub = (float(longitud_transformador_empalme) - float(longitud_subterraneo_medidor2))  # metros de PVC enterrado = tramo total menos la subida galvanizada al medidor
                cantidad_pvc_sub = int(math.ceil(longitud_pvc_sub / 3))  # el conduit viene en tiras de 3 mts, se redondea hacia arriba la cantidad de tiras
                # fila de materiales: conduit PVC para el tramo subterráneo de la acometida
                add_row(desc=f"Conduit PVC {diam_pvc_sub}mm, 3mts",
                    marcas_txt=marcas.get("Conduit PVC", ""),  # marcas sugeridas
                    norma="SEC",  # criterio SEC
                    circuito="Acometida subterránea",  # solo el tramo bajo tierra de la acometida
                    unidad="u",  # se vende por tira
                    k=1,  # sin holgura
                    longitud_m=f"{round(longitud_pvc_sub, 2)} m",  # largo real del tramo, para referencia
                    cantidad=cantidad_pvc_sub)  # tiras de 3mts que van a la compra

        # Abrazadera adicional para PVC subterráneo acometida, distinta medida que las del alimentador
        # abrazadera extra, de medida distinta a las del alimentador, solo si hubo conduit subterráneo en la acometida
        if "sub" in str(tipo_acometida).strip().lower() and diam_pvc_sub:
            # fila de materiales: abrazadera para el conduit PVC de la acometida subterránea
            add_row(desc=f"Abrazadera conduit PVC {diam_pvc_sub}mm",  # abrazadera del mismo diámetro que el conduit subterráneo
                marcas_txt=marcas.get("Abrazadera conduit PVC alimentador", ""),  # marcas sugeridas
                norma="SEC",  # criterio SEC
                circuito="Acometida subterránea",  # va en el tramo subterráneo de la acometida
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m="1 Unid",  # basta con una
                cantidad=1)  # cantidad fija

        # Conduit PVC según canalización alimentador (mínimo 32 mm)
        # Se define ANTES del bloque cámara para que diam_conduit esté disponible
        _m_diam_conduit2 = re.search(r'(\d+)\s*mm', str(canalizacion_txt))  # busca los milímetros (número + 'mm') dentro del texto de la canalización elegida
        diam_conduit = max(32, int(_m_diam_conduit2.group(1))) if _m_diam_conduit2 else 32  # diámetro del conduit del alimentador: nunca menos de 32mm (mínimo de norma)
        # Longitud conduit según tipo alimentador
        # tramo aéreo: se mide la llegada aérea al tablero
        if "aer" in str(tipo_alimentador).strip().lower():
            metros_cond = math.ceil(float(longitud_llegada_aerea_tda))  # metros del conduit de la llegada aérea al TDA
        # tramo subterráneo: se descuenta el pedazo que ya lleva su propio conduit (el de la acometida)
        elif "sub" in str(tipo_alimentador).strip().lower():
            metros_cond = math.ceil(float(longitud_alimentador) - float(longitud_subterraneo_medidor2))  # descuenta el tramo que ya va con su propio tubo
        # cualquier otro caso (en ducto): se usa el largo completo del alimentador
        else:
            metros_cond = math.ceil(float(longitud_alimentador))  # sin tramo aéreo ni subterráneo: todo el alimentador va en conduit

        # si hay metros de conduit que poner, calcula cuántas tiras de 3mts se necesitan
        if metros_cond > 0:
            cantidad_conduit_alim = int(math.ceil(metros_cond / 3))  # el conduit viene en tiras de 3mts
            # fila de materiales: conduit PVC para el alimentador
            add_row(
                desc=f'Conduit de PVC de {diam_conduit}mm, 3mts',  # conduit del diámetro calculado
                marcas_txt=marcas.get("Conduit PVC", ""),  # marcas sugeridas
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito="Alimentador",  # es del alimentador
                unidad="u",  # se vende por tira
                k=1,  # sin holgura
                longitud_m=f"{metros_cond} m",  # metros que se van a canalizar
                cantidad=cantidad_conduit_alim  # tiras que van a la compra
            )

        # separación máxima entre abrazaderas según el diámetro del conduit (Tabla N°4.24 del RIC)
        #Abrazaderas para canalización del alimentador según Tabla N°4.24
        sep_abraz_alim = 1.20 if diam_conduit <= 25 else 1.50  # Tabla N°4.24. como diam_conduit nunca baja de 32mm, en la práctica siempre queda 1.50 m
        # tramo aéreo: abrazaderas a lo largo de todo el conduit calculado arriba
        if "aer" in str(tipo_alimentador).strip().lower():
            longitud_abrazaderas = metros_cond  # las abrazaderas cubren los mismos metros del conduit
            cantidad_abrazaderas_alim = int(math.ceil(longitud_abrazaderas / sep_abraz_alim))  # cantidad de abrazaderas según la separación
            circ_abraz_alim = "Alimentador (tramo llegada aérea al TDA, poste)"  # texto del circuito para el informe
        # tramo subterráneo: abrazaderas a lo largo del tramo aéreo hasta el poste, más una extra en el poste
        elif "sub" in str(tipo_alimentador).strip().lower():
            longitud_abrazaderas = float(longitud_abrazaderas_alimentador)  # tramo visible del alimentador: desde donde sale de la tierra hasta el TDA
            cantidad_abrazaderas_alim = int(math.ceil(longitud_abrazaderas / sep_abraz_alim)) + 1  # +1 del poste
            circ_abraz_alim = "Alimentador (tramo salida subterránea al TDA, poste)"  # texto del circuito para el informe
        # en ducto: abrazaderas a lo largo de todo el conduit
        else:
            cantidad_abrazaderas_alim = int(math.ceil(metros_cond / sep_abraz_alim))  # abrazaderas a lo largo de todo el conduit
            circ_abraz_alim = "Alimentador (tramo en ducto)"  # texto del circuito para el informe
        # si se necesita al menos una abrazadera, agrega la fila de materiales
        if cantidad_abrazaderas_alim > 0:
            # fila de materiales: abrazaderas para el conduit del alimentador
            add_row(
                desc=f"Abrazadera conduit de PVC {diam_conduit}mm",  # abrazadera del mismo diámetro del conduit
                marcas_txt=marcas.get("Abrazadera conduit PVC alimentador", ""),  # marcas sugeridas
                norma="RIC 4 (Tabla N°4.24)",  # tabla de separación entre abrazaderas
                circuito=circ_abraz_alim,  # texto armado según el tipo de tramo
                unidad="u",  # por unidad
                k=1,  # sin holgura
                longitud_m=f"{cantidad_abrazaderas_alim} unid",  # texto con la cantidad
                cantidad=cantidad_abrazaderas_alim  # abrazaderas que van a la compra
            )
        # =========================================================
        # CÁMARA TIPO C — CANALIZACIÓN SUBTERRÁNEA RESIDENCIAL
        # RIC N°04, sección 7.9, 7.9.5, 7.9.8.4.3 y Anexo 4.5
        # Regla RIC N°4 art. 7.9.7.8 / 7.9.7.9 / 7.9.7.10:
        #   Cada tramo (acometida / alimentador) se calcula INDEPENDIENTE:
        #   - Si L es 20 m o menos: 0 cámaras (forma U, RIC N°4 art. 7.9.7.10)
        #   - Si L es mayor a 20 m: ceil(L / 90) cámaras
        #   El total es la SUMA de cámaras de cada tramo.
        # =========================================================
        _hay_sub_acom = "sub" in str(tipo_acometida).strip().lower()  # true si la acometida es subterránea
        _hay_sub_alim = "sub" in str(tipo_alimentador).strip().lower()  # true si el alimentador es subterráneo

        # si ningún tramo es subterráneo, no hace falta cámara ni boquillas: se salta todo este bloque
        if _hay_sub_acom or _hay_sub_alim:  # si no hay tramo enterrado no van ni cámaras ni boquillas

            # --- Longitudes por tramo ---
            _long_acom = max(0.0,  # largo del tramo de acometida que va bajo tierra (0 si la acometida no es subterránea)
                float(longitud_transformador_empalme) - float(longitud_subterraneo_medidor2)  # tramo transformador - empalme menos lo que sale al medidor
            ) if _hay_sub_acom else 0.0  # si la acometida no es subterránea el tramo queda en 0

            _long_alim = max(0.0,  # largo del tramo de alimentador que va bajo tierra (0 si el alimentador no es subterráneo)
                float(longitud_alimentador) - float(longitud_subterraneo_medidor2)  # largo del alimentador menos lo que sale al medidor
            ) if _hay_sub_alim else 0.0  # si el alimentador no es subterráneo el tramo queda en 0

            # --- Cámaras por tramo (independientes) ---
            _cam_acom = int(math.ceil(_long_acom / 90.0)) if _long_acom > 20.0 else 0  # cámaras que necesita la acometida: 0 si el tramo mide 20m o menos
            _cam_alim = int(math.ceil(_long_alim / 90.0)) if _long_alim > 20.0 else 0  # cámaras que necesita el alimentador: 0 si el tramo mide 20m o menos
            _camaras_c = _cam_acom + _cam_alim  # total de cámaras a instalar

            # --- Diámetros de boquilla por tramo ---
            # Si la acometida es subterránea, usa diam_pvc_sub; si no está definido, usa 25mm (mínimo RIC N°4 art. 7.9.7.1)
            _diam_boq_acom = None  # diámetro de boquilla para la acometida (se llena abajo)
            # usa el mismo diámetro que el conduit PVC de la acometida subterránea; si no hay, usa 25mm mínimo
            try:  # diam_pvc_sub siempre existe (parte en None más arriba); el try queda solo por seguridad
                _diam_boq_acom = diam_pvc_sub if diam_pvc_sub else 25  # usa el diámetro del conduit de la acometida, o 25mm de piso
            # rama muerta: diam_pvc_sub ya viene inicializado más arriba
            except NameError:
                _diam_boq_acom = 25  # mínimo de norma para la boquilla
            # Si el alimentador es subterráneo, usa diam_conduit (ya definido antes del bloque)
            _diam_boq_alim = diam_conduit  # diámetro de boquilla para el alimentador: el mismo del conduit del alimentador

            # arma el texto del circuito juntando los tramos que sí son subterráneos
            _circ_cam = []  # acá se van juntando los tramos subterráneos
            if _hay_sub_acom:  # agrega la acometida al texto si corresponde
                _circ_cam.append("Acometida subterránea")  # la acometida va bajo tierra
            if _hay_sub_alim:  # agrega el alimentador al texto si corresponde
                _circ_cam.append("Alimentador subterráneo")  # el alimentador va bajo tierra
            _circ_cam_txt = " / ".join(_circ_cam)  # queda un texto tipo Acometida subterránea / Alimentador subterráneo

            # si se necesita al menos una cámara, se agregan la cámara, el marco y las boquillas
            if _camaras_c > 0:
                # fila de materiales: la cámara de hormigón en sí (una por cada cámara calculada)
                add_row(  # fila de materiales: la cámara de hormigón
                    desc="Cámara tipo C de hormigón prefabricado con tapa de acero diamantado 440x440mm",  # cámara tipo C con tapa de acero
                    marcas_txt=marcas.get("Camara tipo C", ""),  # marcas sugeridas
                    norma="RIC 4 (7.9, 7.9.5, 7.9.7.8, 7.9.7.10, 7.9.8.4.3, Anexo 4.5)",  # artículos del RIC que fijan tipo y ubicación de cámaras
                    circuito=_circ_cam_txt,  # tramos subterráneos que la usan
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=(  # texto con el desglose de cámaras por tramo
                        f"{_camaras_c} unid "
                        f"(Acom={round(_long_acom,1)}m→{_cam_acom}u / "
                        f"Alim={round(_long_alim,1)}m→{_cam_alim}u)"
                    ),
                    cantidad=_camaras_c  # cámaras que van a la compra
                )
                # fila de materiales: el marco metálico que va con cada cámara
                add_row(  # fila de materiales: el marco de la cámara
                    desc="Marco metálico galvanizado para cámara tipo C 440x440mm",  # marco galvanizado del mismo tamaño de la tapa
                    marcas_txt=marcas.get("Marco metalico camara C", ""),  # marcas sugeridas
                    norma="RIC 4 (7.9.8, Anexo 4.5)",  # artículo del RIC del marco y la tapa
                    circuito=_circ_cam_txt,  # tramos subterráneos que la usan
                    unidad="u",  # por unidad
                    k=1,  # sin holgura
                    longitud_m=f"{_camaras_c} unid",  # un marco por cámara
                    cantidad=_camaras_c  # marcos que van a la compra
                )

                # --- Boquillas por tramo ---
                # CASO 1: ambos sub, ambos con cámara y distinto diámetro: van 2 filas separadas
                # si ambos tramos son subterráneos, ambos necesitan cámara y además tienen distinto
                # diámetro de boquilla, hay que poner 2 filas separadas (una por tramo)
                if _hay_sub_acom and _hay_sub_alim \
                        and _cam_acom > 0 and _cam_alim > 0 \
                        and _diam_boq_acom != _diam_boq_alim:  # y las boquillas de cada tramo son de distinto diámetro
                    # fila de materiales: boquillas para el tramo de acometida
                    add_row(
                        desc=(
                            f"Boquilla de PVC ø{_diam_boq_acom}mm con borde redondeado "
                            f"para entrada/salida conduit en cámara tipo C (acometida)"
                        ),
                        marcas_txt=marcas.get("Boquilla camara tipo C", ""),  # marcas que se aceptan para esta boquilla
                        norma="RIC 4 (7.9.8.9, 5.14)",  # artículo del RIC que la pide
                        circuito="Acometida subterránea",  # se carga al tramo de acometida
                        unidad="u",  # se compra por unidad, no por metro
                        k=1,  # sin factor de pérdida
                        longitud_m=f"{_cam_acom * 2} unid",  # texto que se muestra en la columna de largo
                        cantidad=_cam_acom * 2  # 2 boquillas por cámara: una de entrada y una de salida
                    )
                    # fila de materiales: boquillas para el tramo de alimentador
                    add_row(
                        desc=(
                            f"Boquilla de PVC ø{_diam_boq_alim}mm con borde redondeado "
                            f"para entrada/salida conduit en cámara tipo C (alimentador)"
                        ),
                        marcas_txt=marcas.get("Boquilla camara tipo C", ""),  # mismas marcas, ahora para el diámetro del alimentador
                        norma="RIC 4 (7.9.8.9, 5.14)",  # misma referencia normativa
                        circuito="Alimentador subterráneo",  # esta fila va al tramo del alimentador
                        unidad="u",  # se cuenta de a una
                        k=1,  # no lleva merma
                        longitud_m=f"{_cam_alim * 2} unid",  # cuántas unidades quedan a la vista en el informe
                        cantidad=_cam_alim * 2  # 2 boquillas por cámara del alimentador
                    )
                # caso contrario: un solo tramo es subterráneo, o ambos comparten el mismo diámetro,
                # entonces basta con una sola fila de boquillas
                else:
                    # CASO 2: solo acometida sub: usa _diam_boq_acom
                    # CASO 3: solo alimentador sub: usa _diam_boq_alim
                    # CASO 4: ambos sub mismo diámetro: una sola fila con el total
                    if _hay_sub_acom and not _hay_sub_alim:  # solo la acometida va enterrada
                        _diam_boq = _diam_boq_acom  # manda el diámetro de boquilla de la acometida
                    elif _hay_sub_alim and not _hay_sub_acom:  # el único tramo enterrado es el alimentador
                        _diam_boq = _diam_boq_alim  # manda el diámetro de boquilla del alimentador
                    else:  # los dos tramos subterráneos y con el mismo diámetro
                        # ambos sub mismo diámetro
                        _diam_boq = _diam_boq_acom  # da lo mismo cuál, los dos diámetros son iguales
                    _boquillas = _camaras_c * 2  # cada cámara lleva 2 boquillas (entrada y salida)
                    # fila de materiales: boquillas para el total de cámaras
                    add_row(
                        desc=(
                            f"Boquilla de PVC ø{_diam_boq}mm con borde redondeado "
                            f"para entrada/salida conduit en cámara tipo C"
                        ),
                        marcas_txt=marcas.get("Boquilla camara tipo C", ""),  # marcas para la boquilla de cámara tipo C
                        norma="RIC 4 (7.9.8.9, 5.14)",  # artículos del RIC que aplican
                        circuito=_circ_cam_txt,  # texto de circuito armado más arriba según qué tramos son subterráneos
                        unidad="u",  # unidad
                        k=1,  # sin merma
                        longitud_m=f"{_boquillas} unid",  # total de boquillas para todas las cámaras
                        cantidad=_boquillas  # cantidad final que se compra
                    )


        # Cajas de paso alimentador (RIC 7.16.1.13)
        # Solo aplica cuando el empalme es en FACHADA (bloque de arriba,
        # línea ~6838) — en poste NUNCA se generan, aunque el alimentador
        # sea "en ducto", porque acá la caja de paso es exclusiva de fachada.
        cajas_paso_alim = 0  # se deja en 0 a propósito: con empalme en poste no hay caja de paso
        if cajas_paso_alim > 0:  # todo lo de adentro queda inactivo mientras siga en 0
            # Caja de paso ESTANCA (el alimentador va exterior, tipo empalme
            # independiente, no corresponde la caja de PVC de interior).
            # No lleva tapa aparte: la caja estanca viene con su tapa
            # integrada de fábrica, no es una pieza separada.
            # Tamaño según diámetro real del conduit del alimentador: con
            # tubos de 32/40mm entra una caja compacta 150x110x70mm; con
            # 50mm hace falta una más grande 190x140x90mm.
            _dc_alim = 0  # diámetro del conduit del alimentador en mm
            try:  # el diámetro puede venir como texto, por eso el try
                _dc_alim = int(diam_conduit)  # intenta convertir el diámetro a entero
            except Exception:  # el diámetro no venía como número
                _dc_alim = 0  # si no se puede convertir, se deja en 0
            if _dc_alim >= 50:  # conduit de 50mm o más
                _medida_caja_paso_alim = "190x140x90 mm"  # conduit grueso: caja grande
            else:  # conduit de 40mm o menos
                _medida_caja_paso_alim = "150x110x70 mm"  # conduit más chico: caja compacta
            # fila de materiales: la caja de paso estanca del alimentador
            add_row(
                desc=f"Caja de paso estanca IP65 de PVC/policarbonato para exterior {_medida_caja_paso_alim} (incluye tapa)",  # caja estanca IP65 porque el alimentador va a la intemperie
                marcas_txt=marcas.get("Cajas de paso estancas", ""),  # marcas de cajas de paso estancas
                norma="RIC 4 (7.16.1.13)",  # artículo del RIC que exige la caja de paso
                circuito="Alimentador - Caja de paso (tramo > 20m)",  # la caja de paso aparece por tramo largo de alimentador
                unidad="u",  # se cuenta por unidad
                k=1,  # sin merma
                longitud_m=f"{cajas_paso_alim} unid",  # cuántas cajas de paso
                cantidad=cajas_paso_alim  # cantidad que se compra
            )
            # Salidas de caja PVC: 2 por caja de paso (entrada + salida conduit)
            salidas_paso_alim = 2 * cajas_paso_alim  # una salida de caja por cada punta de conduit que entra o sale
            # fila de materiales: las salidas de caja de ese tramo
            add_row(
                desc=f"Salida de caja conduit de PVC de {diam_conduit}mm",  # salida de caja del mismo diámetro que el conduit
                marcas_txt=marcas.get("Salida de caja conduit", ""),  # marcas de salidas de caja conduit
                norma="RIC 4.7.2",  # artículo del RIC de canalizaciones
                circuito="Alimentador - Caja de paso (tramo > 20m)",  # mismo circuito que la caja de paso
                unidad="u",  # se cuentan de a una
                k=1,  # sin merma
                longitud_m=f"{salidas_paso_alim} unid",  # total de salidas
                cantidad=salidas_paso_alim  # cantidad que se compra
            )
            # Fijación de la caja de paso (fila propia, porque el acumulador
            # general de tornillos ya se calculó y escribió antes de llegar
            # acá). En empalme independiente la caja va fijada al poste, así
            # que el criterio es tipo_poste (madera/metálico) — un poste
            # nunca es fibrocemento, no aplica tarugo acá.
            _tornillos_paso_alim = 4 * cajas_paso_alim  # 4 tornillos por caja
            if "madera" in str(tipo_poste).strip().lower():  # poste de madera: tirafondo para madera
                add_row(
                    desc='Tornillo para madera galvanizado 1/4" x 1 1/2" (caja de paso alimentador)',  # tornillo para madera, sin tarugo
                    marcas_txt=marcas.get("Tirafondo hexagonal madera", ""),  # marcas de tirafondo hexagonal
                    norma="-",  # no hay artículo del RIC para la tornillería
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # se carga a la caja de paso del alimentador
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{_tornillos_paso_alim} unid",  # total de tornillos
                    cantidad=_tornillos_paso_alim  # cantidad que se compra
                )
            elif "metal" in str(tipo_poste).strip().lower():  # poste metálico: hay que perforar el fierro
                add_row(
                    desc='Tornillo autoperforante punta de broca 10 x 3/4" (caja de paso alimentador)',  # tornillo autoperforante para el poste metálico
                    marcas_txt=marcas.get("Tornillo autoperforante broca", ""),  # marcas de autoperforantes
                    norma="-",  # sin norma asociada
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # misma caja de paso del alimentador
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{_tornillos_paso_alim} unid",  # misma cantidad de tornillos
                    cantidad=_tornillos_paso_alim  # cantidad que se compra
                )
            else:  # no calzó ningún tipo de poste de los de arriba, queda por definir en terreno
                add_row(
                    desc="Tornillo (definir según tipo de poste) — caja de paso alimentador",  # no se reconoció el tipo de poste, queda para definir en terreno
                    marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas de tornillos
                    norma="-",  # sin norma
                    circuito="Alimentador - Caja de paso (tramo > 20m)",  # mismo circuito
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{_tornillos_paso_alim} unid",  # total de tornillos
                    cantidad=_tornillos_paso_alim  # cantidad que se compra
                )
            # Sumar a acumuladores para otros usos (espuma SIP, etc.) aunque no
            # llegue a tiempo para la fila general de tornillos
            _cajas_total += cajas_paso_alim  # suma al total de cajas, pero llega tarde: la espuma del panel SIP ya se cálculo más arriba

        # caja de derivación metálica con tapa
        # Acometida aérea + alimentador aéreo
        cantidad_caja_derivacion = 1  # valor por defecto, por si ningún caso de abajo aplica
        if ("aer" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower()):  # acometida aérea y alimentador aéreo
            cantidad_caja_derivacion = 1  # todo aéreo: basta una sola caja de derivación
        # Acometida aérea + alimentador subterráneo
        elif ("aer" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower()):  # baja aérea y sigue enterrado
            cantidad_caja_derivacion = 2  # una caja más en el cambio de canalización
        # Acometida subterránea + alimentador aéreo
        elif ("sub" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower()):  # entra enterrado y sube aéreo, mismo criterio
            cantidad_caja_derivacion = 2  # dos cajas de derivación
        # Acometida subterránea + alimentador subterráneo
        elif ("sub" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower()):  # los dos tramos enterrados
            cantidad_caja_derivacion = 3  # una caja por cada cambio de tramo enterrado
        # fila de materiales: las cajas de derivación metálicas
        add_row(
            desc="Caja de derivación metálica pregalvanizada 100x65x65mm con tapa",  # caja metálica pregalvanizada, la de siempre para empalme
            marcas_txt=marcas.get("Caja derivacion metalica", ""),  # marcas de cajas de derivación metálicas
            norma="SEC",  # exigencia de la SEC, no del RIC
            circuito="Empalme y/o alimentador",  # se carga al empalme o al alimentador según el caso
            unidad="u",  # unidad
            k=1,  # sin merma
            longitud_m=f"{cantidad_caja_derivacion} unid",  # cuántas cajas según los casos de arriba
            cantidad=cantidad_caja_derivacion  # cantidad que se compra
        )
        # fila de materiales: el tornillo de tierra de cada caja metálica
        add_row(
            desc="Tornillo autoperforante punta broca 8 x 1/2\" cabeza lenteja + golilla (conexión tierra caja metálica)",  # tornillo con golilla para aterrizar la caja metálica
            marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas de tornillería
            norma="RIC 4 (5.13)",  # artículo del RIC de puesta a tierra de partes metálicas
            circuito="Empalme y/o alimentador (caja de empalme + caja(s) de derivación)",  # cubre la caja de empalme más las de derivación
            unidad="u",  # unidad
            k=1,  # sin merma
            longitud_m=f"{1 + cantidad_caja_derivacion} unid",  # 1 de la caja de empalme más 1 por cada caja de derivación
            cantidad=1 + cantidad_caja_derivacion  # cantidad que se compra
        )
        #Terminal de compresión tipo ojo: 1 por caja de empalme + 1 por cada caja de derivación
        cantidad_terminal_ojo = 1 + cantidad_caja_derivacion  # 1 por la caja de empalme más 1 por cada caja de derivación
        # fila de materiales: los terminales de ojo para la puesta a tierra
        add_row(
            desc=f"Terminal de compresión tipo ojo {sec_pt}mm",  # terminal del mismo calibre que el conductor de tierra
            marcas_txt=marcas.get("Terminal compresion tipo ojo", ""),  # marcas de terminales de compresión
            norma="SEC",  # exigencia SEC
            circuito="Tramos metálicos (caja metálica empalme, caja de derivación, acometida, puesta a tierra 1 y 2)",  # todas las partes metálicas que hay que aterrizar
            unidad="u",  # unidad
            k=1,  # sin merma
            longitud_m=f"{cantidad_terminal_ojo} unid",  # total de terminales
            cantidad=cantidad_terminal_ojo  # cantidad que se compra
        )
        # Material según forrado EXTERIOR de la casa
        hay_madera = "madera" in material_forrado_exterior  # la fachada es de madera
        hay_pvc = "pvc" in material_forrado_exterior  # ej. "siding pvc" — usa el mismo tornillo de madera, sin tarugo
        hay_siding_metalico = ("metál" in material_forrado_exterior) or ("metal" in material_forrado_exterior)  # ej. "siding metálico" o "siding metálico" (sin tilde)
        hay_fibro = "fibro" in material_forrado_exterior  # fibrocemento: es el único forrado que obliga a tarugo
        hay_forrado_valido = any(x in material_forrado_exterior for x in ("madera", "fibro", "siding"))  # si no reconoce ningún forrado conocido, mejor no inventar el tornillo
        # Fila de fachada: SOLO su propio aporte (abrazaderas pvc del
        # alimentador + caja(s) de derivación de la fachada + caddy de PT2) —
        # las abrazaderas/cajas/caddy que van EN EL POSTE ya tienen su propia
        # fila arriba, con el material del poste, no de la fachada.
        # a pesar del nombre, cant_grupo_B_total NO es el grupo B del poste:
        # es la copia de cant_tornillos_fachada calculada mucho más arriba
        # (abrazaderas PVC de fachada x2 + cajas de fachada x4 + caddy de PT2 x1).
        cant_grupo_B_total = int(cant_tornillos_fachada)  # total de tornillos que se fijan contra la fachada
        # Según el material del forrado exterior de la casa se elige el tipo
        # de tornillo (y si hace falta tarugo) para fijar en la fachada.
        if (hay_madera or hay_pvc) and hay_forrado_valido:  # madera o siding pvc: mismo tornillo punta fina, sin tarugo
            if cant_grupo_B_total > 0:  # si no hay nada que fijar en fachada, no se agrega la fila
                add_row(
                    desc='Tornillo 8x1" cabeza lenteja punta fina',  # tornillo punta fina, entra directo en la madera
                    marcas_txt=marcas.get("Tornillo punta fina madera", ""),  # marcas de tornillo punta fina
                    norma="-",  # sin norma
                    circuito="Fachada: abrazaderas pvc del alimentador, caja(s) de derivación, y abrazadera caddy del tramo de puesta a tierra N°2",  # detalle de que se fija con estos tornillos
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{cant_grupo_B_total} unid",  # total de tornillos de fachada
                    cantidad=cant_grupo_B_total  # cantidad que se compra
                )
        # ---- CONTEXTO DE ESTE TRAMO ----
        # Todo lo que sigue, hasta el "return pd.DataFrame(...)", es el CIERRE de
        # build_materiales_df(). Sigue dentro del elif del empalme "independiente"
        # (el empalme va en su propio poste, no pegado a la fachada) y dentro de la
        # ÚLTIMA sección abierta, que es add_section("Empalme"): por eso todas estas
        # filas salen en el Excel bajo el título "Empalme".
        # Este if/elif elige con que se fija lo que va contra la FACHADA, según
        # material_forrado_exterior: madera o siding pvc -> tornillo punta fina;
        # siding metálico -> punta broca; fibrocemento -> tornillo largo + tarugo;
        # forrado reconocido pero distinto -> fila "definir según material de tabique".
        elif hay_siding_metalico and hay_forrado_valido:  # siding metálico: hay que perforar la plancha
            if cant_grupo_B_total > 0:  # sin fijaciones en fachada no se agrega nada
                add_row(
                    desc='Tornillo cabeza lenteja 8x1 1/4" punta broca',  # tornillo punta broca para la plancha metálica
                    marcas_txt=marcas.get("Tornillo autoperforante broca", ""),  # marcas de autoperforantes
                    norma="-",  # sin norma
                    circuito="Fachada: abrazaderas pvc del alimentador, caja(s) de derivación, y abrazadera caddy del tramo de puesta a tierra N°2",  # mismo detalle de fijaciones de fachada
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{cant_grupo_B_total} unid",  # misma cantidad
                    cantidad=cant_grupo_B_total  # cantidad que se compra
                )
        elif hay_fibro and hay_forrado_valido:  # fachada de fibrocemento
            # fibrocemento: necesita tarugo + tornillo juntos
            if cant_grupo_B_total > 0:  # solo si hay algo que fijar
                add_row(
                    desc='Tornillo 8x1 1/2" punta fina cabeza lenteja',  # tornillo más largo porque tiene que atravesar la plancha
                    marcas_txt=marcas.get("Tornillo punta fina madera", ""),  # marcas de tornillo punta fina
                    norma="-",  # sin norma
                    circuito="Fachada: abrazaderas pvc del alimentador, caja(s) de derivación, y abrazadera caddy del tramo de puesta a tierra N°2",  # fijaciones de fachada
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{cant_grupo_B_total} unid",  # total de tornillos
                    cantidad=cant_grupo_B_total  # cantidad que se compra
                )
                # el tarugo va junto al tornillo de arriba, uno por cada fijación
                add_row(
                    desc="Tarugo paloma N°8",  # tarugo paloma, el que sirve en plancha hueca
                    marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas de tarugo paloma
                    norma="-",  # sin norma
                    circuito="Fachada: abrazaderas pvc del alimentador, caja(s) de derivación, y abrazadera caddy del tramo de puesta a tierra N°2",  # las mismas fijaciones del tornillo de arriba
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{cant_grupo_B_total} unid",  # un tarugo por cada tornillo
                    cantidad=cant_grupo_B_total  # cantidad que se compra
                )
        elif not (hay_madera or hay_pvc or hay_siding_metalico or hay_fibro) and hay_forrado_valido:  # el forrado se reconoce pero no cae en ninguno de los materiales de arriba
            if cant_grupo_B_total > 0:  # solo si hay fijaciones
                add_row(
                    desc='Tornillo (definir según material de tabique)',  # queda indicado para definirlo en terreno
                    marcas_txt=marcas.get("Tornillos", ""),  # marcas genéricas
                    norma="-",  # sin norma
                    circuito="Fachada: abrazaderas pvc del alimentador, caja(s) de derivación, y abrazadera caddy del tramo de puesta a tierra N°2",  # fijaciones de fachada
                    unidad="u",  # unidad
                    k=1,  # sin merma
                    longitud_m=f"{cant_grupo_B_total} unid",  # total de tornillos
                    cantidad=cant_grupo_B_total  # cantidad que se compra
                )

        # Conductor puesta a tierra camarilla N°2 (TDA - puesta a tierra 2)
        # +2m: 1 chicote en la camarilla N°2 + 1 chicote en el TDA
        metros_pt2 = math.ceil(float(dist_tda_pt2)) + 2  # largo del cable de tierra entre el TDA y la camarilla N°2, redondeado hacia arriba
        # fila de materiales: el conductor verde de la puesta a tierra N°2
        add_row(
            desc=f"Conductor THWN-2 {sec_pt}mm^2 verde",  # conductor verde de tierra; sec_pt es TEXTO ("4","6","10","16","25"), se fijó más arriba desde la sección del alimentador
            marcas_txt=marcas.get("Conductor THWN-2 verde", ""),  # marcas de THWN-2 verde
            norma="SEC",  # exigencia SEC
            circuito="Puesta a tierra N°2 (Tda - Camarilla N°2)",  # tramo TDA - camarilla N°2
            unidad="m",  # este va por metro
            k=1,  # sin merma
            longitud_m=f"{metros_pt2} m",  # metros que se compran
            cantidad=metros_pt2  # cantidad final en metros
        )

        # Tarugos solo si el forrado exterior es fibrocemento, y solo para lo que se fija en la fachada
        # hay_fibro ya se definió más arriba, leyendo material_forrado_exterior
        # OJO: si el forrado es fibrocemento, estas dos filas se SUMAN a la fila
        # "Tarugo paloma N°8" del bloque de arriba. Aquella usa cant_grupo_B_total
        # (= abrazaderas_fachada*2 + cajas_fachada*4 + caddy de PT2) y ésta usa
        # abrazaderas_fachada*2 + cajas_fachada*4: son las mismas fijaciones sin el
        # caddy, así que el listado queda con tarugos contados dos veces.
        cantidad_tarugos = 0  # sin fibrocemento no se usa ningún tarugo
        if hay_fibro:  # solo el fibrocemento necesita tarugo paloma
            # 2 tarugos por abrazadera + 4 por caja, más 1 extra si la
            # acometida es subterránea y tiene ducto PVC
            cantidad_tarugos = int((abrazaderas_fachada * 2) + (cajas_fachada * 4))  # 2 tarugos por abrazadera y 4 por caja metálica
            if "sub" in str(tipo_acometida).strip().lower() and diam_pvc_sub:  # acometida enterrada con ducto pvc en fachada
                cantidad_tarugos += 1  # una fijación más para ese ducto
            # fila de materiales: los tarugos de fachada
            add_row(
                desc="Tarugo paloma 6mm",  # tarugo paloma de 6mm para la fachada
                marcas_txt=marcas.get("Tarugo paloma", ""),  # marcas de tarugo paloma
                norma="-",  # sin norma
                circuito="Abrazaderas PVC y caja metálica en fachada",  # lo que se fija en la fachada con tarugo
                unidad="u",  # unidad
                k=1,  # sin merma
                longitud_m=f"{cantidad_tarugos} unid",  # total de tarugos
                cantidad=cantidad_tarugos  # cantidad que se compra
            )
            # cada tarugo necesita su tornillo
            add_row(
                desc='Tornillo punta plana 6x1 1/4" para tarugo',  # tornillo que calza con el tarugo paloma
                marcas_txt=marcas.get("Tornillo para tarugo paloma", ""),  # marcas de tornillo para tarugo
                norma="-",  # sin norma
                circuito="Abrazaderas PVC y caja metálica en fachada",  # mismas fijaciones de fachada
                unidad="u",  # unidad
                k=1,  # sin merma
                longitud_m=f"{cantidad_tarugos} unid",  # un tornillo por tarugo
                cantidad=cantidad_tarugos  # cantidad que se compra
            )

        # Pilar para empalme independiente
        # Calcula qué tan alto debe quedar visible el poste según el tipo de
        # acometida y de alimentador (aéreo/subterráneo), y de ahí saca el
        # largo total del poste a comprar (con un 20% extra para enterrarlo).
        # Criterio de los 3 primeros casos: dist_empalme_pt1 (metros del empalme a
        # la camarilla N°1, dato que pregunta el usuario) + la altura que manda el
        # tramo aéreo que exista. El 4° caso no depende de ninguna distancia.
        altura_visible_poste = 0  # altura del poste que queda sobre el suelo, en metros
        if ("aer" in str(tipo_acometida).strip().lower()and "aer" in str(tipo_alimentador).strip().lower()):  # acometida aérea y alimentador aéreo
            altura_visible_poste = (float(dist_empalme_pt1) + float(altura_acometida_aerea))  # distancia empalme→camarilla N°1 más la altura de la acometida aérea
        elif ("aer" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower()):  # acometida aérea y alimentador enterrado
            altura_visible_poste = (float(dist_empalme_pt1) + float(altura_acometida_aerea))  # igual manda la altura de la acometida aérea
        elif ("sub" in str(tipo_acometida).strip().lower() and "aer" in str(tipo_alimentador).strip().lower()):  # acometida enterrada y alimentador aéreo
            altura_visible_poste = (float(dist_empalme_pt1) + float(longitud_poste_alimentador_aereo))  # acá la altura la da el poste del alimentador aéreo
        elif ("sub" in str(tipo_acometida).strip().lower() and "sub" in str(tipo_alimentador).strip().lower()):  # los dos tramos enterrados
            altura_visible_poste = 5  # 5m fijos (criterio propio): con todo enterrado el poste solo sostiene el medidor, no sigue ninguna distancia ingresada
        longitud_total_poste = math.ceil(altura_visible_poste * 1.2)  # +20% para la parte enterrada
        if "madera" in str(tipo_poste).strip().lower():  # poste de madera
            # fila de materiales: el poste de madera del empalme
            add_row(
                desc=f"Poste de madera {longitud_total_poste}m, pino impregnado",  # poste de pino impregnado del largo calculado
                marcas_txt=marcas.get("Poste madera", ""),  # marcas de postes de madera
                norma="-",  # sin norma
                circuito="Empalme independiente",  # va en el empalme independiente
                unidad="u",  # se compra por unidad
                k=1,  # sin merma
                longitud_m=f"{longitud_total_poste} m",  # largo total del poste
                cantidad=1  # siempre es un poste
            )
        elif "metal" in str(tipo_poste).strip().lower():  # pilar metálico
            # fila de materiales: el pilar metálico del empalme
            add_row(
                desc=f"Pilar metálico cuadrado {longitud_total_poste}m, 100x100x3mm",  # perfil cuadrado 100x100x3mm
                marcas_txt=marcas.get("Pilar metalico", ""),  # marcas de pilar metálico
                norma="-",  # sin norma
                circuito="Empalme independiente",  # empalme independiente
                unidad="u",  # unidad
                k=1,  # sin merma
                longitud_m=f"{longitud_total_poste} m",  # largo total del pilar
                cantidad=1  # un solo pilar
            )

        # Portafusible aéreo loza según interruptor termomagnético del empalme
        # Criterio: el fusible va un escalón por sobre el TM del empalme
        # (TM 25A -> fusible 30A; TM 32A o 40A -> fusible 60A). Con cualquier otro
        # calibre (o si no se pudo leer el TM) fusible_A queda en None y NO se
        # agrega ninguna fila de portafusible.
        tm_empalme_A = parse_in_tm(interruptor_texto)  # saca el amperaje del TM del empalme (ej: "1x25A" -> 25)
        fusible_A = None  # si el TM no cae en los casos de abajo, no se agrega portafusible
        if tm_empalme_A == 25:  # empalme con TM de 25A
            fusible_A = 30  # con TM de 25A corresponde fusible de 30A
        elif tm_empalme_A in [32, 40]:  # empalme con TM de 32A o 40A
            fusible_A = 60  # con TM de 32A o 40A corresponde fusible de 60A
        if fusible_A:  # solo si quedó definido un amperaje de fusible
            # fila de materiales: el portafusible del empalme
            add_row(
                desc=f"Portafusible de loza con fusibles cartucho {fusible_A}A",  # portafusible de loza con el fusible que corresponde
                marcas_txt=marcas.get("Portafusible de loza", ""),  # marcas de portafusible de loza
                norma="SEC",  # exigencia SEC
                circuito="Empalme",  # va en el empalme
                unidad="u",  # unidad
                k=1,  # sin merma
                longitud_m="1 unid",  # siempre es uno
                cantidad=1  # cantidad que se compra
            )


    # =========================
    # REORDENAR SECCIONES
    # =========================
    # A esta altura "filas" tiene todas las filas de materiales agregadas
    # con add_row, pero en el orden en que se fueron calculando (no es el
    # orden en que deben verse en el Excel). Acá se reordenan por sección.
    ORDEN_SECCIONES = [
        "Empalme",
        "Protecciones",
        "Borneras de conexión",
        "Cableado interior tablero",
        "Terminales ferrul interiores",
        "Canalizaciones",
        "Conductores",
        "Accesorios",
        "Conectores cónicos",
        "Iluminarias",
        "Tornillería",
        "Tarugos",
        "Sellos / Aislación (Panel SIP)",
    ]  # orden final en que deben aparecer las secciones en el Excel
    # Ojo: las secciones que realmente se crean con add_section() en toda la
    # función son Canalizaciones, Conductores, Protecciones, Accesorios,
    # Conectores cónicos, Iluminarias, Terminales ferrul interiores, Cableado
    # interior tablero, Tornillería, Sellos / Aislación (Panel SIP) y Empalme.
    # "Borneras de conexión" y "Tarugos" están en esta lista pero nunca se
    # generan, así que esos dos nombres quedan sin efecto.

    # Asignar grupo a cada fila según la última sección vista
    grupos = []  # sección a la que pertenece cada fila, en el mismo orden que filas
    seccion_actual = ""  # sección que se viene arrastrando mientras se recorre
    for fila in filas:  # recorre las filas en el orden en que se fueron calculando
        desc = str(fila.get("Ítem", ""))  # variable auxiliar que queda sin uso: la decisión de abajo se toma leyendo fila["Ítem"] directo
        # fila de sección: Ítem está vacío y Descripción técnica es el nombre de la sección
        if fila["Ítem"] == "":  # sin número de item quiere decir que es una cabecera de sección
            seccion_actual = fila["Descripción técnica"]  # de acá para abajo las filas pertenecen a esta sección
        grupos.append(seccion_actual)  # cada fila queda marcada con la sección a la que pertenece

    # Separar filas en bloques por sección (cabecera + sus filas)
    bloques = {}  # {"nombre de la sección": [fila_titulo, fila_material_1, fila_material_2, ...]}
    seccion_actual = ""  # se recorre de nuevo, ahora armando los bloques
    for fila, grupo in zip(filas, grupos):  # empareja cada fila con la sección que le tocó
        if fila["Ítem"] == "" and fila["Descripción técnica"] != "":  # cabecera con nombre: arranca un bloque nuevo
            seccion_actual = fila["Descripción técnica"]  # nombre de la sección que se esta armando
            bloques[seccion_actual] = [fila]  # nueva sección: empieza con su fila de título (si un mismo nombre se abriera 2 veces, el segundo pisaría al primero)
        elif seccion_actual:  # fila normal, siempre que ya haya una sección abierta
            bloques[seccion_actual].append(fila)  # fila normal: se suma al bloque de su sección

    # Reconstruir filas en el orden deseado
    filas_ordenadas = []  # acá se van dejando los bloques ya ordenados
    for nombre in ORDEN_SECCIONES:  # recorre las secciones en el orden definido arriba
        if nombre in bloques and len(bloques[nombre]) > 1:  # solo si tiene filas reales (no solo el título)
            filas_ordenadas.extend(bloques[nombre])  # pega la cabecera y sus filas al listado final
    # Agregar secciones que no estén en ORDEN_SECCIONES (por si acaso)
    for nombre, bloque in bloques.items():  # barrida final por si quedó alguna sección fuera de la lista
        if nombre not in ORDEN_SECCIONES and len(bloque) > 1:  # solo las que no se agregaron antes y tienen filas reales
            filas_ordenadas.extend(bloque)  # se agregan al final para no perderlas

    # Renumerar Ítems (después de reordenar, los números tienen que quedar correlativos de nuevo)
    n = 1  # contador del número de item
    for fila in filas_ordenadas:  # recorre las filas ya en el orden final
        if fila["Ítem"] != "":  # las cabeceras de sección no se numeran
            fila["Ítem"] = n  # le pone el correlativo nuevo
            n += 1  # avanza el correlativo de item del listado

    # Devuelve el listado final de materiales como DataFrame, con las
    # columnas que van a mostrarse en la hoja "Materiales" del Excel.
    return pd.DataFrame(
        filas_ordenadas,  # filas ya agrupadas por sección, en el orden de ORDEN_SECCIONES
        columns=[  # encabezados de la tabla de materiales tal como salen en el excel
            "Ítem",  # número correlativo del item
            "Descripción técnica",  # que es el material (sección, tipo, uso)
            "Marcas sugeridas",  # marcas que se sugieren para ese material
            "Sello SEC",  # valor provisorio puesto por add_row; al escribir el Excel se recalcula con _requiere_sello_sec(descripción, norma)
            "Norma / RIC",  # texto de la norma; al escribir el Excel, si el material tiene fila en la hoja Base Normativa la celda se reemplaza por "Ver normativa" con link (si no la encuentra, conserva el texto)
            "Circuito",  # a que circuito pertenece el material
            "Unidad",  # unidad de venta: metro, unidad, rollo, etc
            "K",  # factor k de pérdida/desperdicio aplicado
            "Longitud (m) / Unidad",  # texto libre con el metraje o las unidades, ej: "12 m" o "4 unid"
            "Cantidad"  # cantidad final ya redondeada a comprar
        ]
    )

# -------- PRIMERO: EMPALME (AUTO) --------
# El empalme se calculará más adelante usando la corriente CON factor de demanda.
# Aquí solo definimos valores iniciales para no bloquear el ingreso de datos.
# Los 4 valores de abajo NO se leen en ninguna parte del ingreso: los cuatro
# quedan sobrescritos de una vez, mucho más abajo, por la línea
#   dif_calibre, max_sum_tm, max_pot_especial, texto_omni = parametros_desde_empalme(interruptor_empalme)
# que corre recién en la PARTE 3.1, cuando ya está calculado el empalme.
proteccion_empalme = None  # (se deja por compatibilidad; NO se usa para cálculos)
dif_calibre = None  # calibre del diferencial general (A); lo define parametros_desde_empalme() según el TM del empalme
max_sum_tm = None  # tope de suma de los TM que pueden colgar de un mismo diferencial (A); también sale de parametros_desde_empalme()
max_pot_especial = float("inf")  # este valor nunca se llega a leer: el único chequeo que lo usaba quedó desactivado (encerrado entre comillas triples, no se ejecuta)
texto_omni = None  # texto del interruptor general omnipolar; parametros_desde_empalme() lo arma como f"2x{Ie}A / 10kA / Curva C", ej: "2x40A / 10kA / Curva C"

# -------- PARTE 1: INGRESO DE AMBIENTES Y CARGAS --------
# Acá arranca el "script" en sí: ya no hay funciones, es código que corre
# de corrido preguntando datos por consola. Primero pregunta cuántos
# ambientes tiene la casa (living, baño, piezas, etc.).
cantidad_ambientes = int(input("\n Ingrese la cantidad de ambientes (living, baño, piezas, etc.): "))  # cuántos ambientes hay que recorrer preguntando

# Los 4 contenedores que llena el ciclo de abajo. La clave de los 3
# diccionarios es el nombre del ambiente TAL COMO lo escribió el usuario (si
# escribe dos veces el mismo nombre, el segundo pisa al primero).
ambientes_detalle = []       # lista de fichas, una por ambiente; termina como el DataFrame ambientes_df, que se escribe en la hoja "Informe" y además se le pasa a build_materiales_df()
componentes_por_ambiente = {}  # {"cocina": [{"id":1,"nombre":"horno","potencia":2000.0}, ...]}
enchufes_por_ambiente = {}     # {"cocina": [{"id":1,"módulos":2,"potencias_modulo":[...],"potencia_total":250.0}, ...]}
luminarias_por_ambiente = {}   # {"cocina": [{"id":1,"tipo":"Foco LED","montaje":"embutida","potencia":9.0,"desc":"..."}, ...]}

# recorre uno por uno todos los ambientes, preguntando sus datos
for i in range(1, cantidad_ambientes + 1):
    print(f"\n Ambiente N°{i}")  # título en pantalla para saber en que ambiente vamos
    nombre = input(f"   - Nombre del ambiente: ")  # nombre del ambiente, después se usa para detectar cocina, baño, etc
    dimension = float(input(f"   - Dimensión del ambiente '{nombre}' en m²: "))  # superficie en m2, define el mínimo de luminarias
    perimetro = float(input(f"   - Perímetro del ambiente '{nombre}' en m: "))  # perímetro en m, se usa para la regla de 1 enchufe cada 8m
    # tipo de tabique del muro: define como se pasa el conduit y que cajas van
    material_tabique = input(
        f" - Tipo de material de construcción del tabique del ambiente '{nombre}' "
        "(Madera, Metalcon, Panel SIP): "
    ).strip().lower()  # deja el material del tabique en minúsculas y sin espacios sobrantes

    # forrado interior del muro, cambia el tipo de caja y el montaje
    material_forrado_interior = input(
        f" - Forrado interior del ambiente '{nombre}' (Volcanita, Madera): "
    ).strip().lower()  # deja el forrado interior en minúsculas y sin espacios sobrantes


    # Iluminación: mínimo de luminarias por área. Es criterio propio del
    # programa (no una tabla del RIC): hasta 10 m² basta 1 luminaria, sobre
    # 10 m² se exigen 2. El número solo se usa para validar lo que ingresa el
    # usuario más abajo y para el texto que se le muestra en pantalla.
    if dimension <= 10:  # ambiente chico: basta con 1 luminaria
        min_luminarias = 1  # mínimo exigido de luminarias para este ambiente
        texto_condicion = "<= 10 m²"  # texto de la condición, solo para mostrarlo en pantalla
    else:  # ambiente más grande de 10 m2
        min_luminarias = 2  # sobre 10 m2 se piden 2 luminarias como mínimo
        texto_condicion = "> 10 m²"  # texto de la condición, solo para mostrarlo en pantalla

    # pide la cantidad de luminarias del ambiente, repite si no cumple el mínimo
    while True:
        try:  # intenta convertir lo que escribió a número entero
            cantidad_luminarias = int(input(  # cantidad de luminarias que va a llevar el ambiente
                f"   - Cantidad de luminarias del ambiente '{nombre}' "
                f"(mínimo {min_luminarias} para {texto_condicion}): "
            ))  # la cantidad queda guardada en cantidad_luminarias
            if cantidad_luminarias < min_luminarias:  # no alcanza el mínimo, se vuelve a preguntar
                print(f"     ! Debe ingresar al menos {min_luminarias} luminaria(s) para {texto_condicion}.")  # avisa cual es el mínimo que falta
            else:  # el número cumple el mínimo
                break  # cumple el mínimo, sigue
        except:  # si no se pudo convertir a entero, cae acá
            print("     ! Ingrese un número entero válido para la cantidad de luminarias.")  # lo que escribió no era un entero

    detalle_luminarias = []  # descripciones armadas de cada luminaria, para el informe
    potencia_luminarias = []  # potencia en w de cada luminaria, para sumarlas después
    luminarias_detalle_amb = []  # detalle completo de cada luminaria del ambiente

    # pregunta los datos de cada luminaria del ambiente, una por una
    for j in range(1, cantidad_luminarias + 1):
        # tipo de luminaria, define la descripción técnica y la marca sugerida en materiales
        tipo = input(
            f"     • Tipo de iluminación {j} "
            f"(Foco LED, aplique LED, Tubo fluorescente, Tubo led, Ampolleta incandescente, ampolleta LED): "
        ).strip()  # saca los espacios de más del tipo de luminaria

        # pregunta si la luminaria va embutida (dentro del cielo/muro) o sobrepuesta
        while True:
            montaje = input(  # pregunta el montaje de esta luminaria
                f"       • ¿La iluminación {j} es embutida o sobrepuesta? (embutida/sobrepuesta): "
            ).strip().lower()  # normaliza la respuesta del montaje
            if montaje in ["embutida", "sobrepuesta"]:  # solo acepta estas dos formas de montaje
                break  # montaje válido, sigue con la siguiente pregunta
            print("         ! Responda solo 'embutida' o 'sobrepuesta'.")  # respuesta inválida, vuelve a preguntar

        # pregunta si el usuario sabe la potencia real de la luminaria
        while True:
            conoce = input(  # pregunta si conoce la potencia real de la luminaria
                f"       ¿Usted conoce el valor de la potencia de la luminaria {j}? (si/no)\n"
                f"       (Si no lo sabe, se ingresará automáticamente 100W por norma): "
            ).strip().lower()  # normaliza el si/no
            if conoce in ["si", "no"]:  # solo acepta si o no
                break  # respuesta válida, sale del bucle
            print("         ! Responda solo 'si' o 'no'.")  # respuesta inválida, vuelve a preguntar

        if conoce == "si":  # si conoce la potencia real, se la pide en watts
            p = float(input(f"       Potencia de la luminaria {j} [W]: "))  # potencia de placa de la luminaria
        else:  # si no conoce la potencia
            p = 100.0  # no se sabe la potencia real, se usa un valor por defecto
            print("       -> No se conoce la potencia, se asignan 100 W por defecto según criterio.")  # avisa que se uso el valor por defecto de norma

        potencia_luminarias.append(p)  # guarda la potencia para el total del ambiente

        desc_final = desc_luminaria_auto(tipo, montaje, p)  # arma la descripción "bonita" de la luminaria

        detalle_luminarias.append(desc_final)  # guarda la descripción para mostrarla en el informe
        luminarias_detalle_amb.append({"id": j, "tipo": tipo, "montaje": montaje, "potencia": p, "desc": desc_final})  # guarda todos los datos de la luminaria por separado

    potencia_total_ilum = sum(potencia_luminarias)  # suma la potencia de todas las luminarias del ambiente
    detalle_iluminacion = "\n".join(detalle_luminarias) if detalle_luminarias else "Ninguna"  # junta las descripciones en un solo texto para la tabla de ambientes
    luminarias_por_ambiente[nombre] = luminarias_detalle_amb[:]  # copia de las luminarias del ambiente; de acá sale luminarias_restantes, la bolsa desde la que se reparten a los circuitos de iluminación en la PARTE 2.1

    # Conmutado: se pregunta una vez por ambiente
    while True:
        conmutado = input(f"   - ¿El ambiente '{nombre}' tiene conmutado (2 puntos)? (si/no): ").strip().lower()  # conmutado = la misma luz se enciende desde dos puntos (9/24)
        if conmutado in ["si", "no"]:  # solo acepta si o no
            break  # respondió si o no, sigue
        print("     ! Responda solo 'si' o 'no'.")  # respuesta inválida, vuelve a preguntar

    # acumuladores de longitudes de cableado (se usan después en build_materiales_df
    # para calcular los metros exactos de cada color de conductor)
    L_viajeros_924 = 0.0  # largo de los viajeros entre los dos interruptores del 9/24
    L_retorno_lampara = 0.0  # largo del retorno del interruptor a la lámpara
    L_fase_caja_primer_int = 0.0  # largo de la fase desde la caja hasta el primer interruptor
    L_troncal_primera_oct_924 = 0.0  # distancia de la caja troncal a la primera caja octogonal (1 vez por ambiente, no por luminaria)

    L_caja_int_fase_ida = 0.0      # suma L_ida por grupo de interruptor (se mantiene por compatibilidad)
    L_troncal_oct1 = 0.0           # suma distancia de la troncal a la primera oct por grupo (idem)
    L_oct1_oct2 = 0.0              # suma distancia entre octs intermedias por grupo (idem)
    grupos_ilum_detalle = []       # detalle por grupo: [{"tipo","n_lum","L_ida","L_tr_oct1","L_o1_o2"}, ...]
                                    # permite calcular extra_R/extra_NT grupo por grupo, en vez de
                                    # con sumas totales (que sobre-cuentan si hay >1 interruptor
                                    # por ambiente con distancias distintas entre sí)

    N_conmutadas_924 = 0  # cuántas luminarias del ambiente quedan conmutadas con 9/24

    if conmutado == "si":  # el ambiente lleva conmutado, hay que preguntar las longitudes del 9/24
        # pregunta cuántas de las luminarias del ambiente van con conmutado 9/24
        while True:  # repite hasta que ingrese un número válido
            try:  # intenta leer la cantidad de conmutadas como entero
                N_conmutadas_924 = int(input(  # cuántas de las luminarias del ambiente van conmutadas
                    f"     • En '{nombre}' hay {cantidad_luminarias} luminarias.\n"
                    f"       ¿Cuántas luminarias serán conmutadas con 9/24? (0 a {cantidad_luminarias}): "
                ))  # la respuesta queda guardada en N_conmutadas_924
                if 0 <= N_conmutadas_924 <= int(cantidad_luminarias):  # el número tiene que quedar entre 0 y el total de luminarias
                    break  # cantidad dentro del rango, sigue
                print(f"       ! Debe estar entre 0 y {cantidad_luminarias}.")  # quedó fuera de rango, vuelve a preguntar
            except:  # si escribió algo que no es número, cae acá
                print("       ! Ingrese un número entero válido.")  # no era un entero, vuelve a preguntar

        if N_conmutadas_924 > 0:  # hay conmutadas, se piden las longitudes del grupo 9/24
            # El 9/24 enciende TODAS las luminarias conmutadas del grupo al
            # mismo tiempo (mismo par de interruptores) — estas 4 longitudes
            # se preguntan 1 sola vez por ambiente, NO por cada luminaria.
            print("     • Conmutado (9/24): ingrese las longitudes REALES del grupo conmutado (ENTER=0.0)")  # aviso de que vienen las longitudes del conmutado
            L_viajeros_924 = pedir_float_opcional(  # largo total entre los dos interruptores del 9/24
                "       • Longitud TOTAL entre interruptores (viajeros 9/24) [m] (ENTER=0.0): ", 0.0, 0.0
            )
            L_retorno_lampara = pedir_float_opcional(  # largo del retorno interruptor a lámpara
                "       • Longitud de RETORNO (interruptor→lámpara) [m] (ENTER=0.0): ", 0.0, 0.0
            )
            L_fase_caja_primer_int = pedir_float_opcional(  # largo de la fase que llega al primer interruptor
                "       • Longitud de FASE (caja→primer interruptor) [m] (ENTER=0.0): ", 0.0, 0.0
            )
            L_troncal_primera_oct_924 = pedir_float_opcional(  # largo de la troncal a la primera caja octogonal
                "       • Longitud de la CAJA TRONCAL a la PRIMERA CAJA OCTOGONAL [m] (ENTER=0.0): ", 0.0, 0.0
            )

        n_no_conmut = int(cantidad_luminarias) - int(N_conmutadas_924)  # luminarias que quedan con interruptor normal (no 9/24)
        if n_no_conmut > 0:  # aparte de las conmutadas quedan luminarias con interruptor normal
            # agrupa las luminarias no conmutadas en interruptores 9/12 (1), 9/15 (2) o 9/32 (3)
            # Descomponer en grupos de interruptor
            c12_g, c15_g, c32_g = descomponer_interruptores(n_no_conmut)
            grupos = []  # lista de grupos: (tipo de interruptor, cuántas luminarias controla)
            for _ in range(c32_g): grupos.append(('9/32', 3))  # un grupo 9/32 controla 3 luminarias
            for _ in range(c15_g): grupos.append(('9/15', 2))  # un grupo 9/15 controla 2 luminarias
            for _ in range(c12_g): grupos.append(('9/12', 1))  # un grupo 9/12 controla 1 luminaria

            # pregunta las longitudes de cada grupo de interruptor no conmutado
            print(f"     • Luminarias NO conmutadas: {n_no_conmut} luminarias → {len(grupos)} grupo(s) de interruptor")
            for g_idx, (tipo_g, n_lum_g) in enumerate(grupos, 1):  # recorre grupo por grupo pidiendo sus distancias
                print(f"       - Grupo {g_idx} ({tipo_g}, {n_lum_g} luminaria{'s' if n_lum_g>1 else ''}):")  # muestra que grupo se esta preguntando
                _L_ida_g = pedir_float_opcional(  # distancia de la caja troncal al interruptor de este grupo
                    f"         • Longitud troncal→interruptor {tipo_g} [m] (ENTER=0.0): ", 0.0, 0.0
                )
                _L_tr1_g = 0.0  # por defecto 0: no aplica si el grupo tiene una sola luminaria
                if n_lum_g > 1:  # solo se pregunta si el grupo controla 2 o más luminarias
                    # Solo aplica a 9/15 y 9/32 (2+ luminarias) — para 9/12
                    # (1 sola luminaria) esta distancia nunca se usa en el
                    # cálculo, así que no se pregunta.
                    _L_tr1_g = pedir_float_opcional(
                        f"         • Longitud troncal→primera oct [m] (ENTER=0.0): ", 0.0, 0.0
                    )
                _L_o12_g = 0.0  # por defecto 0: no aplica si el grupo tiene 2 o menos luminarias
                if n_lum_g > 2:  # solo aplica al 9/32, que lleva 3 luminarias
                    # Solo aplica a 9/32 (3+ luminarias) — para 9/15 (2
                    # luminarias) esta distancia nunca se usa en el cálculo,
                    # así que no se pregunta.
                    for j in range(1, n_lum_g - 1):  # recorre las cajas octogonales intermedias del grupo
                        _L_o12_g += pedir_float_opcional(  # suma la distancia entre cada par de octogonales seguidas
                            f"         • Longitud oct{j}→oct{j+1} [m] (ENTER=0.0): ", 0.0, 0.0
                        )
                # Mantener las sumas totales por compatibilidad con otros usos
                L_caja_int_fase_ida += _L_ida_g  # acumula la distancia troncal a interruptor
                L_troncal_oct1 += _L_tr1_g  # acumula la distancia troncal a primera octogonal
                L_oct1_oct2 += _L_o12_g  # acumula la distancia entre octogonales
                # Y guardar el detalle por grupo, para el cálculo grupo-por-grupo
                grupos_ilum_detalle.append({
                    "tipo": tipo_g, "n_lum": n_lum_g,  # tipo de interruptor del grupo y cuántas luminarias cuelga
                    "L_ida": _L_ida_g, "L_tr_oct1": _L_tr1_g, "L_o1_o2": _L_o12_g  # las tres distancias del grupo, para calcular metros de cable
                })  # cierra el detalle de este grupo de interruptor

    # Si el ambiente NO tiene conmutado, todas las luminarias se reparten en
    # grupos de interruptor normal (9/12, 9/15 o 9/32), con la misma lógica
    # que se usó arriba para las luminarias no conmutadas de un ambiente con conmutado.
    else:
        # sin conmutado: todas las luminarias van en grupos de interruptor normales
        n_no_conmut = int(cantidad_luminarias)  # todas las luminarias de este ambiente van a interruptores normales (no hay conmutadas)
        c12_g, c15_g, c32_g = descomponer_interruptores(n_no_conmut)  # separa la cantidad en grupos de 1, 2 o 3 luminarias por interruptor (9/12, 9/15, 9/32)
        grupos = []  # lista de grupos: cada grupo es (tipo_interruptor, cantidad_de_luminarias)
        for _ in range(c32_g): grupos.append(('9/32', 3))  # un grupo 9/32 controla 3 luminarias
        for _ in range(c15_g): grupos.append(('9/15', 2))  # un grupo 9/15 controla 2 luminarias
        for _ in range(c12_g): grupos.append(('9/12', 1))  # un grupo 9/12 controla 1 luminaria

        print(f"     • Sin conmutado: {n_no_conmut} luminarias → {len(grupos)} grupo(s) de interruptor")  # avisa en pantalla cuántos grupos de interruptor se armaron para este ambiente
        # Recorre cada grupo de interruptor pidiendo las distancias de cableado
        # necesarias para calcular después los metros de cada color de conductor.
        for g_idx, (tipo_g, n_lum_g) in enumerate(grupos, 1):
            print(f"       - Grupo {g_idx} ({tipo_g}, {n_lum_g} luminaria{'s' if n_lum_g>1 else ''}):")  # muestra que grupo se esta preguntando
            # distancia entre la caja troncal (o el tablero) y el interruptor de este grupo
            _L_ida_g = pedir_float_opcional(
                f"         • Longitud troncal→interruptor {tipo_g} [m] (ENTER=0.0): ", 0.0, 0.0
            )
            _L_tr1_g = 0.0  # valor por defecto: no aplica si el grupo tiene una sola luminaria
            if n_lum_g > 1:  # solo se pregunta si el grupo controla 2 o más luminarias
                # Solo aplica a 9/15 y 9/32 (2+ luminarias) — para 9/12
                # (1 sola luminaria) esta distancia nunca se usa en el
                # cálculo, así que no se pregunta.
                _L_tr1_g = pedir_float_opcional(
                    f"         • Longitud troncal→primera oct [m] (ENTER=0.0): ", 0.0, 0.0
                )
            _L_o12_g = 0.0  # valor por defecto: no aplica si el grupo tiene 2 o menos luminarias
            if n_lum_g > 2:  # solo se pregunta si el grupo controla 3 luminarias (9/32)
                # Solo aplica a 9/32 (3+ luminarias) — para 9/15 (2
                # luminarias) esta distancia nunca se usa en el cálculo,
                # así que no se pregunta.
                for j in range(1, n_lum_g - 1):  # recorre las cajas octogonales intermedias del grupo
                    # suma la distancia entre cada par de cajas octogonales consecutivas del grupo
                    _L_o12_g += pedir_float_opcional(
                        f"         • Longitud oct{j}→oct{j+1} [m] (ENTER=0.0): ", 0.0, 0.0
                    )
            # Mantener las sumas totales por compatibilidad con otros usos
            L_caja_int_fase_ida += _L_ida_g  # acumula la distancia troncal→interruptor de este grupo
            L_troncal_oct1 += _L_tr1_g  # acumula la distancia troncal→primera octogonal de este grupo
            L_oct1_oct2 += _L_o12_g  # acumula la distancia entre octogonales de este grupo
            # Y guardar el detalle por grupo, para el cálculo grupo-por-grupo
            grupos_ilum_detalle.append({
                "tipo": tipo_g, "n_lum": n_lum_g,  # tipo de interruptor del grupo y cuántas luminarias cuelga
                "L_ida": _L_ida_g, "L_tr_oct1": _L_tr1_g, "L_o1_o2": _L_o12_g  # las tres distancias del grupo, para calcular metros de cable
            })  # cierra el detalle de este grupo de interruptor

    # Los componentes especiales son equipos que se enchufan y consumen bastante
    # potencia (ej: aire acondicionado, microondas, calefont eléctrico). Se suman
    # aparte de la iluminación y de los enchufes comunes del ambiente.
    # Componentes especiales
    componentes_nombres = []  # nombres de los componentes especiales ingresados en este ambiente
    componentes_potencias = []  # potencia en W de cada componente especial
    componentes_detalle_amb = []  # detalle completo (id, nombre, potencia) de cada componente

    # pregunta si el ambiente tiene algún componente especial conectado a enchufe
    tiene_componentes = input(
        f"   - ¿El ambiente '{nombre}' tiene componentes especiales conectados a enchufes? (si/no): "
    ).strip().lower()  # normaliza el si/no de los componentes especiales

    if tiene_componentes == 'si':  # si contestó que sí, se pide el detalle de cada componente, uno por uno
        idx_comp = 1  # id correlativo para identificar cada componente ingresado
        # pregunta uno por uno los equipos especiales, hasta que el usuario escriba "fin"
        while True:  # bucle que sigue pidiendo componentes hasta que el usuario escriba 'fin'
            nombre_comp = input("     • Ingrese nombre del componente (o 'fin' para terminar): ")  # pregunta el nombre del componente especial
            if nombre_comp.lower() == 'fin':  # 'fin' termina el ingreso de componentes de este ambiente
                break  # escribió fin, corta el ingreso de componentes

            # Si es aire acondicionado, se pide el BTU y se convierte con un EER típico de 3.2
            # (dato estimado para potencia del ambiente; la placa real se ingresa en el circuito)
            _nombre_lower = nombre_comp.strip().lower()  # normaliza el texto para comparar sin importar mayúsculas
            _es_clima_comp = any(k in _nombre_lower for k in ("aire", "clima", "split", "ac ", "a/c"))  # si el nombre contiene alguna de estas palabras, se asume que es un aire acondicionado
            if _es_clima_comp:  # si es aire acondicionado, se pide la capacidad en BTU/h en vez de la potencia en W
                # bucle de validación: repite hasta que se ingrese un BTU/h válido (> 0)
                while True:
                    try:  # intenta leer los btu como número
                        _btu = float(input(f"       Capacidad de '{nombre_comp}' en BTU/h (ej: 9000, 12000, 18000, 24000): "))  # convierte el texto a número y valida que sea mayor que 0
                        if _btu <= 0:  # el btu tiene que ser mayor que cero
                            print("         ! Debe ser mayor que 0.")  # aviso y vuelve a preguntar
                            continue  # btu inválido, vuelve a pedirlo
                        break  # btu válido, sigue
                    except:  # no era número, cae acá
                        print("         ! Ingrese un número válido.")  # no era un número, vuelve a preguntar
                # Conversión con EER típico 3.2 (estimación; dato real se ingresa en el circuito)
                potencia_comp = round(_btu / (3.2 * 3.412), 1)  # fórmula: Watts = BTU/h dividido por (EER x 3.412)
                print(f"       → Potencia estimada: {potencia_comp} W (dato referencial, se precisa en el circuito)")  # muestra la potencia estimada; el valor real y definitivo se ingresa luego en el circuito
            else:  # si no es aire acondicionado, se pregunta la potencia directamente en W
                potencia_comp = float(input(f"       Potencia de '{nombre_comp}' en W: "))  # potencia de placa del equipo, en watts
            # Nota: no limitamos aquí por empalme, porque el empalme se calcula al final con factor de demanda.
            # La validación final se muestra como aviso (no bloqueante).
            componentes_nombres.append(nombre_comp)  # guarda el nombre del componente
            componentes_potencias.append(potencia_comp)  # guarda la potencia del componente
            # guarda el detalle completo de este componente (id, nombre, potencia),
            # para usarlo más adelante en el listado de materiales
            componentes_detalle_amb.append({
                "id": idx_comp,  # id correlativo del componente dentro del ambiente
                "nombre": nombre_comp,  # nombre tal como lo escribió el usuario
                "potencia": potencia_comp  # potencia en w que se suma al total del ambiente
            })  # cierra el registro de este componente especial
            idx_comp += 1  # avanza el contador para el próximo componente
    else:  # si contestó que no, no hay componentes especiales que agregar
        print("     • No se ingresaron componentes especiales para este ambiente.")  # deja constancia en pantalla de que el ambiente no tiene equipos especiales

    componentes_por_ambiente[nombre] = componentes_detalle_amb[:]  # guarda el detalle de componentes de este ambiente para usarlo más adelante
    potencia_total_comp = sum(componentes_potencias)  # potencia total sumando todos los componentes especiales del ambiente

    # ---------- ENCHUFES COMUNES ----------
    nombre_lower = nombre.strip().lower()  # normaliza el nombre del ambiente para comparar sin importar mayúsculas
    es_cocina = nombre_lower.startswith("cocina")  # es cocina si el nombre empieza con 'cocina'
    es_lavadero = nombre_lower.startswith("lavadero")  # es lavadero si el nombre empieza con 'lavadero'

    # ambientes que requieren 1 enchufe DOBLE/TRIPLE cada 8m de perímetro
    es_dormitorio = ("dormitorio" in nombre_lower) or nombre_lower.startswith("pieza") or nombre_lower.startswith("habitacion")  # detecta dormitorio/pieza/habitación por el nombre del ambiente
    es_living = nombre_lower.startswith("living")  # detecta living
    es_comedor = nombre_lower.startswith("comedor")  # detecta comedor
    es_sala_estar = ("sala de estar" in nombre_lower) or ("estar" in nombre_lower)  # detecta sala de estar
    es_bano = (nombre_lower.startswith("baño") or nombre_lower.startswith("bano"))  # detecta baño
    es_pasillo = nombre_lower.startswith("pasillo")  # detecta pasillo

    # Regla 1 enchufe cada 8m SOLO para: dormitorios/living/comedor/estar
    # NO aplicar en pasillo ni baño
    aplica_regla_dt = (es_dormitorio or es_living or es_comedor or es_sala_estar) and (not es_pasillo) and (not es_bano)  # esta regla exige enchufes dobles/triples cada 8m, pero solo en estos ambientes
    req_dt = int(math.ceil(perimetro / 8.0)) if (aplica_regla_dt and perimetro > 0) else 0  # cantidad mínima de enchufes exigida por el perímetro (1 cada 8m, redondeado hacia arriba)
    # mínimo por perímetro (solo si aplica regla)
    if aplica_regla_dt and perimetro > 0:  # el mínimo por perímetro solo se calcula si el ambiente entra en la regla de los 8m
        min_ench_ric = max(1, math.ceil(perimetro / 8.0))  # al menos 1 enchufe si aplica la regla y el ambiente tiene perímetro
    else:  # el ambiente no entra en la regla de los 8m
        min_ench_ric = 0  # si no aplica la regla, no hay mínimo por perímetro

    # CASO ESPECIAL: PASILLO
    if es_pasillo:  # el pasillo tiene una regla distinta: puede no tener enchufes
        while True:  # repite hasta que responda si o no
            # pregunta si el pasillo tiene enchufes (a diferencia de otros ambientes,
            # en el pasillo los enchufes son opcionales)
            resp_pasillo = input(
                f"   - ¿El pasillo '{nombre}' tiene enchufes? (si/no): "
            ).strip().lower()  # normaliza el si/no del pasillo

            if resp_pasillo in ["si", "no"]:  # respuesta válida, sigue
                break  # respondió si o no, sigue
            print("     ! Responda solo 'si' o 'no'.")  # respuesta inválida, vuelve a preguntar

        if resp_pasillo == "no":  # si no tiene enchufes, la cantidad queda en 0 y no se pregunta más
            cantidad_enchufes = 0  # pasillo sin enchufes, queda en cero
        else:  # si sí tiene, ahora se pregunta cuántos
           # si dice que SI, recién preguntamos cuántos
            while True:  # repite hasta que ingrese un entero válido
                try:  # intenta leer la cantidad de enchufes del pasillo
                    # pide la cantidad de enchufes del pasillo (debe ser un entero >= 0)
                    cantidad_enchufes = int(
                        input(f"   - Cantidad de enchufes del pasillo '{nombre}': ")  # pregunta cuántos enchufes lleva el pasillo
                    )
                    if cantidad_enchufes < 0:  # no puede haber una cantidad negativa de enchufes
                        print("     ! No puede ser negativa.")  # aviso y vuelve a preguntar
                        continue  # cantidad negativa, vuelve a preguntar
                    break  # cantidad válida, sigue
                except:  # no era un entero, cae acá
                    print("     ! Ingrese un número entero válido.")  # no era un entero, vuelve a preguntar
    else:  # si NO es pasillo, se usa la pregunta normal de cantidad de enchufes
        while True:  # repite hasta que la cantidad cumpla todas las reglas
            try:  # intenta leer la cantidad de enchufes del ambiente
                cantidad_enchufes = int(input(f"   - Cantidad de enchufes comunes del ambiente '{nombre}': "))  # pregunta cuántos enchufes comunes tiene el ambiente
                if cantidad_enchufes < 0:  # no puede haber una cantidad negativa de enchufes
                    print("     ! La cantidad de enchufes no puede ser negativa.")  # aviso y vuelve a preguntar
                    continue  # cantidad negativa, vuelve a preguntar

                if es_cocina and cantidad_enchufes < 3:  # la cocina debe tener mínimo 3 enchufes según RIC N°10
                    print("     ! Según RIC N°10, la cocina debe tener al menos 3 enchufes.")  # aviso del mínimo de enchufes de la cocina
                    continue  # la cocina no llega a los 3 enchufes, vuelve a preguntar

                if es_lavadero and cantidad_enchufes < 1:  # el lavadero debe tener mínimo 1 enchufe
                    print("     ! El ambiente 'Lavadero' debe tener al menos 1 enchufe.")  # aviso del mínimo de enchufes del lavadero
                    continue  # el lavadero no llega a 1 enchufe, vuelve a preguntar

                if aplica_regla_dt and req_dt > 0 and cantidad_enchufes < req_dt:  # si aplica la regla de 1 cada 8m, valida que alcancen los enchufes ingresados
                    print(f"     ! Según RIC, en '{nombre}' necesitas al menos {req_dt} enchufe(s) "  # avisa cuántos enchufes dobles o triples exige el perímetro
                        f"(doble o triple) (1 por cada 8m de perímetro).")
                    continue  # faltan enchufes dobles/triples, vuelve a preguntar

                if min_ench_ric > 0 and cantidad_enchufes < min_ench_ric:  # vuelve a validar el mínimo por perímetro
                    print(f"     ! Según RIC, en '{nombre}' se requiere al menos "  # aviso del mínimo por perímetro
                        f"{min_ench_ric} enchufe(s) (1 cada 8 m de perímetro o fracción).")
                    continue  # no alcanza el mínimo por perímetro, vuelve a preguntar
                break  # la cantidad paso todas las validaciones ric
            except:  # lo ingresado no era un entero
                print("     ! Ingrese un número entero válido para la cantidad de enchufes.")  # no era un entero, vuelve a preguntar

    potencia_enchufes = []  # potencia de cada enchufe común, se completa más abajo
    enchufes_detalle_amb = []  # detalle completo (id, módulos, potencias) de cada enchufe común
    dobles_triples_cocina = 0  # cuenta cuántos enchufes dobles/triples lleva la cocina (necesita al menos 3)
    lavadero_doble_triple = False  # indica si el lavadero ya tiene su enchufe doble/triple obligatorio
    # contador de enchufes dobles/triples para dormitorios/living/comedor/estar
    dobles_triples_dt = 0  # cuenta cuántos enchufes dobles/triples lleva el ambiente con regla de 8m

    # Pregunta, enchufe por enchufe, cuántos módulos tiene (simple/doble/triple)
    # y su potencia, validando en el camino las reglas de cocina, lavadero y
    # la regla de 1 doble/triple cada 8m de perímetro.
    for j in range(1, cantidad_enchufes + 1):
        while True:  # repite hasta que ingrese 1, 2 o 3
            try:  # intenta leer los módulos del enchufe
                mod = int(input(f"     • Módulos del enchufe común {j} (1=simple, 2=doble, 3=triple): "))  # pregunta si el enchufe es simple, doble o triple
            except:  # no era un número válido, cae acá
                print("         ! Valor inválido. Ingrese 1, 2 o 3.")  # valor inválido, vuelve a preguntar
                continue  # vuelve a pedir el dato
            if mod not in [1, 2, 3]:  # solo se aceptan simple (1), doble (2) o triple (3)
                print("         ! Ingrese solo 1, 2 o 3.")  # avisa que ese número de módulos no sirve
                continue  # vuelve a preguntar los módulos

            if es_lavadero:  # en el lavadero, el último enchufe no puede ser simple si aún falta el doble/triple obligatorio
                restantes = cantidad_enchufes - j  # cuántos enchufes quedan por ingresar después de este
                if (not lavadero_doble_triple) and (restantes == 0) and (mod == 1):  # es el último enchufe y todavía no hay ninguno doble o triple
                    print("         ! En Lavadero debe existir al menos 1 enchufe DOBLE (2) o TRIPLE (3) de 16A.")  # le recuerda la exigencia del lavadero
                    print("           → Este último NO puede ser simple. Ingrese 2 o 3.")  # le avisa que este no puede quedar simple
                    continue  # vuelve a pedir los módulos de este enchufe

            if es_cocina:  # en la cocina se exige que al menos 3 enchufes sean dobles o triples
                faltantes = 3 - dobles_triples_cocina  # cuántos dobles/triples faltan todavía para cumplir el mínimo de 3
                restantes_incluyendo_este = cantidad_enchufes - j + 1  # cuántos enchufes quedan por ingresar, contando este
                max_dobles_posibles = restantes_incluyendo_este - 1  # cuántos de los que quedan podrían ser dobles/triples como máximo

                if mod == 1 and faltantes > max_dobles_posibles:  # poniendo simple ya no alcanzan los que quedan para llegar a 3 dobles/triples
                    print("         ! En cocina, para cumplir RIC N°10, aquí debes ingresar DOBLE (2) o TRIPLE (3).")  # obliga a doble o triple en cocina
                    continue  # vuelve a pedir el dato

            if aplica_regla_dt and req_dt > 0:  # si aplica la regla de 1 doble/triple cada 8m, valida que se pueda cumplir todavía
                faltan_dt = req_dt - dobles_triples_dt  # cuántos dobles/triples faltan para cumplir la regla de 8m
                # cuántos enchufes quedan contando este (incluye el actual)
                restantes_incluyendo_este = cantidad_enchufes - j + 1  # cuántos enchufes quedan por ingresar, contando este
                # máximo de dobles/triples que aún podrías lograr si desde ahora TODOS fueran 2 o 3
                max_dt_posible = restantes_incluyendo_este
                # si el usuario pone simple y con eso ya no alcanza a cumplir, se bloquea
                if mod == 1 and faltan_dt > (max_dt_posible - 1):
                    print(f"         ! En '{nombre}' debes cumplir {req_dt} enchufe(s) DOBLE/TRIPLE "  # le recuerda cuántos dobles/triples pide el perímetro
                        f"(1 por cada 8m de perímetro).")
                    print("           → Aquí NO puedes ingresar simple. Ingrese 2 o 3.")  # le avisa que aquí no puede ir simple
                    continue  # vuelve a pedir los módulos
            break  # los módulos pasaron todas las validaciones

        if es_cocina and mod >= 2:  # en cocina, un doble o triple suma al mínimo de 3
            dobles_triples_cocina += 1  # suma un enchufe doble/triple más para la cocina

        if es_lavadero and mod >= 2:  # el lavadero ya cumple con su doble/triple obligatorio
            lavadero_doble_triple = True  # marca que el lavadero ya tiene su doble/triple obligatorio

        if aplica_regla_dt and mod >= 2:  # este enchufe sirve para la regla de 1 doble/triple cada 8m
            dobles_triples_dt += 1  # suma un enchufe doble/triple más para la regla de 8m

        # Ahora se pregunta la potencia de este enchufe: si el usuario no la sabe,
        # se asignan 250W por norma; si la sabe, se pide módulo por módulo.
        while True:  # se repite hasta que la potencia del enchufe sea válida
            # UNA sola pregunta por enchufe
            while True:  # se repite hasta que responda si o no
                # pregunta si conoce la potencia total del enchufe; si dice que no,
                # se le asignan los 250W por norma sin preguntar módulo por módulo
                conoce_total = input(
                    f"       ¿Conoce la potencia TOTAL del enchufe {j}? (si/no)\n"
                    f"       (Si no lo sabe, se asignarán 250W por norma al enchufe completo): "
                ).strip().lower()  # deja el si/no de la potencia en minúsculas
                if conoce_total in ["si", "no"]:  # respuesta válida
                    break  # sigue con la potencia
                print("         ! Responda solo 'si' o 'no'.")  # cualquier otra cosa no sirve

            potencias_modulo = []  # acá se guardan las potencias de cada módulo si el usuario las conoce

            if conoce_total == "no":  # no sabe la potencia del enchufe
                # 250W POR ENCHUFE (no por módulo)
                potencia_total_enchufe = 250.0  # 250 W por enchufe completo, por norma
                print("         -> No se conoce la potencia, se asignan 250 W por defecto (por enchufe).")  # le avisa el valor que se le asigno
            else:  # si conoce la potencia, se pide módulo por módulo
                # solo si conoce: preguntar por cada módulo
                for m in range(1, mod + 1):  # pregunta la potencia de cada módulo por separado
                    p_mod = float(input(f"         Potencia del módulo {m} [W]: "))  # potencia que consume ese módulo, en watts
                    potencias_modulo.append(p_mod)  # la va guardando para sumarla después
                potencia_total_enchufe = sum(potencias_modulo)  # la potencia total del enchufe es la suma de sus módulos

            if (es_cocina or es_lavadero) and potencia_total_enchufe > 3200:  # en cocina y lavadero, cada enchufe no puede superar 3200W (para que alcance el circuito de 16A)
                print("       ! En cocina y lavadero la potencia total de cada enchufe "  # le explica por qué no se acepta ese enchufe
                      "no puede ser mayor a 3200 W para asegurar circuitos de 16A.")
                print("         → Vuelva a ingresar los módulos y potencias de este enchufe.")  # le pide reingresar el enchufe completo
                continue  # vuelve al inicio de este enchufe
            break  # potencia aceptada, sale del ciclo

        potencia_enchufes.append(potencia_total_enchufe)  # guarda la potencia total de este enchufe en la lista del ambiente
        # guarda el detalle completo de este enchufe (id, módulos, potencias),
        # para usarlo más adelante en el listado de materiales
        enchufes_detalle_amb.append({
            "id": j,  # número del enchufe dentro del ambiente
            "modulos": mod,  # 1 simple, 2 doble, 3 triple
            "potencias_modulo": potencias_modulo,  # potencias de cada módulo, queda vacío si se uso el 250W por norma
            "potencia_total": potencia_total_enchufe  # potencia del enchufe completo, en watts
        })  # cierra el registro de este enchufe del ambiente

    if es_cocina and dobles_triples_cocina < 3:  # si la cocina no llegó a 3 enchufes dobles/triples, se avisa (falta reingresar)
        print("\n      Según RIC N°10, en la cocina debe haber al menos 3 enchufes DOBLES o TRIPLES.")  # avisa que la cocina quedó bajo lo que pide RIC N°10
        print("     → Debes reingresar los enchufes de cocina.")  # tiene que volver a ingresarlos

    if es_lavadero and not lavadero_doble_triple:  # si el lavadero no tiene su doble/triple obligatorio, se avisa
        print("\n     ! Según norma, en Lavadero debe haber al menos 1 enchufe DOBLE o TRIPLE 16A.")  # avisa que falta el doble/triple del lavadero
        print("       → Debes reingresar los enchufes del Lavadero.")  # tiene que reingresar los del lavadero

    if aplica_regla_dt and req_dt > 0 and dobles_triples_dt < req_dt:  # si no se cumplió la regla de 1 doble/triple cada 8m, se avisa
        print(f"\n     ! Según RIC, en '{nombre}' debe haber al menos {req_dt} enchufe(s) DOBLE o TRIPLE "  # avisa cuántos dobles/triples exige el perímetro del ambiente
            f"(1 por cada 8m de perímetro).")
        print("       → Debes reingresar los enchufes de este ambiente.")  # tiene que reingresar los enchufes del ambiente

    potencia_total_ench = sum(potencia_enchufes)  # suma la potencia de todos los enchufes de este ambiente
    detalle_enchufes = "\n".join([f"Enchufe {idx+1} ({p}W)" for idx, p in enumerate(potencia_enchufes)]) if potencia_enchufes else "Ninguno"  # texto con el detalle de cada enchufe, para mostrar en el informe
    enchufes_por_ambiente[nombre] = enchufes_detalle_amb[:]  # guarda copia de los enchufes de este ambiente, para repartirlos después en los circuitos

    nombres_componentes = "\n".join([f"{n} ({p}W)" for n, p in zip(componentes_nombres, componentes_potencias)]) if componentes_nombres else "Ninguno"  # texto con los componentes especiales (horno, calefont, etc.) de este ambiente
    potencia_total_ambiente = potencia_total_ilum + potencia_total_ench + potencia_total_comp  # potencia total del ambiente: iluminación + enchufes + componentes especiales

    # --- CONTEOS REALES PARA MATERIALES ---
    # Cantidades resumidas del ambiente. De todas éstas, build_materiales_df()
    # solo lee después las columnas "Cantidad luminarias (u)" y "Cantidad
    # enchufes (u)" del DataFrame de ambientes; las demás quedan como columnas
    # informativas de la hoja "Informe".
    n_luminarias = cantidad_luminarias  # cuántas luminarias quedaron en el ambiente
    n_enchufes = cantidad_enchufes  # cuántos enchufes quedaron en el ambiente
    n_modulos_enchufe = sum(e.get("modulos", 1) for e in enchufes_detalle_amb) if enchufes_detalle_amb else 0  # suma los módulos de cada enchufe (simple=1, doble=2, triple=3); queda solo como dato del informe
    n_componentes_especiales = len(componentes_detalle_amb) if componentes_detalle_amb else 0  # cuántos equipos especiales tiene el ambiente (horno, calefont, etc.)

    # Cuenta los enchufes por tipo (simple/doble/triple) y amperaje: 10/16A si
    # el ambiente es cocina o lavadero, 10A en el resto. Estos 6 contadores
    # terminan como columnas de la hoja "Informe" y nada más: el amperaje del
    # enchufe que se compra lo vuelve a decidir build_materiales_df() circuito
    # por circuito, según el TM de cada circuito.
    ench_simple_10A = 0  # enchufes de 1 módulo, 10A (ambiente que no es cocina ni lavadero)
    ench_doble_10A = 0  # enchufes dobles de 10A
    ench_triple_10A = 0  # enchufes triples de 10A
    ench_simple_1016 = 0  # enchufes de 1 módulo, 10/16A (solo cocina y lavadero)
    ench_doble_1016 = 0  # enchufes dobles de 10/16A
    ench_triple_1016 = 0  # enchufes triples de 10/16A

    es_coc_lav_amb = es_cocina or es_lavadero  # true si el ambiente es cocina o lavadero (llevan enchufes 10/16A)
    for e in enchufes_detalle_amb:  # recorre los enchufes ya ingresados del ambiente
        mod = int(e.get("modulos", 1))  # cantidad de módulos de este enchufe: 1=simple, 2=doble, 3=triple
        if es_coc_lav_amb:  # cocina y lavadero llevan enchufes 10/16A
            # en cocina/lavadero los enchufes son 10/16A
            if mod == 1:  # un módulo: enchufe simple
                ench_simple_1016 += 1  # suma a los simples 10/16A
            elif mod == 2:  # dos módulos: enchufe doble
                ench_doble_1016 += 1  # suma a los dobles 10/16A
            elif mod == 3:  # tres módulos: enchufe triple
                ench_triple_1016 += 1  # suma a los triples 10/16A
        else:  # resto de los ambientes, ahí los enchufes son de 10A
            # en el resto de los ambientes los enchufes son 10A
            if mod == 1:  # un módulo: enchufe simple
                ench_simple_10A += 1  # suma a los simples 10A
            elif mod == 2:  # dos módulos: enchufe doble
                ench_doble_10A += 1  # suma a los dobles 10A
            elif mod == 3:  # tres módulos: enchufe triple
                ench_triple_10A += 1  # suma a los triples 10A

    # guarda todos los datos calculados de este ambiente, para armar después el DataFrame ambientes_df
    ambientes_detalle.append({
        "Ambiente": nombre,  # nombre del ambiente tal como lo escribió el usuario
        "Área (m²)": dimension,  # superficie en m2; sumada con la de los demás ambientes da area_total_vivienda, que define min_circuitos
        "Perímetro (m)": perimetro,  # perímetro en m; queda como dato del informe (la regla de 1 doble/triple cada 8m ya se aplicó arriba con la variable perímetro)
        "Material tabique": material_tabique,  # de que es el tabique, define el tipo de fijación
        "Material forrado interior": material_forrado_interior,  # forrado interior, también influye en los materiales
        "Detalle iluminación": detalle_iluminacion,  # texto con las luminarias del ambiente
        "Potencia iluminación (W)": potencia_total_ilum,  # watts totales de iluminación del ambiente

        "N_conmutadas_924 (u)": int(N_conmutadas_924),  # cuántas conmutadas 9/24 quedaron en el ambiente

        "L_viajeros_924 (m)": L_viajeros_924,  # metros de cable viajero entre los dos conmutadores
        "L_retorno_lampara (m)": L_retorno_lampara,  # metros de retorno desde el interruptor a la lámpara
        "L_fase_caja_primer_int (m)": L_fase_caja_primer_int,  # metros de fase desde la caja hasta el primer interruptor
        "L_troncal_primera_oct_924 (m)": L_troncal_primera_oct_924,  # metros del troncal hasta la primera octogonal de la conmutada

        "L_caja_int_fase_ida (m)": L_caja_int_fase_ida,  # metros de ida de la fase desde la caja al interruptor
        "L_troncal_oct1 (m)": L_troncal_oct1,  # metros del troncal hasta la primera caja octogonal
        "L_oct1_oct2 (m)": L_oct1_oct2,  # suma de todos los tramos entre octogonales seguidas del ambiente (de todos los grupos de interruptor)
        "Grupos_interruptor_ilum": grupos_ilum_detalle,  # que luminarias controla cada interruptor

        "Detalle enchufes comunes": detalle_enchufes,  # texto con cada enchufe y su potencia
        "Potencia enchufes (W)": potencia_total_ench,  # watts totales de los enchufes del ambiente
        "Componentes especiales": nombres_componentes,  # nombres de los equipos especiales del ambiente
        "Potencia comp. especiales (W)": potencia_total_comp,  # watts de los equipos especiales
        "Total por ambiente (W)": potencia_total_ambiente,  # watts totales del ambiente, para el cuadro de cargas
        "Cantidad luminarias (u)": n_luminarias,  # cantidad de luminarias; build_materiales_df() sí lee esta columna
        "Cantidad enchufes (u)": n_enchufes,  # cantidad de enchufes; build_materiales_df() sí lee esta columna (cajas y tapas)
        "Cantidad módulos enchufe (u)": n_modulos_enchufe,  # módulos totales del ambiente; dato informativo, no se usa para comprar
        "Cantidad comp. especiales (u)": n_componentes_especiales,  # cantidad de equipos especiales; dato informativo
        "Enchufes simples 10A (u)": ench_simple_10A,  # solo informativo, no lo lee el cálculo de materiales
        "Enchufes dobles 10A (u)": ench_doble_10A,  # solo informativo
        "Enchufes triples 10A (u)": ench_triple_10A,  # solo informativo
        "Enchufes simples 10/16A (u)": ench_simple_1016,  # solo informativo (los 10/16A que finalmente se compran los define el cálculo de materiales)
        "Enchufes dobles 10/16A (u)": ench_doble_1016,  # solo informativo
        "Enchufes triples 10/16A (u)": ench_triple_1016,  # solo informativo
    })  # cierra la ficha con todos los datos de este ambiente

# Elementos restantes para repartir entre circuitos
# (a medida que se van armando los circuitos, se van sacando de acá los que ya se usaron)
enchufes_restantes = {k.lower(): v[:] for k, v in enchufes_por_ambiente.items()}  # copia de enchufes por ambiente, con nombre en minúscula como clave
luminarias_restantes = {k.lower(): v[:] for k, v in luminarias_por_ambiente.items()}  # lo mismo con las luminarias que todavía no se asignan a ningún circuito
componentes_restantes = {k.lower(): v[:] for k, v in componentes_por_ambiente.items()}  # lo mismo con los equipos especiales que faltan repartir

# -------- PARTE 2: DATOS ADICIONALES DEL SISTEMA --------
# Termina la parte de "ambientes" (iluminación/enchufes por pieza) y empieza
# a preguntar datos generales de toda la instalación: zona, canalización,
# cantidad de circuitos, etc.
print("\n Ahora se solicitarán datos generales del sistema")  # avisa que se pasa a los datos generales de la instalación

area_total_vivienda = sum(a["Área (m²)"] for a in ambientes_detalle)  # suma el área de todos los ambientes ingresados
min_circuitos = 2 if area_total_vivienda < 30 else 3  # criterio propio por superficie: bajo 30 m² se exigen 2 circuitos, de 30 m² para arriba 3; es el piso que se válida al preguntar cantidad_circuitos

print(f"\n Superficie total aproximada de la vivienda: {area_total_vivienda:.2f} m²")  # le muestra la superficie total calculada
print(f" Según criterio, se requieren al menos {min_circuitos} circuitos.")  # le muestra el mínimo de circuitos que exige el criterio

zona = input("- Ingrese las características de la zona (húmeda, seca): ")  # zona húmeda o seca: solo cambia el tipo de cable, la tabla de corriente es la misma
while True:  # se repite hasta que escriba una canalización válida
    tipo_canalizacion = input("- Ingrese el tipo de canalización (embutida, sobrepuesta): ").strip().lower()  # embutida o sobrepuesta, cambia el conduit y los accesorios
    if "embut" in tipo_canalizacion or "sobre" in tipo_canalizacion:  # acepta que escriba solo 'embut' o 'sobre'
        break  # canalización válida
    print("   ! Debe ingresar 'embutida' o 'sobrepuesta'.")  # cualquier otra palabra no sirve
material_forrado_exterior = input("- Material del forrado exterior de la casa (Fibrocemento, Madera, Siding PVC, Siding metálico): ").strip().lower()  # de esto depende el tipo de tornillo y tarugo que se calcula después

while True:  # se repite hasta que la cantidad de circuitos sea válida
    try:  # por si escribe algo que no es número
        cantidad_circuitos = int(input("- Ingrese la cantidad de circuitos (máx. 10): "))  # cuántos circuitos va a llevar el tablero
        if cantidad_circuitos < min_circuitos:  # no puede quedar bajo el mínimo según la superficie
            print(f"   ! La vivienda requiere al menos {min_circuitos} circuitos.")  # le avisa el mínimo
            continue  # vuelve a preguntar
        if cantidad_circuitos > 10:  # el tablero de este informe se limita a 10 circuitos
            print("   ! El máximo permitido es 10 circuitos.")  # le avisa el tope
            continue  # vuelve a preguntar
        break  # cantidad válida
    except:  # no ingreso un número
        print("   ! Ingrese un número entero válido.")  # le pide un entero

amb_idx = {a["Ambiente"].lower(): a for a in ambientes_detalle}  # para buscar rápido un ambiente por nombre
nombres_amb = [a["Ambiente"] for a in ambientes_detalle]  # lista simple con los nombres de los ambientes, tal como los escribió el usuario

def normaliza_seleccion_ambientes(texto, disponibles):  # deja lo que escribió el usuario como nombres de ambiente válidos
    # toma lo que escribió el usuario (ambientes separados por coma) y
    # deja solo los nombres que existen de verdad, sin repetir
    disp_lower = {d.lower(): d for d in disponibles}  # relaciona el nombre en minúscula con el nombre real (con mayúsculas)
    salida = []  # acá se van guardando los ambientes que si existen
    for t in [s.strip() for s in texto.split(",") if s.strip()]:  # separa por coma y limpia espacios
        key = t.lower()  # el nombre en minúscula, para poder comparar
        if key in disp_lower and disp_lower[key] not in salida:  # existe y no está repetido
            salida.append(disp_lower[key])  # guarda el nombre real del ambiente
    return salida  # devuelve la lista limpia de ambientes

circuitos = []  # lista global donde se van guardando todos los circuitos creados

# función corta para no repetir el mismo diccionario cada vez que se
# crea un circuito nuevo
def add_circuito(nombre_circ, longitud, potencia_estimada, es_ilumin, es_enchufe, es_especial, items=None, detalle_asignacion="", tiene_tramo_20m=True):
    # agrega un circuito nuevo a la lista global "circuitos", que después
    # se convierte en circuitos_df y se usa en build_materiales_df()
    # tiene_tramo_20m: si el circuito NO tiene ningún tramo continuo >=20m
    # (aunque la longitud TOTAL del circuito sea mayor a 20m), no debe
    # contarse ninguna caja de paso (RIC 7.16.1.13) — por defecto True para
    # no romper los llamados que todavía no preguntan esto explícitamente.
    circuitos.append({
        "Circuito": nombre_circ,  # nombre del circuito, ej: C1 enchufes
        "Longitud (m)": float(longitud),  # largo total del circuito, en metros
        "Potencia estimada (W)": round(float(potencia_estimada), 1),  # potencia que va a alimentar, en watts
        "Detalle asignación": detalle_asignacion,  # texto de que quedó conectado en este circuito
        "es_ilumin": es_ilumin,  # marca si es circuito de iluminación
        "es_enchufe": es_enchufe,  # marca si es circuito de enchufes
        "tipo_dif": "especial" if es_especial else "general",  # define qué tipo de diferencial le corresponde
        "_items": items[:] if items else [],  # copia de los items (luminarias/enchufes) del circuito
        "_cajas_adic": _cajas_adic if es_enchufe else 0,  # cajas de derivación extra, solo aplica a enchufes
        "_tiene_tramo_20m": bool(tiene_tramo_20m),  # si no hay ningún tramo de 20m, después no se cuentan cajas de paso
    })  # cierra el registro de este circuito

# cocina/lavadero por ambiente (cocina 1, cocina 2, etc.)
enchufes_coc_lav_por_amb = {}   # {"cocina 1":[...], "cocina 2":[...], "lavadero":[...]}
long_real_coc_lav_por_amb = {}  # {"cocina 1": 25.0, "cocina 2": 18.0, ...}
coc_lav_creado_manual = set()  # ambientes de cocina/lavadero a los que el usuario ya les creo el circuito a mano


CALIBRES_TM_CLIMA = [6, 10, 16, 20, 25, 32, 40, 50, 63]  # calibres comerciales de termomagnético (A). Lo usan tanto climatización como agua caliente, a través de siguiente_calibre_clima()

def siguiente_calibre_clima(corriente):  # elige el TM comercial que le corresponde al circuito de clima
    # busca el calibre de TM comercial más chico que aguante esta corriente
    for cal in CALIBRES_TM_CLIMA:  # va probando de menor a mayor
        if cal >= corriente:  # este calibre ya queda sobre la corriente
            return cal  # este calibre ya aguanta la corriente
    return CALIBRES_TM_CLIMA[-1]  # ninguno alcanzó, se usa el más grande

def calcular_circuito_climatizacion(datos):  # todo el cálculo del circuito de aire acondicionado
    """
    Calcula conductor, termomagnético y diferencial según RIC N°07.
    Lógica TM:
      - Conductor: I_max × 1.25 (RIC 7.3.4)
      - Curva: C para inverter (arranque suave, 5-10× In) / D para on-off
        (arranque fuerte, 10-20× In) — RIC N°10
      - TM: calibre sobre I_diseño, luego verificado contra LRA (RIC 5.6.2.2)
        Si el LRA real viene de la placa del equipo, se usa directo
        Si no hay LRA, se estima según el tipo de compresor on/off
        Si es inverter, no hay LRA, la curva C es suficiente
      - Si el LRA supera el umbral magnético del TM elegido (según su curva), se sube al siguiente calibre
    """
    V             = float(datos.get("tension", 220))  # tensión de alimentacion del equipo, normalmente 220V
    I_max         = float(datos.get("corriente_maxima", 0))       # corriente máxima de placa (si la tiene)
    I_nom         = float(datos.get("corriente_nominal", 0))      # corriente nominal de placa
    P_nom_w       = float(datos.get("potencia_nominal_w", 0))     # potencia nominal, por si no hay corriente
    fp            = float(datos.get("factor_potencia", 0.95)) or 0.95  # factor de potencia, 0.95 si no viene en la placa
    tecnologia    = str(datos.get("tecnologia", "inverter")).lower()  # inverter u on/off
    lra           = datos.get("lra")   # corriente de arranque real (si el fabricante la entrega)
    mocp          = datos.get("mocp")  # protección máxima que indica el fabricante (si la entrega)
    tipo_compresor = str(datos.get("tipo_compresor", "rotativo")).lower()  # rotativo o de pistón, cambia el LRA estimado
    temp_sobre_30 = datos.get("temp_sobre_30", False)  # si el equipo queda expuesto a más de 30°C
    temp_ambiente = datos.get("temp_ambiente", 25.0)  # temperatura del lugar donde queda instalado el equipo

    # --- Paso 1: corriente base desde placa (RIC 5.3.1) ---
    # usa la corriente máxima de placa si existe; si no, la calcula desde la
    # potencia nominal, y si tampoco hay potencia, usa la corriente nominal
    I_base = I_max if I_max > 0 else (P_nom_w / (V * fp) if P_nom_w > 0 else I_nom)

    # --- Paso 2: corriente de diseño conductor (RIC 7.3.4) ---
    I_diseno = I_base * 1.25  # se suma un 25% de margen sobre la corriente base

    # --- Paso 3: factor corrección temperatura (RIC N°4 art. 6.2.6 / Tabla N°4.7) ---
    # Fórmula RIC: Ic = Iz × ft  (ft MULTIPLICA sobre Iz, NO divide I_diseno)
    # La verificación real la hace la PARTE 3.2 con su propio ft (temperatura
    # global del proyecto y método de canalización); este ft queda informativo.
    if temp_sobre_30:  # solo se corrige si el equipo trabaja sobre 30°C
        ft = factor_temperatura_ft(temp_ambiente, "B1")  # busca el factor de la Tabla 4.7 para método B1; queda como dato, no elige la sección
    else:  # temperatura normal, no se corrige nada
        ft = 1.0  # sin corrección, temperatura normal
    I_diseno_corr = I_diseno  # NO se divide por ft — ft va sobre Iz en verificación

    # --- Paso 4: sección mínima conductor ---
    seccion = 2.5  # mínimo RIC 7.3.4; más abajo, en la PARTE 3.2, se puede subir por caída de tensión

    # --- Paso 5: termomagnético ---
    # 5a. Si el fabricante indica un MOCP, ese valor no se puede superar (RIC 5.6.2.3)
    if mocp and float(mocp) > 0:  # el fabricante indico la protección máxima (MOCP)
        In_tm = siguiente_calibre_clima(float(mocp))  # primer calibre comercial que llega al MOCP (ojo: puede quedar por encima, no lo toma como tope)
        curva_tm = "C" if tecnologia == "inverter" else "D"  # inverter parte suave, el on/off necesita curva D
        nota_tm = f"Protección máxima indicada por fabricante: {float(mocp)} A → calibre comercial {In_tm} A (RIC 5.6.2.3)"  # texto para el informe explicando de donde salió ese TM
        aviso_lra = "Protección máxima según fabricante"  # no hace falta revisar LRA, manda lo del fabricante
        # Diferencial y retorno directo
        In_dif = 25 if In_tm <= 25 else (40 if In_tm <= 40 else 63)  # calibre del diferencial según el TM elegido
        I_cuadro = I_max if I_max > 0 else (I_nom if I_nom > 0 else (P_nom_w / (V * fp) if P_nom_w > 0 else I_base))  # corriente que se lleva al cuadro de cargas
        return {  # devuelve el circuito de clima ya resuelto con el tope del fabricante
            "I_base_A": round(I_base, 2), "I_diseno_A": round(I_diseno_corr, 2),  # corriente base de placa y corriente de diseño del conductor
            "seccion_mm2": seccion, "In_tm_A": In_tm, "curva_tm": curva_tm,  # sección del conductor, calibre y curva del TM
            "nota_tm": nota_tm, "aviso_lra": aviso_lra, "In_dif_A": In_dif,  # textos del informe y calibre del diferencial
            "I_cuadro_A": round(I_cuadro, 2), "ft": ft, "temp_sobre_30": temp_sobre_30,  # corriente del cuadro y factor de corrección por temperatura
        }

    # 5b. Calcular TM inicial sobre corriente de diseño
    In_tm = siguiente_calibre_clima(I_diseno_corr)  # primer TM tanteado según la corriente de diseño
    # Inverter: arranque suave, no tiene LRA real, curva C directo.
    # On/Off: se calcula el LRA y se prueba primero con curva C (5-10× In,
    # dispara más rápido); si el LRA no cabe en ese umbral, se pasa a curva D
    # (10-20× In) antes de subir de calibre — RIC N°10.
    curva_tm = "C" if tecnologia == "inverter" else None  # on/off se define más abajo

    # --- Paso 6: verificación LRA (RIC 5.6.2.2) ---
    # Curva C: disparo magnético instantáneo entre 5× y 10× In_tm
    # Curva D: disparo magnético instantáneo entre 10× y 20× In_tm
    aviso_lra = ""  # acá queda el texto sobre la corriente de arranque

    if tecnologia == "inverter":  # equipo inverter
        # Inverter: arranque progresivo, no existe LRA real
        aviso_lra = "Inverter — arranque progresivo"  # no se revisa LRA porque el inverter parte de a poco

    else:  # equipo on/off
        # On/off: existe corriente de arranque real
        if lra and float(lra) > 0:  # el fabricante entrego el LRA de placa
            # LRA real desde placa
            lra_val = float(lra)  # corriente de arranque real, en amperes
            fuente_lra = f"LRA real de placa: {lra_val} A"  # deja anotado que el LRA es de placa
        else:  # no hay LRA de placa, hay que estimarlo
            # Estimar LRA según tipo de compresor (RIC 5.6.2.2 — criterio conservador)
            if "piston" in tipo_compresor or "pistón" in tipo_compresor:  # compresor de pistón, arranca más fuerte
                factor_lra = 7.0  # compresor de pistón: LRA típico 6-8× I_nom
                fuente_lra = f"LRA estimado (compresor pistón, factor 7×): {round(I_nom * factor_lra, 1)} A"  # deja anotado que el LRA es estimado
            else:  # cualquier otro caso se toma como rotativo
                factor_lra = 5.5  # compresor rotativo: LRA típico 4-6× I_nom
                fuente_lra = f"LRA estimado (compresor rotativo, factor 5.5×): {round(I_nom * factor_lra, 1)} A"  # deja anotado el LRA estimado del rotativo
            lra_val = round(I_nom * factor_lra, 1)  # LRA estimado a partir de la corriente nominal

        # Se prueba primero curva C (dispara más rápido, más protección) y
        # solo si el LRA no cabe en su umbral se pasa a curva D — subiendo de
        # calibre dentro de cada curva antes de descartarla.
        curva_tm = None  # todavía no se sabe que curva va a quedar
        for curva_candidata in ("C", "D"):  # primero se prueba C y si no da, se pasa a D
            In_tm_prueba = In_tm  # parte del TM calculado por corriente de diseño
            intentos = 0  # tope de intentos para no quedarse pegado subiendo calibres
            while intentos < len(CALIBRES_TM_CLIMA):  # prueba calibre por calibre dentro de la misma curva
                umbral_max_c = In_tm_prueba * (20 if curva_candidata == "D" else 10)  # umbral de disparo magnético según la curva
                if lra_val <= umbral_max_c:  # el arranque no alcanza a hacer disparar el magnético
                    curva_tm = curva_candidata  # esta curva sirve
                    In_tm = In_tm_prueba  # y este es el calibre que queda
                    break  # listo, no hay que seguir subiendo
                idx_actual = CALIBRES_TM_CLIMA.index(In_tm_prueba) if In_tm_prueba in CALIBRES_TM_CLIMA else 0  # posición del calibre actual dentro de la lista
                if idx_actual < len(CALIBRES_TM_CLIMA) - 1:  # todavía quedan calibres más grandes por probar
                    In_tm_prueba = CALIBRES_TM_CLIMA[idx_actual + 1]  # el TM actual no aguanta el LRA, se prueba con el siguiente calibre
                    intentos += 1  # cuenta el intento
                else:  # ya no hay calibre más grande
                    break  # sale y prueba con la otra curva
            if curva_tm:  # ya se encontró curva y calibre que aguantan
                break  # no hace falta probar la curva D

        if curva_tm:  # hubo combinación que aguanta el arranque
            # se encontró una curva/calibre que sí aguanta el LRA: se arma el aviso final
            subio = In_tm != siguiente_calibre_clima(I_diseno_corr)  # avisa si el TM tuvo que subir por sobre lo que pedia la corriente de diseño
            aviso_lra = f"{fuente_lra}" + (f" | TM subido a {In_tm} A por corriente de arranque" if subio else "")  # texto final del LRA para dejar en el informe
        else:  # ninguna curva aguanto el LRA
            # Ni curva C ni D, en ningún calibre disponible, aguantan el LRA
            curva_tm = "D"  # se deja la curva más permisiva como último recurso
            aviso_lra = (f"AVISO: {fuente_lra} — supera el umbral incluso en el TM más grande "  # aviso de que ni el TM más grande aguanta el arranque
                         f"disponible ({In_tm} A, curva {curva_tm}) — revisar manualmente")

    # --- Paso 7: diferencial exclusivo (RIC 7.4.5) ---
    # el diferencial se elige por tramos según el TM: 25, 40 o 63 A
    In_dif = 25 if In_tm <= 25 else (40 if In_tm <= 40 else 63)  # diferencial exclusivo del circuito de clima, en amperes

    # --- Paso 8: corriente cuadro de cargas ---
    # Se usa I_max (corriente absorbida máxima de placa) cuando está disponible
    # porque es el valor real de operación que define la demanda (RIC 7.3.4)
    I_cuadro = I_max if I_max > 0 else (I_nom if I_nom > 0 else (P_nom_w / (V * fp) if P_nom_w > 0 else I_base))  # corriente real de operación que se lleva al cuadro de cargas

    # texto que explica por qué se eligió esa curva de TM, para dejarlo en el informe
    nota_tm = (
        f"Curva {curva_tm} — arranque suave inverter (RIC 5.6.2.1)" if tecnologia == "inverter"  # inverter: la curva C alcanza porque parte suave
        else f"Curva {curva_tm} — verificado contra LRA {'real' if lra and float(lra)>0 else 'estimado'} (RIC 5.6.2.2)"  # on/off: la curva queda justificada por el LRA
    )

    # resultado final del cálculo del circuito de climatización, listo para usar
    # en el informe y en la lista de materiales
    return {
        "I_base_A":      round(I_base, 2),  # corriente base sacada de la placa
        "I_diseno_A":    round(I_diseno_corr, 2),  # corriente de diseño del conductor
        "seccion_mm2":   seccion,  # sección del conductor, en mm^2
        "In_tm_A":       In_tm,  # calibre del termomagnético
        "curva_tm":      curva_tm,  # curva C o D del termomagnético
        "nota_tm":       nota_tm,  # explicacion de por que quedó esa curva
        "aviso_lra":     aviso_lra,  # aviso sobre la corriente de arranque
        "In_dif_A":      In_dif,  # calibre del diferencial exclusivo
        "I_cuadro_A":    round(I_cuadro, 2),  # corriente para el cuadro de cargas
        "ft":            ft,  # factor de corrección por temperatura
        "temp_sobre_30": temp_sobre_30,  # si el equipo queda expuesto a más de 30°C
    }

# Pregunta por consola todos los datos de un equipo de climatización (aire
# acondicionado) para un circuito ya identificado (nombre, ambientes, longitud
# ya vienen calculados desde donde se recorren los circuitos). Al final calcula
# TM, diferencial y sección llamando a calcular_circuito_climatizacion, y
# devuelve todo junto en un diccionario.
def ingresar_equipo_climatizacion_inline(nombre_sugerido, ambientes_str, longitud):
    """
    Flujo de preguntas de climatización integrado en la parte que recorre los circuitos.
    nombre_sugerido, ambientes_str y longitud ya vienen de ahí.
    """
    nombre_circ = nombre_sugerido  # el nombre del circuito ya viene armado desde el recorrido de circuitos

    # pregunta si el equipo es inverter u on/off, porque cambia toda la lógica de cálculo
    while True:
        tec = input(  # guarda la opción de tecnología que elige el usuario
            "   - Tipo de tecnología del equipo:\n"
            "     (1) Inverter\n"
            "     (2) On/Off\n"
            "   Opción: "
        ).strip()  # deja la opción sin espacios de más
        if tec == "1":  # opción 1: equipo inverter
            tecnologia = "inverter"  # queda guardada la tecnología porque cambia el cálculo del TM
            break  # respuesta válida, sale del ciclo
        elif tec == "2":  # opción 2: equipo on/off
            tecnologia = "on/off"  # el on/off arranca de golpe, más abajo se le pide el LRA
            break  # respuesta válida, sale del ciclo
        print("     ! Ingrese 1 o 2.")  # respuesta inválida, vuelve a preguntar

    print("\n   Ingrese los datos de la PLACA o FICHA TÉCNICA del equipo")  # avisa que los datos que siguen salen de la placa del equipo
    print("   (use el modo FRÍO que es el peor caso eléctrico):")  # en frío el equipo consume más, por eso se pide ese modo

    # corriente nominal de placa, en modo frío (el peor caso eléctrico)
    I_nom = pedir_float_positivo(
        "   - Corriente nominal de placa — modo frío [A]: "
    )
    # corriente máxima que puede llegar a consumir el equipo
    I_max = pedir_float_positivo(
        "   - Corriente máxima de placa — modo frío [A] "
        "(es la más alta que entrega la ficha): "
    )
    # potencia absorbida nominal, en Watts
    P_nom_w = pedir_float_positivo(
        "   - Potencia absorbida nominal — modo frío [W]: "
    )

    print("\n   Factor de potencia:")  # ahora viene el factor de potencia del equipo
    print("   (Si no tiene el dato en la placa, ingrese 0.95 como valor típico)")  # si la placa no lo trae, 0.95 es lo típico en splits
    fp = pedir_float_positivo("   - Factor de potencia (ej: 0.95): ")  # factor de potencia del equipo, sirve para pasar de W a A

    # LRA y tipo de compresor solo para on/off
    lra = None  # queda en None si es inverter o si la placa no trae el dato
    tipo_compresor = "rotativo"  # default
    if tecnologia == "on/off":  # solo los on/off tienen arranque directo, ahí importa el LRA
        # los inverter no necesitan esto (arranque progresivo, no hay LRA real)
        print("\n   Tipo de compresor (para estimar corriente de arranque):")  # explica para qué sirve saber el tipo de compresor
        while True:  # repite hasta que elija 1 o 2
            tc = input(  # guarda el tipo de compresor elegido
                "   - Tipo de compresor:\n"
                "     (1) Rotativo — splits residenciales modernos (LRA típico 4-6× I_nom)\n"
                "     (2) Pistón — equipos más antiguos o industriales (LRA típico 6-8× I_nom)\n"
                "   Opción: "
            ).strip()  # limpia los espacios de lo que tecleó
            if tc == "1":  # opción 1: compresor rotativo
                tipo_compresor = "rotativo"  # el rotativo arranca entre 4 y 6 veces la nominal
                break  # tipo de compresor guardado, sale del ciclo
            elif tc == "2":  # opción 2: compresor de pistón
                tipo_compresor = "piston"  # el de pistón tira más al arrancar, 6 a 8 veces la nominal
                break  # tipo de compresor guardado, sale del ciclo
            print("     ! Ingrese 1 o 2.")  # respuesta inválida, vuelve a preguntar

        print("\n   Corriente de arranque del compresor (LRA):")  # ahora pide el LRA, la corriente con el rotor trabado
        print("   Busque en la placa de la unidad EXTERIOR el campo LRA o")  # el LRA viene en la placa de la unidad exterior, no en la interior
        print("   'Locked Rotor Amps'. Si no aparece, responda no y el")  # sigue diciéndole dónde buscar el dato en la placa
        print("   programa lo estimará automáticamente.")  # si no está el dato, más abajo se estima según el compresor
        # pregunta si la placa exterior tiene el dato de LRA (corriente de arranque)
        while True:
            tiene_lra = input("   - ¿La placa exterior indica LRA? (si/no): ").strip().lower()  # respuesta si/no de si la placa trae el LRA
            if tiene_lra in ("si", "no"):  # solo acepta si o no
                break  # respondió si o no, sigue adelante
            print("     ! Responda solo si o no.")  # respuesta inválida, vuelve a preguntar
        if tiene_lra == "si":  # el usuario tiene el dato de placa, se lo pide
            while True:  # repite hasta que el LRA ingresado tenga sentido
                lra = pedir_float_positivo("     • Ingrese el valor LRA de la placa [A]: ")  # corriente de arranque del compresor, en amperes
                # Validación física: LRA siempre debe ser mayor que I_nom
                # (LRA es la corriente con rotor trabado, siempre mayor que la de operación)
                if lra <= I_nom:
                    print(f"     ! El LRA ({lra} A) no puede ser menor o igual a la")  # le avisa que ese número no puede ser el LRA
                    print(f"       corriente nominal ({I_nom} A). Esto indica que el dato")  # le muestra la nominal para que compare los dos números
                    print(f"       ingresado no es el LRA sino otro parámetro (RLA, FLA, etc).")  # lo más probable es que haya copiado la RLA o la FLA de la placa
                    print(f"       Revise la placa y vuelva a ingresar, o responda no.")  # le ofrece corregirlo o dejar que el programa lo estime
                    while True:  # pregunta si quiere reingresar el valor
                        reintentar = input("     • ¿Desea reingresar el LRA? (si/no): ").strip().lower()  # respuesta si/no de si vuelve a escribir el LRA
                        if reintentar in ("si", "no"):  # solo sale del ciclo con una respuesta válida
                            break  # ya contestó si o no, sale de la repregunta
                    if reintentar == "no":  # se rindió con la placa, no va a ingresar el LRA
                        lra = None  # sin LRA de placa, el arranque se estima por tipo de compresor
                        print("     • Se usará estimación automática según tipo de compresor.")  # le avisa que la estimación la hace el programa
                        break  # sale del ciclo del LRA, ya quedó en None
                    # si reintentar == "si", no se hace nada más acá: vuelve a subir a la
                    # pregunta "Ingrese el valor LRA de la placa" y la repite
                else:  # el LRA es mayor que la nominal, o sea el dato está bien
                    break  # LRA válido

    # MOCP para ambos tipos
    print("\n   Protección máxima recomendada por el fabricante:")
    print("   (MOCP = protección de sobrecorriente máxima — a veces indicada en placa)")  # el MOCP es el máximo que el fabricante permite de protección
    print("   (Si no aparece en la placa ni en la ficha, responda no)")  # si no aparece, se sigue sin ese tope
    tiene_mocp = input("   - ¿La placa o ficha indica un valor MÁXIMO de protección? (si/no): ").strip().lower()  # respuesta si/no de si la placa indica MOCP
    mocp = None  # queda en None si no hay tope del fabricante
    if tiene_mocp == "si":  # solo pide el valor si dijo que si
        mocp = pedir_float_positivo(  # protección máxima que indica el fabricante; sirve de referencia para elegir el calibre del TM
            "     • Ingrese ese valor máximo de protección [A]: "
        )

    # Temperatura (RIC 7.5.2 — factor corrección si supera 30°C)
    temp_sobre_30 = False  # por defecto asume recinto bajo 30°C
    temp_ambiente = 25.0  # valor por defecto
    print("\n   Temperatura ambiente del lugar de instalación:")  # pregunta por la temperatura del recinto donde va el equipo
    t_amb = input(  # respuesta si/no sobre si el recinto pasa los 30°C
        "   - ¿La temperatura ambiente del recinto supera los 30°C? (si/no)\n"
        "     (RIC 7.5.2 exige corrección si supera 30°C): "
    ).strip().lower()  # deja la respuesta en minúsculas para poder compararla
    if t_amb == "si":  # si pasa los 30°C hay que corregir la ampacidad del conductor
        temp_sobre_30 = True  # marca que después se aplica el factor de temperatura
        # se insiste hasta que ingresen un número válido y mayor que 30°C
        while True:
            try:  # si escriben cualquier cosa que no sea número, no se cae
                temp_ambiente = float(input(  # temperatura máxima del recinto, en °C
                    "   - Ingrese la temperatura ambiente máxima del recinto (°C): "
                ).strip().replace(",", "."))  # acepta la coma como separador decimal
                if temp_ambiente > 30:  # tiene que ser mayor a 30, si no no calza con lo que respondió
                    break  # temperatura válida, sigue
                print("     ! Debe ser mayor que 30°C.")  # valor fuera de rango, vuelve a pedirlo
            except:  # escribió algo que no es un número
                print("     ! Ingrese un número válido.")  # lo ingresado no era un número, vuelve a pedirlo

    # junta todos los datos de placa que se preguntaron, tal como los necesita
    # la función de cálculo
    datos_placa = {
        "tension":           220.0,  # en vivienda todo esto va en 220 V monofásico
        "corriente_nominal":  I_nom,  # corriente nominal de placa, modo frío
        "corriente_maxima":   I_max,  # corriente máxima de la ficha, es la que manda para el TM
        "potencia_nominal_w": P_nom_w,  # potencia absorbida, sirve para el cuadro de cargas
        "factor_potencia":    fp,  # fp de placa, o el 0.95 típico
        "tecnologia":         tecnologia,  # inverter u on/off
        "tipo_compresor":     tipo_compresor,  # rotativo o pistón, solo pesa en los on/off
        "lra":                lra,  # corriente de arranque de placa, o None si hay que estimarla
        "mocp":               mocp,  # tope de protección del fabricante, o None
        "temp_sobre_30":      temp_sobre_30,  # si el recinto pasa los 30°C
        "temp_ambiente":      temp_ambiente,  # temperatura para sacar el factor de corrección
    }  # junta todo lo que se preguntó, para pasárselo al cálculo

    resultado = calcular_circuito_climatizacion(datos_placa)  # acá se calcula TM, diferencial y sección

    # junta todos los datos del equipo en un solo paquete, listo para
    # usarse en build_materiales_df
    return {
        "nombre_circ":   nombre_circ,  # nombre con el que aparece el circuito en el informe
        "ambientes_str": ambientes_str,  # ambientes que alimenta, en un solo texto
        "longitud":      longitud,  # metros de canalización hasta el tablero
        "P_nom_w":       P_nom_w,  # potencia para el cuadro de cargas
        "fp":            fp,  # factor de potencia del equipo
        "I_cuadro":      resultado["I_cuadro_A"],  # corriente que se muestra en el cuadro de cargas
        "I_diseno":      resultado["I_diseno_A"],  # corriente con la que se dimensiona el circuito
        "In_tm":         resultado["In_tm_A"],  # calibre del TM elegido, en amperes
        "curva_tm":      resultado["curva_tm"],  # curva del TM: C en inverter, C o D en on/off según el LRA
        "In_dif":        resultado["In_dif_A"],  # calibre del diferencial
        "seccion_mm2":   resultado["seccion_mm2"],  # sección del conductor, en mm^2
        "nota_tm":       resultado["nota_tm"],  # texto que explica por qué quedó ese TM
        "aviso_lra":     resultado["aviso_lra"],  # aviso si el LRA fue estimado y no sacado de placa
        "tecnologia":    tecnologia,  # se repite acá para no tener que abrir el resultado después
        "temp_sobre_30": temp_sobre_30,  # deja constancia de si hubo corrección por temperatura
    }


circuitos_climatizacion = []  # [{"nombre_circ":..., "In_tm":..., "curva_tm":..., ...}, ...] un dict por equipo de clima; se llena en la PARTE 2.1 y después se le pasa entero a build_materiales_df()

# =========================================================
# AGUA CALIENTE — lista de equipos (análoga a circuitos_climatizacion)
# =========================================================
circuitos_agua_caliente = []  # se va llenando más abajo, mientras se recorren los circuitos


# Calcula corriente, TM y diferencial de un equipo de agua caliente (ducha
# eléctrica, termoelectro, calefón). Recibe un diccionario con los datos de
# placa y devuelve un diccionario con los resultados del cálculo.
def calcular_circuito_agua_caliente(datos):
    """
    Calcula TM, diferencial y corriente de diseño para equipos de agua
    caliente (duchas eléctricas, termoelectros, calefones eléctricos).

    Reglas normativas aplicadas:
      - Conductor: Iz ≥ I_nom  (carga resistiva pura — sin factor 1.25)
        El factor 1.25 del RIC N°7 art. 7.3.4 aplica a motores y cargas
        de arranque difícil, no a cargas resistivas.
      - La sección no se decide acá: se calcula después, en la PARTE 3.2
      - fp = 1.0 (carga resistiva pura — sin componente reactiva)
      - Factor temperatura si T > 30°C (RIC N°07 art. 7.5.2 / RIC N°04 art. 6.2.6, Tabla 4.7)
      - TM: calibre comercial sobre I_diseño_corr, Curva C
      - Diferencial:
            ≤ 10 mA si equipo en Volumen 1 (interior ducha) — RIC N°11 art. 6.4.3
            ≤ 30 mA en cualquier otro caso                  — RIC N°07 art. 7.4.5
      - Calibre diferencial ≥ calibre TM
    """
    # datos de entrada, con valores por defecto si no vienen
    V             = float(datos.get("tension", 220))  # tensión de la red, 220 V monofásico
    P_nom_w       = float(datos.get("potencia_nominal_w", 0))  # potencia de placa del calentador, en watts
    temp_sobre_30 = datos.get("temp_sobre_30", False)  # si el recinto pasa los 30°C
    temp_ambiente = float(datos.get("temp_ambiente", 25.0))  # temperatura del recinto, para el factor de corrección
    vol1_bano     = datos.get("vol1_bano", False)   # si es True, va diferencial de 10 mA

    # --- Corriente nominal (fp=1.0, carga resistiva) ---
    I_nom = P_nom_w / V if V > 0 else 0.0  # como el fp es 1, la corriente sale directo de P/V

    # --- Corriente de diseño: si la carga es resistiva pura, I_diseño = I_nom
    # El factor 1.25 del RIC N°7 7.3.4 aplica a motores y cargas de arranque
    # difícil, NO a cargas resistivas (ducha, termoelectro, calefón).
    I_diseno = I_nom  # sin el recargo del 25%, esto no es motor

    # --- Factor corrección temperatura (RIC N°4 art. 6.2.6 / Tabla N°4.7) ---
    # Fórmula RIC: Ic = Iz × ft  (ft MULTIPLICA sobre Iz, NO divide I_diseno).
    # Más abajo si aparece un I_diseno/ft, pero eso es solo una referencia para
    # ir a buscar el calibre del TM, no para dimensionar el conductor.
    if temp_sobre_30:  # solo corrige si el recinto pasa los 30°C
        ft = factor_temperatura_ft(temp_ambiente, "B1")  # factor de la tabla 4.7 para canalización tipo B1
    else:  # recinto bajo 30°C, no hay nada que corregir
        ft = 1.0  # sin corrección, el conductor va con su ampacidad completa
    I_diseno_corr = I_diseno  # NO se divide por ft

    # --- TM: seleccionar calibre tal que Iz_TM × ft >= I_diseno ---
    # Equivalente a: In_TM >= I_diseno / ft  (referencia para buscar calibre)
    _ref_tm = (I_diseno / ft) if ft > 0 else I_diseno  # corriente de referencia para ir a buscar el calibre del TM
    In_tm = siguiente_calibre_clima(_ref_tm)   # primer calibre tal que Iz_TM × ft >= I_diseno

    # --- Diferencial: calibre ≥ TM, sensibilidad según volumen ---
    sensibilidad_dif = "10mA" if vol1_bano else "30mA"  # 10 mA si el equipo queda dentro de la ducha, si no 30 mA
    CALIBRES_DIF = [16, 25, 40, 63]  # calibres comerciales disponibles para el diferencial
    In_dif = next((c for c in CALIBRES_DIF if c >= In_tm), CALIBRES_DIF[-1])  # primer calibre comercial que alcanza al TM

    # resultado final del cálculo, listo para el informe y los materiales
    return {
        "I_nom_A":        round(I_nom, 2),  # corriente nominal del equipo, redondeada
        "I_diseno_A":     round(I_diseno_corr, 2),  # corriente de diseño, igual a la nominal por ser carga resistiva
        "In_tm":          In_tm,  # calibre del TM
        "curva_tm":       "C",  # curva C, la que se usa en estos equipos
        "In_dif":         In_dif,  # calibre del diferencial, nunca menor al del TM
        "sensibilidad_dif": sensibilidad_dif,  # 10 o 30 mA según el volumen del baño
        "ft":             ft,  # factor de temperatura que se aplicó
        "temp_sobre_30":  temp_sobre_30,  # queda registrado si hubo corrección
    }


# Pregunta por consola los datos de un equipo de agua caliente para un
# circuito ya identificado. Detecta si el ambiente es un baño (para aplicar
# el diferencial de 10 mA de Volumen 1) y devuelve todo el paquete de datos.
def ingresar_equipo_agua_caliente_inline(nombre_sugerido, ambientes_str, longitud):
    """
    Flujo de preguntas para circuitos de agua caliente.
    Retorna todos los datos necesarios juntos en un solo paquete, o None si hay error.
    """
    nombre_circ = nombre_sugerido  # nombre que ya venía armado desde el recorrido de circuitos
    amb_lower   = ambientes_str.lower()  # ambientes en minúsculas, para buscar palabras clave
    es_bano     = any(k in amb_lower for k in ("baño", "bano", "bathroom"))  # detecta si el ambiente es un baño

    # 1. Tipo de equipo
    # le muestra al usuario las 4 opciones de equipo de agua caliente
    print(
        "\n   Tipo de equipo:"
        "\n     (1) Ducha eléctrica"
        "\n     (2) Termoelectro"
        "\n     (3) Calefón eléctrico"
        "\n     (4) Otro calentador de agua"
    )
    # pide la opción hasta que sea una de las 4 válidas
    while True:
        tipo = input("   Opción: ").strip()  # guarda la opción elegida, como texto
        if tipo in ("1", "2", "3", "4"):  # valida que sea 1, 2, 3 o 4
            break  # opción válida, sale del ciclo
        print("     ! Ingrese 1, 2, 3 o 4.")  # opción inválida, vuelve a preguntar
    # traduce el número elegido a un texto legible
    tipo_map   = {"1": "Ducha eléctrica", "2": "Termoelectro",
                  "3": "Calefón eléctrico", "4": "Calentador de agua"}  # el 4 es el cajón de sastre para cualquier otro calentador
    tipo_equipo = tipo_map[tipo]  # nombre del tipo de equipo, ya listo para mostrar/guardar

    # Validar: ducha eléctrica solo en baño
    if tipo == "1" and not es_bano:  # tipo 1 = ducha eléctrica, pero el ambiente no es un baño
        # explica por que no se puede instalar aquí, citando el artículo del RIC
        print(f"\n   ! La ducha eléctrica SOLO puede instalarse en un BAÑO.")
        print(f"     (RIC N°11 art. 6 + RIC N°07 art. 7.4.5)")  # deja citada la norma que lo prohíbe fuera del baño
        print(f"     Ambiente indicado: '{ambientes_str}'. Corrija e intente de nuevo.")  # le muestra el ambiente que puso, para que lo corrija
        return None  # no cumple la norma, se cancela el ingreso

    # 2. Potencia nominal
    # pide la potencia hasta que ingrese un número válido mayor que 0
    while True:
        try:  # si escribe cualquier cosa que no sea número, no se cae
            # convierte lo ingresado a número, cambiando la coma decimal por punto
            P_nom_w = float(
                input("\n   - Potencia del equipo [W] (ej: 3300, 4400, 5500): ")  # potencia de placa del calentador, en watts
                .strip().replace(",", ".")  # acepta la coma como separador decimal
            )
            if P_nom_w > 0:  # la potencia debe ser mayor que 0
                break  # potencia válida, sigue
            print("     ! Debe ser mayor que 0.")  # número no positivo, vuelve a pedir
        except:  # lo ingresado no era un número, vuelve a pedir
            print("     ! Ingrese un número válido.")  # lo ingresado no era un número, vuelve a pedir

    # 3. Volumen RIC N°11 (solo para equipos en baño que no sean ducha eléctrica)
    vol1_bano = False  # por defecto asume que NO está en el Volumen 1 (diferencial de 30 mA)
    if tipo == "1":  # tipo 1 = ducha eléctrica
        # Ducha eléctrica: siempre Volumen 1, siempre diferencial de 10 mA
        vol1_bano = True  # la ducha va dentro de la ducha misma, obligado 10 mA
    elif es_bano:  # está en un baño pero no es ducha eléctrica
        # está en un baño pero no es ducha: hay que preguntar en qué volumen queda
        # explica que es el Volumen 1 para que el usuario pueda responder bien
        print(
            "\n   ¿El equipo está instalado dentro del Volumen 1?"
            "\n   (Vol.1 = dentro del perímetro de la ducha/bañera hasta 2,25m de altura)"
            "\n   → SÍ: diferencial 10 mA  |  NO: diferencial 30 mA"
        )
        # pide si/no hasta obtener una respuesta válida
        while True:
            resp = input("   (si/no): ").strip().lower()  # respuesta del usuario sobre el volumen 1
            if resp in ("si", "no"):  # solo acepta si o no
                vol1_bano = (resp == "si")  # True si contesto 'si': va diferencial de 10 mA
                break  # ya quedó definido el volumen, sale del ciclo
            print("     ! Responda si o no.")  # respuesta inválida, vuelve a preguntar

    # 4. Tablero externo de desconexión
    # pregunta si el equipo trae su propio interruptor o si hay que agregar un tablero aparte
    print("\n   ¿El equipo trae interruptor incorporado?")
    print("   → SÍ: sin tablero externo  |  NO: se genera tablero externo de desconexión")  # le explica las dos alternativas antes de responder
    while True:  # repite hasta que responda si o no
        resp_tab = input("   (si/no): ").strip().lower()  # respuesta sobre si el equipo trae interruptor incorporado
        if resp_tab in ("si", "no"):  # valida que sea si o no
            lleva_tablero_externo = (resp_tab == "no")  # si NO trae interruptor propio, hay que agregar un tablero externo
            break  # ya se sabe si hay que sumar tablero externo, sale
        print("     ! Responda si o no.")  # respuesta inválida, vuelve a preguntar

    # 5. Temperatura ambiente
    print("\n   ¿La temperatura del recinto supera los 30°C?")
    while True:  # repite hasta que responda si o no
        t_amb = input("   (si/no): ").strip().lower()  # respuesta sobre la temperatura del recinto
        if t_amb in ("si", "no"):  # valida que sea si o no
            break  # respondió si o no, sigue adelante
        print("     ! Responda si o no.")  # respuesta inválida, vuelve a preguntar
    temp_sobre_30 = False  # por defecto asume que NO supera los 30 grados C
    temp_ambiente = 25.0  # valor por defecto cuando no hay corrección (la base de la Tabla 4.7 es 30°C)
    if t_amb == "si":  # hay que corregir por temperatura, se pide el valor
        temp_sobre_30 = True  # se usara para corregir la capacidad del conductor por temperatura
        # pide la temperatura máxima hasta que sea mayor a 30 grados C
        while True:
            try:  # si escribe cualquier cosa que no sea número, no se cae
                # convierte lo ingresado a número, cambiando la coma decimal por punto
                temp_ambiente = float(
                    input("   - Temperatura máxima [°C]: ").strip().replace(",", ".")  # temperatura máxima del recinto, en °C
                )
                if temp_ambiente > 30:  # tiene que ser mayor a 30, si no no calza con lo que respondió
                    break  # temperatura válida, sigue
                print("     ! Debe ser mayor que 30°C.")  # debe ser mayor a 30 grados porque el usuario dijo que si lo superaba
            except:  # lo ingresado no era un número, vuelve a pedir
                print("     ! Ingrese un número válido.")  # lo ingresado no era un número, vuelve a pedir

    # Cálculo eléctrico
    # arma el diccionario de datos y llama a la función que calcula corriente, TM y diferencial
    resultado = calcular_circuito_agua_caliente({
        "tension":            220.0,  # 220 V monofásico
        "potencia_nominal_w": P_nom_w,  # potencia de placa del equipo
        "temp_sobre_30":      temp_sobre_30,  # para aplicar o no el factor de temperatura
        "temp_ambiente":      temp_ambiente,  # temperatura máxima del recinto
        "vol1_bano":          vol1_bano,  # define si el diferencial va de 10 o de 30 mA
    })  # cierra los datos que se le pasan al cálculo del circuito

    # junta todos los datos del equipo en un solo paquete, listo para
    # usarse en build_materiales_df
    return {
        "nombre_circ":          nombre_circ,  # nombre del circuito para el informe
        "tipo_equipo":          tipo_equipo,  # ducha, termoelectro, calefón o calentador
        "ambientes_str":        ambientes_str,  # ambientes que alimenta
        "longitud":             longitud,  # metros hasta el tablero
        "P_nom_w":              P_nom_w,  # potencia para el cuadro de cargas
        "fp":                   1.0,  # carga resistiva pura, fp = 1
        "I_cuadro":             resultado["I_nom_A"],  # corriente que se muestra en el cuadro de cargas
        "I_diseno":             resultado["I_diseno_A"],  # corriente con la que se dimensiona el circuito
        "In_tm":                resultado["In_tm"],  # calibre del TM
        "curva_tm":             "C",  # curva C
        "In_dif":               resultado["In_dif"],  # calibre del diferencial
        "sensibilidad_dif":     resultado["sensibilidad_dif"],  # 10 mA si quedó en volumen 1, 30 mA en el resto
        "ft":                   resultado["ft"],  # factor de temperatura que se aplicó
        "temp_sobre_30":        temp_sobre_30,  # deja constancia de si hubo corrección
        "vol1_bano":            vol1_bano,  # si el equipo quedó dentro del volumen 1 del baño
        "lleva_tablero_externo": lleva_tablero_externo,  # si hay que sumar tablero de desconexión a los materiales
    }
# -------- PARTE 2.1: CONSTRUCCIÓN DE CIRCUITOS --------
# recorre uno por uno los circuitos que pidió el usuario, preguntando su
# nombre, ambientes, longitud, y detectando de qué tipo es (iluminación,
# enchufes, climatización, agua caliente, o especial genérico)
for i in range(1, cantidad_circuitos + 1):

    # pregunta el nombre del circuito número i
    base = input(
        f"\n   - Ingrese el nombre del circuito N°{i} "
        "(Ejemplos: iluminacion, enchufes generales, cocina-encimera, baño-lavadora, etc.): "
    ).strip()  # deja el nombre sin espacios al principio ni al final

    base = limpiar_nombre_circuito(base)  # solo deja un espacio entre palabras y arregla el tipeo "eenc..." -> "enc..."; NO toca mayúsculas ni tildes
    base = sugerir_nombre_circuito(base)  # si el nombre no trae ninguna palabra clave, PREGUNTA por consola "¿quisiste decir iluminación/enchufes/climatización?" y solo lo reemplaza si el usuario responde "si"
    base_lower = base.lower()  # versión en minúsculas, para buscar palabras clave más abajo

    print("     • Ambientes disponibles:", ", ".join(nombres_amb))  # muestra la lista de ambientes ya definidos, para que el usuario elija

    # pide los ambientes del circuito hasta que sean válidos
    while True:
        sel = input("     • Escriba los ambientes que componen el circuito (ej: Living, Cocina): ").strip()  # texto crudo con los ambientes que escribió el usuario
        sel_list = normaliza_seleccion_ambientes(sel, nombres_amb)  # convierte el texto ingresado en una lista de ambientes válidos
        if sel_list:  # si reconoció al menos un ambiente válido, sigue
            break  # reconoció ambientes válidos, sigue
        print("       ! Debe ingresar al menos un ambiente válido (respetando nombres).")  # no reconoció ningún ambiente válido, vuelve a pedir

    ambientes_str = ", ".join(sel_list)  # ambientes del circuito, unidos en un solo texto (ej: 'Living, Cocina')

    # Detectar tipo de circuito ANTES de preguntar longitud para ajustar el texto
    _base_lower_prev = base.lower()  # nombre en minúsculas, para detectar si es un equipo de agua caliente
    # True si el nombre del circuito contiene alguna palabra de agua caliente
    _es_agua_prev = any(k in _base_lower_prev for k in
                        ("ducha", "termo", "termoelectro", "calefon", "calefón",  # palabras que delatan un equipo de agua caliente
                         "calentador", "agua caliente"))  # cierra la lista de palabras clave de agua caliente

    if len(sel_list) == 1:  # un solo ambiente: la longitud se conoce exacta
        # un solo ambiente: la longitud es "real", conocida con precisión
        if _es_agua_prev:  # equipo de agua caliente: se pregunta la distancia hasta el tablero
            # agua caliente: la longitud se mide desde el tablero hasta el equipo
            longitud = float(input(
                f"     • Longitud del circuito '{base} ({ambientes_str})' en metros\n"
                f"       (desde el tablero principal hasta el equipo): "
            ))  # cierra el input de la longitud hasta el equipo
        else:  # el circuito no es de agua caliente
            # no es agua caliente: se pide la longitud real de la canalización según el RIC
            longitud = float(input(
                f"     • Ingrese la longitud REAL de la canalización en (m) del circuito '{base} ({ambientes_str})'\n"
                f"       (En tramos horizontales: recorridos a 0,30m del cielo y 0,20m del piso - RIC N°4 7.16.1.16): "
            ))  # cierra el input de la longitud de la canalización
        longitud_unica = True  # marca que esta longitud es exacta (un solo ambiente); bandera que hoy no lee nadie más en el script
    else:  # el circuito abarca más de un ambiente
        # varios ambientes: la longitud es una estimación (recorre varios lugares)
        if _es_agua_prev:  # mismo caso que arriba, pero el circuito abarca varios ambientes
            longitud = float(input(  # misma pregunta de antes: del tablero al equipo de agua caliente
                f"     • Longitud del circuito '{base} ({ambientes_str})' en metros\n"
                f"       (desde el tablero principal hasta el equipo): "
            ))  # cierra el input de la longitud hasta el equipo
        else:  # tampoco es agua caliente, se pide el recorrido real de la canalización
            # longitud real del recorrido de la canalización, no la distancia en línea recta
            longitud = float(input(
                f"     • Ingrese la longitud REAL de la canalización en (m) del circuito '{base} ({ambientes_str})'\n"
                f"       (En tramos horizontales: recorridos a 0,30m del cielo y 0,20m del piso - RIC N°4 7.16.1.16): "
            ))  # cierra el input de la longitud de la canalización
        longitud_unica = False  # marca que esta longitud es una estimación; igual que arriba, la bandera no se usa después

    # Cajas de paso (RIC 7.16.1.13): solo se consideran si el circuito tiene
    # un tramo CONTINUO >=20m — no basta con que la longitud total del
    # circuito supere 20m (puede ser la suma de varios tramos cortos).
    # Se pregunta solo cuando la longitud total ya da para sospechar que
    # podría haber un tramo así (si es menor a 20m, es matemáticamente
    # imposible que exista un tramo continuo >=20m dentro de ese circuito).
    tiene_tramo_20m = False  # por defecto asume que no hay tramos continuos largos
    # OJO: esta respuesta solo se le pasa a add_circuito() en los circuitos
    # ESPECIALES y en los de enchufes/iluminación. Los circuitos de
    # climatización, agua caliente y cocina/lavadero automático llaman a add_circuito() sin este
    # argumento, así que se quedan con el valor por defecto (True) aunque el
    # usuario haya contestado que no hay ningún tramo continuo de 20m.
    if longitud >= 20:  # solo tiene sentido preguntar si la longitud total ya alcanza 20m
        _resp_tramo20 = input(  # pregunta si dentro del circuito hay un tramo continuo largo
            f"     • ¿Existe algún tramo continuo dentro del circuito '{base} ({ambientes_str})' "
            f"mayor o igual a 20m? (si/no): "
        ).strip().lower()  # deja la respuesta sin espacios y en minúsculas para compararla
        tiene_tramo_20m = _resp_tramo20 in ("si", "sí", "s", "y", "yes")  # True si el usuario confirmo que existe un tramo continuo >=20m

    # detecta el tipo de circuito buscando palabras clave en el nombre
    es_ilumin = "ilumin" in base_lower  # True si el nombre incluye 'ilumin' (circuito de iluminación)
    es_enchufe = "enchufe" in base_lower  # True si el nombre incluye 'enchufe'
    es_clima_circ = any(k in base_lower for k in ("clima", "aire", "split", "ac ", "a/c"))  # True si el nombre sugiere climatización (clima, aire, split, ac, a/c)
    # True si el nombre sugiere un equipo de agua caliente (ducha, termo, calefón, etc.)
    es_agua_circ  = any(k in base_lower for k in
                        ("ducha", "termo", "termoelectro", "calefon", "calefón",  # duchas eléctricas, termos y calefones
                         "calentador", "agua caliente"))  # calentadores y cualquier nombre que hable de agua caliente

    # ---- CIRCUITO DE CLIMATIZACIÓN detectado por nombre ----
    if es_clima_circ:  # el nombre trae clima/aire/split, así que se arma como circuito de climatización
        # delega todas las preguntas de climatización a otra función (definida más arriba en el archivo)
        datos_clima = ingresar_equipo_climatizacion_inline(
            nombre_sugerido=base,  # nombre base que escribió el usuario
            ambientes_str=ambientes_str,  # ambientes donde va el equipo
            longitud=longitud  # longitud ya preguntada más arriba
        )
        circuitos_climatizacion.append(datos_clima)  # guarda este equipo en la lista global de climatización, para reportes/resumenes
        items_clima = [{"amb": datos_clima["ambientes_str"], "potencia": float(datos_clima["P_nom_w"]), "nombre": datos_clima["nombre_circ"]}]  # arma el item que se usara para el detalle de materiales de este circuito
        # Formato detalle igual al resto: "ambiente: descripción"
        _amb_lower = datos_clima["ambientes_str"].lower()  # nombre del ambiente en minúsculas, para armar el texto de detalle
        _lra_txt = f" | {datos_clima['aviso_lra']}" if datos_clima["aviso_lra"] else ""  # agrega el aviso de corriente de arranque (LRA) si el equipo lo tiene
        # arma el texto que describe este circuito en el informe
        _detalle_clima = (
            f"{_amb_lower}: climatización {datos_clima['tecnologia'].upper()} "
            f"({datos_clima['I_cuadro']}A nominal{_lra_txt})"
        )
        # registra el circuito de climatización en la lista general de circuitos
        add_circuito(
            nombre_circ       = datos_clima["nombre_circ"] + " (" + datos_clima["ambientes_str"] + ")",  # nombre del circuito con sus ambientes entre paréntesis
            longitud          = datos_clima["longitud"],  # largo de la canalización hasta el equipo
            potencia_estimada = datos_clima["P_nom_w"],  # potencia nominal del equipo en W
            es_ilumin         = False,  # no es circuito de iluminación
            es_enchufe        = False,  # no es circuito de enchufes
            es_especial       = True,  # clima siempre va como circuito especial (dedicado)
            items             = items_clima,  # componentes del circuito, para el detalle de materiales
            detalle_asignacion = _detalle_clima  # texto que describe el circuito en el informe
        )
        # a partir de aquí se completan, en el último circuito agregado, los datos propios de climatización
        circuitos[-1]["Disyuntor termomagnético"] = f"1x{datos_clima['In_tm']}A / 6kA / Curva {datos_clima['curva_tm']}"  # TM del clima: 1 polo, poder de corte 6kA y curva según el arranque del equipo
        circuitos[-1]["_In_TM"]            = datos_clima["In_tm"]  # guarda el calibre del TM en amperes, para el cuadro de cargas
        circuitos[-1]["_In_dif_clima"]     = datos_clima["In_dif"]  # corriente nominal del diferencial de este equipo
        circuitos[-1]["_es_climatizacion"] = True  # bandera para que el resto del script trate este circuito como climatización
        circuitos[-1]["_I_diseno_clima"]   = datos_clima["I_diseno"]  # corriente de diseño del equipo, ya con sus factores aplicados
        circuitos[-1]["Corriente estimada (A)"] = round(float(datos_clima["I_cuadro"]), 2)  # corriente nominal del equipo, redondeada a 2 decimales
        continue  # este circuito ya quedó armado, pasa al siguiente

    # ---- CIRCUITO DE AGUA CALIENTE detectado por nombre ----
    elif es_agua_circ:  # el nombre trae ducha/termo/calefón, así que es un equipo de agua caliente
        # ── Validación de ambientes ──────────────────────────────────────────
        # Validación previa: rechaza ambientes claramente inválidos para
        # cualquier equipo de agua caliente (living, dormitorio, etc.)
        # La validación específica de que la ducha solo va en baño se hace dentro de
        # ingresar_equipo_agua_caliente_inline una vez que el usuario elige el tipo.
        # lista de palabras que indican ambientes donde NO se puede instalar agua caliente
        _AMBIENTES_INVALIDOS_AGUA = (
            "living", "dormitorio", "dorm", "bedroom",  # ambientes de estar y de dormir
            "comedor", "dining", "estar", "sala",  # comedor y salas de estar
            "hall", "estudio", "oficina", "terraza",  # circulaciones, lugares de trabajo y terrazas
            "lounge", "studio", "office",  # los mismos ambientes escritos en ingles
        )

        _ambs_invalidos_encontrados = []  # aquí se van guardando los ambientes que no son válidos
        for _amb_check in sel_list:  # revisa cada ambiente elegido para este circuito
            _amb_low = _amb_check.lower().strip()  # nombre del ambiente normalizado para comparar
            if any(k in _amb_low for k in _AMBIENTES_INVALIDOS_AGUA):  # el ambiente contiene alguna palabra prohibida para agua caliente
                _ambs_invalidos_encontrados.append(_amb_check)  # guarda el ambiente que no sirve, para avisarlo después

        if _ambs_invalidos_encontrados:  # se encontró al menos un ambiente inválido
            # explica al usuario por que el ambiente no sirve y cuáles si son válidos
            print(f"\n   ! ADVERTENCIA: Los siguientes ambientes NO son válidos")
            print(f"     para instalar equipos de agua caliente:")  # aclara para que tipo de equipos aplica el rechazo
            for _a in _ambs_invalidos_encontrados:  # va nombrando uno por uno los ambientes que no sirven
                print(f"       - {_a}")  # muestra el ambiente rechazado
            print(f"   ! Los equipos de agua caliente solo pueden instalarse en:")  # le indica donde si puede quedar el equipo
            print(f"     baño, lavadero, cocina, bodega, pasillo de servicio o exterior.")  # ambientes permitidos para agua caliente
            print(f"   ! La ducha eléctrica SOLO puede instalarse en baño.")  # la ducha es el caso más estricto: solo en baño
            print(f"   ! Corrija el nombre del circuito o los ambientes e intente de nuevo.")  # le pide corregir el circuito y volver a ingresarlo
            continue  # vuelve a pedir los datos de este circuito desde el principio
        # ────────────────────────────────────────────────────────────────────
        # delega las preguntas especificas de agua caliente a la función definida más arriba
        datos_agua = ingresar_equipo_agua_caliente_inline(
            nombre_sugerido=base,  # nombre base que escribió el usuario
            ambientes_str=ambientes_str,  # ambientes donde va el equipo
            longitud=longitud  # longitud ya preguntada más arriba
        )
        if datos_agua is None:  # la función devolvió None: el ambiente no servia para ese tipo de equipo
            # Ambiente inválido para el tipo de equipo: el circuito se descarta
            continue
        circuitos_agua_caliente.append(datos_agua)  # guarda este equipo en la lista global de agua caliente, para reportes/resumenes
        # arma el item que se usara para el detalle de materiales de este circuito
        items_agua = [{"amb":     datos_agua["ambientes_str"],
                       "potencia": float(datos_agua["P_nom_w"]),  # potencia nominal del equipo en W
                       "nombre":   datos_agua["tipo_equipo"]}]  # tipo de equipo: ducha, termo, calefón eléctrico, etc.
        # arma el texto que describe este circuito en el informe
        _detalle_agua = (
            f"{datos_agua['ambientes_str'].lower()}: {datos_agua['tipo_equipo']} "
            f"({datos_agua['P_nom_w']:.0f} W)"
        )
        # registra el circuito de agua caliente en la lista general de circuitos
        add_circuito(
            nombre_circ        = datos_agua["nombre_circ"] + " (" + datos_agua["ambientes_str"] + ")",  # nombre del circuito con sus ambientes entre paréntesis
            longitud           = datos_agua["longitud"],  # largo de la canalización hasta el equipo
            potencia_estimada  = datos_agua["P_nom_w"],  # potencia nominal del equipo en W
            es_ilumin          = False,  # no es circuito de iluminación
            es_enchufe         = False,  # no es circuito de enchufes
            es_especial        = True,  # agua caliente siempre va en circuito dedicado
            items              = items_agua,  # componentes del circuito, para el detalle de materiales
            detalle_asignacion = _detalle_agua  # texto que describe el circuito en el informe
        )
        # a partir de aquí se completan, en el último circuito agregado, los datos propios de agua caliente
        circuitos[-1]["Disyuntor termomagnético"]  = f"1x{datos_agua['In_tm']}A / 6kA / Curva C"  # TM del equipo de agua caliente: 1 polo, 6kA y curva C
        circuitos[-1]["_In_TM"]                      = datos_agua["In_tm"]  # calibre del TM en amperes, para el cuadro de cargas
        circuitos[-1]["_In_dif_agua"]                = datos_agua["In_dif"]  # corriente nominal del diferencial de este equipo
        circuitos[-1]["_sensibilidad_dif_agua"]       = datos_agua["sensibilidad_dif"]  # sensibilidad del diferencial (10mA en ducha, 30mA en el resto)
        circuitos[-1]["_es_agua_caliente"]            = True  # bandera para que el resto del script trate este circuito como agua caliente
        circuitos[-1]["_I_diseno_agua"]               = datos_agua["I_diseno"]  # corriente de diseño del equipo, ya con sus factores aplicados
        circuitos[-1]["_lleva_tablero_externo_agua"]  = datos_agua["lleva_tablero_externo"]  # True si el equipo necesita su propio tablero fuera del baño
        circuitos[-1]["_vol1_bano_agua"]              = datos_agua["vol1_bano"]  # True si el equipo queda dentro del volumen 1 del baño (zona mojada)
        circuitos[-1]["_tipo_equipo_agua"]            = datos_agua["tipo_equipo"]  # tipo de equipo, se usa después en el informe
        circuitos[-1]["Corriente estimada (A)"]       = round(float(datos_agua["I_cuadro"]), 2)  # corriente nominal del equipo, redondeada a 2 decimales
        continue  # este circuito ya quedó armado, pasa al siguiente


    # ni iluminación, ni enchufes, ni climatización, ni agua caliente:
    # se pregunta directamente si es un circuito especial (horno, lavadora, etc.)
    while True:  # repite la pregunta hasta que conteste si o no
        resp = input("     • ¿Este circuito es ESPECIAL? (si/no): ").strip().lower()  # pregunta si el circuito es especial (horno, lavadora, calefactor, etc.)
        # solo se acepta 'si' o 'no' como respuesta
        if resp in ["si", "no"]:
            break  # respuesta válida, sale del bucle
        print("       ! Responda solo 'si' o 'no'.")  # la respuesta no sirvio, vuelve a preguntar
    es_especial = (resp == "si")  # True si el usuario contesto 'si'

    if es_especial:  # circuito especial: se arma con los componentes que el usuario ya cargo por ambiente
        # ---------- CIRCUITO ESPECIAL ----------
        # se le van mostrando al usuario los componentes especiales que ya
        # ingresó por ambiente, para que elija cuáles van en este circuito
        componentes_elegidos = []  # aquí se van guardando los componentes que el usuario elige
        # recorre cada ambiente seleccionado para este circuito
        for amb in sel_list:
            key = amb.lower()  # normaliza el nombre para buscarlo en el diccionario
            disponibles = componentes_restantes.get(key, [])  # componentes especiales que aun quedan libres en este ambiente
            if not disponibles:  # el ambiente quedó sin componentes especiales libres
                # no queda ningún componente especial disponible en este ambiente
                print(f"       ! El ambiente '{amb}' no tiene componentes especiales disponibles.")
                continue  # pasa al siguiente ambiente

            # arma un texto con id, nombre y potencia de cada componente disponible
            lista_str = ", ".join([f"{e['id']} - {e['nombre']} ({e['potencia']}W)" for e in disponibles])
            print(f"       Componentes especiales disponibles en '{amb}': {lista_str}")  # muestra la lista al usuario

            while True:  # repite hasta que escriba una selección válida
                # pide los números de componentes a incluir (ej: '1,3') o 0 si ninguno,
                # repite hasta que la entrada sea válida
                entrada = input(
                    "       • Ingrese los números de componentes que van en este circuito (ej: 1,3) o 0 si ninguno: "
                ).strip()  # saca los espacios de lo que escribió
                # entrada vacía o '0': no se elige ningún componente
                if entrada in ["", "0"]:
                    ids_sel = []  # no eligió ningún componente de este ambiente
                    break  # entrada válida, sale del bucle
                try:  # intenta interpretar los números que escribió
                    # convierte el texto en una lista de números únicos y ordenados
                    ids_sel = sorted(set(int(x.strip()) for x in entrada.split(",") if x.strip()))
                except:  # escribió algo que no son números
                    print("         ! Formato inválido. Use números separados por coma (ej: 1,2,3) o 0.")  # escribió cualquier cosa, se le explica el formato
                    continue  # vuelve a pedir la entrada

                ids_disponibles = {e["id"] for e in disponibles}  # conjunto con los ids que realmente existen
                # revisa que todos los números ingresados sean componentes válidos
                if all(idx in ids_disponibles for idx in ids_sel):
                    break  # todos los números existen, sigue adelante
                print("         ! Uno o más números no corresponden a componentes disponibles.")  # puso un número que no está en la lista, vuelve a pedir

            # si el usuario eligió algo, se procesa la eleccion
            if ids_sel:
                # saca los componentes elegidos de la lista de "disponibles", para
                # que no se puedan volver a asignar a otro circuito por error
                asignados = [e for e in disponibles if e["id"] in ids_sel]  # los componentes que el usuario escogio
                componentes_restantes[key] = [e for e in disponibles if e["id"] not in ids_sel]  # en el ambiente quedan solo los componentes que no se eligieron
                # guarda cada componente elegido junto con su ambiente
                for e in asignados:
                    componentes_elegidos.append({  # agrega el componente a los que van en este circuito
                        "amb": amb,  # ambiente de donde salió el componente
                        "id": e["id"],  # número con que se mostro en la lista
                        "nombre": e["nombre"],  # nombre del componente (horno, lavadora, etc.)
                        "potencia": e["potencia"]  # potencia del componente en W
                    })  # cierra el componente elegido

        # si no se eligió ningún componente, no se crea el circuito
        if not componentes_elegidos:
            print("       ! No se eligieron componentes especiales. No se crea este circuito.")  # avisa que el circuito quedaria vacío
            continue  # vuelve a pedir los datos de este circuito desde el principio

        # tanteo de TM que hoy no se usa en ninguna parte: el TM definitivo se calcula más abajo, en la PARTE 3
        tension_ref = 220.0  # tensión de referencia para estimar la corriente (V)
        calibres_tm_ref = [6, 10, 16, 20, 25, 32, 40, 50, 63]  # calibres de TM disponibles, en amperes
        potencia_total_comp = sum(e["potencia"] for e in componentes_elegidos)  # suma la potencia de todos los componentes elegidos
        i_est_approx = potencia_total_comp / tension_ref  # corriente estimada: I = P / V
        i_necesaria = i_est_approx * 1.10  # 10% de margen
        tm_aprox = next((cal for cal in calibres_tm_ref if cal >= i_necesaria), calibres_tm_ref[-1])  # elige el primer calibre de TM que alcanza para esa corriente

        # OJO:
        # Aquí todavía NO se ha calculado el empalme final.
        # Por eso NO se debe comparar tm_aprox con max_sum_tm en este punto.
        # Primero se crean todos los circuitos, luego se calcula la demanda,
        # y recién después se obtiene el empalme calculado y sus protecciones.
        nombre_circ = f"{base} ({ambientes_str})"  # nombre final del circuito, con sus ambientes
        items = [{"amb": e["amb"], "potencia": float(e["potencia"]), "nombre": e["nombre"]} for e in componentes_elegidos]  # componentes en el formato que espera add_circuito
        detalle = resumen_items_por_ambiente(items, modo="especial")  # texto resumen de que componente hay en cada ambiente

        # registra el circuito especial ya armado
        add_circuito(
            nombre_circ,  # nombre del circuito con sus ambientes
            longitud,  # largo de la canalización de este circuito
            potencia_total_comp,  # suma de potencias de los componentes elegidos
            es_ilumin=False,  # no es circuito de iluminación
            es_enchufe=False,  # no es circuito de enchufes
            es_especial=True,  # va marcado como especial
            items=items,  # componentes del circuito, para el detalle de materiales
            detalle_asignacion=detalle,  # texto que dice que componente quedó en cada ambiente
            tiene_tramo_20m=tiene_tramo_20m  # avisa si hay tramo continuo >=20m, para las cajas de paso
        )

    else:  # no es especial: es un circuito común de iluminación o de enchufes
        # ---------- CIRCUITO GENERAL (iluminación o enchufes) ----------
        potencia_estimada = 0.0  # se vuelve a sumar con lo que el usuario elija en este circuito
        ambientes_str_final = ambientes_str  # por defecto; puede cambiar más abajo según lo elegido (enchufes)
        items = []  # acá se van juntando las luminarias o enchufes que entren a este circuito

        if es_ilumin:  # circuito de iluminación: se eligen luminarias ambiente por ambiente
            # va mostrando las luminarias disponibles de cada ambiente y
            # dejando que el usuario elija cuáles van en este circuito
            # mismo patrón que con los componentes especiales: recorre cada
            # ambiente y deja elegir cuáles luminarias entran a este circuito
            for amb in sel_list:
                key = amb.lower()  # normaliza el nombre para buscarlo en el diccionario
                disponibles = luminarias_restantes.get(key, [])  # luminarias que aun quedan libres en este ambiente
                if not disponibles:  # el ambiente quedó sin luminarias libres
                    # no queda ninguna luminaria disponible en este ambiente
                    print(f"       ! El ambiente '{amb}' no tiene luminarias disponibles.")
                    continue  # pasa al siguiente ambiente

                # arma un texto con id y potencia de cada luminaria disponible
                lista_str = ", ".join([f"{e['id']} ({e['potencia']}W)" for e in disponibles])
                print(f"       Luminarias disponibles en '{amb}': {lista_str}")  # muestra la lista al usuario

                while True:  # repite hasta que escriba una selección válida
                    # pide los números de luminarias a incluir, o 0 si ninguna
                    entrada = input(
                        "       • Ingrese los números de luminaria que van en este circuito (ej: 1,3) o 0 si ninguna: "
                    ).strip()  # saca los espacios de lo que escribió
                    # entrada vacía o '0': no se elige ninguna luminaria
                    if entrada in ["", "0"]:
                        ids_sel = []  # no eligió ninguna luminaria de este ambiente
                        break  # entrada válida, sale del bucle
                    try:  # intenta interpretar los números que escribió
                        # convierte el texto en una lista de números únicos y ordenados
                        ids_sel = sorted(set(int(x.strip()) for x in entrada.split(",") if x.strip()))
                    except:  # escribió algo que no son números
                        print("         ! Formato inválido.")  # escribió cualquier cosa, se le vuelve a pedir
                        continue  # vuelve a pedir los números

                    ids_disponibles = {e["id"] for e in disponibles}  # conjunto con los ids que realmente existen
                    # revisa que todos los números ingresados sean luminarias válidas
                    if all(idx in ids_disponibles for idx in ids_sel):
                        break  # todos los números existen, sigue adelante
                    print("         ! Uno o más números no corresponden a luminarias disponibles.")  # puso un número que no está en la lista, vuelve a pedir

                # si el usuario eligió algo, se procesa la eleccion
                if ids_sel:
                    asignados = [e for e in disponibles if e["id"] in ids_sel]  # las luminarias que el usuario escogio
                    potencia_estimada += sum(e["potencia"] for e in asignados)  # suma la potencia de las luminarias elegidas al total del circuito
                    # guarda cada luminaria elegida con su tipo y descripción
                    for e in asignados:
                        items.append({  # guarda la luminaria entre los items de este circuito
                            "amb": amb,  # ambiente de donde salió la luminaria
                            "potencia": float(e["potencia"]),  # potencia de la luminaria en W
                            "tipo_lum": str(e.get("tipo", "")).strip().lower(),  # tipo de luminaria (led, dicroico, etc.), se usa para los materiales
                            "desc_lum": str(e.get("desc", "")).strip().lower()  # descripción de la luminaria tal como la escribió el usuario
                        })  # cierra la luminaria agregada al circuito
                    luminarias_restantes[key] = [e for e in disponibles if e["id"] not in ids_sel]  # ya no están disponibles para otro circuito

        elif es_enchufe:  # circuito de enchufes: mismo flujo que iluminación pero con enchufes
            ambientes_con_enchufe_general = []  # ambientes que aportaron enchufes normales (no cocina/lavadero) a este circuito
            agrego_coc_lav = False  # True si se agrego algún enchufe de cocina o lavadero
            items_coc_lav = []  # guarda por separado los enchufes de cocina/lavadero, por si hay que crear su circuito aparte

            # mismo patrón que luminarias: muestra los enchufes disponibles y
            # deja que el usuario elija cuáles van en este circuito
            for amb in sel_list:
                key = amb.lower()  # normaliza el nombre para buscarlo en el diccionario
                disponibles = enchufes_restantes.get(key, [])  # enchufes que aun quedan libres en este ambiente
                if not disponibles:  # el ambiente quedó sin enchufes libres
                    # no queda ningún enchufe disponible en este ambiente
                    print(f"       ! El ambiente '{amb}' no tiene enchufes disponibles.")
                    continue  # pasa al siguiente ambiente

                # arma un texto con id y potencia de cada enchufe disponible
                lista_str = ", ".join([f"{e['id']} ({e['potencia_total']}W)" for e in disponibles])
                print(f"       Enchufes disponibles en '{amb}': {lista_str}")  # muestra la lista al usuario

                while True:  # repite hasta que escriba una selección válida
                    # pide los números de enchufes a incluir, o 0 si ninguno
                    entrada = input(
                        "       • Ingrese los números de enchufe que van en este circuito (ej: 1,3) o 0 si ninguno: "
                    ).strip()  # saca los espacios de lo que escribió
                    # entrada vacía o '0': no se elige ningún enchufe
                    if entrada in ["", "0"]:
                        ids_sel = []  # no eligió ningún enchufe de este ambiente
                        break  # entrada válida, sale del bucle
                    try:  # intenta interpretar los números que escribió
                        # convierte el texto en una lista de números únicos y ordenados
                        ids_sel = sorted(set(int(x.strip()) for x in entrada.split(",") if x.strip()))
                    except:  # escribió algo que no son números
                        print("         ! Formato inválido.")  # escribió cualquier cosa, se le vuelve a pedir
                        continue  # vuelve a pedir los números

                    ids_disponibles = {e["id"] for e in disponibles}  # conjunto con los ids que realmente existen
                    # revisa que todos los números ingresados sean enchufes válidos
                    if all(idx in ids_disponibles for idx in ids_sel):
                        break  # todos los números existen, sigue adelante
                    print("         ! Uno o más números no corresponden a enchufes disponibles.")  # puso un número que no está en la lista, vuelve a pedir

                # si el usuario eligió algo, se procesa la eleccion
                if ids_sel:
                    asignados = [e for e in disponibles if e["id"] in ids_sel]  # los enchufes que el usuario escogio
                    aporto_general = False  # True si algún enchufe elegido queda en este circuito general (no cocina/lavadero)

                    # cocina y lavadero se tratan aparte porque el RIC exige circuitos
                    # dedicados para sus enchufes
                    for e in asignados:
                        kamb = key.strip().lower()  # nombre del ambiente en minúscula, para comparar
                        if kamb.startswith("cocina") or kamb.startswith("lavadero"):  # los enchufes de cocina y lavadero no pueden ir con el resto
                            # acumular por ambiente (cocina 1, cocina 2, lavadero, etc.)
                            if amb not in enchufes_coc_lav_por_amb:  # primera vez que aparece este ambiente, se le crea su lista
                                enchufes_coc_lav_por_amb[amb] = []  # lista vacía donde se acumulan los enchufes de ese ambiente
                            enchufes_coc_lav_por_amb[amb].append({  # guarda el enchufe en la bolsa de cocina/lavadero
                                "amb": amb,  # ambiente exacto (cocina 1, lavadero, etc.)
                                "potencia": float(e["potencia_total"]),  # potencia total del enchufe en W
                                "longitud": float(longitud),  # longitud del circuito, se hereda al circuito dedicado
                                "id_ench": int(e.get("id", 0)),  # número con que se mostro el enchufe en la lista
                                "modulos": int(e.get("modulos", 1))  # cuántos módulos tiene el enchufe (simple, doble, triple)
                            })  # cierra el enchufe que queda en la bolsa de cocina/lavadero
                            agrego_coc_lav = True  # marca que este circuito reunio enchufes de cocina o lavadero

                            # también se guarda acá para poder crear el circuito manual si hace falta
                            items_coc_lav.append({
                                "amb": amb,  # ambiente exacto de donde salió
                                "potencia": float(e["potencia_total"]),  # potencia total del enchufe en W
                                "n_ench": 1,  # cada entrada corresponde a un enchufe
                                "id_ench": int(e.get("id", 0)),  # número con que se mostro el enchufe en la lista
                                "modulos": int(e.get("modulos", 1))  # cuántos módulos tiene el enchufe
                            })  # cierra la copia que se usa para el circuito dedicado
                        else:  # el enchufe no es de cocina ni de lavadero
                            # enchufe normal: entra directo a este circuito
                            p = float(e["potencia_total"])  # potencia del enchufe normal
                            potencia_estimada += p  # se suma al total del circuito general
                            items.append({  # guarda el enchufe entre los items del circuito general
                                "amb": amb,  # ambiente de donde salió el enchufe
                                "potencia": p,  # potencia del enchufe en W
                                "n_ench": 1,  # cada entrada corresponde a un enchufe
                                "id_ench": int(e.get("id", 0)),  # número con que se mostro el enchufe en la lista
                                "modulos": int(e.get("modulos", 1))  # cuántos módulos tiene el enchufe
                            })  # cierra el enchufe agregado al circuito general
                            aporto_general = True  # este ambiente aporto al menos un enchufe al circuito general

                    # si este ambiente aporto enchufes normales, queda en la lista
                    # de ambientes del circuito general
                    if aporto_general:
                        ambientes_con_enchufe_general.append(amb)  # este ambiente queda dentro del circuito general de enchufes

                    enchufes_restantes[key] = [e for e in disponibles if e["id"] not in ids_sel]  # los enchufes elegidos ya no están disponibles para otro circuito

            # calcula si hacen falta cajas de paso adicionales cuando el circuito
            # reune enchufes de más de un ambiente
            if ambientes_con_enchufe_general:
                ambientes_str_final = ", ".join(sorted(set(ambientes_con_enchufe_general)))  # texto final con los ambientes que efectivamente aportaron enchufes
                _n_amb = len(set(ambientes_con_enchufe_general))  # cuántos ambientes distintos aportaron enchufes
                _n_ench = len([it for it in items if it.get("amb", "") in ambientes_con_enchufe_general])  # cuántos enchufes hay en total en esos ambientes
                # un solo ambiente, o como máximo un enchufe por ambiente:
                # no hace falta caja adicional
                if _n_amb <= 1 or _n_ench <= _n_amb:
                    _cajas_adic = 0  # no hace falta caja adicional
                else:  # hay varios ambientes y más enchufes que ambientes: se reparten cajas
                    _cajas_adic = max(0, _n_amb - 1)  # 1 caja extra por cada ambiente adicional que se conecta
            else:  # ningún ambiente aporto enchufes normales a este circuito
                # no hubo enchufes en el circuito general: no hay caja adicional
                _cajas_adic = 0

            if agrego_coc_lav:  # hubo enchufes de cocina o lavadero: van a su propio circuito
                # pedir longitud real por cada cocina/lavadero que entró
                for amb_cl in [a for a in sel_list if a.lower().startswith("cocina") or a.lower().startswith("lavadero")]:  # recorre solo los ambientes de tipo cocina o lavadero
                    # todavía no se pidió la longitud real de este ambiente
                    if amb_cl not in long_real_coc_lav_por_amb:  # solo pregunta una vez por ambiente
                        # un solo ambiente: se usa la longitud ya ingresada del circuito
                        if len(sel_list) == 1:  # si el circuito tiene un solo ambiente no hay nada que preguntar
                            long_real_coc_lav_por_amb[amb_cl] = longitud  # la longitud del circuito pasa a ser la de ese ambiente
                        # varios ambientes: se pregunta la longitud real de cada uno
                        else:  # el circuito trae varios ambientes
                            long_real_coc_lav_por_amb[amb_cl] = pedir_longitud_sub(f"Enchufes {amb_cl}", longitud)  # guarda la longitud real que respondió el usuario para ese ambiente
        else:  # no era iluminación, enchufes ni especial
            # circuito especial "genérico" que no calzó en ninguna categoría anterior:
            # se le asigna toda la potencia del ambiente completo
            for amb in sel_list:  # recorre los ambientes elegidos para este circuito genérico
                p = float(amb_idx[amb.lower()]["Total por ambiente (W)"])  # potencia total del ambiente completo, sin desglosar por componente
                potencia_estimada += p  # se suma al total del circuito
                items.append({"amb": amb, "potencia": p})  # se guarda como un único item para ese ambiente

        # permite crear el circuito manualmente si solo tiene enchufes de cocina/lavadero
        if es_enchufe and potencia_estimada == 0 and (len(items) == 0) and (len(items_coc_lav) > 0):  # no quedó potencia asignada pero si hay enchufes de cocina/lavadero juntados aparte
            potencia_estimada = sum(it["potencia"] for it in items_coc_lav)  # la potencia del circuito es la suma de los enchufes de cocina/lavadero
            items = items_coc_lav[:]  # copia la lista de enchufes de cocina/lavadero como items del circuito
            # marcar cocina/lavadero como creado manualmente
            ambs_cl = set(it["amb"] for it in items_coc_lav)  # ambientes involucrados en esta creacion manual
            for a in ambs_cl:  # recorre cada uno de esos ambientes
                coc_lav_creado_manual.add(a.strip().lower())  # los marca para no volver a ofrecerlos como cocina/lavadero pendiente

        # Subdivisión inmediata si supera potencia máxima permitida
        if potencia_estimada > 0:  # solo sigue si al circuito le quedó alguna carga
            # 220 V fija a propósito: la tensión real todavía NO se pregunta
            # (se pide recién en la PARTE 3, más abajo). Acá solo se necesita
            # una referencia para decidir si el circuito hay que partirlo.
            tension_ref = 220.0  # tensión de referencia para estimar la corriente (V)

            # calibres de TM permitidos según el tipo de circuito
            if es_enchufe:  # circuitos de enchufes
                allowed = [10, 16]  # enchufes: solo se permiten TM de 10 o 16A
            elif es_ilumin:  # circuitos de iluminación
                allowed = [6, 10, 16]  # iluminación: se permiten TM de 6, 10 o 16A
            else:  # especiales y cualquier otro caso
                allowed = [16]  # cualquier otro caso: TM de 16A

            I_max = allowed[-1]  # el calibre más grande permitido
            # P = I x V x 0,9. El 0,9 es criterio del script (no sale de un
            # artículo del RIC): deja el circuito cargado como máximo al 90%
            # del TM, o sea un 10% de holgura. Como el mayor calibre es 16A
            # tanto en enchufes ([10,16]) como en iluminación ([6,10,16]), da
            # P_max = 16 x 220 x 0,9 = 3168 W para enchufes e iluminación.
            P_max = I_max * tension_ref * 0.9  # potencia máxima que soporta ese calibre (con 10% de margen)
            # evita división por cero o un máximo negativo más abajo
            if P_max <= 0:  # si el margen quedó en cero o negativo no sirve para dividir
                P_max = potencia_estimada  # deja el máximo igual a la potencia, así no queda en cero


            # Si la potencia estimada supera lo que soporta el TM (interruptor termomagnético)
            # más grande permitido, hay que partir el circuito en varios subcircuitos.
            if potencia_estimada > P_max and len(items) > 1:  # se pasó del TM más grande y hay más de un item para repartir
                # se pasó del máximo: hay que partirlo en varios subcircuitos
                bins = binpack_items(items, P_max)  # reparte los items en grupos (bins) que no superen P_max cada uno
                # recorre cada grupo armado por binpack_items: cada uno será un subcircuito
                for idx_bin, b in enumerate(bins, start=1):  # idx_bin es el número del subcircuito, b es el grupo de items
                    pot_bin = float(b["potencia_total"])  # potencia total de este subcircuito
                    ambs_bin = sorted(set(it["amb"] for it in b["items"]))  # ambientes únicos que quedaron en este subcircuito
                    ambs_str_bin = ", ".join(ambs_bin)  # los junta en un solo texto, separado por comas
                    nombre_sub = f"{base} ({ambs_str_bin}) - subcircuito {idx_bin}"  # nombre del subcircuito, con número correlativo
                    long_real = pedir_longitud_sub(nombre_sub, longitud)  # cada subcircuito pregunta su propia longitud real
                    # texto con el detalle de items por ambiente (se llena más abajo)
                    detalle = ""  # se llena según el tipo de circuito
                    # arma el detalle según el tipo de circuito (enchufe, iluminación o especial)
                    if es_enchufe:  # subcircuito de enchufes
                        detalle = resumen_items_por_ambiente(b["items"], modo="enchufe")  # resumen de enchufes por ambiente para el informe
                    elif es_ilumin:  # subcircuito de iluminación
                        detalle = resumen_items_por_ambiente(b["items"], modo="iluminacion")  # resumen de centros de iluminación por ambiente
                    elif es_especial:  # subcircuito especial
                        detalle = resumen_items_por_ambiente(items, modo="especial")  # ojo: acá pasa items completo y no b['items']
                    # crea el subcircuito con sus datos: nombre, longitud real, potencia e items
                    add_circuito(nombre_sub, long_real, pot_bin,  # nombre del subcircuito, su longitud real y la potencia del bin
                                 es_ilumin=es_ilumin, es_enchufe=es_enchufe, es_especial=False,  # el subcircuito ya no se marca como especial
                                 items=b["items"], detalle_asignacion=detalle,  # items del bin y el texto de detalle para el informe
                                 tiene_tramo_20m=tiene_tramo_20m)  # arrastra si el circuito tiene algún tramo sobre 20 m
            else:  # la potencia cabe en un solo TM, no hay que partirlo
                # cabe en un solo circuito, no hace falta dividir
                nombre_circ = f"{base} ({ambientes_str_final})"  # nombre final del circuito, con sus ambientes entre paréntesis
                detalle = ""  # texto con el detalle de items por ambiente, se llena más abajo si corresponde
                # solo arma el detalle si quedó algún item asignado a este circuito
                if items:  # sin items no hay nada que detallar
                    if es_enchufe:  # circuito de enchufes
                        detalle = resumen_items_por_ambiente(items, modo="enchufe")  # resumen de enchufes por ambiente
                    elif es_ilumin:  # circuito de iluminación
                        detalle = resumen_items_por_ambiente(items, modo="iluminacion")  # resumen de centros por ambiente
                    elif es_especial:  # circuito especial
                        detalle = resumen_items_por_ambiente(items, modo="especial")  # resumen del especial (lavadora, horno, etc.)

                long_final = longitud  # por defecto usa la misma longitud que se ingresó al principio del circuito
                # para enchufes: decide qué longitud real usar, según si el circuito
                # quedó compuesto solo por cocina/lavadero o mezclado con otros ambientes
                if es_enchufe and len(items) > 0:  # en enchufes la longitud puede cambiar si hay cocina/lavadero
                    ambs_items = {str(it.get("amb", "")).strip().lower() for it in items if str(it.get("amb", "")).strip()}  # ambientes en minúscula y sin espacios, sin contar los vacíos
                    # true si TODOS los ambientes de este circuito son cocina o lavadero
                    solo_coc_lav = all(  # all() con lista vacía daria true, por eso el else de más abajo
                        a.startswith("cocina") or a.startswith("lavadero")  # basta con que el nombre del ambiente parta con cocina o lavadero
                        for a in ambs_items  # revisa uno por uno los ambientes del circuito
                    ) if ambs_items else False  # si no quedó ningún ambiente, no cuenta como cocina/lavadero
                    # Si el circuito final es solo cocina/lavadero, usar la longitud ya ingresada antes
                    if solo_coc_lav and len(ambs_items) == 1:  # un solo ambiente y además es cocina o lavadero
                        amb_unico = next(iter(ambs_items))  # el único ambiente cocina/lavadero que quedó en este circuito
                        long_final = float(long_real_coc_lav_por_amb.get(amb_unico, longitud))  # usa la longitud real que se preguntó antes, para ese ambiente
                    # Si quedó un circuito general restante (baño/living/pasillo, etc.), preguntar su longitud real
                    elif agrego_coc_lav:  # el circuito mezcla otros ambientes, hay que preguntar aparte
                        long_final = pedir_longitud_sub(nombre_circ, longitud)  # pregunta la longitud real de este circuito general (no es solo cocina/lavadero)
                # crea el circuito final, ya no hace falta subdividirlo
                add_circuito(nombre_circ, long_final, potencia_estimada,  # nombre, longitud final y potencia total del circuito
                            es_ilumin=es_ilumin, es_enchufe=es_enchufe, es_especial=False,  # conserva el tipo del circuito, ya no es especial
                            items=items, detalle_asignacion=detalle,  # items asignados y el texto de detalle
                            tiene_tramo_20m=tiene_tramo_20m)  # marca si hay algún tramo de más de 20 m
        # si no se pudo estimar ninguna potencia, no se crea el circuito
        else:  # no se pudo estimar ninguna potencia
            print("       ! No se asignaron cargas a este circuito (solo tenía enchufes de cocina/lavadero o nada). No se crea.")  # avisa que el circuito quedó vacío y no se agrega


# Circuitos por ambiente para enchufes de cocina/lavadero
# ---------------------------------------------------------------------------
# Esto corre DESPUÉS de terminar de recorrer todos los circuitos. Mientras el
# usuario armaba circuitos, cada enchufe de cocina o lavadero que eligió NO se
# metio en el circuito general: se fue guardando aparte en
# enchufes_coc_lav_por_amb = {"cocina 1": [ {amb, potencia, longitud,
# id_ench, módulos}, ... ], "lavadero": [...] }.
# Acá se recorre esa bolsa ambiente por ambiente y se le crea un circuito
# dedicado a cada uno, salvo que el usuario ya lo haya armado a mano (esos
# ambientes quedaron anotados en coc_lav_creado_manual).
# OJO con el "16A": acá NO se fija ningún calibre. El TM de 16A se decide más
# abajo, en el bucle de la PARTE 3 que asigna termomagnéticos: ahí se fuerza
# 16A si el NOMBRE del circuito contiene "cocina" o "lavadero", y el nombre
# que se arma acá (nombre_circ) siempre los contiene.
if enchufes_coc_lav_por_amb:  # solo entra si quedaron enchufes de cocina/lavadero anotados
    for amb_cl, lista_ench in enchufes_coc_lav_por_amb.items():  # amb_cl es el ambiente y lista_ench sus enchufes pendientes

        # si ya se creó manual, NO crear automático
        if amb_cl.strip().lower() in coc_lav_creado_manual:  # ese ambiente ya tiene un circuito armado a mano
            continue  # pasa al siguiente ambiente sin crear nada

        if not lista_ench:  # no hay enchufes pendientes para este ambiente
            continue  # sin enchufes pendientes no hay circuito que armar

        potencia_total = sum(float(e["potencia"]) for e in lista_ench)  # suma la potencia de todos los enchufes pendientes de este ambiente
        if potencia_total <= 0:  # no hay nada que instalar
            continue  # potencia en cero, no vale la pena crear el circuito

        long_max = max(float(e.get("longitud", 0.0)) for e in lista_ench)  # la longitud más larga entre esos enchufes, por si no se preguntó una longitud real
        long_final = float(long_real_coc_lav_por_amb.get(amb_cl, long_max))  # usa la longitud real si ya se había preguntado antes, si no, la más larga

        nombre_circ = f"Enchufes {amb_cl} (cocina/lavadero)"  # nombre del circuito automático

        # arma la lista de items en el formato que espera add_circuito
        items_cl = [{  # cada enchufe pendiente pasa a ser un item del circuito
            "amb": e["amb"],  # ambiente al que pertenece el enchufe
            "potencia": float(e["potencia"]),  # potencia del enchufe en watts
            "n_ench": 1,  # cada item representa un solo enchufe
            "id_ench": int(e.get("id_ench", 0)),  # id del enchufe, para no perderle el rastro
            "modulos": int(e.get("modulos", 1))  # módulos que ocupa el enchufe en la caja
        } for e in lista_ench]  # uno por cada enchufe pendiente del ambiente

        detalle_cl = resumen_items_por_ambiente(items_cl, modo="enchufe")  # texto con el resumen de enchufes por ambiente

        # crea el circuito automático de enchufes de cocina/lavadero
        add_circuito(  # crea el circuito dedicado de cocina/lavadero
            nombre_circ, long_final, potencia_total,  # nombre, longitud y potencia total del circuito
            es_ilumin=False, es_enchufe=True, es_especial=False,  # es de enchufes, no de iluminación ni especial
            items=items_cl, detalle_asignacion=detalle_cl  # items del circuito y el texto de detalle
        )
        # Marca heredada: pone la bandera "es_coc_lav" en el circuito recién creado.
        # OJO 1: 'items' y 'es_enchufe' son los que quedaron colgando de la ÚLTIMA
        #        vuelta del bucle de circuitos de más arriba, no de este bucle; o sea
        #        la condición depende del último circuito que armo el usuario.
        # OJO 2: la bandera "es_coc_lav" no la lee nadie en el resto del script; se
        #        borra antes de exportar (ver el pop de campos internos en la PARTE 3.3).
        #        Lo que de verdad fuerza el TM de 16A es el nombre del circuito.
        if es_enchufe and any(  # ojo: items viene del bucle anterior, no de este circuito
            it["amb"].lower().startswith(("cocina", "lavadero"))  # el ambiente parte con cocina o lavadero
            for it in items  # recorre los items que quedaron de la vuelta anterior
        ):
            circuitos[-1]["es_coc_lav"] = True  # bandera informativa; el TM 16A lo decide el nombre del circuito, no esta marca

        # Paso 4: si el usuario ya armo a mano un circuito con esos mismos enchufes
        # de cocina/lavadero, se vacía su lista pendiente para no crear un segundo
        # circuito automático con la misma carga.
        # Ojo: items_coc_lav también viene de la última vuelta del bucle anterior.
        if es_enchufe and (len(items_coc_lav) > 0):  # el circuito manual anterior ya cubria cocina/lavadero
            ambs_cl = set(it["amb"] for it in items_coc_lav)  # ambientes que ya quedaron cubiertos por el circuito manual
            for amb_cl in ambs_cl:  # recorre los ambientes que ya quedaron cubiertos
                if amb_cl in enchufes_coc_lav_por_amb:  # solo si ese ambiente tenía enchufes anotados como pendientes
                    enchufes_coc_lav_por_amb[amb_cl] = []  # vacía la lista para no volver a crear un circuito automático duplicado

        # Paso 5: repite la misma marca del paso anterior (queda redundante, porque
        # más abajo (línea ~10607) se pone igual y sin condición).
        if es_enchufe and (len(items_coc_lav) > 0):  # mismo chequeo del paso anterior
            circuitos[-1]["es_coc_lav"] = True  # repite la marca; no cambia nada del cálculo

        # Marca final, esta vez sin condición: todo circuito creado en este bucle
        # queda con es_coc_lav=True y con una copia de sus enchufes.
        # Los dos campos se borran antes de exportar (pop de campos internos en la
        # PARTE 3.3) y ninguna otra parte del script los lee.
        circuitos[-1]["es_coc_lav"] = True  # bandera informativa, no se usa después
        circuitos[-1]["enchufes_coc_lav"] = lista_ench[:]  # copia de los enchufes del ambiente; hoy no la consume nadie

# =========================================================
# RESPALDO DE _items Y CAJAS ADICIONALES POR CIRCUITO (antes de la PARTE 3)
# =========================================================


# Guardar _items por nombre de circuito ANTES de PARTE 3
# En la PARTE 3.3 los campos internos (los que empiezan con "_") se borran de
# cada circuito con c.pop(...), así que después de eso ya no hay forma de
# recuperar el detalle de luminarias/enchufes. Por eso se congela acá.
import copy as _copy  # para poder hacer copias profundas de listas/diccionarios
# items_por_nombre     = {"nombre del circuito": [ {amb, potencia, ...}, ... ]}
# cajas_adic_por_nombre = {"nombre del circuito": n° de cajas de derivación extra}
# Los dos se le pasan después a build_materiales_df(), que los usa para contar
# chicotes, conectores cónicos, cajas y tapas.
# Recordatorio: "_cajas_adic" lo escribe add_circuito() leyendo la variable
# GLOBAL _cajas_adic que quedó de la última vuelta del bucle de circuitos, y
# solo la guarda cuando el circuito es de enchufes.
items_por_nombre = {str(c.get("Circuito","")): _copy.deepcopy(c.get("_items", [])) for c in circuitos}  # copia profunda, para no compartir referencias
cajas_adic_por_nombre = {str(c.get("Circuito","")): int(c.get("_cajas_adic", 0)) for c in circuitos}  # guarda cuántas cajas adicionales necesita cada circuito, por nombre


# A partir de aquí el script deja de armar circuitos y pasa a preguntar los
# datos generales: tensión, empalme, acometida, alimentador y distancias.
# Estos datos se usan más adelante para calcular el TM general y el alimentador.
# -------- PARTE 3: TABLAS + INTERRUPTOR TERMOMAGNÉTICO --------
tension = float(input("\n- Ingrese la tensión nominal [V]: "))  # tensión de la red, en volts (ej: 220)
factor_potencia = float(input("- Ingrese el factor de potencia (ej. 0.92): "))  # factor de potencia de la instalación (cos fi), se usa para calcular corrientes
temperatura = float(input("- Ingrese la temperatura ambiente típica de la zona: "))  # temperatura ambiente típica de la zona, corrige la capacidad de los conductores
longitud_alimentador = float(input("Ingrese la longitud del alimentador (Empalme → Tablero) en metros: "))  # tramo entre el empalme y el tablero de la casa
longitud_transformador_empalme = float(input("Ingrese la longitud de la acometida (Transformador → Empalme) en metros: "))  # tramo entre el transformador de la red pública y el empalme (acometida)
# pregunta el tipo de acometida (aérea o subterránea) hasta que sea una respuesta válida
while True:  # repite hasta que la respuesta sirva
    tipo_acometida = input("Ingrese el tipo de acometida (Aérea / subterránea): ").strip().lower()  # lo que escribió el usuario, sin espacios y en minúscula
    if "aer" in tipo_acometida or "sub" in tipo_acometida:  # acepta cualquier respuesta que traiga 'aer' o 'sub'
        tipo_acometida = "aerea" if "aer" in tipo_acometida else "subterranea"  # deja el valor fijo en 'aérea' o 'subterránea'
        break  # respuesta válida, sale del bucle
    print("  Valor no reconocido. Ingrese 'aerea' o 'subterranea'.")  # no entendió la respuesta, vuelve a preguntar
# Temperatura del suelo: se pregunta más adelante, después de ingresar tipo_alimentador,
# ya que puede aplicar a acometida subterránea, alimentador subterráneo, o ambos.
# Ver bloque "TEMPERATURA DEL SUELO" más abajo (RIC 4 punto 6.2.5, nota tabla 4.7).
temperatura_suelo = temperatura  # valor por defecto hasta que se calcule abajo
# =========================
# DATOS DE UBICACIÓN DEL EMPALME
# =========================
# pregunta si el empalme queda a menos de 15 m del acceso a la propiedad
while True:  # repite hasta que responda si o no
    empalme_dentro_15m = input("¿El empalme se ubica dentro de 15 m del acceso a la propiedad? (si/no): ").strip().lower()  # distancia del empalme al acceso de la propiedad
    if "si" in empalme_dentro_15m or "no" in empalme_dentro_15m:  # acepta si la respuesta trae 'si' o 'no'
        empalme_dentro_15m = "si" if "si" in empalme_dentro_15m else "no"  # deja el valor fijo en 'si' o 'no'
        break  # respuesta válida, sigue
    print("  Valor no reconocido. Ingrese 'si' o 'no'.")  # no entendió, pregunta de nuevo
# pregunta si el empalme va en la fachada de la casa o en una estructura aparte (poste)
while True:  # repite hasta que la respuesta calce con el RIC
    tipo_instalacion_empalme = input("¿El empalme se instalará en fachada o en estructura independiente? (fachada/independiente): ").strip().lower()  # donde se va a montar el empalme
    if "fachada" in tipo_instalacion_empalme or "independiente" in tipo_instalacion_empalme:  # acepta si trae 'fachada' o 'independiente'
        tipo_instalacion_empalme = "fachada" if "fachada" in tipo_instalacion_empalme else "independiente"  # deja el valor fijo en 'fachada' o 'independiente'
        # Validación RIC N°01: la distancia al acceso obliga a un tipo de instalación específico
        if empalme_dentro_15m == "si" and tipo_instalacion_empalme == "independiente":  # esta combinación no la permite el RIC
            print("  Según RIC N°01 punto 7.2, un empalme dentro del radio de 15 m debe instalarse")  # avisa la regla del RIC N01 punto 7.2
            print("    en la fachada de la vivienda. Por favor seleccione 'fachada'.")  # le pide que corrija a fachada
            continue  # vuelve a preguntar, esa combinación no se puede
        if empalme_dentro_15m == "no" and tipo_instalacion_empalme == "fachada":  # esta combinación tampoco la permite el RIC
            print("  Según RIC N°01 punto 7.3, un empalme fuera del radio de 15 m debe instalarse")  # avisa la regla del RIC N01 punto 7.3
            print("    en estructura independiente cerca del cierre de la propiedad. Por favor seleccione 'independiente'.")  # le pide que corrija a independiente
            continue  # vuelve a preguntar
        break  # la combinación pasa la validación del RIC
    print("  Valor no reconocido. Ingrese 'fachada' o 'independiente'.")  # respuesta no reconocida, insiste
tipo_poste = ""  # solo se usa si el empalme es independiente (con poste)
altura_acometida_aerea = 0  # solo aplica si el poste tiene acometida aérea
longitud_subterraneo_medidor2 = 0  # solo aplica si el poste tiene acometida subterránea
# si el empalme va en un poste aparte, hay que preguntar más datos del poste
if tipo_instalacion_empalme == "independiente":  # el empalme va en poste, hay que pedir más datos
    # pregunta el material del poste
    while True:  # repite hasta que responda madera o metálico
        tipo_poste = input("¿El poste será de madera o metálico? (madera/metalico): ").strip().lower()  # material del poste, sirve para la lista de materiales
        if "madera" in tipo_poste or "metal" in tipo_poste:  # acepta 'madera' o cualquier cosa con 'metal'
            tipo_poste = "madera" if "madera" in tipo_poste else "metalico"  # deja el valor fijo en 'madera' o 'metálico'
            break  # respuesta válida
        print("  Valor no reconocido. Ingrese 'madera' o 'metalico'.")  # no entendió, pregunta otra vez
    # Solo para acometida aérea
    if "aer" in tipo_acometida:  # el poste recibe la acometida por arriba
        # altura del tramo aéreo, desde el empalme hasta donde llega la acometida
        altura_acometida_aerea = float(input(  # largo del tramo aéreo que sube por el poste
                "Ingrese la altura desde el empalme hasta el punto de llegada de la acometida aérea en metros: "))  # cierra el input del tramo aéreo
    # Solo para acometida subterránea
    elif "sub" in tipo_acometida:  # la acometida llega enterrada al poste
        # tramo subterráneo desde la salida de la acometida hasta el medidor
        longitud_subterraneo_medidor2 = float(input(  # largo del tramo enterrado hasta el medidor
                "Ingrese la longitud desde la salida subterránea de la acometida hasta el medidor en metros: "))  # cierra el input del tramo enterrado
requiere_mastil = ""  # solo se usa si el empalme va en fachada
longitud_mastil = 0  # largo del mástil, si es que corresponde
# en fachada, según el tipo de acometida, puede ser obligatorio usar mástil
if tipo_instalacion_empalme == "fachada" and ("aer" in tipo_acometida or "sub" in tipo_acometida):  # el empalme va en fachada, puede necesitar mástil
    while True:  # repite hasta que la respuesta sea compatible
        tipo_txt = "aérea" if "aer" in tipo_acometida else "subterránea"  # solo se usa para mostrar el mensaje con el texto correcto
        requiere_mastil = input(f"¿La acometida {tipo_txt} requiere mástil? (si/no): ").strip().lower()  # responde si el empalme lleva mástil o no
        if "sub" in tipo_acometida and requiere_mastil == "si":  # no se permite mástil si la acometida es subterránea
            print(" No se puede utilizar mástil en un empalme en fachada con acometida subterránea.")  # combinación prohibida: acometida subterránea con mástil
            continue  # vuelve a preguntar
        break  # respuesta aceptada
    if requiere_mastil == "si":  # solo si lleva mástil se pide su largo
        longitud_mastil = float(input("Ingrese el largo del mástil (Caja empalme - extremo del mástil en metros): "))  # distancia desde la caja de empalme hasta la punta del mástil

# tipo de alimentador: cada combinación con el tipo de empalme tiene reglas propias
# pregunta el tipo de alimentador y valida que sea compatible con el tipo de empalme
while True:  # repite hasta que el alimentador calce con el tipo de empalme
    # primero valida que la respuesta sea una de las tres opciones válidas
    while True:  # bucle interno: solo revisa que el texto ingresado se entienda
        _resp_alim = input("Ingrese el tipo de alimentador (Aéreo / Subterráneo / En ducto): ").strip().lower()  # lo que escribió el usuario, en minúscula y sin espacios
        if "aer" in _resp_alim:  # respuesta con 'aer'
            tipo_alimentador = "aereo"  # alimentador aéreo
            break  # ya quedó el tipo, sale del bucle interno
        elif "sub" in _resp_alim:  # respuesta con 'sub'
            tipo_alimentador = "subterraneo"  # alimentador enterrado
            break  # ya quedó el tipo, sale del bucle interno
        elif "duct" in _resp_alim or "duc" in _resp_alim:  # respuesta con 'duct' o 'duc'
            tipo_alimentador = "en ducto"  # alimentador dentro de ducto
            break  # ya quedó el tipo, sale del bucle interno
        else:  # escribió cualquier otra cosa
            print("  Valor no reconocido. Ingrese 'aereo', 'subterraneo' o 'en ducto'.")  # vuelve a preguntar el tipo de alimentador
    # Si es fachada, solo se permite en ducto
    if (tipo_instalacion_empalme == "fachada" and ("aer" in tipo_alimentador or "sub" in tipo_alimentador)):  # en fachada sin mástil no se admite aéreo ni subterráneo
        print(" No se puede poner un alimentador aéreo o subterráneo en fachada sin mástil. Debe seleccionar: En ducto.")  # le explica por qué no se puede
        continue  # vuelve a preguntar desde el bucle interno
    # Si es independiente, no se permite en ducto
    elif (tipo_instalacion_empalme == "independiente" and "duct" in tipo_alimentador):  # el poste independiente no admite ducto
        print("No se puede poner un alimentador en ducto en empalme independiente. Debe seleccionar: aéreo o subterráneo.")  # le explica por qué no se puede
        continue  # vuelve a preguntar
    break  # la combinación empalme-alimentador es válida
# Poste independiente con acometida AÉREA + alimentador SUBTERRÁNEO también
# necesita longitud_subterraneo_medidor2: la usan el tubo galvanizado del
# poste (la acometida usa altura_acometida_aerea y el alimentador subterráneo
# usa este tramo, cada uno por su lado; solo caen en la misma fila si coincide
# el diámetro), el conduit PVC del
# alimentador (se resta este tramo) y las cámaras tipo C del alimentador.
if (tipo_instalacion_empalme == "independiente" and "aer" in tipo_acometida and "sub" in tipo_alimentador):  # poste con acometida aérea y alimentador enterrado
    longitud_subterraneo_medidor2 = float(input(  # tramo enterrado del alimentador hasta el medidor
        "Ingrese la longitud desde la salida subterránea del alimentador hasta el medidor en metros: "))  # cierra el input del tramo enterrado del alimentador
#Alimentador aéreo, empalme independiente
# TDA = tablero de distribucion de la vivienda (donde llega el alimentador y
# de donde salen los circuitos). Es el mismo "Tablero" del input del alimentador.
longitud_llegada_aerea_tda = 0  # tramo desde la llegada del alimentador aéreo en la casa hasta el TDA, en metros
longitud_poste_alimentador_aereo = 0  # tramo de subida del alimentador aéreo en el poste
if (tipo_instalacion_empalme == "independiente" and "aer" in tipo_alimentador):  # poste independiente con alimentador aéreo
    longitud_llegada_aerea_tda = float(  # tramo desde la llegada del alimentador aéreo hasta el TDA
    input("Ingrese la longitud desde la llegada del alimentador aéreo en la casa hasta el TDA en metros: "))  # pregunta el tramo de llegada del alimentador aéreo hasta el TDA
    longitud_poste_alimentador_aereo = float(  # tramo del alimentador que sube por el poste
    input("Ingrese la longitud de subida del tramo del alimentador aéreo en el poste en metros: "))  # pregunta el tramo de subida del alimentador por el poste
# Alimentador subteraneo, empalme independiente, para cálculo de abrazaderas pvc
longitud_abrazaderas_alimentador = 0  # tramo subterráneo desde la salida del alimentador hasta el TDA
if (tipo_instalacion_empalme == "independiente" and "sub" in tipo_alimentador):  # poste independiente con alimentador enterrado
    longitud_abrazaderas_alimentador = float(  # largo del tramo subterráneo, define cuántas abrazaderas van
        input("Ingrese la longitud desde la salida subterránea del alimentador hasta el TDA en metros: "))  # pregunta el tramo enterrado del alimentador hasta el TDA
# Longitud salida subterránea en empalme en fachada
longitud_subterraneo_medidor = 0  # tramo subterráneo desde la salida de la acometida hasta el medidor, en fachada sin mástil
if (tipo_instalacion_empalme == "fachada" and requiere_mastil == "no" and "sub" in tipo_acometida  # empalme en fachada sin mástil y acometida enterrada
    and "duct" in tipo_alimentador):  # además el alimentador tiene que ir en ducto
    longitud_subterraneo_medidor = float(  # largo del tramo enterrado hasta el medidor
    input("Ingrese la longitud desde la salida subterránea de la acometida hasta el medidor en metros: ")  # pregunta el tramo enterrado de la acometida hasta el medidor
    )
# Distancia vertical acometida aérea en fachada sin mástil
dist_vertical_acometida = 0  # distancia vertical entre la caja de empalme y el punto de llegada de la acometida
if (tipo_instalacion_empalme == "fachada" and requiere_mastil == "no" and "aer" in tipo_acometida):  # fachada sin mástil y acometida aérea
    dist_vertical_acometida = float(  # distancia vertical que hay que subir hasta la acometida
        input("Ingrese la distancia vertical (caja del empalme - punto de llegada de la acometida en metros: "))  # pregunta la distancia vertical hasta el punto de llegada
dist_empalme_pt1 = float(input("Ingrese la distancia del empalme a la puesta a tierra 1 (camarilla N°1) en metros: ") or 0)  # distancia del empalme a la primera puesta a tierra (camarilla N°1)
dist_tda_pt2 = float(input("Ingrese la distancia del TDA a la puesta a tierra 2 (camarilla N°2) en metros: ") or 0)  # distancia del TDA a la segunda puesta a tierra (camarilla N°2)

# =========================
# TEMPERATURA DEL SUELO (RIC 4 punto 6.2.5, nota tabla 4.7)
# Se pregunta si acometida O alimentador son subterráneos
# =========================
_acom_sub = "sub" in str(tipo_acometida).lower()  # true si la acometida es subterránea
_alim_sub = "sub" in str(tipo_alimentador).lower()  # true si el alimentador es subterráneo

# solo pregunta la temperatura del suelo si hay algún tramo subterráneo
if _acom_sub or _alim_sub:  # solo tiene sentido preguntar si hay algo enterrado
    _tramos_sub = []  # junta los nombres de los tramos subterráneos, para armar el mensaje
    if _acom_sub:  # la acometida va enterrada
        _tramos_sub.append("acometida subterránea")  # la agrega al texto del aviso
    if _alim_sub:  # el alimentador va enterrado
        _tramos_sub.append("alimentador subterráneo")  # lo agrega al texto del aviso
    _tramos_str = " y ".join(_tramos_sub)  # arma el texto final, ej: 'acometida subterránea y alimentador subterráneo'
    print(f"\n  Tiene tramos subterráneos: {_tramos_str}.")  # le muestra qué tramos van bajo tierra
    print("  La temperatura del suelo afecta el factor de corrección de ampacidad (ft) del método D1.")  # recuerda para qué se usa la temperatura del suelo
    # repite hasta que ingresen un número válido
    while True:  # repite hasta que escriba un número
        try:  # por si escribe letras en vez de número
            # si aprieta ENTER sin escribir nada, usa la temperatura ambiente como valor por defecto
            temperatura_suelo = float(input(  # lee la temperatura del suelo para los tramos enterrados
                f"  Ingrese la temperatura del suelo para tramos subterráneos [°C] "
                f"(RIC 4 punto 6.2.5, nota tabla 4.7) [ENTER = {temperatura}°C]: "
            ).strip() or temperatura)  # si aprieta ENTER, queda la temperatura ambiente
            break  # ya quedó la temperatura del suelo, sale
        except ValueError:  # lo que escribió no se pudo pasar a número
            print("  ! Ingrese un número válido.")  # avisa y vuelve a preguntar
else:  # no hay tramos subterráneos, la temperatura del suelo queda igual a la ambiente
    temperatura_suelo = temperatura  # si no hay tramos subterráneos, queda igual a la T° ambiente

# =========================
# PUESTA A TIERRA - CÁLCULO RIC 6
# =========================
# Se dimensionan DOS puestas a tierra independientes:
#   PT N°1: la del empalme  (camarilla N°1, a dist_empalme_pt1 metros)
#   PT N°2: la del tablero  (camarilla N°2, a dist_tda_pt2 metros)
# Para cada una se pregunta la resistividad del terreno y se calcula cuántas
# barras copperweld hacen falta. Los resultados (n_barras_pt1 / n_barras_pt2)
# los usa después build_materiales_df para contar barras, camarillas,
# conectores y metros de conductor desnudo.

# Largo de barra copperweld
# Pregunta el largo de barra a usar y repite hasta que responda 1 o 2
print("\n  Largo de barra copperweld:")  # menú del largo de la barra copperweld
print("    1) 3 metros (recomendado)")  # opción recomendada
print("    2) 1,5 metros (solo si hay restricción de profundidad)")  # opción corta, solo si el terreno no da para 3 m
while True:  # repite hasta que elija 1 o 2
    op_barra = input("  Seleccione opción (ENTER = 1): ").strip()  # opción elegida para el largo de la barra
    if op_barra in ["", "1"]:  # opción por defecto: 3 metros
        largo_barra_pt = 3.0  # barra estándar de 3 metros
        desc_barra_pt  = "Barra copperweld 5/8 3mts + conector de bronce"  # texto para el informe y la lista de materiales
        break  # ya quedó la barra de 3 m
    elif op_barra == "2":  # opción alternativa: barra más corta
        largo_barra_pt = 1.5  # barra corta, para terrenos con poca profundidad
        desc_barra_pt  = "Barra copperweld 5/8 1,5mts + conector de bronce"  # texto para el informe y la lista de materiales
        break  # ya quedó la barra de 1,5 m
    else:  # escribió cualquier otra cosa
        print("  ! Ingrese 1 o 2.")  # opción inválida, vuelve a preguntar

# Pregunta la resistividad del terreno (rho) para una puesta a tierra (PT).
# Devuelve una tupla (valor numérico en Ohm.m, texto descriptivo para el informe).
def _pedir_resistividad(etiqueta):  # etiqueta indica de qué puesta a tierra se está hablando
    """Pide resistividad del terreno para una PT dada."""
    print(f"\n  Resistividad del terreno (ρ) - {etiqueta}:")  # título con la PT que se está preguntando
    print("    1) Ingresar valor medido en Ohm·m")  # opción 1: ya midió la resistividad en terreno
    print("    2) Estimar por tipo de terreno (Tabla 6.3 RIC 6)")  # opción 2: estimarla según el tipo de terreno
    # repite la pregunta hasta que el usuario elija una opción válida
    while True:  # repite hasta que elija 1 o 2
        op_rho = input("  Seleccione opción: ").strip()  # opción elegida en el  menú
        if op_rho == "1":  # eligió ingresar el valor medido
            # el usuario ya midió la resistividad en terreno
            while True:  # repite hasta que ingrese un valor mayor que 0
                try:  # por si escribe letras en vez de número
                    rho = float(input("  Ingrese resistividad medida (Ohm·m): ").strip())  # convierte lo ingresado a número
                    if rho > 0:  # la resistividad debe ser positiva
                        return rho, f"{rho} Ohm·m - Medido"  # devuelve el valor y una descripción para el informe
                    print("  ! Debe ser mayor que 0.")  # valor cero o negativo, no sirve
                except:  # el usuario no ingreso un número válido
                    print("  ! Ingrese un número válido.")  # no era un número, lo pide de nuevo
        elif op_rho == "2":  # eligió estimar por tipo de terreno
            # no tiene medición: usa un valor típico según el tipo de terreno
            print("    1) Terreno fértil / húmedo         →  50 Ohm·m")  # valores típicos de la tabla 6.3 del RIC 6
            print("    2) Terraplén poco fértil            → 500 Ohm·m")  # terraplén poco fértil
            print("    3) Pedregoso / arena seca           → 3000 Ohm·m")  # pedregoso o arena seca, el peor caso
            while True:  # repite hasta que elija un tipo de terreno
                op_ter = input("  Seleccione tipo de terreno: ").strip()  # tipo de terreno elegido
                if op_ter == "1":  # terreno fértil o húmedo
                    return 50.0, "50 Ohm·m - Fértil/húmedo"  # terreno fértil: resistividad típica baja
                elif op_ter == "2":  # terraplén poco fértil
                    return 500.0, "500 Ohm·m - Poco fértil"  # terraplén: resistividad típica media
                elif op_ter == "3":  # pedregoso o seco
                    return 3000.0, "3000 Ohm·m - Pedregoso/seco"  # terreno pedregoso o seco: resistividad alta, cuesta más aterrizar
                else:  # escribió otra cosa
                    print("  ! Ingrese 1, 2 o 3.")  # opción fuera de rango, pregunta de nuevo
        else:  # no eligió 1 ni 2
            print("  ! Ingrese 1 o 2.")  # vuelve al menú de resistividad

# Resistividad PT1 (empalme - camarilla N°1)
rho_pt1, desc_rho_pt1 = _pedir_resistividad("PT N°1 (empalme - camarilla N°1)")  # resistividad y descripción de la PT del empalme
# Resistividad PT2 (tablero - camarilla N°2)
rho_pt2, desc_rho_pt2 = _pedir_resistividad("PT N°2 (tablero - camarilla N°2)")  # resistividad y descripción de la PT del tablero

# Alias que quedaron de una versión anterior con una sola puesta a tierra.
# Hoy NO los lee nadie más en el script: la tabla resumen de más abajo arma sus
# filas con desc_rho_pt1 y desc_rho_pt2 directamente.
rho_terreno = rho_pt1  # copia de la resistividad de PT1; sin uso posterior
desc_rho    = desc_rho_pt1  # copia del texto de PT1, ej: "50 Ohm·m - Fértil/húmedo"; sin uso posterior

# Dimensiona una puesta a tierra: cuántas barras copperweld hacen falta para
# quedar bajo 20 Ohm (límite que exige el RIC 6 para la tierra de protección).
# Versión simplificada, a propósito:
#   R de una barra  = rho / largo   (aproximación, no la fórmula completa de
#                                    Dwight con el diámetro de la barra)
#   n barras en paralelo -> R / n   (ideal: no descuenta la interferencia
#                                    entre barras vecinas, así que el R real
#                                    en terreno sale algo más alto)
def _calcular_pt(rho, largo_barra):  # rho en Ohm·m y el largo de la barra copperweld elegida
    """Calcula N° barras, R final, separación y longitud conductor desnudo."""
    R_1b  = rho / largo_barra  # resistencia estimada si se usara solo 1 barra, en Ohm
    n     = max(1, math.ceil(R_1b / 20.0))  # barras necesarias para que R_1b/n quede en 20 Ohm o menos
    R_fin = R_1b / n  # resistencia final con esa cantidad de barras (en paralelo ideal)
    sep   = 2 * largo_barra  # separación mínima entre barras: 2 veces el largo (RIC 6, punto 8.3.2)
    long_desnudo = (n - 1) * sep  # metros de conductor desnudo para unir las barras en línea
    return n, R_1b, R_fin, sep, long_desnudo  # devuelve barras, R de 1 barra, R final, separación y metros de conductor desnudo

# Cálculo PT1
n_barras_pt1, R_1barra_pt1, R_final_pt1, sep_min_pt1, long_cond_desnudo_pt1 = _calcular_pt(rho_pt1, largo_barra_pt)  # dimensiona la PT del empalme (camarilla N°1)
# Cálculo PT2
n_barras_pt2, R_1barra_pt2, R_final_pt2, sep_min_pt2, long_cond_desnudo_pt2 = _calcular_pt(rho_pt2, largo_barra_pt)  # dimensiona la PT del tablero (camarilla N°2)

# Más alias de la versión de una sola puesta a tierra. Ninguna de estas cinco
# variables se vuelve a leer en el resto del script: tanto el resumen del Excel
# como la lista de materiales trabajan con los valores _pt1 y _pt2 por separado.
n_barras_pt       = max(n_barras_pt1, n_barras_pt2)  # la mayor entre PT1 y PT2; sin uso posterior
R_1barra          = R_1barra_pt1  # R de 1 barra de PT1; sin uso posterior
R_final_pt        = R_final_pt1  # R final de PT1; sin uso posterior
sep_min_pt        = sep_min_pt1  # separación mínima de PT1; sin uso posterior
long_cond_desnudo_pt = long_cond_desnudo_pt1  # metros de desnudo de PT1; sin uso posterior

# Aviso en pantalla si el terreno obliga a demasiadas barras (más de 6).
# Es solo informativo: no cambia el cálculo ni la lista de materiales, igual se
# siguen usando las n barras que salieron.
for n, label in [(n_barras_pt1, "PT1"), (n_barras_pt2, "PT2")]:  # revisa PT1 y PT2 por separado
    if largo_barra_pt == 1.5 and n > 6:  # con barras de 1,5m, más de 6 ya es demasiado
        print(f"\n  ADVERTENCIA {label}: Se requieren {n} barras de 1,5m.")  # avisa en pantalla cuántas barras salieron
        print("    Se recomienda usar barras de 3m o realizar un estudio especial de puesta a tierra.")  # sugiere pasar a barras de 3m
    elif largo_barra_pt == 3.0 and n > 6:  # mismo caso pero con barras de 3m
        print(f"\n  ADVERTENCIA {label}: Se requieren {n} barras de 3m.")  # avisa la cantidad de barras de 3m
        print("    Se recomienda realizar un estudio especial de puesta a tierra (RIC 6, punto 5.1).")  # a esta altura conviene un estudio de puesta a tierra

# potencia_total: solo alimenta la fila "Potencia total estimada (W)" del
# resumen del Excel. NO se usa para elegir el empalme (para eso esta
# potencia_total_instalacion_w, más abajo, en el bloque de factor de demanda).
# Dos diferencias con ese otro total, a tener presente:
#   - acá NO se descuenta lo que el usuario ya había anotado como clima o agua
#     caliente dentro de "Componentes especiales" del ambiente, así que un
#     equipo de clima puede quedar sumado dos veces (el estimado del ambiente
#     más la potencia de placa que se suma en la línea siguiente);
#   - acá NO se suma la potencia de los equipos de agua caliente.
potencia_total = sum(a["Total por ambiente (W)"] for a in ambientes_detalle)  # suma la potencia de todos los ambientes (iluminación, enchufes, etc.)
# Agregar potencia de equipos de climatización (no están en ambientes_detalle)
if circuitos_climatizacion:  # hay equipos de climatización cargados
    potencia_total += sum(float(eq["P_nom_w"]) for eq in circuitos_climatizacion)  # suma la potencia nominal (de placa) de cada equipo de climatización
if tension <= 0:  # la tensión quedó en 0 o negativa (dato mal ingresado)
    tension = 1.0  # evita división por cero más adelante

# Calibres comerciales de TM. Esta lista completa solo se usa para los circuitos
# ESPECIALES (horno, lavadora, etc.); iluminación y enchufes tienen su propia
# lista corta más abajo ([6,10,16] y [10,16]).
calibres_tm = [6, 10, 16, 20, 25, 32, 40, 50, 63]  # calibres comerciales de termomagnéticos disponibles, en Amperes

# Tabla de corriente admisible (ampacidad) para conductores de cobre con
# aislación de 70°C, para los circuitos interiores — Tabla N°4.4 del RIC N°4.
# Forma: {seccion_mm2: {"A1": amperes, "B1": amperes}}
#   A1 = conductor dentro de tubo embutido en muro aislante
#   B1 = conductor dentro de tubo o canaleta sobrepuesta
# Los valores son a 30°C; la corrección por temperatura se aplica después en la
# PARTE 3.2 multiplicando por factor_temperatura_ft() (RIC N°4 Tabla N°4.7).
tabla_70 = {
    1.5:  {"A1": 14, "B1": 16},  # 1,5 mm^2: sección mínima, se usa en iluminación
    2.08: {"A1": 16, "B1": 19},
    2.5:  {"A1": 18, "B1": 21},  # 2,5 mm^2: la típica de enchufes
    3.31: {"A1": 21, "B1": 25},
    4.0:  {"A1": 24, "B1": 28},  # 4 mm^2: para circuitos especiales chicos
    5.26: {"A1": 28, "B1": 34},
    6.0:  {"A1": 31, "B1": 36},  # 6 mm^2: cargas más grandes
    8.37: {"A1": 38, "B1": 45},
    10.0: {"A1": 42, "B1": 50},  # 10 mm^2: ya se usa más bien como alimentador
    13.3: {"A1": 50, "B1": 60},
    16.0: {"A1": 56, "B1": 68},
    21.1: {"A1": 66, "B1": 80},
    25.0: {"A1": 73, "B1": 89},
    26.7: {"A1": 76, "B1": 93},
    33.6: {"A1": 87, "B1": 108},
    35.0: {"A1": 89, "B1": 110},
    42.4: {"A1": 99, "B1": 125},
    50.0: {"A1": 108, "B1": 134},
    53.5: {"A1": 116, "B1": 144},
    67.4: {"A1": 133, "B1": 167},
}

# ---- Asignación del TM (termomagnético) circuito por circuito ----
# Cada vuelta agrega UN elemento a las dos listas y en el mismo orden, así que
# tm_corrientes[i] es el calibre del TM de nuevo_circuitos[i]. Ese calzado es
# lo que después le permite a optimizar_agrupacion() sumar corrientes por índice.
nuevo_circuitos = []  # copia de cada circuito, ya con "Disyuntor termomagnético" y "_In_TM"
tm_corrientes = []  # calibre del TM (en A) de cada circuito, en el mismo orden que nuevo_circuitos

# Orden de prioridad de los casos de más abajo:
#   1) climatización  -> TM ya venía calculado, no se toca
#   2) agua caliente  -> TM ya venía calculado, no se toca
#   3) enchufes de cocina / lavadero / baño -> TM fijo 16A
#   4) iluminación o enchufes normales -> calibre por corriente + 10%
#   5) todo lo demás -> calibre por corriente + 10%, con piso de 16A solo
#      si el circuito tiene tipo_dif == "especial"
for c in circuitos:
    pot = float(c.get("Potencia estimada (W)", 0.0))  # potencia del circuito, en Watts
    es_ilumin = c.get("es_ilumin", False)  # True si es un circuito de iluminación
    es_enchufe = c.get("es_enchufe", False)  # True si es un circuito de enchufes
    tipo_dif = c.get("tipo_dif", "general")  # a que tipo de diferencial va asociado este circuito
    curva = "B" if es_ilumin else "C"  # iluminación usa curva B, el resto curva C

    nombre = str(c.get("Circuito", "")).lower()  # nombre del circuito en minúsculas, para comparar texto
    items_c = c.get("_items", [])  # detalle de los ambientes/items que alimenta este circuito

    # CLIMATIZACIÓN: TM ya calculado por RIC N°07 — NO recalcular
    if c.get("_es_climatizacion", False):  # el circuito es de climatización
        in_tm_clima = int(c.get("_In_TM", 16))  # usa el TM que ya venía calculado para climatización
        i_cuadro = float(c.get("Corriente estimada (A)", 0.0))  # corriente ya calculada para este equipo
        if i_cuadro <= 0:  # el circuito venía sin corriente guardada
            i_cuadro = float(c.get("Potencia estimada (W)", 0.0)) / tension if tension else 0.0  # si no venía calculada, la estima con potencia y tensión
        nuevo = c.copy()  # copia el circuito para no modificar el original
        nuevo["Corriente estimada (A)"] = round(i_cuadro, 2)  # guarda la corriente redondeada
        nuevo["_In_TM"] = in_tm_clima  # deja guardada la corriente del TM dentro del circuito
        nuevo["_I_diseno_clima"] = c.get("_I_diseno_clima", in_tm_clima)  # corriente de diseño usada por el cálculo de climatización
        nuevo_circuitos.append(nuevo)  # agrega el circuito ya procesado a la lista final
        tm_corrientes.append(in_tm_clima)  # guarda la corriente del TM para el cálculo de diferenciales
        continue  # ya quedó listo, pasa al siguiente circuito

    # AGUA CALIENTE: el TM ya lo cálculo calcular_circuito_agua_caliente() cuando
    # se creo el circuito (carga resistiva: I = P/V, sin el 1,25 del RIC N°07
    # 7.3.4, que es para motores). Acá solo se copia — NO recalcular.
    if c.get("_es_agua_caliente", False):  # el circuito es de agua caliente
        in_tm_agua = int(c.get("_In_TM", 20))  # usa el TM que ya venía calculado para agua caliente
        i_cuadro = float(c.get("Corriente estimada (A)", 0.0))  # corriente ya calculada para este equipo
        if i_cuadro <= 0:  # el circuito venía sin corriente guardada
            i_cuadro = float(c.get("Potencia estimada (W)", 0.0)) / tension if tension else 0.0  # si no venía calculada, la estima con potencia y tensión
        nuevo = c.copy()  # copia el circuito para no modificar el original
        nuevo["Corriente estimada (A)"] = round(i_cuadro, 2)  # guarda la corriente redondeada
        nuevo["_In_TM"]        = in_tm_agua  # deja guardada la corriente del TM dentro del circuito
        nuevo["_I_diseno_agua"] = c.get("_I_diseno_agua", in_tm_agua)  # corriente de diseño usada por el cálculo de agua caliente
        nuevo_circuitos.append(nuevo)  # agrega el circuito ya procesado a la lista final
        tm_corrientes.append(in_tm_agua)  # guarda la corriente del TM para el cálculo de diferenciales
        continue  # ya quedó listo, pasa al siguiente circuito

    # detectar si el circuito incluye baño (por nombre o por ambientes en items)
    tiene_bano = ("baño" in nombre) or ("bano" in nombre) or any(
        isinstance(it, dict) and str(it.get("amb", "")).strip().lower().startswith(("baño", "bano"))  # revisa si algún item del circuito queda en un baño
        for it in (items_c if isinstance(items_c, list) else [])  # recorre los items del circuito, si es que hay lista
    )  # cierra la búsqueda de baño
    # Enchufes de cocina, lavadero o baño: TM fijo de 16A, sin mirar la potencia.
    # Es el criterio del script para los circuitos dedicados de esos recintos, y
    # es LO QUE DE VERDAD hace que los circuitos automáticos de cocina/lavadero
    # armados más arriba queden en 16A (se detecta por el nombre del circuito,
    # no por la bandera "es_coc_lav").
    # El baño se detecta por nombre o por el ambiente de alguno de sus items.
    if tipo_dif == "general" and es_enchufe and (("cocina" in nombre) or ("lavadero" in nombre) or tiene_bano):  # enchufes de cocina, lavadero o baño: el TM queda fijo en 16A
        i_est = pot / tension  # corriente estimada del circuito, en Amperes
        nuevo = c.copy()  # copia el circuito para no modificar el original
        nuevo["Corriente estimada (A)"] = round(i_est, 2)  # guarda la corriente calculada
        nuevo["Disyuntor termomagnético"] = "1x16A / 6kA / Curva C"  # cocina, lavadero y baño van en circuito dedicado con TM fijo de 16A
        nuevo["_In_TM"] = 16  # deja guardado el TM de 16A dentro del circuito
        nuevo_circuitos.append(nuevo)  # agrega el circuito ya procesado
        tm_corrientes.append(16)  # guarda la corriente del TM para el cálculo de diferenciales
        continue  # ya quedó listo, pasa al siguiente circuito

    # circuitos normales de iluminación o enchufes (los que no son especiales,
    # climatización, agua caliente, ni forzados a 16A): calcula el TM según la
    # corriente estimada más un 10% de margen
    if tipo_dif == "general" and (es_ilumin or es_enchufe):  # circuito común de iluminación o enchufes
        i_est = pot / tension  # corriente estimada del circuito, en Amperes
        # RIC N°10 art. 5.1.4.1: la capacidad del circuito de alumbrado estará
        # determinada por la potencia requerida más un 10% de capacidad adicional.
        # El valor nominal del TM será el valor nominal de corriente de la protección
        # inmediatamente superior disponible en el mercado.
        # El script aplica ese mismo 10% también a los circuitos de enchufes y a
        # los especiales (el artículo habla de alumbrado), y más abajo lo vuelve a
        # aplicar sobre la corriente de demanda para elegir el empalme.
        i_necesaria = i_est * 1.10  # corriente estimada más el 10% que pide la norma
        allowed = [10, 16] if es_enchufe else [6, 10, 16]  # calibres de TM permitidos según el tipo de circuito

        # Si ningún calibre de la lista alcanza (i_necesaria > 16 A), queda el más
        # grande: el circuito se acepta igual, aunque el TM quede justo o corto.
        # La subdivisión por P_max de la PARTE 2 debería evitar llegar a este caso.
        in_tm = allowed[-1]  # por defecto, el calibre más grande de la lista
        for cal in allowed:  # busca el calibre comercial más chico que alcance
            if cal >= i_necesaria:  # primer calibre >= i_est × 1.10
                in_tm = cal  # se queda con este calibre de TM
                break  # ya encontró el calibre, no sigue buscando

        nuevo = c.copy()  # copia el circuito para no modificar el original
        nuevo["Corriente estimada (A)"] = round(i_est, 2)  # guarda la corriente calculada
        nuevo["Disyuntor termomagnético"] = f"1x{in_tm}A / 6kA / Curva {curva}"  # arma el texto del TM elegido, para el informe
        nuevo["_In_TM"] = in_tm  # deja guardado el TM elegido dentro del circuito
        nuevo_circuitos.append(nuevo)  # agrega el circuito ya procesado
        tm_corrientes.append(in_tm)  # guarda la corriente del TM para el cálculo de diferenciales
    else:  # no es iluminación ni enchufes: se trata como circuito especial
        # circuitos especiales genéricos: mismo criterio, pero con toda la lista de calibres
        i_est = pot / tension  # corriente estimada del circuito, en Amperes
        i_necesaria = i_est * 1.10  # igual criterio: 10% de margen sobre la corriente estimada

        in_tm = calibres_tm[-1]  # por defecto, el calibre más grande disponible
        for cal in calibres_tm:  # recorre los calibres comerciales de menor a mayor
            if cal >= i_necesaria:  # primer calibre que alcanza a cubrir la corriente necesaria
                in_tm = cal  # se queda con este calibre de TM
                break  # ya encontró el calibre, no sigue buscando

        # Piso de 16A para los especiales (horno, lavadora, encimera, etc.):
        # criterio del script, para que un circuito dedicado no quede con un TM
        # de 6 o 10 A aunque la carga sea chica.
        if c.get("tipo_dif") == "especial" and in_tm < 16:  # los circuitos especiales no bajan de 16A
            in_tm = 16  # los especiales nunca van con TM menor a 16A

        nuevo = c.copy()  # copia el circuito para no modificar el original
        nuevo["Corriente estimada (A)"] = round(i_est, 2)  # guarda la corriente calculada
        nuevo["Disyuntor termomagnético"] = f"1x{in_tm}A / 6kA / Curva {curva}"  # arma el texto del TM elegido, para el informe
        nuevo["_In_TM"] = in_tm  # deja guardado el TM elegido dentro del circuito
        nuevo_circuitos.append(nuevo)  # agrega el circuito ya procesado
        tm_corrientes.append(in_tm)  # guarda la corriente del TM para el cálculo de diferenciales

circuitos = nuevo_circuitos  # desde acá, "circuitos" es la lista con TM asignado; tm_corrientes calza índice a índice

# -------- PARTE 3.1: DIFERENCIALES SEGÚN EMPALME --------
# arma los grupos de circuitos que van a compartir un mismo diferencial:
# junta los que quepan sin pasar de 3 circuitos ni de la suma máxima de
# corriente de TM, usando la menor cantidad posible de diferenciales
def optimizar_agrupacion(indices, max_sum):
    """
    Agrupa indices minimizando la cantidad de diferenciales,
    restricciones:
      - máx 3 circuitos por diferencial
      - suma In_TM por diferencial <= max_sum
    """
    items = sorted(indices, key=lambda i: tm_corrientes[i], reverse=True)  # de mayor a menor corriente
    best_bins = None  # acá se va guardando la mejor agrupación encontrada

    def score(bins):  # puntaje para comparar dos agrupaciones y ver cual conviene
        # bins = la agrupación completa: una lista de grupos.
        # cada grupo (b) es una lista con los circuitos que comparten
        # el mismo diferencial. Ej: bins = [[0,1], [2,3,4]] son 2 grupos:
        # el primero con 2 circuitos, el segundo con 3.
        #
        # entre 2 opciones de "bins", gana la que use menos diferenciales
        # (o sea, menos grupos da un len(bins) más chico).
        # el segundo termino solo cuenta los circuitos ya ubicados; lo que decide de verdad es len(bins)
        # (el "-" es solo para que "más lleno" cuente como "mejor puntaje")
        return (len(bins), -sum(len(b) for b in bins))

    def can_place(bin_list, idx):  # revisa si un circuito cabe en un grupo (diferencial)
        # bin_list = un grupo (lista de circuitos que ya están juntos en un diferencial)
        # idx = el circuito que se quiere agregar a ese grupo
        # tm_corrientes = lista con la corriente del TM de cada circuito
        #
        # revisa si el circuito "idx" cabe en este grupo, sin pasarse
        # de 3 circuitos ni de la corriente máxima del diferencial
        if len(bin_list) >= 3:  # ya hay 3 circuitos colgando de ese diferencial
            return False  # no cabe, hay que probar otro grupo
        return (sum(tm_corrientes[j] for j in bin_list) + tm_corrientes[idx]) <= max_sum  # cabe solo si la suma de los TM no pasa el tope del diferencial

    def backtrack(pos, bins):  # prueba todas las formas de repartir los circuitos en grupos
        # va probando dónde meter cada circuito (en los grupos ya armados o
        # en uno nuevo) y se queda con la mejor combinación encontrada
        nonlocal best_bins  # permite modificar, desde acá adentro, la variable definida afuera

        if best_bins is not None and score(bins) >= score(best_bins):  # compara esta combinación parcial con la mejor que ya se tenía
            return  # esta combinación ya no puede ser mejor que la que tenemos, corta acá (poda)

        if pos == len(items):  # ya ubico todos los circuitos de la lista
            best_bins = [b[:] for b in bins]  # llegó al final: guarda esta combinación como la mejor
            return  # vuelve atras a seguir probando otras combinaciones

        idx = items[pos]  # el circuito que toca ubicar ahora


        # probar primero en bins existentes (más llenos primero)
        # arma el orden en que se van a probar los grupos existentes
        order = sorted(
            range(len(bins)),  # índices de todos los grupos ya armados
            key=lambda k: (len(bins[k]), sum(tm_corrientes[j] for j in bins[k])),  # ordena por cantidad de circuitos y por suma de TM del grupo
            reverse=True  # de más lleno a más vacío
        )  # ordena los grupos ya armados, del más lleno al más vacío

        # recorre los grupos (en el orden ya calculado) buscando donde meter el circuito 'idx'
        for k in order:
            if can_place(bins[k], idx):  # revisa si el circuito cabe en este grupo
                bins[k].append(idx)          # prueba metiendo el circuito en este grupo
                backtrack(pos + 1, bins)      # sigue probando con el siguiente circuito
                bins[k].pop()                 # deshace, para probar otra combinación

        # abrir bin nuevo
        bins.append([idx])       # prueba abriendo un grupo nuevo solo para este circuito
        backtrack(pos + 1, bins)  # sigue probando con el siguiente circuito
        bins.pop()                # deshace, para no arrastrar el grupo nuevo a otras pruebas

    backtrack(0, [])  # arranca la búsqueda de la mejor combinación

    if best_bins is None:  # no se encontró ninguna combinación válida (caso raro)
        return [[i] for i in items]  # no encontró nada (no debería pasar), cada circuito en su propio grupo

    return [sorted(b) for b in best_bins]  # devuelve la mejor agrupación encontrada


# Con el interruptor termomagnético (TM) del empalme, define el calibre
# del diferencial general y otros topes que se usan más adelante para armar el tablero.
def parametros_desde_empalme(interruptor_empalme):
    """
    Deriva los parámetros del tablero a partir del interruptor termomagnético del empalme (A).
    - diferencial recomendado
    - suma máxima de ITM por diferencial (para tu agrupación)
    - potencia máxima sugerida para un componente especial (solo referencia/validación)
    - texto del omnipolar general
    """
    try:  # el interruptor del empalme puede venir como texto
        Ie = int(interruptor_empalme)  # corriente del interruptor del empalme
    except:  # el dato vino mal escrito o vacío
        Ie = 25  # si viene mal, usa un valor por defecto razonable

    # según el calibre del empalme, define el calibre del diferencial general,
    # el tope de suma de TM por diferencial y una potencia especial de referencia
    if Ie <= 25:  # empalme chico: diferencial general de 25A
        dif_calibre = 25  # diferencial general de 25A
        max_sum_tm = 25  # tope de suma de TM por diferencial: 25A
        max_pot_especial = 5000  # potencia máxima de referencia para un componente especial (W)
    elif Ie <= 32:  # empalme mediano: diferencial general de 40A, tope de suma 32A
        dif_calibre = 40  # diferencial general de 40A
        max_sum_tm = 32  # tope de suma de TM por diferencial: 32A
        max_pot_especial = 6400  # potencia máxima de referencia para un componente especial (W)
    elif Ie <= 40:  # empalme mediano-grande: diferencial 40A, tope de suma 40A
        dif_calibre = 40  # diferencial general de 40A
        max_sum_tm = 40  # tope de suma de TM por diferencial: 40A
        max_pot_especial = 8000  # potencia máxima de referencia para un componente especial (W)
    else:  # empalme más grande de lo que cubre este programa
        # Si en el futuro agregas empalmes mayores, ajusta aquí.
        dif_calibre = 63  # diferencial general de 63A
        max_sum_tm = Ie  # el tope de suma queda igual al calibre del empalme
        max_pot_especial = Ie * 220  # aproximación a 220V

    texto_omni = f"2x{Ie}A / 10kA / Curva C"  # texto del interruptor general omnipolar
    return dif_calibre, max_sum_tm, max_pot_especial, texto_omni  # entrega los 4 valores calculados a quien llamo la función


# =========================================================
# EMPALME Y PROTECCIONES (AUTO DESDE FACTOR DE DEMANDA)
# =========================================================
# Nota: aquí se calcula el empalme con la corriente con factor de demanda.
# Luego se derivan diferencial/omnipolar y se asignan a los circuitos ANTES
# de crear dataframes/materiales.

# --- Factor de demanda (RIC 3 Tabla 3.1) ---
# Criterio:
#   corriente_sin = suma de corrientes individuales del cuadro de cargas
#                   (cada circuito ya tiene su corriente calculada con su fp real)
#   I_primeros    = si alumbrado+enchufes <= 3kW, la corriente medida
#                   (corriente_sin_alum_ench); si pasa de 3kW,
#                   min(kW, 3) x 1000 / tensión. Siempre al 100%
#   I_resto       = (resto_kw   / pot_kw) x corriente_sin, multiplicado por 0,35
#   corriente_con = I_primeros x 1,0 + I_resto x 0,35
tension_nominal = float(tension)  # tensión de la instalación (V), la ingreso el usuario antes
fp = float(factor_potencia) if factor_potencia not in [None, ""] else 1.0  # factor de potencia; si no se ingreso, se asume 1.0 (carga resistiva)
if fp == 0:  # evita dividir por 0 más adelante si quedó en 0
    fp = 1.0  # 

# Potencias separadas (para tabla Excel)
# potencia_sin_clima_w debe excluir climatización y agua caliente — esos
# componentes quedan guardados dentro de "Total por ambiente (W)" con la
# potencia ESTIMADA que se ingresó al crear el ambiente (no la real de la
# placa, que se pide después al crear el circuito), así que se restan acá
# para que pot_kw_alum_ench_global (base de primeros/resto) no las arrastre.
import re  # libreria para buscar patrones de texto (expresiones regulares)
# palabras clave que identifican equipos de climatización/agua caliente
# dentro del texto libre de 'Componentes especiales' de cada ambiente
_palabras_clima_agua = ("aire", "clima", "split", "ac ", "a/c", "ducha",
                         "termo", "calefon", "calefón", "calentador", "agua caliente")  # cierra las palabras clave de clima y agua caliente
potencia_clima_agua_en_ambientes = 0.0  # acumulador: potencia (W) de clima/agua caliente ya contada en los ambientes
for _amb in ambientes_detalle:  # recorre cada ambiente ingresado por el usuario
    for _linea in str(_amb.get("Componentes especiales", "") or "").split("\n"):  # recorre cada línea de texto de los componentes especiales de ese ambiente
        _linea_l = _linea.strip().lower()  # texto en minúscula, sin espacios sobrantes, para comparar más fácil
        if any(palabra in _linea_l for palabra in _palabras_clima_agua):  # la línea menciona clima o agua caliente
            _match_w = re.search(r'\(([\d.,]+)\s*W\)', _linea, re.IGNORECASE)  # busca el número de watts que viene entre paréntesis, ej: '(1200 W)'
            if _match_w:  # encontró un valor de potencia en el texto
                try:  # por si el número viene mal escrito y falla la conversión
                    potencia_clima_agua_en_ambientes += float(_match_w.group(1).replace(",", "."))  # suma esa potencia (acepta coma decimal) al acumulador
                except ValueError:  # el texto no se pudo convertir a número
                    pass  # no hace nada, sigue con la siguiente línea
potencia_sin_clima_w = sum(a["Total por ambiente (W)"] for a in ambientes_detalle) - potencia_clima_agua_en_ambientes  # potencia de iluminación+enchufes: total de ambientes menos clima/agua caliente
potencia_clima_w = 0.0  # acumulador de potencia de climatización (W)
if circuitos_climatizacion:  # hay circuitos de climatización ingresados
    for eq_cl in circuitos_climatizacion:  # recorre cada equipo de climatización
        potencia_clima_w += float(eq_cl.get("P_nom_w", 0))  # suma la potencia nominal (real, de la placa) de cada equipo
# Potencia agua caliente (no está en ambientes_detalle)
potencia_agua_w = 0.0  # acumulador de potencia de agua caliente (W)
if circuitos_agua_caliente:  # hay circuitos de agua caliente ingresados
    for eq_ac in circuitos_agua_caliente:  # recorre cada equipo de agua caliente
        potencia_agua_w += float(eq_ac.get("P_nom_w", 0))  # suma la potencia nominal (real, de la placa) de cada equipo
potencia_total_instalacion_w = potencia_sin_clima_w + potencia_clima_w + potencia_agua_w  # potencia total de toda la instalación (alumbrado+enchufes+clima+agua)
pot_kw = potencia_total_instalacion_w / 1000.0  # lo mismo pero en kW

# kW primeros/resto — solo alumbrado+enchufes (para tabla Excel y corriente_con)
pot_kw_alum_ench_global = potencia_sin_clima_w / 1000.0  # potencia de alumbrado+enchufes en kW (sin clima ni agua caliente)
primeros_kw = min(pot_kw_alum_ench_global, 3.0)  # los primeros 3kW van al 100%
resto_kw    = max(pot_kw_alum_ench_global - 3.0, 0.0)  # el resto va al 35%
factor_primero = 1.0  # los primeros 3kW se cuentan al 100%
factor_resto   = 0.35  # el resto se cuenta al 35%
con_factor_primero = primeros_kw * factor_primero  # kW de los primeros 3kW ya con su factor aplicado
con_factor_resto   = resto_kw   * factor_resto  # kW del resto ya con su factor aplicado
total_kw = con_factor_primero + con_factor_resto  # demanda total en kW (alumbrado+enchufes) después de aplicar el factor

# --- Corriente alumbrado+enchufes SIN fd (factor de demanda) ---
# mismas palabras clave que arriba, pero para filtrar la lista de circuitos
# (cada circuito tiene su nombre en 'Circuito', ej: 'Iluminación 1', 'Aire acondicionado')
_keywords_clima_agua = ("climatiz","aire","split","ac ","a/c",
                        "ducha","termo","calefon","calefón","calentador","agua caliente")  # cierra las palabras clave de clima y agua caliente
# suma la corriente de los circuitos que NO son clima ni agua caliente
# (o sea, iluminación y enchufes)
corriente_sin_alum_ench = sum(
    float(c.get("Corriente estimada (A)", 0.0)) for c in circuitos  # toma la corriente estimada de cada circuito
    if not any(k in str(c.get("Circuito","")).lower() for k in _keywords_clima_agua)  # deja fuera climatización y agua caliente
)
# suma la corriente de los circuitos de climatización
corriente_sin_clima = sum(
    float(c.get("Corriente estimada (A)", 0.0)) for c in circuitos  # toma la corriente estimada de cada circuito
    if any(k in str(c.get("Circuito","")).lower() for k in ("climatiz","aire","split","ac ","a/c"))  # se queda solo con los circuitos de climatización
)
# suma la corriente de los circuitos de agua caliente
corriente_sin_agua = sum(
    float(c.get("Corriente estimada (A)", 0.0)) for c in circuitos  # toma la corriente estimada de cada circuito
    if any(k in str(c.get("Circuito","")).lower() for k in ("ducha","termo","calefon","calefón","calentador","agua caliente"))  # se queda solo con los circuitos de agua caliente
)
corriente_sin = corriente_sin_alum_ench + corriente_sin_clima + corriente_sin_agua  # corriente total sin aplicar ningún factor de demanda (suma simple)

# --- Corriente CON factor de demanda (RIC 3 art. 6.1, 6.2, 6.3) ---
# Alumbrado+enchufes: aplica Tabla N°3.1 (fd primeros 3kW=1,0; resto=0,35)
# Climatización y agua caliente: corriente plena (fd=1,0) — RIC 7 para clima, RIC 11 para agua caliente
pot_kw_alum_ench = pot_kw_alum_ench_global  # kW de alumbrado+enchufes (mismo valor que pot_kw_alum_ench_global)
if pot_kw_alum_ench <= 3.0:  # toda la demanda de alumbrado+enchufes entra en el tramo de 100%
    I_alum_primeros = corriente_sin_alum_ench  # usa la corriente real medida, sin recalcular desde kW
    I_alum_resto    = 0.0  # no hay tramo 'resto' porque no se pasa de 3kW
else:  # se paso de 3kW: hay que separar en tramo 100% y tramo 35%
    I_alum_primeros = (min(pot_kw_alum_ench, 3.0) * 1000.0) / tension_nominal if tension_nominal else 0.0  # corriente que corresponde a los primeros 3kW, recalculada desde la potencia
    I_alum_resto    = (max(pot_kw_alum_ench - 3.0, 0.0) * 1000.0) / tension_nominal if tension_nominal else 0.0  # corriente que corresponde al resto (sobre 3kW), recalculada desde la potencia

# corriente final que se usa para elegir el interruptor del empalme:
# alumbrado+enchufes con su factor de demanda, más clima y agua caliente a corriente plena
corriente_con = (I_alum_primeros * factor_primero  # primeros 3kW con su factor de demanda
               + I_alum_resto    * factor_resto  # resto sobre los 3kW, con su propio factor
               + corriente_sin_clima   # fd=1,0
               + corriente_sin_agua)   # fd=1,0

# --- Selección interruptor empalme (normalizado a calibre comercial) ---
# Se aplica un 10% extra de holgura sobre la corriente con factor de
# demanda antes de elegir el calibre comercial.
corriente_empalme = corriente_con * 1.1  # se agrega el 10% de holgura antes de buscar el calibre comercial

# recorre los calibres comerciales de interruptor de mayor uso y se queda
# con el primero que alcanza para la corriente calculada
if corriente_empalme <= 25:  # la demanda no pasa de 25A
    interruptor_empalme = 25  # empalme con interruptor de 25A
elif corriente_empalme <= 32:  # hasta 32A
    interruptor_empalme = 32  # empalme con interruptor de 32A
elif corriente_empalme <= 40:  # hasta 40A
    interruptor_empalme = 40  # empalme con interruptor de 40A
elif corriente_empalme <= 50:  # hasta 50A
    interruptor_empalme = 50  # empalme con interruptor de 50A
elif corriente_empalme <= 63:  # hasta 63A
    interruptor_empalme = 63  # empalme con interruptor de 63A
else:  # se paso de 63A, fuera del alcance del programa
    interruptor_empalme = 63  # tope con el que sigue el cálculo; sobre 40A igual avisa que hay que revisar el proyecto a mano

# ── Aviso límite del programa ─────────────────────────────────────────────
if interruptor_empalme > 40:  # el empalme calculado supera lo que este programa esta preparado para manejar
        # imprime en pantalla un aviso para que el usuario revise el proyecto a mano
    print(  # imprime el aviso de que el empalme se pasa del alcance del programa
        f"\n{'='*65}"
        f"\n  AVISO — Empalme calculado: {interruptor_empalme} A"
        f"\n  Este programa está diseñado para instalaciones con empalme"
        f"\n  hasta 40 A (tipo A-9 o S-9)."
        f"\n  La instalación ingresada supera ese límite ({corriente_empalme:.1f} A de demanda)."
        f"\n  Se continuará el cálculo con {interruptor_empalme} A, pero se recomienda"
        f"\n  revisar el proyecto manualmente."
        f"\n{'='*65}\n"
    )

interruptor_texto = f"1x{interruptor_empalme}A / 6kA / Curva D"  # texto del disyuntor termomagnético del empalme (el omnipolar es texto_omni)
dif_calibre, max_sum_tm, max_pot_especial, texto_omni = parametros_desde_empalme(interruptor_empalme)  # obtiene diferencial, tope de suma y potencia especial según el empalme ya calculado

# Bloque desactivado (queda como texto, no se ejecuta): validación
# informativa que compara componentes especiales contra el empalme
# calculado, solo para referencia, no detiene el programa.
"""
try:
    if isinstance(componentes_potencias, list) and len(componentes_potencias) > 0:
        pmax_esp = max([float(p) for p in componentes_potencias])
        if pmax_esp > float(max_pot_especial):
            print(
                f"\n[AVISO] Hay un componente especial de {pmax_esp:.0f} W que supera "
                f"la referencia ({max_pot_especial} W) para un empalme calculado de {interruptor_empalme}A."
            )
            print("       Revisa si corresponde circuito dedicado y/o aumentar empalme según criterio/SEC-RIC.")
except Exception:
    pass
"""

# =========================================================
# DIFERENCIALES + OMNIPOLAR (DESDE EMPALME CALCULADO)
# =========================================================
# le pregunta al usuario como quiere organizar los diferenciales:
# uno por circuito, o agrupando hasta 3 circuitos por diferencial
modo_dif = input(  # pregunta cómo se agrupan los circuitos en los diferenciales
    "\n- ¿Desea 1 diferencial por circuito (1) o agrupar hasta 3 circuitos por diferencial (3)?: "
).strip()  # deja la respuesta sin espacios

group_info = {}  # diccionario: guarda, para cada 'grupo' (diferencial), que circuitos tiene y sus datos
gid = 0  # id del grupo/diferencial que se va armando (0, 1, 2, ...)

if modo_dif == "1":  # el usuario eligió 1 diferencial por circuito
    # cada circuito con su propio diferencial exclusivo
    for idx in range(len(circuitos)):  # recorre todos los circuitos, uno por uno
        es_clima = circuitos[idx].get("_es_climatizacion", False)  # el circuito es de climatización
        es_agua  = circuitos[idx].get("_es_agua_caliente", False)  # el circuito es de agua caliente
        if es_clima:  # climatización: usa un diferencial dedicado, respetando el mínimo del tablero
            dif_val = max(circuitos[idx].get("_In_dif_clima", dif_calibre), dif_calibre)  # calibre del diferencial exclusivo para este equipo de climatización
            group_info[gid] = {"indices": [idx], "dif": dif_val, "sensibilidad_dif": "30mA"}  # guarda el grupo: un solo circuito con su diferencial y sensibilidad
        elif es_agua:  # agua caliente: usa su propio diferencial (calibre y sensibilidad según el equipo)
            dif_val = circuitos[idx].get("_In_dif_agua", dif_calibre)  # calibre del diferencial exclusivo para este equipo de agua caliente
            sens    = circuitos[idx].get("_sensibilidad_dif_agua", "30mA")  # sensibilidad del diferencial (normalmente 30mA, puede variar según el equipo)
            group_info[gid] = {"indices": [idx], "dif": dif_val, "sensibilidad_dif": sens}  # guarda el grupo: un solo circuito con su diferencial y sensibilidad
        else:  # circuito normal (iluminación/enchufes/especiales)
            group_info[gid] = {"indices": [idx], "dif": dif_calibre, "sensibilidad_dif": "30mA"}  # usa el diferencial general calculado desde el empalme
        gid += 1  # avanza al siguiente id de grupo
else:  # el usuario eligió agrupar hasta 3 circuitos por diferencial
    # Separar circuitos de climatización y agua caliente (siempre diferencial
    # exclusivo) del resto — los "especiales" (horno, lavadora, encimera,
    # etc.) ahora SÍ se agrupan igual que iluminación/enchufes.
# listas de índices de circuitos según su tipo, para procesarlos por separado
    idx_normales = [i for i in range(len(circuitos))  # circuitos que no son clima ni agua caliente: estos sí se pueden agrupar
                    if not circuitos[i].get("_es_climatizacion", False)  # descarta los circuitos de climatización
                    and not circuitos[i].get("_es_agua_caliente", False)]  # y también los de agua caliente
    idx_clima    = [i for i in range(len(circuitos)) if circuitos[i].get("_es_climatizacion", False)]  # índices de los circuitos de climatización
    idx_agua     = [i for i in range(len(circuitos)) if circuitos[i].get("_es_agua_caliente", False)]  # índices de los circuitos de agua caliente

    # Agrupar circuitos normales + especiales (usa el algoritmo de
    # backtracking para minimizar diferenciales)
    if idx_normales:  # hay circuitos normales/especiales para agrupar
        grupos_norm = optimizar_agrupacion(idx_normales, max_sum_tm)  # usa el backtracking para armar los grupos que minimizan diferenciales
        for g in grupos_norm:  # recorre cada grupo ya armado
            group_info[gid] = {"indices": g, "dif": dif_calibre, "sensibilidad_dif": "30mA"}  # guarda el grupo con el diferencial general y sensibilidad estándar
            gid += 1  # avanza al siguiente id de grupo

    # Cada circuito de climatización tiene diferencial EXCLUSIVO (RIC 7.1.2 + 7.4.5)
    for idx in idx_clima:  # recorre cada circuito de climatización, uno por uno (grupo de a 1)
        dif_val = max(circuitos[idx].get("_In_dif_clima", dif_calibre), dif_calibre)  # calibre del diferencial exclusivo, respetando el mínimo del tablero
        group_info[gid] = {"indices": [idx], "dif": dif_val, "sensibilidad_dif": "30mA"}  # guarda el grupo: un solo circuito con su diferencial
        gid += 1  # avanza al siguiente id de grupo

    # Cada circuito de agua caliente tiene diferencial EXCLUSIVO (RIC N°07 7.4.5 + RIC N°11 6.4.3)
    for idx in idx_agua:  # recorre cada circuito de agua caliente, uno por uno (grupo de a 1)
        dif_val = circuitos[idx].get("_In_dif_agua", dif_calibre)  # calibre del diferencial exclusivo para este equipo
        sens    = circuitos[idx].get("_sensibilidad_dif_agua", "30mA")  # sensibilidad del diferencial para este equipo
        group_info[gid] = {"indices": [idx], "dif": dif_val, "sensibilidad_dif": sens}  # guarda el grupo: un solo circuito con su diferencial y sensibilidad
        gid += 1  # avanza al siguiente id de grupo

# Reordenar circuitos para que queden agrupados por diferencial
# (así en el Excel salen juntos los circuitos que comparten diferencial)
idx_to_gid = {}  # diccionario: para cada índice de circuito, a que grupo (diferencial) pertenece
for g, meta in group_info.items():  # recorre cada grupo armado
    for idx in meta["indices"]:  # recorre los circuitos de ese grupo
        idx_to_gid[idx] = g  # guarda a que grupo pertenece cada circuito

new_order = sorted(range(len(circuitos)), key=lambda i: (idx_to_gid.get(i, 999), i))  # nuevo orden de circuitos: primero por grupo, y dentro del grupo por índice original
circuitos = [circuitos[i] for i in new_order]  # reordena la lista de circuitos según el nuevo orden
tm_corrientes = [tm_corrientes[i] for i in new_order]  # reordena también las corrientes de TM en el mismo orden, para que sigan calzando


# Recalcular group_info con nuevos índices (porque el orden de los circuitos cambió)
old_to_new = {old_idx: new_pos for new_pos, old_idx in enumerate(new_order)}  # mapa: índice viejo -> índice nuevo
new_group_info = {}  # aquí se arma group_info de nuevo, pero con los índices ya reordenados
new_gid = 0  # número de grupo (diferencial) que se va asignando, empezando de 0
for g, meta in group_info.items():  # recorre los grupos viejos para rearmarlos con los índices nuevos
    nuevos = sorted(old_to_new[i] for i in meta["indices"])  # traduce cada índice viejo a su nueva posición
    new_group_info[new_gid] = {  # mismo diferencial, pero apuntando a los circuitos ya reordenados
        "indices":          nuevos,  # circuitos que cuelgan de este diferencial
        "dif":              meta["dif"],  # calibre en amperes del diferencial del grupo
        "sensibilidad_dif": meta.get("sensibilidad_dif", "30mA"),  # sensibilidad de fuga, 30mA si no viene otra cosa
    }
    new_gid += 1  # pasa al siguiente número de grupo
group_info = new_group_info  # group_info queda actualizado con los índices correctos

# ── REGLA: mínimo 2 diferenciales en la instalación ──────────────────────────
# Si quedó solo 1 grupo (1 diferencial), se divide en 2 grupos
if len(group_info) == 1:  # quedó un solo diferencial y se exige mínimo 2 en la instalación
    gid_unico = list(group_info.keys())[0]  # id del único grupo que quedó
    meta_unica = group_info[gid_unico]  # datos de ese grupo: circuitos, calibre y sensibilidad
    indices = meta_unica["indices"]  # circuitos que están colgando de ese diferencial
    dif_val = meta_unica["dif"]  # calibre que se va a repetir en los dos diferenciales
    if len(indices) >= 2:  # con 2 o más circuitos sí se puede repartir en dos diferenciales
        # Dividir a la mitad
        mitad = len(indices) // 2  # punto de corte: la mitad de los circuitos a cada diferencial
        sens_val = meta_unica.get("sensibilidad_dif", "30mA")  # misma sensibilidad para los dos diferenciales
        group_info = {  # reemplaza el grupo único por dos grupos
            0: {"indices": indices[:mitad], "dif": dif_val, "sensibilidad_dif": sens_val},  # primer diferencial: la primera mitad de los circuitos
            1: {"indices": indices[mitad:], "dif": dif_val, "sensibilidad_dif": sens_val},  # segundo diferencial: la otra mitad
        }
    # Si solo hay 1 circuito total, no se puede dividir, así que queda 1 diferencial
    # (caso borde: instalación con 1 solo circuito)

# Asignar texto de diferencial y omnipolar a cada circuito (para Excel y materiales)
for g, meta in group_info.items():  # recorre cada diferencial ya definido
    val  = meta["dif"]  # calibre en amperes del diferencial
    sens = meta.get("sensibilidad_dif", "30mA")  # sensibilidad de fuga del diferencial
    dif_text = f"2X{val} {sens} / Tipo A"  # ej: "2X25 30mA / Tipo A"
    for idx in meta["indices"]:  # a cada circuito que cuelga de ese diferencial
        circuitos[idx]["Interruptor diferencial"] = dif_text  # guarda el texto del diferencial en cada circuito del grupo

# a todos los circuitos les toca el mismo interruptor general omnipolar (protección de cabecera)
for c in circuitos:  # recorre todos los circuitos
    c["Interruptor general omnipolar"] = texto_omni  # el omnipolar de cabecera es el mismo para toda la instalación

ambientes_df = pd.DataFrame(ambientes_detalle)  # tabla de ambientes lista para exportar al Excel

# -------- PARTE 3.2: CÁLCULO DE SECCIÓN Y CAÍDA DE TENSIÓN --------
# Para cada circuito hay que elegir el grosor (sección) del conductor:
# tiene que soportar la corriente y además no perder más de 3% de tensión
# en el trayecto (caída de tensión), según el RIC.
zona_lower = zona.strip().lower()  # zona en minúsculas, para comparar si es húmeda o seca
canal_lower = tipo_canalizacion.strip().lower()  # tipo de canalización en minúsculas, para poder comparar el texto

if "humed" in zona_lower:  # zona húmeda: baños, exteriores, lavaderos
    tabla_corriente = tabla_70  # zona húmeda: se usa igual la tabla de 70°C, la única cargada en el programa
    tipo_cable_default = "THWN-2"  # cable que aguanta humedad
else:  # zona seca
    tabla_corriente = tabla_70  # zona seca usa la tabla de aislación 70°C
    tipo_cable_default = "H07Z1-K"  # cable libre de halógenos para interior seco

# método de instalación para buscar en la tabla de ampacidad:
# A1 = conductor embutido, B1 = conductor sobrepuesto (canaleta)
metodo_inst = "A1" if "embut" in canal_lower else "B1"

secciones_ordenadas = sorted(tabla_corriente.keys())  # secciones comerciales disponibles, de menor a mayor
# resistividad del cobre, usada para calcular la caída de tensión
rho_cobre = 0.0179  # [ohm·mm^2/m]

# recorre cada circuito y le calcula la sección de conductor que cumple
# tanto la capacidad de corriente como la caída de tensión máxima
for c in circuitos:  # recorre circuito por circuito para dimensionar su conductor
    I_circ = float(c.get("Corriente estimada (A)", 0.0))  # corriente estimada del circuito, en amperes
    L_circuito = float(c.get("Longitud (m)", 0.0))  # largo de la canalización del circuito, en metros
    In_tm = float(c.get("_In_TM", 0.0))  # corriente nominal del TM que protege este circuito
    es_ilumin = c.get("es_ilumin", False)  # el circuito es de iluminación
    es_enchufe = c.get("es_enchufe", False)  # el circuito es de enchufes
    es_clima = c.get("_es_climatizacion", False)  # el circuito es de climatización

    # La sección/caída de tensión se calcula con la longitud REAL de la
    # canalización (L_circuito) — los chicotes son cable extra dentro de
    # las cajas, sin distancia eléctrica real, así que no deben inflar el
    # cálculo de caída de tensión. (El largo con chicotes, para saber
    # cuántos metros comprar, se recalcula aparte más adelante.)

    # Para climatización usar corriente de diseño (I_max × 1.25 ya calculada)
    # que quedó guardada en _I_diseno_clima
    if es_clima:  # climatización: la corriente de diseño (I_max x 1,25) se usa solo para verificar ampacidad; la caída de tensión sigue con la estimada
        I_circ_diseno = float(c.get("_I_diseno_clima", In_tm))  # corriente de diseño del equipo (I_max x 1.25)
        if I_circ <= 0:  # circuito sin corriente: no hay nada que calcular
            # sin corriente no hay circuito real: se deja todo vacío y se pasa al siguiente
            c["Conductor"] = ""  # sin conductor asignado
            c["Caída de tensión (%)"] = np.nan  # no hay caída de tensión que informar
            c["Canalización"] = ""  # tampoco lleva canalización
            continue  # pasa al siguiente circuito
    # Para agua caliente usar corriente de diseño (I_nom × 1.25 / ft ya calculada)
    elif c.get("_es_agua_caliente", False):  # agua caliente: también se dimensiona con corriente de diseño
        I_circ_diseno = float(c.get("_I_diseno_agua", In_tm))  # corriente de diseño del equipo de agua caliente
        if I_circ <= 0:  # equipo sin corriente: se deja el circuito vacío
            c["Conductor"] = ""  # sin conductor asignado
            c["Caída de tensión (%)"] = np.nan  # no hay caída de tensión que informar
            c["Canalización"] = ""  # tampoco lleva canalización
            continue  # pasa al siguiente circuito
    else:  # resto de los circuitos (iluminación, enchufes, especiales)
        I_circ_diseno = I_circ  # circuitos normales: usan directo su corriente estimada

    if I_circ <= 0 or L_circuito <= 0 or In_tm <= 0:  # sin corriente, sin largo o sin TM no hay cómo dimensionar
        # faltan datos básicos (corriente, largo o TM): no se puede calcular, se deja vacío
        c["Conductor"] = ""  # sin conductor asignado
        c["Caída de tensión (%)"] = np.nan  # no hay caída de tensión que informar
        c["Canalización"] = ""  # tampoco lleva canalización
        continue  # pasa al siguiente circuito

    sec_min = 1.5 if es_ilumin else 2.5  # sección mínima permitida por el RIC según el tipo de circuito

    # ===== PASO 1: SECCIÓN MÍNIMA POR CAÍDA DE TENSIÓN =====
    dv_max_volts = tension * 0.03  # 3% de la tensión nominal, el máximo permitido de caída
    # despeje de la fórmula de caída de tensión: sección mínima para no pasarse del 3%
    Smin = (2 * L_circuito * I_circ * rho_cobre) / dv_max_volts if dv_max_volts > 0 else sec_min

    # respetar mínimo normativo
    Sbase = max(sec_min, Smin)  # nunca bajar de la sección mínima que exige el RIC

    # ===== PASO 2: BUSCAR SECCIÓN COMERCIAL =====
    # busca la primera sección de la tabla que sea igual o mayor a lo que se necesita
    seccion = None  # acá va a quedar la sección elegida
    for s in secciones_ordenadas:  # recorre las secciones comerciales de menor a mayor
        if s >= Sbase:  # la primera que alcanza para lo calculado
            seccion = s  # se queda con esa sección
            break  # no sigue buscando

    if seccion is None:  # ninguna sección de la tabla alcanzó
        seccion = secciones_ordenadas[-1]  # si nada alcanza, usar la sección más grande disponible

    # ===== PASO 3: VERIFICAR AMPACIDAD Y CAÍDA =====
    # va probando la sección elegida y, si no cumple ampacidad o caída de tensión,
    # sube a la siguiente sección comercial hasta que cumpla (o se acabe la tabla)
    while True:  # repite hasta dar con una sección que cumpla las dos condiciones
        Iz = tabla_corriente[seccion][metodo_inst]  # ampacidad de tabla para esta sección y método de instalación

        # Corrección temperatura RIC N°4 art. 6.2.6 / Tabla N°4.7: Ic = Iz × ft
        ft_circ = factor_temperatura_ft(temperatura, metodo_inst)  # factor de corrección por temperatura ambiente
        Ic = Iz * ft_circ  # ampacidad real corregida (fórmula RIC correcta)

        delta_v = 2 * I_circ * rho_cobre * L_circuito / seccion  # caída de tensión en volts con esta sección
        pct = (delta_v / tension) * 100.0 if tension > 0 else 0.0  # caída de tensión en porcentaje

        cumple_ampacidad = (Ic > I_circ_diseno) and (Ic >= In_tm)  # el conductor aguanta la corriente y la TM
        cumple_caida = pct <= 3.0  # no se pasa del 3% de caída permitido

        if cumple_ampacidad and cumple_caida:  # la sección aguanta la corriente y la caída está dentro del 3%
            break  # sale del ciclo con esa sección

        idx_sec = secciones_ordenadas.index(seccion)  # posición de la sección actual dentro de las comerciales
        if idx_sec == len(secciones_ordenadas) - 1:  # ya está en la sección más grande de la tabla
            break  # ya se llegó a la sección más grande de la tabla, no hay más para probar

        seccion = secciones_ordenadas[idx_sec + 1]  # probar con la siguiente sección más grande

    c["Conductor"] = f"{tipo_cable_default} {seccion} mm^2"  # texto del conductor que va al Excel, ej: H07Z1-K 2.5 mm^2
    c["Caída de tensión (%)"] = pct  # caída de tensión final del circuito, en porcentaje
    # --- N° conductores para canalización ---
    # el diámetro del conduit/canaleta depende de cuántos conductores van adentro
    if es_ilumin:  # iluminación: los conductores dependen de los centros y comandos
        n_cond = n_conductores_iluminacion_para_circuito(c.get("_items", []), ambientes_df)  # cuenta los conductores del circuito de iluminación
    elif c.get("es_enchufe", False):  # circuito de enchufes
        # Enchufes: máximo 6 conductores en el tramo entre cajas
        # (3 que llegan + 3 que salen hacia siguiente caja)
        # Si solo hay 1 enchufe, son 3 conductores
        n_ench_circ = sum(  # cuenta cuántos enchufes tiene el circuito en total
            int(it.get("n_ench", 1) or 1)  # cada item aporta su cantidad de enchufes, mínimo 1
            for it in c.get("_items", [])  # recorre los items asignados al circuito
            if isinstance(it, dict) and ("id_ench" in it or "modulos" in it or "n_ench" in it)  # solo cuenta los items que de verdad son enchufes
        )
        n_cond = 6 if n_ench_circ > 1 else 3  # más de un enchufe: 6 conductores entre cajas; si es uno solo, 3
    else:  # climatización, agua caliente y especiales
        n_cond = 3  # climatización, agua caliente, especiales: 3 conductores (F + N + T)

    # ── Agua caliente en Vol.1: canalización SIEMPRE conduit PVC ────────────
    # RIC N°11 art. 6.4.3 tabla Vol.1: exige mínimo IPX4, y la canaleta PVC no cumple ese requisito
    # Art. 6.5.3: cable bajo tubo aislante garantizando IPX5
    # Aplica a: ducha eléctrica (siempre Vol.1) y cualquier equipo en Vol.1
    _es_agua_vol1 = (  # equipo de agua caliente que además queda dentro del Volumen 1 del baño
        c.get("_es_agua_caliente", False) and  # es un equipo de agua caliente
        c.get("_vol1_bano_agua", False)  # y además queda dentro del Volumen 1
    )
    if _es_agua_vol1 and "sobre" in tipo_canalizacion.strip().lower():  # agua caliente en Vol.1 pero la instalación es sobrepuesta
        _tipo_can_agua = "Embutida"   # forzar conduit PVC
    else:  # todos los demás casos
        _tipo_can_agua = tipo_canalizacion  # mantiene el tipo de canalización que eligió el usuario
    c["Canalización"] = canalizacion_recomendada_por_conductores(_tipo_can_agua, seccion, n_cond)  # elige el conduit o canaleta según la sección y cuántos conductores van dentro

# -------- PARTE 3.3: DATAFRAMES --------
# arma la tabla de resumen general que va en la hoja "Informe" del Excel
# cada fila es un par [nombre del parámetro, valor]; las filas que empiezan
# con "**SEP**" son títulos de sección, no datos (se detectan más adelante
# al darle formato al Excel)
parametros_generales = pd.DataFrame([  # tabla resumen con los datos generales de la instalación
    ["Zona", zona],  # zona seca o húmeda
    ["Tipo de canalización", tipo_canalizacion],  # embutida o sobrepuesta
    ["Protección de empalme (A) (calculada)", interruptor_empalme],  # protección del empalme ya calculada
    ["Tensión nominal (V)", tension],  # tensión nominal, normalmente 220V monofásico
    ["Factor de potencia", factor_potencia],  # factor de potencia usado en los cálculos
    ["Temperatura ambiente", temperatura],  # temperatura ambiente, la que define el factor de corrección
    ["Potencia total estimada (W)", potencia_total],  # suma de potencias de toda la instalación
    ["**SEP** PUESTA A TIERRA  (RIC 6, Tabla 6.4)", ""],  # título de la sección de puesta a tierra
    ["Largo barra copperweld (m)", largo_barra_pt],  # largo de la barra copperweld que va enterrada
    ["**SEP** PT N°1 (empalme - camarilla N°1)", ""],  # título: puesta a tierra del empalme
    ["Resistividad terreno PT1 (Ohm·m)", f"{desc_rho_pt1}"],  # resistividad del terreno donde va la PT1
    ["R por barra PT1 (Ohm)", round(R_1barra_pt1, 2)],  # resistencia que da una sola barra en PT1
    ["N° barras PT1", n_barras_pt1],  # cuántas barras se necesitan en PT1
    ["R final PT1 (Ohm)", f"{round(R_final_pt1, 2)} ≤ 20 Ohm  cumple" if R_final_pt1 <= 20 else f"{round(R_final_pt1, 2)} > 20 Ohm  NO cumple"],  # resistencia final de PT1 y si cumple el límite de 20 ohm
    *([ ["Separación mín. entre barras PT1 (m)", f"{sep_min_pt1:.1f}  (RIC 6, punto 8.3.2)"] ] if n_barras_pt1 > 1 else []),  # la separación entre barras solo se muestra si hay más de una
    ["**SEP** PT N°2 (tablero - camarilla N°2)", ""],  # título: puesta a tierra del tablero
    ["Resistividad terreno PT2 (Ohm·m)", f"{desc_rho_pt2}"],  # resistividad del terreno donde va la PT2
    ["R por barra PT2 (Ohm)", round(R_1barra_pt2, 2)],  # resistencia de una sola barra en PT2
    ["N° barras PT2", n_barras_pt2],  # cuántas barras se necesitan en PT2
    ["R final PT2 (Ohm)", f"{round(R_final_pt2, 2)} ≤ 20 Ohm  cumple" if R_final_pt2 <= 20 else f"{round(R_final_pt2, 2)} > 20 Ohm  NO cumple"],  # resistencia final de PT2 y si cumple los 20 ohm
    *([ ["Separación mín. entre barras PT2 (m)", f"{sep_min_pt2:.1f}  (RIC 6, punto 8.3.2)"] ] if n_barras_pt2 > 1 else []),  # igual que en PT1, solo si quedó más de una barra
], columns=["Parámetro", "Valor"])  # la tabla queda con dos columnas: Parámetro y Valor

circuitos_df = pd.DataFrame(circuitos)  # pasa la lista de circuitos a tabla de pandas, que es lo que se escribe al Excel
ambientes_df = pd.DataFrame(ambientes_detalle)  # lo mismo con los ambientes

# convertir a número las columnas que deben ser numéricas (por si quedaron como texto)
for columna in ["Longitud (m)", "Potencia estimada (W)", "Corriente estimada (A)"]:  # columnas que tienen que quedar numéricas sí o sí
    if columna in circuitos_df.columns:  # solo si la columna existe en la tabla
        circuitos_df[columna] = pd.to_numeric(circuitos_df[columna], errors="coerce")  # lo que no se pueda convertir queda como NaN

# Guardar tipo_dif antes del pop para usarlo en build_materiales_df
tipo_dif_por_circ = {str(c.get("Circuito","")).strip(): c.get("tipo_dif","general") for c in circuitos}  # guarda qué tipo de diferencial le tocó a cada circuito, buscando por nombre

# Limpiar auxiliares antes de exportar (NO borrar "Detalle asignación")
# estos campos con "_" al inicio eran solo para uso interno del cálculo,
# no deben quedar visibles en la hoja Excel de circuitos
for c in circuitos:  # recorre cada circuito
    for k in ["es_ilumin", "es_enchufe", "tipo_dif", "_In_TM", "es_coc_lav", "enchufes_coc_lav", "_items",  # campos internos del cálculo que no deben salir en el Excel
              "_In_dif_clima", "_es_climatizacion", "_I_diseno_clima",
              "_In_dif_agua", "_sensibilidad_dif_agua", "_es_agua_caliente",
              "_I_diseno_agua", "_lleva_tablero_externo_agua", "_vol1_bano_agua", "_tipo_equipo_agua"]:  # y los campos internos del agua caliente
        c.pop(k, None)  # los borra si están, y si no están no reclama

# circuitos_df_con_items: después del pop para tener Canalización
# _items se provee por items_por_nombre (el deepcopy está en la línea ~10627)
circuitos_df_con_items = pd.DataFrame(circuitos)  # copia que se pasa a materiales; ojo que ya no trae _items (se borraron arriba), esos van por items_por_nombre

# Re-hacer circuitos_df sin _items (para hoja Informe)
circuitos_df = pd.DataFrame(circuitos)  # tabla de circuitos ya limpia, para la hoja Informe
for columna in ["Longitud (m)", "Potencia estimada (W)", "Corriente estimada (A)"]:  # de nuevo pasa a número las columnas numéricas
    if columna in circuitos_df.columns:  # solo si la columna está
        circuitos_df[columna] = pd.to_numeric(circuitos_df[columna], errors="coerce")  # el texto que no sea número queda como NaN

# Reorden columnas (si existen)
# deja "Interruptor general omnipolar" justo después de "Interruptor diferencial",
# y después de esas dos pone Conductor / Caída de tensión / Canalización en orden
columnas = list(circuitos_df.columns)  # lista de columnas tal como quedaron
if "Interruptor diferencial" in columnas and "Interruptor general omnipolar" in columnas:  # solo reordena si están las dos columnas de protecciones
    for extra_col in ["Conductor", "Caída de tensión (%)", "Canalización"]:  # estas tres se sacan de donde estén
        if extra_col in columnas:  # solo si la columna existe
            columnas.remove(extra_col)  # la quita para reinsertarla después en el orden que se quiere

    columnas.remove("Interruptor general omnipolar")  # saca el omnipolar de su posición actual
    idx_ins = columnas.index("Interruptor diferencial") + 1  # posición justo después del diferencial
    columnas.insert(idx_ins, "Interruptor general omnipolar")  # y ahí mete el omnipolar

    insert_pos = idx_ins + 1  # desde ahí siguen conductor, caída y canalización
    for extra_col in ["Conductor", "Caída de tensión (%)", "Canalización"]:  # las vuelve a meter en ese orden
        if extra_col in circuitos_df.columns:  # solo las que existan en la tabla
            columnas.insert(insert_pos, extra_col)  # la inserta en la posición que corresponde
            insert_pos += 1  # corre la posición para la siguiente columna

    circuitos_df = circuitos_df[columnas]  # aplica el nuevo orden de columnas

# Eliminar cualquier columna interna residual (que empiece con _) antes de exportar
cols_visibles = [c for c in circuitos_df.columns if not str(c).startswith("_")]  # se queda solo con las columnas que no son internas
circuitos_df = circuitos_df[cols_visibles]  # tabla final de circuitos, lista para exportar

# nombre del Excel de salida, con fecha y hora para no pisar informes anteriores
nombre_archivo = f"Informe_Instalacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# -------- PARTE 4: FORMATO + AUTOAJUSTE DE ANCHO --------
with pd.ExcelWriter(nombre_archivo, engine="openpyxl") as writer:  # abre el archivo Excel para ir escribiendo las hojas
    # a partir de acá se escribe todo al Excel: primero los DataFrames (ambientes,
    # parámetros generales, circuitos, materiales), y después se les da formato
    # con openpyxl (colores, bordes, hipervínculos, anchos de columna, etc.)
    start_amb = 0  # la tabla de ambientes parte en la primera fila de la hoja
    ambientes_df.to_excel(writer, sheet_name="Informe", index=False, startrow=start_amb)  # escribe la tabla de ambientes en la hoja Informe

    start_param = start_amb + len(ambientes_df) + 3  # deja 3 filas de espacio entre la tabla de ambientes y la de parámetros
    parametros_generales.to_excel(writer, sheet_name="Informe", index=False, startrow=start_param)  # escribe más abajo la tabla de parámetros generales
    ws = writer.sheets["Informe"]  # toma la hoja Informe para darle formato con openpyxl
    # =========================================================
    # SECCIÓN FACTOR DE DEMANDA (al lado de Parámetro / Valor)
    # =========================================================
    # acá no se escribe nada todavía: solo se importan los estilos de openpyxl
    # y se dejan listos los que se reutilizan más abajo. El cuadro propiamente
    # tal (el del factor de demanda) se dibuja en el bloque siguiente, a la
    # derecha de la tabla Parámetro / Valor (que ocupa las columnas A y B).
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side  # estilos de openpyxl para pintar, bordear y alinear celdas
    thin = Side(border_style="thin", color="AAAAAA")  # borde delgado gris; se usa en la tabla de fórmulas y en las tablas de abajo
    # OJO: este verde no llega a verse en el Excel. La variable fill_header se
    # reasigna dos veces más abajo (a celeste E6F2FF, en la sección del empalme
    # y en la hoja Materiales) ANTES de usarse por primera vez, así que ningún
    # encabezado queda pintado con este C6E0B4.
    fill_header = PatternFill(start_color="C6E0B4",
                              end_color="C6E0B4",
                              fill_type="solid")  # verde claro que queda sin uso (ver nota de arriba)
    # =========================================================
    # CUADRO CÁLCULO PROTECCIÓN GENERAL
    # =========================================================
    # Qué resuelve este bloque: dibujar en la hoja Informe (columnas D a I) el
    # cuadro que justifica el empalme. Criterio: alumbrado+enchufes se parten en
    # dos tramos (primeros 3 kW al 100%, resto al 35%), y climatización y agua
    # caliente van con factor de demanda 1,0, cada equipo en su propia fila.
    # Acá se rehacen las mismas sumas del bloque global, solo para mostrarlas
    # en el cuadro; el resultado es informativo y no cambia el empalme ya
    # elegido. Se re-expresa en kW y en A para
    # mostrarla.
    # --- Potencia alumbrado + enchufes (sin clima ni agua caliente) ---
    # primeros_kw, resto_kw, factor_primero (1.0), factor_resto (0.35),
    # con_factor_primero y con_factor_resto vienen del bloque global
    # "EMPALME Y PROTECCIONES (AUTO DESDE FACTOR DE DEMANDA)", donde se
    # calculan a partir de pot_kw_alum_ench_global (potencia de la casa en kW
    # ya descontando climatización y agua caliente).


    # Corrientes alumbrado + enchufes
    # El nombre engaña: "sin" es "sin factor de demanda" y "alum" es
    # "alumbrado+enchufes". O sea, corriente_sin_alum = corriente cruda de los
    # circuitos de alumbrado y enchufes. Es el mismo número que
    # corriente_sin_alum_ench del bloque global del factor de demanda, pero acá
    # se vuelve a sumar en vez de reutilizarlo.
    corriente_sin_alum = sum(  # suma la corriente de alumbrado y enchufes, sin clima ni agua caliente
        float(c.get("Corriente estimada (A)", 0.0)) for c in circuitos  # toma la corriente ya calculada de cada circuito
        # descarta los circuitos cuyo nombre contiene alguna de estas palabras clave
        if not any(k in str(c.get("Circuito","")).lower()
                   for k in ("climatiz","aire","split","ac ","a/c",
                              "ducha","termo","calefon","calefón","calentador","agua caliente"))  # palabras clave de climatización y agua caliente
    )
    # Si toda la carga cabe en los primeros 3kW, no hay "resto":
    # toda la corriente real medida va a los primeros 3kW.
    # Si hay carga más allá de los 3kW, se convierte cada tramo
    # de potencia a corriente usando la tensión nominal.
    if pot_kw_alum_ench_global <= 3.0 or resto_kw <= 0.0:  # toda la carga cabe en los primeros 3kW?
        I_prim_sin = corriente_sin_alum  # toda la corriente medida va al tramo de los primeros 3kW
        I_rest_sin = 0.0  # no hay corriente en el tramo "resto"
    else:  # la carga pasa de 3kW: se reparte en dos tramos
        I_prim_sin = (primeros_kw * 1000.0) / tension_nominal if tension_nominal else 0.0  # kW de los primeros 3kW pasados a Amperes (P = V . I)
        I_rest_sin = (resto_kw   * 1000.0) / tension_nominal if tension_nominal else 0.0  # kW del resto pasados a Amperes
    I_prim_con = I_prim_sin * factor_primero  # aplica el factor de demanda (100%) al tramo de los primeros 3kW
    I_rest_con = I_rest_sin * factor_resto  # aplica el factor de demanda (35%) al tramo del resto

    # A partir de acá arma una fila por cada equipo de climatización y de
    # agua caliente, para mostrarlos después en la tabla del cuadro.
    # --- Filas dinámicas clima (fd=1,0) ---
    filas_clima = []  # aquí se va a guardar una fila por cada equipo de climatización
    for eq in (circuitos_climatizacion or []):  # recorre cada equipo de climatización ingresado (o ninguno si no hay)
        kw_eq = float(eq.get("P_nom_w", 0)) / 1000.0  # potencia nominal del equipo, de W a kW
        # corriente del equipo: usa la ya calculada (I_cuadro o I_diseno) o,
        # si no hay ninguna, la calcula a partir de la potencia.
        # En la práctica siempre gana I_cuadro: ingresar_equipo_climatizacion_inline()
        # deja esa clave en todos los equipos, así que los dos respaldos no se usan.
        I_eq  = float(eq.get("I_cuadro", eq.get("I_diseno", kw_eq * 1000.0 / tension_nominal if tension_nominal else 0.0)))
        nombre_eq = str(eq.get("nombre_circ", "Climatización")).split("(")[0].strip()  # nombre del equipo, sin lo que va entre paréntesis
        filas_clima.append((nombre_eq, kw_eq, I_eq, 1.0, kw_eq, I_eq))  # guarda: (nombre, kW sin fd, A sin fd, factor demanda=1.0, kW con fd, A con fd)

    # --- Filas dinámicas agua caliente (fd=1,0) ---
    filas_agua = []  # aquí se va a guardar una fila por cada equipo de agua caliente
    for eq in (circuitos_agua_caliente or []):  # recorre cada equipo de agua caliente ingresado
        kw_eq = float(eq.get("P_nom_w", 0)) / 1000.0  # potencia nominal del equipo, de W a kW
        I_eq  = kw_eq * 1000.0 / tension_nominal if tension_nominal else 0.0  # corriente del equipo (P = V . I)
        nombre_eq = str(eq.get("tipo_equipo", "Agua caliente")).capitalize()  # nombre del tipo de equipo, con mayuscula inicial
        filas_agua.append((nombre_eq, kw_eq, I_eq, 1.0, kw_eq, I_eq))  # misma estructura de 6 datos que filas_clima

    # --- Totales ---
    # cada fila de filas_clima/filas_agua trae 6 datos juntos, en este orden:
    # (nombre, kw_sin_fd, I_sin_fd, factor_demanda, kw_con_fd, I_con_fd)
    # o sea f[1]/f[2] = sin factor de demanda, f[4]/f[5] = con factor de demanda

    # kW totales SIN factor de demanda (alumbrado+enchufes + clima + agua, tal cual consumen)
    # f[1] = kW de cada equipo (clima/agua)
    total_kw_sin = pot_kw_alum_ench_global + sum(f[1] for f in filas_clima) + sum(f[1] for f in filas_agua)
    # corriente total SIN factor de demanda
    # f[2] = corriente de cada equipo (clima/agua)
    total_I_sin  = corriente_sin_alum + sum(f[2] for f in filas_clima) + sum(f[2] for f in filas_agua)
    # kW totales CON factor de demanda ya aplicado (esto es lo que realmente exige el empalme)
    # f[4] = kW con fd de cada equipo (clima/agua)
    total_kw_fd  = con_factor_primero + con_factor_resto + sum(f[4] for f in filas_clima) + sum(f[4] for f in filas_agua)
    # corriente total CON factor de demanda ya aplicado
    # f[5] = corriente con fd de cada equipo (clima/agua)
    # Reproduce el mismo criterio de corriente_con (bloque del factor de demanda),
    # que es la corriente con la que ya se eligió interruptor_empalme más arriba.
    # Diferencia: acá clima y agua se toman del equipo (I_cuadro / P_nom_w) y allá
    # se toman de la "Corriente estimada (A)" del circuito, así que el total puede
    # diferir en decimales. Este número es solo informativo, no cambia el empalme.
    total_I_con  = I_prim_con + I_rest_con + sum(f[5] for f in filas_clima) + sum(f[5] for f in filas_agua)

    # Estilos para pintar el cuadro en Excel: colores de fondo y bordes de las celdas
    fill_titulo  = PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")  # celeste para el título del cuadro
    fill_gris    = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # gris para la fila de encabezados
    fill_blanco  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # blanco para las filas de datos
    borde_gris   = Border(top=Side(style="thin", color="BFBFBF"),  # borde gris fino para todas las celdas del cuadro
                          left=Side(style="thin", color="BFBFBF"),
                          right=Side(style="thin", color="BFBFBF"),
                          bottom=Side(style="thin", color="BFBFBF"))  # el mismo gris fino abajo

    row_base   = start_param + 1  # fila donde arranca el cuadro (justo debajo de los parámetros generales)
    col_inicio = 4  # columna D

    # Dibuja el título del cuadro, combinando (merge) varias celdas en una sola
    # ---- TÍTULO ----
    ws.merge_cells(start_row=row_base, start_column=col_inicio,  # combina las celdas del título, de la columna D a la I de esa fila
                   end_row=row_base,   end_column=col_inicio + 5)  # el título abarca las 6 columnas de la tabla
    tc = ws.cell(row=row_base, column=col_inicio)  # celda donde va el título
    tc.value     = "CUADRO CÁLCULO PROTECCIÓN GENERAL"  # texto del título
    tc.font      = Font(bold=True, color="000000")  # negrita
    tc.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    tc.fill      = fill_titulo  # color de fondo celeste
    for c in range(col_inicio, col_inicio + 6):  # pinta el borde y el fondo celeste en toda la fila del título
        ws.cell(row=row_base, column=c).border = borde_gris  # borde gris en cada celda del título
        ws.cell(row=row_base, column=c).fill   = fill_titulo  # y el mismo fondo celeste en todas
    ws.row_dimensions[row_base].height = 18  # alto de la fila del título

    # ---- ENCABEZADOS ----
    # "KW" e "In [A]" salen dos veces a propósito: las columnas 2 y 3 son SIN
    # factor de demanda y las 5 y 6 son las mismas magnitudes ya multiplicadas
    # por el f/d de la columna 4.
    col_headers = ["Tablero", "KW", "In [A]", "f/d", "KW", "In [A]"]  # nombres de las columnas de la tabla
    for j, h in enumerate(col_headers):  # recorre cada encabezado con su posición
        cell = ws.cell(row=row_base+1, column=col_inicio+j)  # celda del encabezado, una fila debajo del título
        cell.value     = h  # escribe el nombre de la columna
        cell.font      = Font(bold=True, color="000000")  # encabezados en negrita
        cell.alignment = Alignment(horizontal="center", vertical="center")  # texto centrado
        cell.fill      = fill_gris  # fondo gris para diferenciarlo del título
        cell.border    = borde_gris  # mismo borde gris que el resto del cuadro

    # Arma la lista de filas que van en el cuadro: primero los 2 tramos de
    # alumbrado+enchufes (primeros 3kW y resto) y después clima/agua caliente
    # ---- FILAS DE DATOS ----
    filas_datos = [  # lista con todas las filas que van dentro del cuadro
        ("T.D.A PRIMEROS 3KW", primeros_kw, I_prim_sin, factor_primero, con_factor_primero, I_prim_con),  # tramo de los primeros 3kW: nombre, kW y A sin/con factor de demanda
        (f"T.D.A RESTO {resto_kw:.3f} KW", resto_kw, I_rest_sin, factor_resto, con_factor_resto, I_rest_con),  # tramo del resto: mismos datos que la fila anterior
    ] + filas_clima + filas_agua  # se agregan las filas de clima y agua caliente ya armadas más arriba

    for i, (nombre, kw_sin, in_sin, fd, kw_con, in_con) in enumerate(filas_datos):  # recorre cada fila de datos para dibujarla en el Excel
        r = row_base + 2 + i  # fila del Excel donde va esta fila de datos
        # redondea los valores antes de mostrarlos
        kw_con_r  = round(kw_con, 3)  # kW con factor de demanda, a 3 decimales
        in_con_r  = round(in_con, 2)  # corriente con factor de demanda, a 2 decimales
        kw_sin_r  = round(kw_sin, 3)  # kW sin factor de demanda, a 3 decimales
        in_sin_r  = round(in_sin, 2)  # corriente sin factor de demanda, a 2 decimales
        valores = [nombre, kw_sin_r, in_sin_r, fd, kw_con_r, in_con_r]  # datos en el mismo orden que las columnas de la tabla
        for j, v in enumerate(valores):  # recorre cada dato para ponerlo en su columna
            cell = ws.cell(row=r, column=col_inicio+j)  # celda de la columna j de esta fila
            cell.value     = v  # escribe el dato
            cell.font      = Font(bold=(j==0), color="000000")  # solo la primera columna (nombre del tablero) va en negrita
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado y con salto de línea si el nombre es largo
            cell.fill      = fill_blanco  # fondo blanco
            cell.border    = borde_gris  # borde gris
        ws.row_dimensions[r].height = 28  # alto de cada fila de datos

    # ---- FILA TOTAL ----
    row_total = row_base + 2 + len(filas_datos)  # fila justo después de la última fila de datos
    # redondea los totales antes de mostrarlos
    total_kw_sin_r = round(total_kw_sin, 3)  # total de kW sin factor de demanda, redondeado
    total_kw_fd_r  = round(total_kw_fd, 3)  # total de kW con factor de demanda, redondeado
    total_I_sin_r  = round(total_I_sin, 2)  # corriente total sin factor de demanda, redondeada
    total_I_con_r  = round(total_I_con, 2)  # corriente total con factor de demanda, redondeada
    total_vals = ["TOTAL", total_kw_sin_r, total_I_sin_r, "", total_kw_fd_r, total_I_con_r]  # fila TOTAL: no lleva factor de demanda (por eso el "" en esa columna)
    for j, v in enumerate(total_vals):  # recorre los datos de la fila TOTAL
        cell = ws.cell(row=row_total, column=col_inicio+j)  # celda de la fila TOTAL
        cell.value     = v  # escribe el total
        cell.font      = Font(bold=True, color="000000")  # la fila TOTAL entera va en negrita
        cell.alignment = Alignment(horizontal="center", vertical="center")  # centrado
        cell.fill      = fill_blanco  # fondo blanco
        cell.border    = borde_gris  # borde gris
    ws.row_dimensions[row_total].height = 18  # alto de la fila TOTAL


    # =========================================================
    # CÁLCULO ACOMETIDA
    # =========================================================
    # Cálculo acometida
    tipo_acometida_calc = "aereo" if tipo_acometida == "aerea" else "subterraneo"  # normaliza el texto: "aérea" -> "aéreo", cualquier otra cosa -> "subterráneo"
    res_acom = seleccionar_acometida(  # calcula la sección del conductor de la acometida según el RIC
        L_m=float(longitud_transformador_empalme),  # distancia entre el transformador y el empalme
        I_empalme_A=float(interruptor_empalme),  # corriente del interruptor del empalme
        fp=float(fp),  # factor de potencia de la instalación
        V_nom=float(tension_nominal),  # tensión nominal, normalmente 220 V
        tipo_acometida=tipo_acometida_calc,  # aéreo o subterráneo, ya normalizado
        temp_override=temperatura_suelo if "sub" in str(tipo_acometida_calc).lower() else None  # corrige por temperatura del suelo solo si la acometida es subterránea
    )
    # Texto final
    # Ojo con el formato: build_materiales_df lo vuelve a leer con
    # split("2x")[1] para recuperar la sección, así que el "2x" y el "mm^2"
    # tienen que quedar tal cual.
    acometida_txt = f"Concéntrico Cu 2x {res_acom['S']} mm^2"  # texto final, por ejemplo: "Concéntrico Cu 2x 10 mm^2"

    # =========================================================
    # CANALIZACIÓN ACOMETIDA
    # =========================================================
    tac = str(tipo_acometida).strip().lower()  # tipo de acometida en minúsculas, sin espacios
    if "sub" in tac:  # solo las acometidas subterráneas llevan tuberia (conduit)
        n_cond_acom = 2  # Concéntrico = 2 conductores (F+N)
        d_ducto_acom = ducto_nominal_tablas(res_acom["S"], n_cond_acom, "subterraneo")  # diámetro de conduit necesario según la sección y N de conductores
        # si la tabla no devuelve diámetro (d_ducto_acom None o 0) la celda del
        # Excel queda vacía, no queda un "-": el "-" es solo para la acometida aérea
        canalizacion_acom_txt = f"Ø PVC Conduit {d_ducto_acom} mm" if d_ducto_acom else ""  # texto final del ducto, ej. "Ø PVC Conduit 25 mm"
    else:  # acometida aérea
        canalizacion_acom_txt = "-"  # acometida aérea: no necesita conduit

    #==========================================================
    #CÁLCULO DE CAÍDA DE TENSIÓN ACOMETIDA
    #==========================================================
    # El título dice "cálculo" pero acá no se calcula nada: la caída ya viene
    # resuelta dentro de seleccionar_acometida() (clave dV_pct). Esto solo la
    # formatea como texto para la celda del Excel.
    caida_acometida_txt = f"{res_acom['dV_pct']:.2f}%"  # caída de tensión de la acometida, en porcentaje

    # =========================================================
    # EMPALME NORMALIZADO
    # =========================================================
    # Determinar número del empalme
    # interruptor_empalme solo puede valer 25, 32, 40, 50 o 63 A (se normalizó
    # así en el bloque "Selección interruptor empalme"), o sea que de estos tres
    # tramos salen únicamente: 25 -> 6, 32 y 40 -> 9, 50 y 63 -> 16.
    if interruptor_empalme <= 30:  # empalmes chicos (hasta 30A) son tipo 6
        emp_num = 6  # empalme N°6
    elif interruptor_empalme <= 40:  # empalmes medianos (31 a 40A) son tipo 9
        emp_num = 9  # empalme N°9
    else:  # empalmes grandes (más de 40A) son tipo 16
        emp_num = 16  # empalme N°16
    # Letra según acometida
    if tipo_acometida == "aerea":  # la letra del empalme depende de si la acometida es aérea o subterránea
        empalme_txt = f"A-{emp_num}"  # "A" de aérea
    else:  # acometida subterránea
        empalme_txt = f"S-{emp_num}"  # "S" de subterránea

    # =========================================================
    # POTENCIA NOMINAL SEGÚN EMPALME
    # =========================================================
    # Tabla fija escrita a mano en el código: corriente normalizada del empalme
    # (A) -> potencia nominal que se declara para ese empalme (kW). El código no
    # cita artículo del RIC, así que no se le atribuye uno acá.
    # Como interruptor_empalme solo puede ser 25/32/40/50/63, las únicas
    # potencias que llegan al Excel son 5, 6.5, 8, 10 y 13 kW; el resto de las
    # entradas de la tabla nunca se usa.
    tabla_potencia_empalme = {6: 1, 10: 2, 16: 3, 20: 4, 25: 5, 30: 6, 32: 6.5, 35: 7, 40: 8, 50: 10, 63: 13}
    potencia_nominal_empalme = tabla_potencia_empalme.get(interruptor_empalme, "")  # potencia (kW) del empalme elegido; "" si el calibre no está en la tabla

    # =========================================================
    # CÁLCULO ALIMENTADOR
    # =========================================================
    # Caídas de tensión de cada circuito (para verificar ΔV total <= 5%).
    # Ojo: seleccionar_alimentador() NO suma esta lista; se queda con la PEOR
    # (el máximo) y exige ΔV_alimentador + ΔV_peor_circuito <= 5%.
    _dv_circs = []  # aquí se guardan las caídas de tensión válidas de cada circuito
    for _c in circuitos:  # recorre los circuitos ya calculados
        _dv = _c.get("Caída de tensión (%)", None)  # caída de tensión de este circuito, si esta calculada
        try:  # la caída puede venir vacía o como texto
            _dv_f = float(_dv)  # intenta pasarla a número, si no se puede se ignora más abajo
            if _dv_f == _dv_f:  # filtrar NaN
                _dv_circs.append(_dv_f)  # guarda solo las caídas válidas
        except (TypeError, ValueError):  # ignora circuitos sin caída de tensión calculada
            pass  # no se guarda nada y sigue con el próximo

    res_alim = seleccionar_alimentador(  # calcula la sección del conductor del alimentador según el RIC
        L_m=float(longitud_alimentador),  # longitud del alimentador (tablero general a TDA)
        I_demanda_A=float(corriente_con),  # corriente total de demanda (con factor de demanda aplicado)
        I_empalme_A=float(interruptor_empalme),  # corriente del empalme, el alimentador no puede quedar por debajo
        fp=float(fp),  # factor de potencia
        V_nom=float(tension_nominal),  # tensión nominal
        tipo_alimentador=tipo_alimentador,  # aéreo, subterráneo o embutido, cambia el método de instalación
        dv_circuitos=_dv_circs,  # lista de caídas de los circuitos; adentro se usa solo la mayor, para no pasar del 5% total
        temp_override=temperatura_suelo if "sub" in str(tipo_alimentador).lower() else None  # corrige por temperatura del suelo si el alimentador es subterráneo
    )
    alim_txt = f"RV-K Cu 3x{res_alim['S']} mm^2"  # texto final: cable y sección del alimentador
    # res_alim['método'] es el método de INSTALACIÓN de la tabla de ampacidad
    # (E aéreo, D1 subterráneo, B1 en ducto), no un "método de cálculo".
    # canal_txt queda sin uso: más abajo, en la fila de datos del alimentador,
    # se vuelve a armar el mismo f-string en vez de leer esta variable.
    canal_txt = f"{tipo_alimentador} (Método {res_alim['metodo']})"  # variable muerta: se arma pero nunca se escribe en el Excel
    caida_txt = f"{res_alim['dV_pct']:.2f}%"  # caída de tensión del alimentador, en porcentaje (este SÍ se usa, en la columna 5)
    # Canalización real (Ø ducto) según tablas
    n_cond_alim = 3  # F + N + PE
    d_ducto = ducto_nominal_tablas(res_alim["S"], n_cond_alim, tipo_alimentador)  # diámetro de conduit necesario para el alimentador
    if d_ducto is None:  # si no hay diámetro es porque el tramo es aéreo
        canalizacion_txt = "-"              # si es aéreo, no lleva canalización
    else:  # tramo con ducto: se arma el texto
        canalizacion_txt = f"Ø PVC Conduit {d_ducto} mm"  # texto final del ducto del alimentador, ej. "Ø PVC Conduit 25 mm"

    # >>> arma la lista de materiales. los _items ya se borraron más arriba,
    # >>> por eso se pasan aparte con items_por_nombre
    # acá se llama a la función más grande del programa: recorre todos los
    # circuitos ya calculados y arma la lista completa de materiales a comprar
    materiales_df = build_materiales_df(circuitos_df_con_items, texto_omni, ambientes_df, group_info, tipo_canalizacion,  # arma el DataFrame de materiales que después va a su propia hoja
                                    tipo_alimentador, tipo_acometida, tipo_instalacion_empalme, requiere_mastil,  # tipos de canalización y si la instalación lleva mástil
                                    acometida_txt, interruptor_texto, longitud_transformador_empalme, canalizacion_txt,  # textos ya armados de acometida, interruptor y canalización
                                    dist_empalme_pt1, dist_tda_pt2, longitud_subterraneo_medidor, longitud_mastil,  # metros a las 2 camarillas de tierra, tramo subterráneo al medidor y largo del mástil
                                    altura_acometida_aerea, longitud_subterraneo_medidor2, longitud_llegada_aerea_tda,  # subida aérea por el poste, 2do tramo subterráneo y llegada aérea al tablero
                                    longitud_alimentador, longitud_abrazaderas_alimentador, longitud_poste_alimentador_aereo,  # metros de alimentador, tramo del que salen las abrazaderas y subida por el poste
                                    dist_vertical_acometida, alim_txt,  # tramo vertical de la acometida y texto del alimentador
                                    circuitos_climatizacion=circuitos_climatizacion,  # equipos de climatización, que llevan materiales aparte
                                    items_por_nombre=items_por_nombre,  # artefactos contados por circuito
                                    cajas_adic_por_nombre=cajas_adic_por_nombre,  # cajas de derivación extra que pidió el usuario
                                    tipo_dif_por_circ=tipo_dif_por_circ,  # tipo de diferencial de cada circuito
                                    n_barras_pt1=n_barras_pt1, n_barras_pt2=n_barras_pt2,  # cuántas barras copperweld lleva cada puesta a tierra
                                    long_cond_desnudo_pt1=long_cond_desnudo_pt1,  # metros de cobre desnudo que unen las barras de la PT1
                                    long_cond_desnudo_pt2=long_cond_desnudo_pt2)  # lo mismo para la segunda puesta a tierra

    # =========================================================
    # TABLA DE FÓRMULAS (al lado derecho del bloque factor demanda)
    # =========================================================
    col_form = col_inicio + 7   # deja espacio después de la tabla de 6 columnas (col_inicio a col_inicio+5) + 1 libre
    # con col_inicio = 4 (columna D), la tabla de fórmulas parte en la columna 11 (K)
    # y ocupa K, L y M

    # Colores para la tabla de fórmulas: gris para encabezados, celeste para el título
    fill_form_hdr  = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # gris de los encabezados de la tabla de fórmulas
    fill_form_title= PatternFill(start_color="9DC3E6", end_color="9DC3E6", fill_type="solid")  # celeste del título de la tabla de fórmulas

    # Tabla de referencia (no se comenta línea por línea porque es repetitiva):
    # cada elemento es (título de la sección, fórmula, descripción de las variables).
    # Esta tabla se imprime tal cual en la hoja de Excel, al lado del cuadro de
    # cálculo de la protección general, para que quede la referencia normativa.
    formulas_data = [  # cada tupla de acá abajo es una fila de la tabla de fórmulas
        # (título de sección, fórmula, descripción corta)
        ("CORRIENTE DE CIRCUITO",
         "I = P / (V · fp)",
         "I: corriente [A]  |  P: potencia [W]  |  V: tensión nominal [V]  |  fp: factor de potencia (1.0 para cargas resistivas)"),
        ("TERMOMAGNÉTICO CIRCUITO",
         "In_TM ≥ I_circ",
         "In_TM: calibre termomagnético [A]  |  I_circ: corriente de circuito [A]  |  Calibres: 6, 10, 16 A (ilum.) / 10, 16 A (ench.)  |  RIC 5, Tabla 5.1"),
        ("TERMOMAGNÉTICO ESPECIAL",
         "In_TM ≥ 1.10 × I_circ  (mín. 16 A)",
         "In_TM: calibre termomagnético [A]  |  I_circ: corriente de circuito [A]  |  10% de reserva  |  Mínimo 16 A para agua caliente y climatización"),
        ("SECCIÓN MÍNIMA POR ΔV",
         "Smin = (2 · L · I · ρ) / ΔVmax",
         "Smin: sección mínima [mm^2]  |  L: longitud tramo [m]  |  I: corriente [A]  |  ρ = 0.0179 Ω·mm^2/m (Cu)  |  ΔVmax = V × 3% [V]  |  Mínimo: 1.5 mm^2 (ilum.) / 2.5 mm^2 (ench.)"),
        ("CAÍDA DE TENSIÓN",
         "ΔV% = (2 · ρ · L · I) / (S · V) × 100",
         "ΔV%: caída de tensión [%]  |  ρ = 0.0179 Ω·mm^2/m (Cu)  |  L: longitud tramo [m]  |  I: corriente [A]  |  S: sección conductor [mm^2]  |  V: tensión nominal [V]  |  Límite: ≤ 3%"),
        ("AMPACIDAD CONDUCTOR",
         "Iz ≥ I_circ  y  Iz ≥ In_TM",
         "Iz: ampacidad corregida [A]  |  Iz = Iz_tabla × ft  |  Iz_tabla: ampacidad tabulada [A]  |  ft: factor corrección temperatura [-]  |  I_circ: corriente circuito [A]  |  In_TM: calibre TM [A]  |  RIC 5, Tabla 5.4"),
        ("FACTOR DE DEMANDA",
         "Pd = P₁ × 1.0 + P₂ × 0.35",
         "Pd: potencia de demanda [W]  |  P₁ = min(Ptotal, 3000) [W]: primeros 3 kW al 100%  |  P₂ = max(Ptotal − 3000, 0) [W]: resto al 35%  |  RIC 3, Art. 6.1, 6.2, 6.3"),
        ("CORRIENTE EMPALME",
         "I_emp = Pd / (V × fp)",
         "I_emp: corriente de empalme [A]  |  Pd: potencia de demanda [W]  |  V: tensión nominal [V]  |  fp: factor de potencia [-]  |  In normalizado: 25 / 32 / 40 / 50 / 63 A"),
        ("ALIMENTADOR (Smin)",
         "Smin = (2 · L · I_dem · fp · ρ) / ΔVmax",
         "Smin: sección mínima [mm^2]  |  L: longitud alimentador [m]  |  I_dem: corriente de demanda [A]  |  fp: factor de potencia [-]  |  ρ = 0.0179 Ω·mm^2/m  |  ΔVmax = V × 3% [V]  |  Cumple: Iz≥I_dem, Iz≥I_emp y ΔV≤3%  |  Mínimo 4 mm^2"),
        ("ACOMETIDA (Smin)",
         "Smin = (2 · L · I_emp · fp · ρ) / ΔVmax",
         "Smin: sección mínima [mm^2]  |  L: longitud acometida [m]  |  I_emp: corriente de empalme [A]  |  fp: factor de potencia [-]  |  ρ = 0.0179 Ω·mm^2/m  |  ΔVmax = V × 3% [V]  |  Método E (aéreo) o D1 (subterráneo)  |  Mínimo 4 mm^2"),
        ("PUESTA A TIERRA (RIC 6)", "", ""),
        ("RESIST. UNA PICA VERTICAL",
         "R₁ = ρ / L",
         "R₁: resistencia una barra [Ohm]  |  ρ: resistividad del terreno [Ohm·m]  |  L: largo de la barra [m]  |  RIC 6, Tabla 6.4"),
        ("N° BARRAS NECESARIAS",
         "N = ⌈R₁ / 20⌉  (mín. 1)",
         "N: número de barras [-]  |  R₁: resistencia una barra [Ohm]  |  20: resistencia máxima permitida [Ohm]  |  Barras en paralelo  |  RIC 6, punto 6.1"),
        ("RESIST. FINAL (N barras)",
         "R_final = R₁ / N",
         "R_final: resistencia sistema [Ohm]  |  R₁: resistencia una barra [Ohm]  |  N: número de barras [-]  |  Separación mínima entre barras = 2 × L  |  RIC 6, punto 8.3.2"),
        ("CONDUCTOR DESNUDO (unión barras)",
         "L_desnudo = (N − 1) × (2 · L_barra)",
         "L_desnudo: longitud conductor desnudo [m]  |  N: número de barras [-]  |  L_barra: largo de cada barra [m]  |  Material: Cu desnudo 16 mm^2  |  RIC 6"),
    ]

    # Título principal de la tabla
    # se fusionan 3 celdas (col_form, col_form+1, col_form+2) en la fila row_base
    # para poner ahí el título "FÓRMULAS UTILIZADAS EN LOS CÁLCULOS"
    ws.merge_cells(start_row=row_base, start_column=col_form,  # fusiona 3 celdas para el título de la tabla de fórmulas
                   end_row=row_base, end_column=col_form + 2)  # el título ocupa las 3 columnas de la tabla
    tc = ws.cell(row=row_base, column=col_form)  # celda donde queda el título (arriba a la izquierda de la fusión)
    tc.value = "FÓRMULAS UTILIZADAS EN LOS CÁLCULOS"  # texto del título
    tc.font = Font(bold=True, color="000000")  # negrita, letra negra
    tc.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    tc.fill = fill_form_title  # color de fondo del título
    tc.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino en los 4 lados
    for _c in range(col_form + 1, col_form + 3):  # recorre las otras 2 celdas de la fusión
        # las otras 2 celdas fusionadas también necesitan el mismo fondo y borde
        _cell = ws.cell(row=row_base, column=_c)  # celda a rellenar
        _cell.fill = fill_form_title  # mismo celeste del título
        _cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # mismo borde fino

    # Encabezados de columna
    # escribe la fila justo debajo del título, con el nombre de cada columna
    for _j, _hdr in enumerate(["Parámetro / Etapa", "Fórmula", "Notas"]):  # las 3 columnas de la tabla de fórmulas
        _cell = ws.cell(row=row_base + 1, column=col_form + _j)  # fila de encabezados, justo debajo del título
        _cell.value = _hdr  # nombre de la columna
        _cell.font = Font(bold=True)  # encabezado en negrita
        _cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado y con ajuste de texto
        _cell.fill = fill_form_hdr  # fondo gris
        _cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino

    # Filas de datos (una fila más ancha y en negrita para el separador "PUESTA A TIERRA")
    # recorre formulas_data (la tabla con etapa/fórmula/notas definida más arriba)
    # y escribe cada fila debajo de los encabezados
    fill_form_sep = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")  # azul clarito para la fila separadora de puesta a tierra
    for _i, (etapa, formula, notas) in enumerate(formulas_data):  # recorre fórmula por fórmula
        _row = row_base + 2 + _i  # fila del Excel donde va esta fórmula
        es_separador = etapa.startswith("PUESTA A TIERRA")  # fila especial que separa las fórmulas eléctricas de las de puesta a tierra
        for _j, _val in enumerate([etapa, formula, notas]):  # escribe las 3 columnas de la fila
            _cell = ws.cell(row=_row, column=col_form + _j)  # celda de la columna que toca
            _cell.value = _val  # escribe el texto
            _cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
            if es_separador:  # fila separadora de puesta a tierra
                # fila separadora: negrita, centrada, con fondo distinto
                _cell.font = Font(bold=True, color="000000")  # negrita en negro
                _cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrada, con salto de línea si no cabe
                _cell.fill = fill_form_sep  # fondo azul clarito
            else:  # fila de fórmula normal
                _cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)  # fila normal: la etapa y las notas van alineadas a la izquierda
                if _j == 1:   # columna fórmula: Calibri, sin negrita, centrada
                    _cell.font = Font(bold=False, name="Calibri")  # la fórmula va sin negrita
                    _cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # y centrada en su columna
        if es_separador:  # la separadora además se fusiona
            # la fila separadora ocupa las 3 columnas fusionadas, como el título
            ws.merge_cells(start_row=_row, start_column=col_form,
                           end_row=_row, end_column=col_form + 2)  # hasta la tercera columna del cuadro

    from openpyxl.utils import get_column_letter as _gcl  # para pasar número de columna a letra (4 -> D)

    # ahora escribe la tabla de circuitos, debajo de todo lo anterior.
    # El +8 es un margen fijo escrito a mano: separa la tabla de circuitos del
    # final de la tabla Parámetro / Valor (columnas A y B). El cuadro de factor
    # de demanda y la tabla de fórmulas viven a la derecha (columnas D en
    # adelante), así que no chocan con ésta aunque bajen más filas.
    start_circ = start_param + len(parametros_generales) + 8  # fila donde parte la tabla de circuitos, dejando aire después de los parámetros
    circuitos_df.to_excel(writer, sheet_name="Informe", index=False, startrow=start_circ)  # vuelca el DataFrame de circuitos a la hoja Informe

    # Vuelve a pedir libro y hoja después del to_excel. La hoja "Informe" ya
    # existía desde el principio del bloque (se tomó en la primera línea con
    # writer.sheets["Informe"]); esto es solo por seguridad, para seguir
    # trabajando sobre el objeto vigente después de que pandas escribió.
    wb = writer.book  # libro completo; queda asignado pero no se usa en ninguna línea posterior
    ws = writer.sheets["Informe"]  # misma hoja Informe de antes, re-tomada tras el volcado de circuitos

    # estos imports ya se hicieron al inicio del bloque; se repiten sin efecto
    # (Python los resuelve igual). get_column_letter es el mismo _gcl de arriba.
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side  # estilos de celda de openpyxl
    from openpyxl.utils import get_column_letter  # para convertir número de columna a letra

    def aplicar_estilo(fila_inicio, nrows, ncols, filtro=False):  # da formato a una tabla ya volcada en la hoja
        # le da formato a una tabla de la hoja "Informe": encabezado en negrita,
        # bordes, ancho de columna automático y filtro si se pide
        # qué recibe:
        #   fila_inicio .. fila de Excel donde arranca la tabla (el encabezado va abajo)
        #   nrows ........ cuántas filas de datos tiene la tabla
        #   ncols ........ cuántas columnas tiene
        #   filtro ....... True si se quiere el filtro de Excel en el encabezado
        # no devuelve nada, solo pinta la hoja.
        hdr = fila_inicio + 1  # fila del encabezado: to_excel() escribe el encabezado en startrow, y openpyxl cuenta desde 1
        fila_final = hdr + nrows  # última fila de datos de la tabla
        thin = Side(border_style="thin", color="AAAAAA")  # estilo de borde fino gris

        # formatea la fila de encabezado: negrita, centrado, fondo celeste, con borde
        for ccol in range(1, ncols + 1):
            cell = ws.cell(row=hdr, column=ccol)  # celda del encabezado de esta columna
            cell.font = Font(bold=True)  # encabezado en negrita
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado y con salto de línea si el título es largo
            cell.fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")  # fondo celeste claro
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino en los 4 lados

        # recorre columna por columna: le pone borde a los datos y calcula el ancho
        for ccol in range(1, ncols + 1):
            col_letter = get_column_letter(ccol)  # letra de la columna, se usa para fijarle el ancho
            max_len = 0  # va guardando el texto más largo de la columna

            for r in range(hdr + 1, fila_final + 1):  # recorre las filas de datos de esta columna
                cell = ws.cell(row=r, column=ccol)  # celda de dato
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # le pone el mismo borde fino
                cell.alignment = Alignment(vertical="top", wrap_text=True)  # texto alineado arriba y con salto de línea

                val = cell.value  # contenido de la celda
                if val is not None:  # si tiene algo, se mide para calcular el ancho
                    max_len = max(max_len, max(len(str(x)) for x in str(val).split("\n")))  # por si el texto tiene saltos de línea

            val_hdr = ws.cell(row=hdr, column=ccol).value  # texto del encabezado de la columna
            if val_hdr is not None:  # si el encabezado tiene texto
                max_len = max(max_len, len(str(val_hdr)))  # el título también cuenta para el ancho

            ws.column_dimensions[col_letter].width = min(max(12, max_len + 4), 80)  # ancho mínimo 12, máximo 80

        if filtro:  # solo si se pidió filtro
            ws.auto_filter.ref = f"A{hdr}:{get_column_letter(ncols)}{fila_final}"  # agrega el autofiltro de Excel

    # aplica el formato de tabla a cada bloque ya escrito en la hoja: ambientes,
    # parámetros generales y (más abajo) circuitos
    aplicar_estilo(start_amb, len(ambientes_df), ambientes_df.shape[1], filtro=True)  # tabla de ambientes, esta si lleva filtro
    aplicar_estilo(start_param, len(parametros_generales), parametros_generales.shape[1])  # tabla de parámetros generales
    fila_ini = start_param + 2  # primera fila de datos de parámetros generales
    fila_fin = start_param + 1 + len(parametros_generales)  # última fila de datos de parámetros generales
    # recorre la tabla de parámetros generales buscando filas marcadas con
    # "**SEP**" (así se armaron en parametros_generales, más arriba): esas
    # filas no son un dato más, son un título de subsección, así que se
    # fusionan en 1 sola celda centrada en vez de quedar como fila normal
    for r in range(fila_ini, fila_fin + 1):  # recorre fila por fila la tabla de parámetros
        val_param = ws.cell(row=r, column=1).value  # nombre del parámetro que hay en esa fila
        if val_param and str(val_param).startswith("**SEP**"):  # fila marcada como separadora de subsección
            # es una fila separadora: le saca el prefijo "**SEP**" y la fusiona
            texto_sep = str(val_param).replace("**SEP** ", "")  # deja solo el título de la subsección, sin el prefijo
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)  # junta la columna del nombre con la del valor
            cell_sep = ws.cell(row=r, column=1)  # celda resultante de la fusión
            cell_sep.value = texto_sep  # escribe el título ya limpio
            cell_sep.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)  # centrado a lo ancho de las 2 columnas
            cell_sep.font = Font(bold=True)  # título de subsección en negrita
            thin = Side(border_style="thin", color="AAAAAA")  # borde fino gris
            cell_sep.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde en los 4 lados
            ws.row_dimensions[r].height = 15  # alto fijo para que no se estire la fila
        else:  # fila normal de parámetro
            # fila normal de parámetro: alinea el nombre a la izquierda y el valor centrado
            ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)  # columna A: nombre del parámetro, a la izquierda
            ws.cell(row=r, column=2).alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)  # columna B: valor, centrado
            ws.row_dimensions[r].height = 15  # mismo alto que las filas separadoras

    aplicar_estilo(start_circ, len(circuitos_df), circuitos_df.shape[1])  # tabla de circuitos, sin filtro
    # Anchos fijos DESPUÉS de todos los aplicar_estilo para que no sean pisados
    ws.column_dimensions["A"].width = 42  # columna A ancha: nombres de circuito y de parámetro
    # CUIDADO: este 55 no es el ancho final. Más abajo, en la sección del
    # alimentador, la columna B se vuelve a fijar en 32 (y la C en 35), así que
    # en el Excel la B termina con 32.
    ws.column_dimensions["B"].width = 55  # ancho provisorio de la columna B (después queda en 32)
    # Anchos tabla de fórmulas (forzar AQUÍ para que no sean pisados por aplicar_estilo)
    ws.column_dimensions[_gcl(col_form)].width     = 32  # columna Parámetro / Etapa
    ws.column_dimensions[_gcl(col_form + 1)].width = 35  # columna Fórmula
    ws.column_dimensions[_gcl(col_form + 2)].width = 65  # columna Notas, la más larga de las tres
    # Columnas E-I
    for _col_letra in ["E", "F", "G", "H", "I"]:  # ancho de E a I; la D del cuadro queda con el ancho que le cálculo aplicar_estilo
        ws.column_dimensions[_col_letra].width = 23  # mismo ancho para las cinco
    # Altura filas tabla de fórmulas (individual)
    # cada número es el alto (en puntos) de una fila de formulas_data, en orden.
    # Son 15 valores para las 15 filas de formulas_data (medidos a ojo, según
    # cuánto texto lleva la columna "Notas"); el 15 corresponde a la fila
    # separadora "PUESTA A TIERRA (RIC 6)", que va vacía.
    _form_heights = [33, 33, 33, 53, 53, 53, 50, 53, 63, 63, 15, 33, 48, 50, 33]
    for _i, _h in enumerate(_form_heights):  # aplica cada alto a la fila que le toca
        ws.row_dimensions[row_base + 2 + _i].height = _h  # las filas de fórmulas parten 2 abajo del título

    # Formato porcentaje numérico para "Caída de tensión (%)"
    # así en Excel el número se ve como "2.35%" en vez de "2.35"
    if "Caída de tensión (%)" in circuitos_df.columns:
        col_caida = list(circuitos_df.columns).index("Caída de tensión (%)") + 1  # posición de la columna de caída de tensión dentro de la tabla
        fila_ini = start_circ + 2  # primera fila con datos de circuitos
        fila_fin = start_circ + 1 + len(circuitos_df)  # última fila con datos de circuitos
        for r in range(fila_ini, fila_fin + 1):  # recorre las filas de circuitos para darles formato
            ws.cell(row=r, column=col_caida).number_format = '0.00"%"'  # muestra la caída con 2 decimales y el signo %

  # =========================================================
    # MERGE de celdas en "Informe" (SIN tocar Materiales)
    # - Omnipolar: una sola celda para todos los circuitos
    # - Diferencial: merge por grupos (mismo texto consecutivo)
    # =========================================================

    def _find_col_idx_by_name(df, nombre_columna):  # busca en qué columna quedó un campo del informe
        """Devuelve índice 1-based de la columna 'nombre_columna' en el DataFrame, o None si no existe."""
        try:  # si la columna no existe, index() revienta
            return list(df.columns).index(nombre_columna) + 1  # +1 porque Excel empieza en columna 1, no 0
        except:  # atrapa el error y sigue sin caerse
            return None  # esa columna no existe en el DataFrame

    # OJO: en este archivo esta función se llama una sola vez y siempre con
    # merge_all=True (columna del omnipolar). La rama "por tramos" queda escrita
    # pero no se ejecuta; los diferenciales se fusionan con
    # merge_diferenciales_por_grupo(), que va por group_info y no por texto igual.
    def merge_vertical_runs(ws, col_idx, first_data_row, last_data_row, merge_all=False):  # fusiona celdas de una columna, todo junto o por tramos iguales
        """
        Combina verticalmente celdas en una columna:
        - Si merge_all=True: merge desde first_data_row a last_data_row.
        - Si merge_all=False: merge por 'runs' consecutivos con el mismo valor (no vacío).

        Qué recibe:
          ws .............. la hoja de Excel donde se va a hacer el merge
          col_idx ......... número de la columna a fusionar (si viene None, no hace nada)
          first_data_row .. primera fila de datos
          last_data_row ... última fila de datos
          merge_all ....... True para fusionar todo el rango de una sola vez
        No devuelve nada, solo fusiona las celdas.
        """
        if col_idx is None:  # sin columna no hay nada que fusionar
            return  # no hay columna, no hace nada

        col_letter = get_column_letter(col_idx)  # letra de la columna, para armar el rango tipo "F5:F9"

        if merge_all:  # caso omnipolar: es el mismo para toda la instalación
            # modo simple: fusiona todo el rango de una sola vez
            if last_data_row > first_data_row:  # solo tiene sentido si hay más de una fila
                ws.merge_cells(f"{col_letter}{first_data_row}:{col_letter}{last_data_row}")  # fusiona el rango completo de una
                c = ws.cell(row=first_data_row, column=col_idx)  # la celda de arriba es la que queda visible
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # texto centrado y ajustado dentro de la celda
            return  # listo, no sigue con el modo por tramos

        # modo "por tramos": va agrupando filas seguidas que tengan el mismo valor
        fila_tramo_inicio = first_data_row  # primera fila del tramo actual
        prev_val = ws.cell(row=first_data_row, column=col_idx).value  # valor del tramo actual

        def _is_empty(v):  # chequeo chico que se usa más abajo
            # revisa si la celda está vacía (None o solo espacios)
            return v is None or str(v).strip() == ""  # vacía si es None o si son puros espacios

        for r in range(first_data_row + 1, last_data_row + 1):  # empieza a comparar desde la segunda fila
            cur_val = ws.cell(row=r, column=col_idx).value  # valor de la fila actual

            if str(cur_val).strip() != str(prev_val).strip():  # compara sin los espacios de los lados
                # cambió el valor: cierra el tramo anterior y lo fusiona (si tenía más de 1 fila)
                if (r - 1) > fila_tramo_inicio and (not _is_empty(prev_val)):  # solo fusiona si el tramo tiene 2 o más filas y no está vacío
                    ws.merge_cells(f"{col_letter}{fila_tramo_inicio}:{col_letter}{r-1}")  # fusiona el tramo que acaba de terminar
                    c = ws.cell(row=fila_tramo_inicio, column=col_idx)  # celda visible del tramo
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrada y con salto de línea
                fila_tramo_inicio = r  # empieza un tramo nuevo
                prev_val = cur_val  # guarda el valor nuevo para seguir comparando

        # cierra el último tramo, después de terminar el for
        if last_data_row > fila_tramo_inicio and (not _is_empty(prev_val)):  # el último tramo queda abierto al salir del for
            ws.merge_cells(f"{col_letter}{fila_tramo_inicio}:{col_letter}{last_data_row}")  # lo fusiona hasta la última fila de datos
            c = ws.cell(row=fila_tramo_inicio, column=col_idx)  # celda visible del último tramo
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # mismo formato que los otros tramos

    # --- calcular filas de la tabla de circuitos en la hoja "Informe"
    hdr_circ_row = start_circ + 1                    # header
    first_data_row = hdr_circ_row + 1                # primera fila datos
    last_data_row = hdr_circ_row + len(circuitos_df) # última fila datos

    # --- ubicar columnas por nombre
    col_dif = _find_col_idx_by_name(circuitos_df, "Interruptor diferencial")  # columna del interruptor diferencial
    col_omni = _find_col_idx_by_name(circuitos_df, "Interruptor general omnipolar")  # columna del interruptor general omnipolar

    # merge omnipolar (una celda para todos los circuitos)
    merge_vertical_runs(ws, col_omni, first_data_row, last_data_row, merge_all=True)  # el omnipolar es uno solo, va todo fusionado

    # merge diferenciales por grupo (group_info)
    def merge_diferenciales_por_grupo(ws, col_idx, first_data_row, group_info):  # fusiona los circuitos que cuelgan del mismo diferencial
        # junta en una sola celda las filas de circuitos que comparten el
        # mismo diferencial, para que no se repita el texto
        # qué recibe:
        #   ws .............. la hoja de Excel
        #   col_idx ......... número de la columna del diferencial
        #   first_data_row .. primera fila de datos de la tabla de circuitos
        #   group_info ...... qué circuitos van en cada grupo (diferencial) y su calibre
        # no devuelve nada, solo fusiona y escribe el texto del diferencial.
        if col_idx is None or not group_info:  # sin columna o sin grupos no hay nada que hacer
            return  # no hay columna o no hay grupos, no hay nada que hacer

        col_letter = get_column_letter(col_idx)  # letra de columna Excel (ej: "F")

        for gid, meta in group_info.items():  # recorre grupo por grupo de diferencial
            idxs = sorted(meta.get("indices", []))  # filas (posiciones) de los circuitos de este grupo
            if not idxs:  # grupo sin circuitos, se salta
                continue  # grupo vacío, se salta

            start_r = first_data_row + idxs[0]  # primera fila del grupo en la hoja
            end_r   = first_data_row + idxs[-1]  # última fila del grupo en la hoja

            # Es el mismo texto que ya trae la columna "Interruptor diferencial"
            # del DataFrame (se armó igual en el bloque de diferenciales, con
            # dif_text). Se reescribe porque al fusionar celdas openpyxl deja
            # solo el valor de la primera y conviene asegurarlo.
            dif_txt = f"2X{meta['dif']} {meta.get('sensibilidad_dif', '30mA')} / Tipo A"  # texto del diferencial, ej: "2X25 30mA / Tipo A"

            if end_r > start_r:  # si el grupo tiene más de un circuito, fusiona
                ws.merge_cells(f"{col_letter}{start_r}:{col_letter}{end_r}")  # fusiona las celdas del grupo

            c = ws.cell(row=start_r, column=col_idx)  # celda de más arriba del grupo
            c.value = dif_txt  # el texto queda solo en la primera celda fusionada
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)  # texto a la izquierda y centrado vertical

    # aplicar merge por grupo
    merge_diferenciales_por_grupo(ws, col_dif, first_data_row, group_info)  # ahora sí, fusiona la columna de diferenciales


    # =========================================================
    # ENCABEZADO SECCIÓN EMPALME (sin merge)
    # =========================================================
    # Debajo de la tabla de circuitos van tres mini-tablas más, escritas a mano
    # celda por celda (no con to_excel): empalme, alimentador y verificación de
    # caída de tensión total. Cada una deja una fila en blanco respecto de la
    # anterior y vuelve a definir thin y fill_header (acá fill_header pasa a
    # celeste E6F2FF, distinto del verde de más arriba).
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side  # estilos de openpyxl para armar las tablas de abajo

    thin = Side(border_style="thin", color="AAAAAA")  # línea fina gris para los bordes
    fill_header = PatternFill(start_color="E6F2FF",  # fondo celeste de los encabezados
                              end_color="E6F2FF",  # mismo color, es un relleno plano
                              fill_type="solid")  # relleno sólido
    # última fila real de la tabla de circuitos
    # (se vuelven a calcular los mismos dos valores del bloque de merges de más
    #  arriba; dan idéntico resultado, solo se repiten para tenerlos a mano acá)
    hdr_circ_row = start_circ + 1  # fila del encabezado de la tabla de circuitos
    last_data_row = hdr_circ_row + len(circuitos_df)  # última fila con datos de circuitos
    # dejar 1 fila en blanco y escribir encabezado
    row_emp = last_data_row + 2  # fila donde parte el encabezado del empalme
    # nombres de columna para la mini-tabla que resume el empalme (la conexión a la red pública)
    headers_emp = [  # columnas de la mini-tabla del empalme
        "Empalme",  # tipo de empalme
        "Tarifa",  # tarifa contratada
        "Pot. Nominal (kW)",  # potencia nominal en kW
        "Acometida",  # conductor de la acometida
        "Longitud (m)",  # largo del tramo
        "Disyuntor termomagnético",  # protección general del empalme
        "Caída de tensión",  # caída de tensión de la acometida
        "Canalización"  # cómo va tendida
    ]
    for columna, txt in enumerate(headers_emp, start=1):  # escribe cada encabezado en su columna
        cell = ws.cell(row=row_emp, column=columna)  # celda del encabezado
        cell.value = txt  # texto del encabezado
        cell.font = Font(bold=True)  # en negrita
        cell.alignment = Alignment(horizontal="center",  # centrado horizontal
                                   vertical="center",  # y centrado vertical
                                   wrap_text=True)  # con salto de línea si no cabe
        cell.fill = fill_header  # fondo celeste
        cell.border = Border(top=thin,  # borde fino en los cuatro lados
                             left=thin,
                             right=thin,
                             bottom=thin)  # cierra el borde del encabezado
        row_emp_data = row_emp + 1  # fila donde van los datos, justo debajo del encabezado
        # (esta asignación está dentro del for: se repite una vez por columna,
        #  siempre con el mismo valor. Queda igual que si estuviera afuera.)

    # escribe los datos del empalme, celda por celda, debajo de cada encabezado
    # columna 1 ('Empalme'): tipo de empalme, ya armado antes como texto
    ws.cell(row=row_emp_data, column=1).value = empalme_txt  # escribe el texto del tipo de empalme
    cell_empalme = ws.cell(row=row_emp_data, column=1)  # guarda la celda para poder darle formato después
    cell_empalme.alignment = Alignment(horizontal="center", vertical="center")  # centra el texto horizontal y verticalmente
    cell_empalme.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino en los 4 lados de la celda
    # columna 2 ('Tarifa'): siempre BT-1, la tarifa residencial estándar
    ws.cell(row=row_emp_data, column=2).value = "BT-1"  # tarifa fija BT-1
    ws.cell(row=row_emp_data, column=2).alignment = Alignment(horizontal="center")  # centra el texto
    # columna 3 ('Pot. Nominal (kW)'): potencia nominal contratada del empalme
    cell_pot = ws.cell(row=row_emp_data, column=3)  # celda de la potencia nominal
    cell_pot.value = potencia_nominal_empalme  # número con la potencia nominal
    cell_pot.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    cell_pot.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # columna 4 ('Acometida'): texto que describe la acometida (conductor y tipo)
    cell_acom = ws.cell(row=row_emp_data, column=4)  # celda de la acometida
    cell_acom.value = acometida_txt  # texto de la acometida
    cell_acom.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    cell_acom.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # columna 6 ('Disyuntor termomagnético'): interruptor general del empalme
    # (el código escribe la columna 6 antes que la 5; el orden en la hoja no cambia)
    ws.cell(row=row_emp_data, column=6).value = interruptor_texto  # disyuntor termomagnético del empalme
    ws.cell(row=row_emp_data, column=6).alignment = Alignment(horizontal="center", vertical="center")  # centrado
    ws.cell(row=row_emp_data, column=6).border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # columna 5 ('Longitud (m)'): largo del tramo del transformador al empalme
    cel_long = ws.cell(row=row_emp_data, column=5)  # celda de la longitud del tramo
    cel_long.value = f"{longitud_transformador_empalme:.2f}"  # número con 2 decimales, como texto
    cel_long.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    cel_long.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # le agrega el borde a la celda de la Tarifa (el valor "BT-1" ya se puso arriba)
    cell_tarifa = ws.cell(row=row_emp_data, column=2)  # celda de la tarifa, solo para ponerle el borde
    cell_tarifa.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino igual que las demás
    # columna 7 ('Caída de tensión'): caída de tensión calculada para la acometida
    cell_dv_acom = ws.cell(row=row_emp_data, column=7)  # celda de la caída de tensión de la acometida
    cell_dv_acom.value = caida_acometida_txt  # texto con la caída ya calculada
    cell_dv_acom.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    cell_dv_acom.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # columna 8 ('Canalización'): como va tendida la acometida (aérea, subterránea, etc.)
    cell_canal_acom = ws.cell(row=row_emp_data, column=8)  # celda de la canalización de la acometida
    cell_canal_acom.value = canalizacion_acom_txt  # texto del tipo de canalización
    cell_canal_acom.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    cell_canal_acom.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino

    # =========================================================
    # SECCIÓN ALIMENTADOR (con 1 fila en blanco antes)
    # =========================================================
    # fila donde empieza esta tabla: 1 fila libre después de los datos del empalme
    row_alim = row_emp_data + 2  # dos filas más abajo, para dejar un espacio
    # nombres de columna de la tabla del alimentador
    headers_alim = ["Alimentador", "Tipo de alimentador", "Canalización", "Longitud (m)", "Caída de tensión"]  # columnas de la tabla del alimentador
    # Encabezado con mismo formato
    # recorre cada encabezado y lo escribe con el mismo estilo que las tablas anteriores
    for columna, txt in enumerate(headers_alim, start=1):  # escribe los encabezados del alimentador
        cell = ws.cell(row=row_alim, column=columna)  # celda del encabezado
        cell.value = txt  # texto de la columna
        cell.font = Font(bold=True)  # negrita
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado y con salto de línea si el texto es largo
        cell.fill = fill_header  # fondo celeste, igual que los otros encabezados
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # Fila de datos debajo (con bordes)
    # fila justo debajo del encabezado, para los datos del alimentador
    row_alim_data = row_alim + 1  # fila de datos del alimentador
    # primero les pone el borde y la alineación a todas las celdas de la fila
    for columna in range(1, len(headers_alim) + 1):  # recorre las columnas de esa fila
        cell = ws.cell(row=row_alim_data, column=columna)  # celda a formatear
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino en los cuatro lados
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrada y con ajuste de texto
    # ahora escribe el valor de cada columna: alimentador, tipo, canalización, largo y caída de tensión
    ws.cell(row=row_alim_data, column=1).value = alim_txt  # columna 1: texto del alimentador
    ws.cell(row=row_alim_data, column=2).value = f"{tipo_alimentador} (Método {res_alim['metodo']})"  # columna 2: tipo de alimentador + método de instalación de la tabla de ampacidad (E, D1 o B1)
    ws.cell(row=row_alim_data, column=3).value = canalizacion_txt  # columna 3: canalización del alimentador
    ws.cell(row=row_alim_data, column=4).value = f"{longitud_alimentador:.2f}"  # columna 4: longitud en metros, con 2 decimales
    ws.cell(row=row_alim_data, column=5).value = caida_txt  # columna 5: caída de tensión del alimentador
    # ancho fijo para que se lean bien los textos largos de esas columnas
    # estos dos anchos son los últimos que se aplican a B y C, así que son los
    # que manda el Excel: pisan el 55 que se le había puesto a la B más arriba
    ws.column_dimensions["B"].width = 32  # columna del tipo de alimentador (ancho final de la B)
    ws.column_dimensions["C"].width = 35  # columna de la canalización

    # =========================================================
    # TABLA VERIFICACIÓN CAÍDA DE TENSIÓN TOTAL (alim + circ <= 5%)
    # =========================================================
    # Qué resuelve: dejar por escrito, circuito por circuito, la suma
    # ΔV del circuito + ΔV del alimentador, y si esa suma respeta el 5% que es
    # el tope de caída acumulada. Los circuitos sin caída calculada (NaN) no
    # aparecen en la tabla.
    # porcentaje de caída de tensión del alimentador (ya calculado antes)
    _dv_alim_pct = res_alim.get("dV_pct", 0.0)  # si no viene calculada, queda en 0%
    # estilos propios de esta tabla: colores de fondo y tipografia Calibri
    _fill_title_dv  = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")  # celeste del título
    _fill_hdr_dv    = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # gris de los encabezados
    _font_calibri_b = Font(name="Calibri", size=10, bold=True)  # Calibri 10 en negrita
    _font_calibri   = Font(name="Calibri", size=10)  # Calibri 10 normal
    # encabezados de la tabla: circuito, caída del circuito, del alimentador, total y si cumple
    _hdrs_dv = ["Circuito", "ΔV circuito (%)", "ΔV alimentador (%)", "ΔV total (%)", "¿Cumple ≤ 5%?"]
    _ncols_dv = len(_hdrs_dv)  # cantidad de columnas de esta tabla: 5

    # fila del título de la tabla, 2 filas después de los datos del alimentador
    row_dv_title = row_alim_data + 2  # deja una fila en blanco antes del título
    # fusiona todas las columnas del título en una sola celda ancha
    ws.merge_cells(start_row=row_dv_title, start_column=1,  # el título ocupa todo el ancho de la tabla
                   end_row=row_dv_title, end_column=_ncols_dv)  # hasta la última columna
    # escribe el título y le da formato
    _tc = ws.cell(row=row_dv_title, column=1)  # celda donde queda el texto del título
    _tc.value     = "VERIFICACIÓN CAÍDA DE TENSIÓN TOTAL"  # título de la tabla
    _tc.font      = Font(name="Calibri", size=10, bold=True)  # negrita
    _tc.alignment = Alignment(horizontal="center", vertical="center")  # centrado
    _tc.fill      = _fill_title_dv  # fondo celeste
    _tc.border    = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
    # le pone el mismo fondo y borde al resto de las celdas fusionadas del título
    for _c in range(2, _ncols_dv + 1):  # recorre las celdas tapadas por el merge para darles el mismo fondo y borde
        _cell = ws.cell(row=row_dv_title, column=_c)  # celda tapada por el merge
        _cell.fill   = _fill_title_dv  # mismo fondo
        _cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # mismo borde

    # fila con los encabezados de columna de la tabla (debajo del título)
    row_dv_hdr = row_dv_title + 1  # encabezados justo debajo del título
    # recorre cada encabezado y lo escribe con su estilo
    for _j, _h in enumerate(_hdrs_dv, start=1):  # escribe cada encabezado de la tabla de caídas
        _cell = ws.cell(row=row_dv_hdr, column=_j)  # celda del encabezado
        _cell.value     = _h  # texto de la columna
        _cell.font      = _font_calibri_b  # negrita Calibri 10
        _cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado y con salto de línea
        _cell.fill      = _fill_hdr_dv  # fondo gris
        _cell.border    = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino

    # a partir de aquí empiezan las filas de datos, una por circuito
    row_dv_data = row_dv_hdr + 1  # primera fila de datos, va avanzando sola
    # solo arma la tabla si el dataframe de circuitos tiene las columnas necesarias
    if "Circuito" in circuitos_df.columns and "Caída de tensión (%)" in circuitos_df.columns:  # sin esas columnas no se puede armar la verificación
        # recorre cada circuito y verifica que la caída de tensión total
        # (circuito + alimentador) no supere el 5% máximo permitido
        # recorre cada circuito del informe
        for _idx, _row in circuitos_df.iterrows():  # una fila por circuito
            _nombre  = _row.get("Circuito", "")  # nombre del circuito
            _dv_circ = _row.get("Caída de tensión (%)", None)  # caída de tensión propia del circuito (%), puede venir vacía
            # intenta convertir la caída de tensión a número; si no se puede o es NaN, se salta este circuito
            try:  # la caída puede venir como texto o vacía
                _dv_circ_f = float(_dv_circ)  # la pasa a número
                if _dv_circ_f != _dv_circ_f:  # es NaN
                    continue  # NaN, no sirve para comparar
            except (TypeError, ValueError):  # no era un número
                continue  # se salta ese circuito
            # caída de tensión TOTAL = caída del circuito + caída del alimentador
            _dv_total_f = round(_dv_circ_f + _dv_alim_pct, 2)
            # el máximo permitido por el RIC es 5% de caída de tensión total.
            # Normalmente todas las filas dicen "cumple", porque
            # seleccionar_alimentador() ya subió de calibre hasta que
            # ΔV_alimentador + ΔV_peor_circuito quedara bajo el 5%. Solo sale
            # "NO cumple" si se agotaron las secciones comerciales (esa función
            # devuelve una advertencia en ese caso).
            _cumple_txt = "cumple" if _dv_total_f <= 5.0 else "NO cumple"
            # arma la fila de valores en el orden de las columnas de esta tabla
            _vals = [_nombre, f"{_dv_circ_f:.2f}%", f"{_dv_alim_pct:.2f}%", f"{_dv_total_f:.2f}%", _cumple_txt]
            # escribe cada valor en su columna, con el mismo estilo que el resto del informe
            for _j, _v in enumerate(_vals, start=1):  # escribe los 5 valores de la fila
                _cell = ws.cell(row=row_dv_data, column=_j)  # celda a escribir
                _cell.value     = _v  # valor
                _cell.font      = _font_calibri  # Calibri 10
                _cell.alignment = Alignment(  # alineación de la celda
                    horizontal="left" if _j == 1 else "center",  # el nombre del circuito a la izquierda, el resto centrado
                    vertical="center", wrap_text=True)  # centrado vertical y con salto de línea
                _cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino
            # pasa a la siguiente fila para el próximo circuito
            row_dv_data += 1
    # =========================
    # HOJA 2: MATERIALES (PRO)
    # =========================
    # De aquí al final se arma la segunda hoja del Excel. Orden de los pasos:
    #   1. recalcular la columna "Sello SEC" con _requiere_sello_sec()
    #   2. acortar la columna "Norma / RIC" con normalizar_ric_materiales()
    #   3. volcar materiales_df a la hoja y darle formato (secciones, anchos,
    #      alturas, filtro, panel fijo)
    #   4. llamar a aplicar_base_normativa_e_hipervinculos(), que crea la hoja
    #      "Base Normativa" y convierte la columna Norma / RIC en enlaces
    # Calcular Sello SEC basado en descripción Y norma original
    # Decide si un material necesita 'Sello SEC' (certificacion obligatoria de la
    # Superintendencia de Electricidad y Combustibles) según su descripción y su
    # norma original. Recibe el texto de la descripción y la norma, y devuelve
    # "SEC", "-" (no necesita) o "" (sin descripción).
    def _requiere_sello_sec(desc, norma):  # decide qué va en la columna "Sello SEC" de cada material
        # decide qué poner en la columna "Sello SEC": si ya venía marcado como
        # SEC lo deja igual, y si es un consumible (cinta, sellador, etc.)
        # no le pide sello
        d = str(desc or "").lower()  # descripción en minúsculas, para comparar sin problemas de mayúsculas
        n = str(norma or "").strip()  # norma original, sin espacios sobrantes
        # si la norma ya decía "SEC", se respeta tal cual.
        # Acá 'norma' todavía es el texto crudo que puso build_materiales_df
        # (ej: "RIC 4 (5.13)" o "SEC"): esta función corre ANTES de que
        # normalizar_ric_materiales() lo acorte, así que la comparación exacta
        # con "SEC" funciona.
        if n == "SEC":  # la norma ya venía marcada como SEC
            return "SEC"  # se respeta tal cual
        # sin descripción no se puede decidir nada
        if not d or d in ["nan", "none", ""]:  # descripción vacía o basura
            return ""  # queda en blanco
        # Los consumibles de instalación no llevan Sello SEC
        if any(x in d for x in ["tubo de estaño", "pasta para soldar", "cinta aislante", "cinta autofundente", "sellador", "espuma", "teflón", "teflon"]):  # cintas, pastas y selladores no se certifican
            return "-"  # guion, o sea no aplica
        # Conductores y cables
        if any(x in d for x in ["conductor", "cable", "thwn-2", "rv-k", "nyyj", "nyy",  # conductores y cables
                                  "concéntrico", "concentrico", "h07z1", "desnudo cu"]):  # sigue la lista de tipos de cable
            return "SEC"  # los cables van certificados
        # Protecciones
        if any(x in d for x in ["interruptor", "diferencial", "termomagnético",  # protecciones del tablero
                                  "termomagnetico", "disyuntor", "omnipolar",  # más nombres de protecciones
                                  "automático", "automatico"]):  # y los automáticos
            return "SEC"  # las protecciones van certificadas
        # Enchufes
        if "enchufe" in d:  # enchufes
            return "SEC"  # certificados
        # Tableros
        if "tablero" in d:  # tableros
            return "SEC"  # certificados
        # Cajas de derivación y empalme
        if any(x in d for x in ["caja de derivación", "caja de derivacion",  # cajas de derivación
                                  "caja de empalme"]):  # y cajas de empalme
            return "SEC"  # certificadas
        # Tapas de cajas
        if "tapa ciega" in d:  # tapas ciegas de las cajas
            return "SEC"  # certificadas
        # Conduit y accesorios
        if any(x in d for x in ["conduit", "canaleta pvc", "tubo conduit",  # tubos conduit y canaletas
                                  "tubo galvanizado", "salida de caja",  # tubo galvanizado y salidas de caja
                                  "unión copla", "union copla",  # uniones copla
                                  "curva", "terminal pvc conduit"]):  # curvas y terminales
            return "SEC"  # certificados
        # Abrazaderas
        if "abrazadera" in d:  # abrazaderas
            return "SEC"  # certificadas
        # Prensaestopas y boquillas
        if any(x in d for x in ["prensaestopa", "boquilla"]):  # prensaestopas y boquillas
            return "SEC"  # certificados
        # Medidor
        if any(x in d for x in ["medidor", "medida monofásica", "medida trifásica"]):  # medidor de la empresa
            return "SEC"  # certificado
        # Portafusibles y fusibles
        if any(x in d for x in ["portafusible", "fusible"]):  # portafusibles y fusibles
            return "SEC"  # certificados
        # Barras repartidoras y unipolares
        if any(x in d for x in ["barra unipolar", "barra repartidora"]):  # barras del tablero
            return "SEC"  # certificadas
        # Borneras
        if "bornera" in d:  # borneras
            return "SEC"  # certificadas
        # Terminales
        if any(x in d for x in ["terminal ferrul", "terminal de compresión",  # terminales ferrul
                                  "terminal de compresion"]):  # y de compresión
            return "SEC"  # certificados
        # Supresor de transiente
        if "supresor" in d:  # supresor de transientes
            return "SEC"  # certificado
        # Protector sobrevoltaje
        if "protector sobrevoltaje" in d or "protector de sobrevoltaje" in d:  # protector de sobrevoltaje
            return "SEC"  # certificado
        # Portalámparas y luminarias
        if any(x in d for x in ["portalámpara", "portalampara", "luminaria",  # portalámparas y luminarias
                                  "ampolleta", "led", "foco"]):  # ampolletas, led y focos
            return "SEC"  # certificados
        # Riel DIN
        if "riel din" in d:  # riel DIN del tablero
            return "SEC"  # certificado
        # Mordaza
        if "mordaza" in d:  # mordazas
            return "SEC"  # certificadas
        # Hub / conector acometida
        if any(x in d for x in ["hub", "conector hub"]):  # hub o conector de la acometida
            return "SEC"  # certificado
        # Alimentador
        if "alimentador" in d:  # alimentador
            return "SEC"  # certificado
        # Punto conexión fija / directa
        if "punto para conexión" in d or "punto para conexion" in d:  # punto de conexión fija o directa
            return "SEC"  # certificado
        # Todo lo demás no requiere Sello SEC
        return "-"  # todo lo demás no necesita sello

    # si la hoja de materiales tiene la columna 'Sello SEC', la rellena usando la función anterior
    if "Sello SEC" in materiales_df.columns:  # solo si la hoja de materiales trae esa columna
        # aplica la función a cada fila para decidir si necesita Sello SEC
        materiales_df["Sello SEC"] = materiales_df.apply(  # recalcula el sello de cada material
            lambda rr: _requiere_sello_sec(rr.get("Descripción técnica", ""), rr.get("Norma / RIC", "")),  # usa la descripción y la norma de esa fila
            axis=1  # fila por fila
        )

    # Limpia la columna "Norma / RIC":
    # Ejemplo: "RIC 4 (5.13)" queda como "RIC 4". Es solo un paso intermedio y
    # ese "RIC 4" no llega a verse en el Excel: al final del bloque,
    # aplicar_base_normativa_e_hipervinculos() recorre la hoja Materiales y le
    # pisa el valor a cada celda con "Ver normativa" + hipervínculo a la hoja
    # Base Normativa. Lo que se ve en la columna es "Ver normativa" o "-".
    if "Norma / RIC" in materiales_df.columns:
        # aplica la función normalizar_ric_materiales fila por fila, usando la norma y la descripción
        materiales_df["Norma / RIC"] = materiales_df.apply(  # reescribe la columna con la norma ya normalizada
            lambda rr: normalizar_ric_materiales(  # deja la norma en formato corto, tipo "RIC 4"
                rr.get("Norma / RIC", ""),  # norma que traía el material
                rr.get("Descripción técnica", "")  # descripción, por si hay que deducir la norma
            ),
            axis=1  # recorre fila por fila
        )

    # vuelca la tabla de materiales al Excel, empezando en la fila 3 (deja espacio para el título)
    materiales_df.to_excel(writer, sheet_name="Materiales", index=False, startrow=2)
    ws2 = writer.sheets["Materiales"]  # hoja de Excel recién creada

    # título de la hoja
    # fusiona A1:J1 para poner el título. La J está escrita a mano y calza
    # porque materiales_df tiene exactamente 10 columnas (Ítem, Descripción
    # técnica, Marcas sugeridas, Sello SEC, Norma / RIC, Circuito, Unidad, K,
    # Longitud (m) / Unidad, Cantidad), como las devuelve build_materiales_df.
    ws2.merge_cells("A1:J1")
    ws2["A1"] = "CUBICACIÓN DE MATERIALES PARA INSTALACIONES RESIDENCIALES EN CASAS PREFABRICADAS"  # texto del título que se ve arriba de la tabla
    ws2["A1"].font = Font(bold=True, size=14)  # título en negrita y más grande que el resto
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")  # título centrado dentro de la celda fusionada
    ws2.row_dimensions[1].height = 24  # fila del título un poco más alta

    # fila donde empieza el encabezado de la tabla de materiales
    header_row = 3
    thin = Side(border_style="thin", color="AAAAAA")  # borde fino, mismo estilo que en la hoja Informe
    fill_header = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")  # fondo celeste para encabezados
    fill_section = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # fondo gris para separar secciones

    ncols2 = materiales_df.shape[1]  # cantidad de columnas de la tabla de materiales
    nrows2 = materiales_df.shape[0]  # cantidad de filas de datos (sin contar encabezado)
    last_row2 = header_row + nrows2  # última fila con datos

    # color gris para pintar la fila que separa la última sección del resto
    gris_fill = PatternFill(
    fill_type="solid",  # relleno sólido, sin degradado
    start_color="F2F2F2",  # gris claro, el mismo que usan las filas de sección
    end_color="F2F2F2"  # color de termino igual al de inicio para que quede parejo
    )

    # Corte visual: busca la primera fila cuya Descripción (columna B) mencione
    # "Espuma expansiva" y pinta de gris la fila de ABAJO, para separar ese
    # bloque del siguiente. Si no aparece esa fila, no pinta nada.
    for fila in range(4, last_row2 + 1):
        texto = ws2.cell(row=fila, column=2).value  # texto de la columna B (Descripción) en esa fila
        if texto and "Espuma expansiva" in str(texto):  # cuando aparece la fila de la espuma expansiva
            fila_gris = fila + 1  # la fila de abajo es la que se pinta de gris
            # pinta de gris todas las columnas de esa fila
            for columna in range(1, ncols2 + 1):
                ws2.cell(row=fila_gris, column=columna).fill = gris_fill  # pinta de gris esa celda
            # una vez encontrada la sección, no hace falta seguir buscando
            break

    # formato del encabezado de la tabla de materiales
    # recorre las columnas del encabezado y les da el mismo formato que en la hoja Informe
    for ccol in range(1, ncols2 + 1):
        cell = ws2.cell(row=header_row, column=ccol)  # celda del encabezado en la columna que toca
        cell.font = Font(bold=True)  # negrita
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)  # centrado, con salto de línea si el texto es largo
        cell.fill = fill_header  # fondo celeste
        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde fino


    # Recorre cada fila de datos y le da formato. Criterio para distinguirlas:
    # una fila es "cabecera de sección" (ej. "CANALIZACIONES", "PROTECCIONES")
    # si viene sin Ítem en la columna A pero con texto en la Descripción
    # (columna B) — así las deja build_materiales_df. Esas se pintan grises, en
    # negrita y fusionadas de A a J; las demás son materiales normales.
    for r in range(header_row + 1, last_row2 + 1):
        # ¿es una fila de "sección"? no tiene Ítem (columna A vacía) pero sí tiene Descripción (columna B)
        is_section = (ws2.cell(row=r, column=1).value in [None, ""]) and (ws2.cell(row=r, column=2).value not in [None, ""])
        max_lineas = 1  # cuántas líneas (separadas por \n) tiene el texto más largo de la fila
        # recorre cada columna de la fila actual para darle formato (borde, alineación, etc.)
        for ccol in range(1, ncols2 + 1):
            cell = ws2.cell(row=r, column=ccol)  # celda actual
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)  # borde delgado alrededor de la celda
            if ccol in [8, 9, 10]:  # columnas H, I y J de Excel: "K" (factor de desperdicio), "Longitud (m) / Unidad" y "Cantidad"
                cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)  # números alineados a la derecha
            else:  # el resto de las columnas son texto
                cell.alignment = Alignment(vertical="top", wrap_text=True)  # el resto del texto alineado arriba, con ajuste de línea
            if ccol == 8 and isinstance(cell.value, (int, float)):  # columna 8 = columna H de Excel, la que tiene por título "K"
                cell.number_format = "0.00"  # el factor K (desperdicio) se muestra con dos decimales
            if is_section:  # las filas de sección se pintan distinto
                cell.fill = fill_section  # fondo gris para destacar la fila de sección
                cell.font = Font(bold=True)  # texto en negrita
            if isinstance(cell.value, str) and "\n" in cell.value:  # si el texto de la celda trae saltos de línea
                max_lineas = max(max_lineas, cell.value.count("\n") + 1)  # guarda cuántas líneas tiene el texto más largo de la fila

        if is_section:  # la fila de sección se arma como un solo bloque
            # las filas de sección ocupan todo el ancho, fusionadas en una sola celda
            ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols2)
            ws2.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")  # texto de la sección alineado a la izquierda
            ws2.row_dimensions[r].height = 18  # altura fija para la fila de sección
        else:  # filas normales de material
            # Altura dinámica: ~15 puntos por línea de texto, para que el
            # detalle multi-línea (ferrules, cónicos, etc.) se vea completo
            # sin tener que agrandar la fila a mano en Excel.
            ws2.row_dimensions[r].height = max(15, max_lineas * 15)

    # Ancho automático según el contenido más largo de cada columna.
    # Ojo: justo después se fijan a mano los anchos de A a J, así que este
    # cálculo automático solo sobrevive si materiales_df tuviera más de 10
    # columnas (hoy tiene exactamente 10).
    for ccol in range(1, ncols2 + 1):
        col_letter = get_column_letter(ccol)  # convierte el número de columna a letra (A, B, C...)
        max_len = 0  # reinicia el largo máximo de texto encontrado en esta columna
        for r in range(1, last_row2 + 1):  # recorre todas las filas de esa columna
            v = ws2.cell(row=r, column=ccol).value  # contenido de la celda, para medir cuánto ocupa
            if v is None:  # celda vacía no aporta al ancho
                continue  # celda vacía, se salta
            max_len = max(max_len, len(str(v)))  # guarda el largo de texto más grande encontrado
        ws2.column_dimensions[col_letter].width = min(max(10, max_len + 4), 55)  # ancho según el texto, entre 10 y 55

    # anchos fijos definidos a mano para las columnas principales (sobreescribe el automático)
    ws2.column_dimensions["A"].width = 6  # Item
    ws2.column_dimensions["B"].width = 48  # Descripción técnica, es la columna más larga
    ws2.column_dimensions["C"].width = 30  # Marcas sugeridas
    ws2.column_dimensions["D"].width = 12  # Sello SEC
    ws2.column_dimensions["E"].width = 20  # Norma / RIC
    ws2.column_dimensions["F"].width = 45  # Circuito
    ws2.column_dimensions["G"].width = 8  # Unidad (u o m)
    ws2.column_dimensions["H"].width = 5  # factor K de desperdicio (columna titulada "K", angosta porque es un número corto)
    ws2.column_dimensions["I"].width = 22  # Longitud (m) / Unidad
    ws2.column_dimensions["J"].width = 10  # Cantidad final a comprar

    ws2.auto_filter.ref = f"A{header_row}:{get_column_letter(ncols2)}{last_row2}"  # agrega filtro automático a toda la tabla
    ws2.freeze_panes = f"A{header_row+1}"  # deja fija la fila de encabezado cuando se mueve hacia abajo en la pantalla

    # =========================
    # HOJA 3: BASE NORMATIVA + HIPERVÍNCULOS DESDE MATERIALES
    # =========================
    # arma la tercera hoja del Excel (Base Normativa) y deja los hipervínculos
    # en la hoja Materiales apuntando al artículo del RIC correspondiente
    aplicar_base_normativa_e_hipervinculos(
        writer,  # el mismo ExcelWriter donde se están escribiendo las hojas
        materiales_df,  # tabla de materiales, de ahí salen los artículos del RIC
        sheet_materiales="Materiales",  # hoja desde donde se hacen los hipervínculos
        sheet_base="Base Normativa"  # hoja nueva con el detalle de cada artículo
    )

# al terminar de escribir el Excel, intenta abrirlo automáticamente
# (funciona distinto según el sistema operativo: Windows, Mac o Linux)
try:
    if sys.platform.startswith("win"):  # Windows
        os.startfile(nombre_archivo)  # abre el archivo con el programa predeterminado (por ejemplo Excel)
    elif sys.platform == "darwin":  # Mac
        os.system(f'open "{nombre_archivo}"')  # comando "open" de macOS
    else:  # Linux
        os.system(f'xdg-open "{nombre_archivo}"')  # comando de Linux para abrir con la app predeterminada
except Exception:  # si el sistema operativo no lo deja abrir, se ignora
    pass  # si no se pudo abrir automáticamente, no es grave, el archivo ya quedó guardado

print(f"\n Informe generado correctamente: {nombre_archivo}")  # avisa al usuario que terminó y le muestra el nombre del archivo