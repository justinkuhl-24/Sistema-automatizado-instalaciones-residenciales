# Sistema automatizado para instalaciones eléctricas en casas prefabricadas

Programa de consola en Python que dimensiona una instalación eléctrica residencial completa y genera el informe técnico y la cubicación de materiales en un archivo Excel, aplicando el Reglamento de Instalaciones de Consumo (RIC) de la SEC de Chile.

Está orientado a viviendas prefabricadas y construcción en panel SIP. El usuario responde una serie de preguntas por consola sobre ambientes, cargas, empalme y canalizaciones, y el programa calcula conductores, protecciones, canalizaciones y puesta a tierra, además de emitir la lista completa de materiales con la referencia normativa de cada ítem.

## Contenido

1. [Qué resuelve](#qué-resuelve)
2. [Qué hace el programa](#qué-hace-el-programa)
3. [Archivo Excel de salida](#archivo-excel-de-salida)
4. [Requisitos](#requisitos)
5. [Instalación](#instalación)
6. [Uso](#uso)
7. [Flujo de ejecución](#flujo-de-ejecución)
8. [Motor de cálculo](#motor-de-cálculo)
9. [Resumen de fórmulas](#resumen-de-fórmulas)
10. [Cubicación de materiales](#cubicación-de-materiales)
11. [Base normativa](#base-normativa)
12. [Estructura del código](#estructura-del-código)
13. [Personalización](#personalización)
14. [Limitaciones](#limitaciones)
15. [Problemas frecuentes](#problemas-frecuentes)
16. [Trabajo pendiente](#trabajo-pendiente)
17. [Contribuir](#contribuir)
18. [Advertencia de uso](#advertencia-de-uso)
19. [Autor](#autor)
20. [Licencia](#licencia)

## Qué resuelve

Proyectar una instalación residencial y cubicarla implica cruzar tablas de ampacidad, verificar caída de tensión, elegir diámetros de conduit, contar cajas, abrazaderas, ferrules y tornillos, y dejar registro del artículo del RIC que respalda cada decisión. Es un trabajo repetitivo y fácil de equivocar.

El programa cubre ese ciclo completo. Recibe los datos del proyecto por consola, ejecuta el cálculo normativo y entrega un Excel de tres hojas que sirve a la vez como memoria de cálculo, lista de compra y respaldo normativo frente a la inspección.

## Qué hace el programa

Levantamiento de ambientes: registra área, perímetro, material de tabique y de forrado interior, luminarias con potencia real o estimada, luminarias conmutadas, enchufes por cantidad de módulos y componentes especiales.

Armado de circuitos: distribuye los ambientes en circuitos de iluminación, enchufes, cocina y lavadero, y circuitos especiales. Cuando corresponde agrupa cargas con un algoritmo de bin packing (First-Fit Decreasing) respetando la potencia máxima por circuito.

Factor de demanda: aplica la Tabla N°3.1 del RIC 3, con los primeros 3 kW al 100% y el resto al 35%. La climatización y el agua caliente quedan a factor 1,0.

Selección del empalme: calcula el interruptor del empalme (25, 32, 40, 50 o 63 A) sobre la corriente con factor de demanda más un 10% de holgura.

Conductores: determina la sección por caída de tensión con límite de 3% y verifica la ampacidad aplicando el factor de corrección por temperatura de la Tabla N°4.7.

Canalizaciones: elige el diámetro de conduit según las Tablas N°4.17, N°4.19, N°4.20 y N°4.29, o dimensiona canaleta por área real de los conductores cuando la canalización es sobrepuesta.

Climatización: genera un circuito exclusivo por equipo, con conductor a 1,25 veces la corriente de placa, termomagnético de curva C o D según sea inverter u on-off, y verificación contra el LRA y el MOCP declarados por el fabricante.

Agua caliente: trata duchas, termos y calefones como carga resistiva pura con factor de potencia 1,0, y asigna diferencial de 10 mA cuando el equipo queda en el Volumen 1 del baño. Si corresponde, agrega el tablero de desconexión externo a la vista del equipo.

Diferenciales: agrupa los circuitos generales mediante backtracking, minimizando la cantidad de diferenciales sin superar los 3 circuitos por unidad ni el tope de suma de calibres derivado del empalme.

Puesta a tierra: calcula dos mallas independientes, una en el empalme y otra en el tablero, entregando número de barras copperweld, resistencia resultante, separación mínima y metros de conductor desnudo.

Cubicación: emite varios cientos de líneas de material con cantidad, unidad, factor de holgura, marcas de referencia, indicación de sello SEC y circuito asociado.

Trazabilidad: cada material queda hipervinculado a la hoja que contiene el artículo del RIC que lo justifica.

## Archivo Excel de salida

Al terminar la ejecución se genera en el directorio de trabajo un archivo con la fecha y hora en el nombre:

```
Informe_Instalacion_20260817_143052.xlsx
```

El programa intenta abrirlo automáticamente en Windows, macOS y Linux.

### Hoja Informe

Es la hoja más densa del archivo. No es una sola tabla, sino ocho bloques distintos repartidos en la hoja: los que ocupan las columnas A y B van apilados de arriba hacia abajo, y a la derecha de los parámetros generales aparecen dos cuadros anexos.

Este es el mapa completo:

| Bloque | Ubicación | Contenido |
|---|---|---|
| 1 | Columnas A en adelante, arriba | Detalle por ambiente, con autofiltro |
| 2 | Columnas A y B, debajo del anterior | Parámetros generales y puesta a tierra |
| 3 | Desde la columna D, a la derecha del bloque 2 | Cuadro de cálculo de la protección general |
| 4 | Desde la columna K, a la derecha del bloque 3 | Tabla de fórmulas utilizadas |
| 5 | Columnas A en adelante, debajo del bloque 2 | Cuadro de cargas y circuitos |
| 6 | Debajo del bloque 5 | Datos del empalme y la acometida |
| 7 | Debajo del bloque 6 | Datos del alimentador |
| 8 | Debajo del bloque 7 | Verificación de caída de tensión total |

#### Bloque 1: detalle por ambiente

| Columna | Contenido |
|---|---|
| Ambiente | Nombre ingresado (Living, Baño, Dormitorio 1) |
| Área (m²), Perímetro (m) | Dimensiones usadas para determinar la cantidad mínima de puntos |
| Material tabique, Material forrado interior | Determinan tornillería, tarugos y sellos del panel SIP |
| Detalle iluminación, Potencia iluminación (W) | Luminarias con tipo, montaje y potencia |
| N_conmutadas_924 (u) | Luminarias comandadas desde dos puntos |
| L_viajeros_924, L_retorno_lampara, L_troncal_oct1, L_oct1_oct2 | Longitudes parciales que alimentan el cálculo de metros de cable |
| Detalle enchufes comunes, Potencia enchufes (W) | Enchufes desglosados por módulos |
| Componentes especiales, Potencia comp. especiales (W) | Cargas puntuales declaradas por el usuario |
| Total por ambiente (W) | Suma de las tres potencias anteriores |
| Cantidad luminarias, enchufes, módulos y comp. especiales | Conteos que se usan en la cubicación |

#### Bloque 2: parámetros generales y puesta a tierra

Tabla de dos columnas, Parámetro y Valor, con zona, tipo de canalización, protección de empalme calculada, tensión nominal, factor de potencia, temperatura ambiente y potencia total estimada.

A continuación viene el largo de la barra copperweld y, separadas por subtítulos que el programa fusiona y centra, las dos mallas de tierra. Para cada una, PT1 en el empalme y PT2 en el tablero, se indica resistividad del terreno, resistencia por barra, número de barras y resistencia final. Esta última no aparece como número seco, sino con el veredicto escrito: "12.5 menor o igual a 20 Ohm cumple", o bien "24.0 mayor a 20 Ohm NO cumple". La fila de separación mínima entre barras solo aparece cuando hace falta más de una barra, y cita el RIC 6 punto 8.3.2.

#### Bloque 3: cuadro de cálculo de la protección general

Este es el cuadro que justifica el calibre del empalme. Arranca en la columna D, a la derecha de los parámetros generales, y tiene seis columnas agrupadas en dos mitades: la carga sin factor de demanda y la carga con el factor ya aplicado.

Estas son sus seis columnas, en el orden en que aparecen:

| Columna | Qué contiene |
|---|---|
| Tablero | Nombre de la partida |
| KW | Potencia antes de aplicar el factor de demanda |
| In [A] | Corriente antes de aplicar el factor de demanda |
| f/d | Factor de demanda que se aplica a esa fila |
| KW | Potencia ya con el factor aplicado |
| In [A] | Corriente ya con el factor aplicado |

Y estas son sus filas:

| Fila | Factor de demanda que lleva |
|---|---|
| T.D.A PRIMEROS 3KW | 1,00 |
| T.D.A RESTO, con los kilowatts que exceden los primeros 3 kW | 0,35 |
| Una fila por cada equipo de climatización, con su nombre | 1,00 |
| Una fila por cada equipo de agua caliente, con su nombre | 1,00 |
| TOTAL, que suma las columnas de todas las filas anteriores | no aplica |

Las filas de climatización y agua caliente son dinámicas: aparece una por cada equipo declarado. Siempre llevan factor 1,00 porque el RIC 7 no permite reducir esas cargas. El valor de la fila TOTAL en la última columna es la corriente que después se multiplica por 1,10 para elegir el calibre del empalme.

#### Bloque 4: tabla de fórmulas

A la derecha del cuadro anterior se escribe una tabla titulada "FÓRMULAS UTILIZADAS EN LOS CÁLCULOS", con tres columnas, Parámetro o Etapa, Fórmula y Notas, y quince filas. Deja registrada dentro del propio informe cada expresión que el programa usó, con sus unidades, sus límites y el artículo del RIC correspondiente. Está reproducida completa en la sección [Resumen de fórmulas](#resumen-de-fórmulas) de este documento.

#### Bloque 5: cuadro de cargas y circuitos

| Columna | Contenido |
|---|---|
| Circuito | Nombre normalizado |
| Longitud (m) | Longitud real de canalización del circuito |
| Potencia estimada (W), Corriente estimada (A) | Carga del circuito |
| Detalle asignación | Ambientes y cargas que quedaron en ese circuito |
| Interruptor termomagnético | Calibre, poder de corte y curva |
| Interruptor diferencial | Calibre y sensibilidad, 30 mA o 10 mA |
| Interruptor general omnipolar | Protección general del tablero |
| Conductor | Tipo y sección, por ejemplo H07Z1-K 2.5 mm² |
| Caída de tensión (%) | Valor calculado, con límite de 3% |
| Canalización | Conduit embutido o canaleta sobrepuesta con su medida |

Dos columnas de esta tabla se presentan fusionadas verticalmente, para que se lea como un plano de tablero y no como una lista repetida. La columna del interruptor omnipolar se fusiona en una sola celda que abarca todos los circuitos, porque hay uno solo para toda la instalación. La columna del diferencial se fusiona por grupo: los circuitos que comparten diferencial quedan bajo una única celda, con el texto en el formato `2X40 30mA / Tipo A`. Así se ve de un vistazo qué circuitos cuelgan de cada diferencial.

#### Bloque 6: empalme y acometida

Debajo de la tabla de circuitos, separado por una fila en blanco, aparece un bloque de ocho columnas con los datos del empalme:

| Columna | Contenido |
|---|---|
| Empalme | Designación normalizada, por ejemplo A-9 o S-9 |
| Tarifa | BT-1 |
| Pot. Nominal (kW) | Potencia nominal que corresponde al calibre del empalme |
| Acometida | Conductor calculado, por ejemplo Concéntrico Cu 2x 4 mm² |
| Longitud (m) | Distancia entre transformador y empalme |
| Disyuntor termomagnético | Por ejemplo 1x25A / 6kA / Curva D |
| Caída de tensión | Porcentaje calculado del tramo de acometida |
| Canalización | Diámetro del conduit, o un guion si la acometida es aérea |

La designación del empalme se arma con la letra según el tipo de acometida, A para aérea y S para subterránea, y un número que sale del calibre: 6 hasta 30 A, 9 hasta 40 A y 16 por sobre eso. De ahí salen las denominaciones A-9 y S-9 que se usan habitualmente.

La potencia nominal se toma de una tabla interna que asocia cada calibre con su potencia en kW: 6 A equivale a 1 kW, 10 A a 2 kW, 16 A a 3 kW, 20 A a 4 kW, 25 A a 5 kW, 30 A a 6 kW, 32 A a 6,5 kW, 35 A a 7 kW, 40 A a 8 kW, 50 A a 10 kW y 63 A a 13 kW.

#### Bloque 7: alimentador

Cinco columnas con el conductor calculado, por ejemplo RV-K Cu 3x6 mm², el tipo de alimentador con el método de instalación entre paréntesis, la canalización con su diámetro, la longitud y la caída de tensión del tramo.

#### Bloque 8: verificación de caída de tensión total

Una tabla final que cruza cada circuito con el alimentador para comprobar el límite acumulado del 5%:

| Circuito | ΔV circuito (%) | ΔV alimentador (%) | ΔV total (%) | ¿Cumple ≤ 5%? |
|---|---|---|---|---|

La última columna dice literalmente "cumple" o "NO cumple" por cada circuito, de modo que el incumplimiento queda visible sin tener que sumar a mano.

### Hoja Materiales

Lleva un título fusionado en la primera fila, "CUBICACIÓN DE MATERIALES PARA INSTALACIONES RESIDENCIALES EN CASAS PREFABRICADAS". Los encabezados van en la fila 3 y los datos desde la fila 4, con el panel congelado ahí para que las columnas se mantengan visibles al desplazarse, y con autofiltro activado sobre todo el rango.

| Columna | Descripción |
|---|---|
| Ítem | Correlativo, vacío en las filas de encabezado de sección |
| Descripción técnica | Nombre completo del material con sus características |
| Marcas sugeridas | Marcas de referencia del mercado chileno |
| Sello SEC | Indica si el producto requiere certificación |
| Norma / RIC | Enlace a la hoja Base Normativa, ver nota abajo |
| Circuito | Circuito o partida a la que pertenece el material |
| Unidad | u de unidad, m de metro, y también rollo, tira o caja |
| K | Factor de holgura aplicado |
| Longitud (m) / Unidad | Base de cálculo antes del factor |
| Cantidad | Cantidad final a comprar |

Sobre la columna Norma / RIC conviene aclarar cómo se comporta. Cuando el material tiene entrada en el catálogo normativo, la celda no muestra el número del RIC sino el texto "Ver normativa" en azul y subrayado, y al hacer clic salta a la fila exacta de la hoja Base Normativa. Cuando no hay entrada, la celda conserva el texto del RIC ya normalizado, sin enlace.

Las filas de encabezado de sección se distinguen porque tienen la columna Ítem vacía: el programa las fusiona a todo el ancho, las pone en negrita y les aplica fondo gris. Las alturas de fila se ajustan al contenido, contando los saltos de línea, para que los detalles multilínea como los ferrules o los conectores cónicos se vean completos sin tener que agrandar la fila a mano.

Las secciones aparecen en este orden: Empalme, Protecciones, Borneras de conexión, Cableado interior tablero, Terminales ferrul interiores, Canalizaciones, Conductores, Accesorios, Conectores cónicos, Iluminarias, Tornillería, Tarugos, y Sellos y aislación de panel SIP.

### Hoja Base Normativa

Lleva un título fusionado, "BASE NORMATIVA RIC PARA TRAZABILIDAD DE MATERIALES", encabezados en la fila 2 y los datos desde la fila 3, con panel congelado. Son dos columnas, Material o Elemento y Norma, con más de 130 materiales y el listado de artículos, puntos y tablas del RIC que respaldan a cada uno.

En la celda D1 hay un enlace de vuelta, "Volver a Materiales", para no perder la navegación al saltar entre hojas.

Un ejemplo de fila:

```
Interruptor diferencial 10mA agua caliente
RIC N°11 (6, 6.4.3, Tabla Vol.1) / RIC N°07 (7.4.5, 7.6.5.4)
Volumen 1 baño: sensibilidad menor o igual a 10 mA
```

## Requisitos

- Python 3.9 o superior
- pandas 1.3 o superior
- numpy 1.21 o superior
- openpyxl 3.0 o superior
- Terminal con soporte UTF-8, porque el programa imprime caracteres como °, ², Ω
- Excel, LibreOffice Calc o Google Sheets para abrir el resultado

## Instalación

```bash
git clone https://github.com/USUARIO/sistema-instalaciones-electricas.git
cd sistema-instalaciones-electricas

python -m venv .venv
source .venv/bin/activate        # en Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

Si el archivo `requirements.txt` no existe, créalo con este contenido:

```
pandas>=1.3
numpy>=1.21
openpyxl>=3.0
```

En Windows, si la consola muestra caracteres incorrectos en lugar de ² o °, ejecuta `chcp 65001` antes de correr el programa, o usa Windows Terminal.

## Uso

```bash
python informe_electrico.py
```

El programa es interactivo. Va preguntando en orden y no avanza hasta recibir un valor válido. Conviene tener a mano antes de empezar:

- Plano o croquis con el área y el perímetro de cada ambiente
- Cantidad y potencia de luminarias y enchufes por ambiente
- Placa de características de los aires acondicionados, duchas, termos o calefones
- Longitudes de recorrido: transformador a empalme, empalme a tablero, y tablero a cada circuito
- Distancias desde el empalme y desde el tablero hasta sus camarillas de tierra
- Tipo de acometida, aérea o subterránea, y ubicación del empalme, en fachada o en poste independiente

Conviene ejecutarlo desde una carpeta de trabajo dedicada. Como el nombre del Excel incluye fecha y hora, se pueden correr varias iteraciones sin sobrescribir los resultados anteriores.

## Flujo de ejecución

### Parte 1: ambientes y cargas

Se pregunta la cantidad de ambientes y luego, por cada uno:

1. Nombre, área en m² y perímetro en m
2. Material del tabique y del forrado interior, que definen tornillería y tarugos
3. Iluminación: cantidad de luminarias, tipo, montaje y potencia. Si el usuario no conoce la potencia, el programa la estima según tipo y montaje
4. Conmutados: si el ambiente tiene puntos comandados desde dos lugares, se pide cuántas luminarias y las longitudes de viajeros, retorno y troncal
5. Enchufes: cantidad y módulos por enchufe, simple, doble o triple, con potencia total o por módulo. En dormitorios, living, comedor y sala de estar el programa calcula además el mínimo por perímetro, uno cada 8 m. En pasillos pregunta primero si llevan enchufes
6. Componentes especiales: equipos con nombre y potencia. En climatización se puede ingresar la capacidad en BTU/h y el programa hace la conversión

### Parte 2: datos del sistema

Zona húmeda o seca, tipo de canalización embutida o sobrepuesta, material del forrado exterior y cantidad de circuitos.

Antes de preguntar por los circuitos, el programa muestra la superficie total de la vivienda y el mínimo de circuitos que corresponde según esa superficie, 2 bajo los 30 m² y 3 desde ahí. No acepta un valor menor a ese mínimo, y el máximo permitido es 10.

### Parte 2.1: construcción de circuitos

Por cada circuito se pide nombre base, ambientes que lo componen, longitud del recorrido, si existe algún tramo continuo de 20 m o más (dato que define si corresponden cajas de paso) y si es un circuito especial con carga dedicada.

Si el nombre del circuito viene mal escrito, el programa sugiere la corrección más probable en lugar de rechazarlo.

Cuando se detecta un equipo de climatización se abre un cuestionario específico: tecnología inverter u on-off, corriente nominal y máxima de placa, LRA declarado, MOCP del fabricante, tipo de compresor y temperatura de trabajo.

Para agua caliente se pregunta el tipo de equipo (ducha, termo o calefón), la potencia, si queda dentro del Volumen 1 del baño, si requiere tablero de desconexión externo y la temperatura ambiente.

### Parte 3: parámetros eléctricos y empalme

Tensión nominal, factor de potencia, temperatura ambiente, longitud del alimentador entre empalme y tablero, longitud de la acometida entre transformador y empalme, y tipo de acometida.

A continuación se define la geometría del empalme, con validaciones normativas cruzadas que el programa aplica de forma automática:

| Situación | Regla |
|---|---|
| Empalme dentro de 15 m del acceso | Debe instalarse en fachada (RIC 1, 7.2) |
| Empalme fuera de los 15 m | Debe instalarse en estructura independiente (RIC 1, 7.3) |
| Fachada con acometida subterránea | No admite mástil |
| Empalme en fachada | El alimentador solo puede ir en ducto |
| Poste independiente | El alimentador debe ser aéreo o subterráneo, nunca en ducto |

Según la combinación elegida, el programa pide únicamente las longitudes que corresponden: altura de la acometida aérea, tramo subterráneo hasta el medidor, largo del mástil, subida por el poste, y así.

### Parte 3.1: puesta a tierra

Largo de la barra copperweld, 3 m recomendado o 1,5 m si hay restricción de profundidad, y resistividad del terreno para cada una de las dos mallas.

### Parte 3.2: cálculo automático

Sin más intervención del usuario, el programa calcula factor de demanda, empalme, agrupación de diferenciales, secciones de conductor, caídas de tensión y canalizaciones.

### Parte 4: generación del Excel

Se arman los DataFrames, se construye la cubicación completa, se aplica el formato y se abre el archivo.

### Orden interno de cálculo

El orden en que el programa resuelve las cosas no es arbitrario, responde a las dependencias entre los cálculos. Vale la pena entenderlo porque explica por qué el empalme, que es lo primero que uno pensaría en definir, en realidad se calcula casi al final.

1. **Ingreso de datos.** No se calcula nada todavía. Al inicio del script hay variables de empalme puestas en nulo, con un comentario que aclara que existen solo por compatibilidad y que no se usan para calcular.
2. **Armado de circuitos.** Cada circuito queda con su potencia y su longitud.
3. **Corriente y termomagnético de cada circuito.** Se puede hacer en este punto porque el calibre del termomagnético depende únicamente de la corriente del propio circuito, no del empalme.
4. **Factor de demanda.** Recién ahora se conocen todas las cargas, así que se puede separar entre los primeros 3 kW y el resto, y sumar aparte la climatización y el agua caliente.
5. **Calibre del empalme.** Sale de la corriente con factor de demanda más un 10%.
6. **Parámetros derivados del empalme.** El calibre del diferencial general, el tope de suma de termomagnéticos y el texto del omnipolar salen del empalme ya calculado.
7. **Agrupación de diferenciales.** Necesita el tope del paso anterior, por eso va después.
8. **Sección de conductor, caída de tensión y canalización de cada circuito.** Necesita el termomagnético asignado en el paso 3.
9. **Acometida y alimentador.** El alimentador necesita conocer la peor caída de tensión de los circuitos para verificar el límite acumulado de 5%, así que va después del paso 8.
10. **Cubicación de materiales.** Necesita absolutamente todo lo anterior. Se ejecuta ya dentro del bloque de escritura del Excel.
11. **Formato y cierre del archivo.**

Un detalle de implementación que conviene tener presente al modificar el código: la cubicación se llama antes de limpiar los campos auxiliares de los circuitos, los que empiezan con guion bajo. Si se invierte ese orden, la función se queda sin los datos que necesita para contar puntos y materiales.

## Motor de cálculo

### Cantidad mínima de puntos y de circuitos

Antes de calcular nada eléctrico, el programa fija los mínimos de diseño a partir de la geometría de la vivienda.

La cantidad mínima de circuitos depende de la superficie total, que se obtiene sumando el área de todos los ambientes ingresados: 2 circuitos si la vivienda tiene menos de 30 m², y 3 circuitos si tiene 30 m² o más. El programa muestra la superficie y el mínimo antes de preguntar cuántos circuitos se van a armar, y rechaza cualquier valor por debajo de ese piso.

La cantidad mínima de enchufes se calcula por perímetro, a razón de uno cada 8 m, redondeando hacia arriba. Esa regla se aplica en dormitorios, living, comedor y sala de estar. No se aplica en baños ni en pasillos, que reciben tratamiento aparte: en el caso del pasillo, el programa primero pregunta si lleva enchufes y solo entonces consulta la cantidad.

Cocina y lavadero se identifican por el nombre del ambiente y reciben tratamiento propio, con sus enchufes agrupados en circuitos dedicados y con longitud real registrada por ambiente. El programa admite varios ambientes de cocina, numerándolos como cocina 1, cocina 2 y así sucesivamente.

Para la iluminación, la cantidad de conductores por tramo se deriva de cómo quedan comandadas las luminarias. La función `descomponer_interruptores()` convierte el número de luminarias no conmutadas de un ambiente en la combinación de interruptores comerciales que las cubre: 9/12 simple, 9/15 doble y 9/32 triple. Esa combinación determina después el número de conductores del tramo y, con él, el diámetro de la canalización.

### Factor de demanda y empalme (RIC 3)

Las cargas se separan en dos familias. Alumbrado y enchufes usan la Tabla N°3.1, con los primeros 3 kW a factor 1,00 y el resto a factor 0,35. Climatización y agua caliente van a factor 1,00, es decir, corriente plena según el RIC 7.

```
I_con_fd  = I_primeros * 1,00 + I_resto * 0,35 + I_clima + I_agua
I_empalme = I_con_fd * 1,10
```

El calibre se normaliza al comercial inmediatamente superior:

| Corriente calculada | Interruptor |
|---|---|
| hasta 25 A | 25 A |
| hasta 32 A | 32 A |
| hasta 40 A | 40 A |
| hasta 50 A | 50 A |
| hasta 63 A | 63 A |

Sobre 40 A el programa emite un aviso en consola. Está calibrado para empalmes tipo A-9 o S-9, de hasta 40 A. El cálculo continúa, pero recomienda revisar el proyecto de forma manual.

Del calibre del empalme se derivan el diferencial general, el tope de suma de termomagnéticos por diferencial y el texto del interruptor omnipolar:

| Empalme | Diferencial | Suma máxima de TM |
|---|---|---|
| hasta 25 A | 25 A | 25 A |
| hasta 32 A | 40 A | 32 A |
| hasta 40 A | 40 A | 40 A |
| sobre 40 A | 63 A | igual al calibre del empalme |

### Circuitos interiores

El cálculo de sección se hace en cuatro pasos.

Primero, sección mínima por caída de tensión, con un límite de 3% de la tensión nominal:

```
dV_max = V_nominal * 0,03

              2 * L * I * rho
S_min    =   ------------------          rho_cobre = 0,0179 ohm*mm²/m
                   dV_max
```

Segundo, se respetan los mínimos normativos: 1,5 mm² en iluminación y 2,5 mm² en enchufes y circuitos especiales.

Tercero, se sube a la sección comercial siguiente.

Cuarto, se verifica la ampacidad contra la Tabla N°4.4, aislación 70 °C, usando el método A1 si la canalización es embutida y B1 si es sobrepuesta o en ducto. El factor de corrección por temperatura de la Tabla N°4.7 se aplica como multiplicador sobre la ampacidad de tabla, no como divisor de la corriente de diseño:

```
Ic = Iz * ft   debe ser mayor o igual a I_diseño
```

Valores del factor según método de instalación:

| Temperatura | ft (A1 y B1) | ft (E, aéreo) | ft (D1, suelo) |
|---|---|---|---|
| hasta 10 °C | 1,22 | 1,22 | 1,07 |
| hasta 20 °C | 1,12 | 1,12 | 1,00 |
| hasta 30 °C | 1,00 | 1,00 | 0,93 |
| hasta 40 °C | 0,87 | 0,82 | 0,85 |
| hasta 50 °C | 0,71 | 0,65 | 0,76 |
| sobre 55 °C | 0,50 | 0,43 | 0,65 |

El tipo de conductor depende de la zona: THWN-2 en zona húmeda y H07Z1-K, libre de halógenos, en zona seca.

### Alimentador y acometida

La selección es iterativa. Se parte de la sección mínima por caída de tensión y se sube hasta cumplir tres condiciones al mismo tiempo:

- La ampacidad corregida debe superar tanto la corriente de demanda como la corriente del interruptor del empalme
- La caída de tensión del alimentador no debe pasar de 3%
- La caída acumulada del alimentador más el peor circuito no debe pasar de 5%

El método de instalación se deduce del tipo de tramo: E para aéreo, D1 para subterráneo y B1 para ducto. En los tramos subterráneos el factor de corrección se calcula con la temperatura del suelo, no con la ambiente.

La sección mínima de ambos tramos es 4 mm², aunque el cálculo por caída de tensión arroje un valor menor.

Si el programa llega a la sección más grande de la lista sin cumplir las tres condiciones, no falla en silencio: devuelve igualmente esa sección, pero acompañada de una advertencia que indica cuál de los requisitos no se pudo satisfacer.

La acometida se expresa como conductor concéntrico de cobre de dos conductores, y el alimentador como RV-K de cobre de tres conductores, fase, neutro y protección.

Las secciones comerciales consideradas son 1,5; 2,08; 2,5; 3,31; 4; 5,26; 6; 8,37; 10; 13,3; 16; 21,1; 25; 26,7; 33,6; 35; 42,4 y 50 mm².

### Canalizaciones (RIC 4)

| Caso | Tabla utilizada |
|---|---|
| Circuitos interiores embutidos | Tabla N°4.17, diámetro según sección y número de conductores, de 1 a 5 |
| Circuitos sobrepuestos | Cálculo dinámico de canaleta por área real de los conductores |
| Alimentador embutido o en ducto | Tabla N°4.19 en mm², con respaldo en la Tabla N°4.20 en AWG y kcmil |
| Acometida y alimentador subterráneos | Tabla N°4.29, hasta 25 mm² y máximo 3 conductores |
| Alimentador aéreo | Sin canalización |

El número de conductores para iluminación se deduce del tipo de comando: 3 conductores para conmutado 9/24 y para interruptor simple 9/12, 4 para interruptor doble 9/15 y 5 para interruptor triple 9/32. Se toma el peor caso dentro del circuito. Para los ductos del alimentador se aplica un mínimo práctico de 32 mm.

### Climatización (RIC 7)

```
I_base   = corriente máxima de placa, o P / (V * fp) si no viene declarada
I_diseño = I_base * 1,25          (RIC 7.3.4)
S_min    = 2,5 mm²
```

Para elegir el termomagnético, el programa sigue esta lógica. Si el fabricante declara un MOCP, ese valor manda y no se supera, según RIC 5.6.2.3. Si no lo declara, se toma el calibre comercial inmediatamente superior a la corriente de diseño.

La curva se decide según la tecnología del equipo y su corriente de arranque. Los equipos inverter van con curva C: tienen arranque suave y no presentan un LRA real. Los equipos on-off se prueban primero también con curva C, que dispara entre 5 y 10 veces la corriente nominal, y solo si el LRA no alcanza a cubrirse con ningún calibre se pasa a curva D, que dispara entre 10 y 20 veces. El cambio de curva se intenta antes de subir de calibre.

Los calibres disponibles son 6, 10, 16, 20, 25, 32, 40, 50 y 63 A.

Cada equipo recibe circuito y diferencial exclusivos, con sensibilidad de 30 mA, además de sus materiales propios: enchufe 2P+T, caja de derivación cercana, canalización exclusiva, conector cónico, boquilla y prensaestopa cuando la conexión es directa.

### Agua caliente (RIC 7 y RIC 11)

Las duchas eléctricas, los termoelectros y los calefones se tratan como carga resistiva pura:

```
fp       = 1,0
I_nom    = P / V
I_diseño = I_nom
S_min    = 2,5 mm²
```

El factor 1,25 del RIC 7 artículo 7.3.4 no se aplica aquí, porque corresponde a motores y cargas de arranque difícil, no a cargas resistivas.

La sensibilidad del diferencial depende de dónde queda el equipo:

| Ubicación | Sensibilidad | Norma |
|---|---|---|
| Volumen 1, interior de la ducha | 10 mA o menos | RIC 11, 6.4.3 |
| Cualquier otra ubicación | 30 mA o menos | RIC 7, 7.4.5 |

Cuando corresponde, el programa agrega un tablero de desconexión externo ubicado fuera de los Volúmenes 0, 1 y 2 y a la vista del equipo, con su termomagnético bipolar, bornera PE y prensaestopas.

### Termomagnético de cada circuito

El calibre se elige sobre la corriente estimada más un 10% de reserva, según el RIC 10 artículo 5.1.4.1, tomando la protección comercial inmediatamente superior. La lista de calibres depende del tipo de circuito:

| Tipo de circuito | Calibres disponibles | Curva |
|---|---|---|
| Iluminación | 6, 10, 16 A | B |
| Enchufes | 10, 16 A | C |
| Especiales genéricos | 6, 10, 16, 20, 25, 32, 40, 50, 63 A, con mínimo de 16 A | C |
| Climatización | 6 a 63 A, según LRA y MOCP | C o D |
| Agua caliente | 6 a 63 A | C |

Los circuitos especiales nunca llevan un termomagnético menor a 16 A, aunque el cálculo por corriente diera menos. El poder de corte de los circuitos es 6 kA, y el texto que aparece en el informe tiene el formato `1x16A / 6kA / Curva C`.

El disyuntor del empalme va con curva D. Con eso, el reparto de curvas de toda la instalación queda en curva B para iluminación, curva C para enchufes y cargas especiales, y curva D para el empalme, que es el criterio de selectividad que sigue el programa.

### Diferenciales y agrupación

Los circuitos de iluminación, de enchufes y los especiales se agrupan minimizando la cantidad de diferenciales, mediante backtracking con poda. Las restricciones son un máximo de 3 circuitos por diferencial y una suma de calibres de termomagnéticos que no supere el límite derivado del empalme.

Conviene aclarar un punto que suele confundirse: los circuitos especiales, como horno, lavadora o encimera, sí se agrupan junto con los de iluminación y enchufes. Los únicos que quedan siempre en un diferencial exclusivo son los de climatización, por RIC 7 puntos 7.1.2 y 7.4.5, y los de agua caliente.

El texto que aparece en el informe tiene el formato `2X40 30mA / Tipo A`, o con 10 mA cuando corresponde al Volumen 1 de un baño.

### Puesta a tierra (RIC 6)

Se calculan dos mallas independientes: PT1 en el empalme, camarilla N°1, y PT2 en el tablero, camarilla N°2.

```
R_1barra  = rho_terreno / L_barra
n_barras  = redondeo hacia arriba de (R_1barra / 20)
R_final   = R_1barra / n_barras
sep_min   = 2 * L_barra                    (RIC 6, 8.3.2)
L_desnudo = (n_barras - 1) * sep_min
```

El informe indica de forma explícita si cumple o no cumple frente al límite de 20 Ω. Si se requieren más de 6 barras, el programa advierte en consola que conviene encargar un estudio especial de puesta a tierra, según RIC 6 punto 5.1.

### Temperatura del suelo en tramos subterráneos

Cuando la acometida o el alimentador van enterrados, el factor de corrección por temperatura no se toma de la temperatura ambiente sino de la del suelo, que suele ser distinta. El programa detecta que hay tramos subterráneos, avisa cuáles son y pide un valor de temperatura de suelo aparte, según el RIC 4 punto 6.2.5 y la nota de la Tabla N°4.7. Si no se ingresa nada, se usa la temperatura ambiente. Ese valor alimenta la columna D1 del factor de corrección.

## Resumen de fórmulas

El programa escribe esta misma tabla dentro de la hoja Informe, en el bloque 4, de modo que el archivo entregado lleva su propia memoria de cálculo y no hace falta recurrir al código para saber de dónde salió cada número.

### Circuitos

| Fórmula | Expresión | Variables |
|---|---|---|
| Corriente de circuito | I = P / (V × fp) | I=corriente [A], P=potencia [W], V=tensión nominal [V], fp=factor de potencia (1,0 en cargas resistivas) |
| Termomagnético de circuito | In<sub>TM</sub> ≥ I<sub>circ</sub> | Calibres 6/10/16 A en iluminación, 10/16 A en enchufes. RIC 5, Tabla 5.1 |
| Termomagnético especial | In<sub>TM</sub> ≥ 1,10 × I<sub>circ</sub>, mínimo 16 A | 10% de reserva de capacidad. Mínimo 16 A en agua caliente y climatización |
| Sección mínima por caída de tensión | S<sub>min</sub> = (2 × L × I × ρ) / ΔV<sub>max</sub> | S<sub>min</sub> [mm²], L=longitud tramo [m], I=corriente [A], ρ<sub>cu</sub>=0,0179 Ω·mm²/m, ΔV<sub>max</sub>=V × 3% [V] |
| Sección mínima normativa | S ≥ 1,5 mm² en iluminación, S ≥ 2,5 mm² en enchufes | Piso que se aplica aunque el cálculo dé menos. RIC 4 |
| Caída de tensión | ΔV% = (2 × ρ × L × I) / (S × V) × 100 | S=sección conductor [mm²], V=tensión nominal [V]. Límite ≤ 3% |
| Ampacidad corregida | I<sub>c</sub> = I<sub>z</sub> × ft ≥ I<sub>diseño</sub> | I<sub>z</sub>=ampacidad de tabla RIC 4, ft=factor de temperatura, Tabla N°4.7 |
| Verificación de ampacidad | I<sub>c</sub> ≥ I<sub>circ</sub> y I<sub>c</sub> ≥ In<sub>TM</sub> | El conductor debe soportar la carga y el calibre de su protección. RIC 5, Tabla 5.4 |

Una precisión sobre la segunda fila. La tabla que se escribe en el Excel enuncia el termomagnético de circuito como In<sub>TM</sub> mayor o igual a la corriente del circuito, sin más. En el código, el 10% de reserva del RIC 10 artículo 5.1.4.1 se aplica en realidad a todos los circuitos, no solo a los especiales: la diferencia entre ambos casos está en la lista de calibres disponibles, no en el factor. Si en algún momento se actualiza el texto de la tabla dentro del script, conviene reflejarlo también acá.

### Demanda y empalme

| Fórmula | Expresión | Variables |
|---|---|---|
| Factor de demanda | P<sub>d</sub> = P<sub>1</sub> × 1,0 + P<sub>2</sub> × 0,35 | P<sub>1</sub>=el menor entre P<sub>total</sub> y 3000 [W], que son los primeros 3 kW al 100%. P<sub>2</sub>=el excedente sobre 3000 [W], o cero si no lo hay, que va al 35%. RIC 3, artículos 6.1, 6.2 y 6.3 |
| Corriente de empalme | I<sub>emp</sub> = P<sub>d</sub> / (V × fp) | Se normaliza a 25 / 32 / 40 / 50 / 63 A |
| Holgura del empalme | I<sub>sel</sub> = I<sub>emp</sub> × 1,10 | 10% adicional antes de elegir el calibre comercial |
| Sección del alimentador | S<sub>min</sub> = (2 × L × I<sub>dem</sub> × fp × ρ) / ΔV<sub>max</sub> | Debe cumplir I<sub>c</sub> ≥ I<sub>dem</sub>, I<sub>c</sub> ≥ I<sub>emp</sub> y ΔV ≤ 3%. Mínimo 4 mm² |
| Sección de la acometida | S<sub>min</sub> = (2 × L × I<sub>emp</sub> × fp × ρ) / ΔV<sub>max</sub> | Método E si es aérea, D1 si es subterránea. Mínimo 4 mm² |
| Caída de tensión acumulada | ΔV<sub>total</sub> = ΔV<sub>alim</sub> + ΔV<sub>circ</sub> ≤ 5% | Se verifica circuito por circuito contra el alimentador |

### Puesta a tierra, RIC 6

| Fórmula | Expresión | Variables |
|---|---|---|
| Resistencia de una pica vertical | R<sub>1</sub> = ρ / L | ρ=resistividad del terreno [Ω·m], L=largo de la barra [m]. RIC 6, Tabla 6.4 |
| Número de barras necesarias | N = R<sub>1</sub> / 20, redondeado hacia arriba, mínimo 1 | 20 Ω es la resistencia máxima permitida, barras en paralelo. RIC 6, punto 6.1 |
| Resistencia final del sistema | R<sub>final</sub> = R<sub>1</sub> / N | Debe resultar ≤ 20 Ω para cumplir |
| Separación mínima entre barras | d = 2 × L<sub>barra</sub> | RIC 6, punto 8.3.2 |
| Conductor desnudo de unión | L<sub>desnudo</sub> = (N − 1) × (2 × L<sub>barra</sub>) | Cobre desnudo de 16 mm² |

### Cubicación

| Fórmula | Expresión | Variables |
|---|---|---|
| Chicotes por circuito | L<sub>chic</sub> = n<sub>puntos</sub> × 0,15 | n<sub>puntos</sub>=puntos de conexión, 15 cm de cable perdido en cada caja |
| Metros de conductor a comprar | L<sub>compra</sub> = (L<sub>recorrido</sub> + L<sub>chic</sub>) × 1,10, redondeado al metro superior | Los chicotes ya vienen sumados en L<sub>chic</sub> y el 10% cubre el desperdicio de corte |
| Tramos de canalización | n<sub>tramos</sub> = L / L<sub>tira</sub>, redondeado hacia arriba | L<sub>tira</sub>=2 m en canaleta, 3 m en conduit |
| Abrazaderas | n = L<sub>efectiva</sub> / d<sub>abraz</sub>, redondeado hacia arriba | d<sub>abraz</sub>=1,20 m hasta 25 mm, 1,50 m sobre 32 mm. RIC 4, Tabla N°4.24 |
| Cajas de paso | n = L / 20, sin decimales | Solo si existe un tramo continuo de 20 m o más. RIC 4, 7.16.1.13 |
| Cámaras tipo C | n = 0 si L ≤ 20 m, si no n = L / 90, redondeado hacia arriba | Tramos subterráneos. Bajo 20 m se resuelve en forma de U |
| Puestos del tablero | P = P<sub>base</sub> + (n<sub>circ</sub> × 0,25, redondeado hacia arriba) × 3 | P<sub>base</sub>=suma de módulos de todos los componentes, más reserva de ampliación |

## Cubicación de materiales

La función `build_materiales_df()` recorre circuitos, ambientes y la geometría del empalme, y va emitiendo líneas de material. Cada línea incluye la cantidad calculada a partir de conteos reales, el factor K de holgura donde corresponde, las marcas de referencia definidas en el diccionario `marcas`, la indicación de sello SEC y el circuito asociado, de modo que la compra se puede filtrar por partida.

### Metros de conductor y chicotes

Los metros de cable no se calculan solo con la longitud del recorrido. El programa cuenta los puntos de conexión reales de cada circuito, es decir, cada caja donde el cable se corta y se vuelve a empalmar, y suma 15 cm por punto. En un ambiente conmutado, por ejemplo, el primer interruptor aporta 3 chicotes, una fase y dos viajeros, y el segundo aporta otros 3, dos viajeros y un retorno. Los circuitos especiales, de climatización y de agua caliente aportan un chicote base en la caja del equipo.

```
chicotes         = n_puntos * 0,15
L_con_chicotes   = L_recorrido + chicotes
metros a comprar = redondeo hacia arriba de (L_con_chicotes * 1,10)
```

El 10% adicional cubre desperdicio de corte. Es importante notar que estos metros se usan solo para la compra: el cálculo de sección y de caída de tensión se hace con la longitud real de la canalización, porque los chicotes son cable dentro de las cajas y no representan distancia eléctrica.

### Cajas de paso

Se cuenta una caja de paso cada 20 m de canalización, según RIC 4 punto 7.16.1.13. La regla tiene una salvedad importante que el programa respeta: un circuito puede sumar más de 20 m en total sin tener ningún tramo continuo de esa longitud, y en ese caso no corresponde ninguna caja de paso. Por eso, al construir cada circuito se pregunta si existe realmente un tramo continuo de 20 m o más.

### Columna Sello SEC

El valor de la columna no se ingresa a mano, lo decide la función `_requiere_sello_sec()` a partir de la descripción del material. Los conductores y cables, las protecciones, los enchufes, los interruptores y los tableros quedan marcados como SEC. Los consumibles de instalación, como el tubo de estaño, la pasta para soldar, la cinta aislante, el sellador de roscas, la espuma expansiva y el teflón, quedan marcados con un guion, porque no requieren certificación.

### Tramos subterráneos

Cuando la acometida o el alimentador van enterrados, la cubicación agrega las cámaras tipo C de hormigón prefabricado de 440 por 440 mm con tapa de acero diamantado, junto con su marco metálico y sus boquillas de PVC, dimensionadas según la longitud del tramo subterráneo.

### Descripción automática de luminarias

El usuario solo ingresa el tipo de luminaria en texto libre, el montaje y la potencia. La función `desc_luminaria_auto()` interpreta ese texto para reconocer si se trata de un foco o panel, un aplique, un tubo o una ampolleta, deduce la tecnología (LED por defecto, salvo que el texto diga incandescente o fluorescente) y arma la descripción técnica completa que aparece en la cubicación. La misma lógica determina si esa luminaria requiere tubo de estaño y pasta para soldar.

### Conectores cónicos

Los conectores cónicos se cubican por combinación de secciones a empalmar dentro de cada caja, no como un total genérico, de modo que la lista distingue las distintas medidas necesarias.

### Lógica de cálculo de cada material

Esta es la parte que más suele preguntarse al revisar una cubicación: de dónde salió cada cantidad. La tabla siguiente recorre las trece secciones del listado y explica la regla exacta que aplica el programa.

#### Canalizaciones

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Canaleta PVC sobrepuesta | unidad | n = L / 2, redondeado hacia arriba, se vende en tramos de 2 m |
| Conduit PVC embutido | unidad | n = L / 3, redondeado hacia arriba, se vende en tiras de 3 m |
| Unión copla para canaleta | unidad | n = n<sub>tramos</sub> − 1, una por cada junta entre tramos |
| Abrazadera conduit | unidad | n = L<sub>efectiva</sub> / d, redondeado hacia arriba, con d = 1,20 m hasta 25 mm y 1,50 m desde 32 mm. L<sub>efectiva</sub> es el largo ya redondeado a tiras completas |
| Curvas internas 90°, iluminación | unidad | n = 4 × n<sub>ambientes</sub> + n<sub>interruptores</sub> |
| Curvas planas 90°, iluminación | unidad | n = 2 × n<sub>interruptores</sub>, una salida de tablero más dos por caja troncal |
| Curvas internas 90°, enchufes | unidad | n = 3 × n<sub>ambientes</sub> |
| Curvas planas 90°, enchufes | unidad | n = 1 + n<sub>ambientes</sub>, la salida del circuito más el último enchufe de cada ambiente |
| Curva T, enchufes | unidad | n = n<sub>enchufes</sub> − 1, mínimo 0 |
| Salida de caja conduit | unidad | Conteo punto por punto. Parte en 1 por la salida del tablero y va sumando según el rol de cada caja: la caja troncal aporta 3 o 4 según sea la última del ambiente o del circuito, la caja de interruptor aporta 1, salvo el conmutado 9/24 que aporta 4 por sus dos cajas, y cada caja octogonal aporta 1 si es la última de la cadena o 2 si continúa |

#### Conductores

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Conductor de circuito | metro | L<sub>compra</sub> = (L<sub>recorrido</sub> + n<sub>puntos</sub> × 0,15) × 1,10, redondeado al metro superior. Los 15 cm por punto son el chicote perdido en cada caja, y el 10% es desperdicio de corte |
| Conductor de tierra de protección | metro | Misma longitud que el circuito, en color verde |

Vale la pena insistir en un punto: los metros que se compran incluyen chicotes y desperdicio, pero el cálculo de sección y de caída de tensión usa la longitud real de la canalización. Los chicotes son cable dentro de las cajas y no representan distancia eléctrica.

#### Protecciones

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Interruptor general omnipolar | unidad | Siempre 1, con el calibre del empalme |
| Disyuntor termomagnético | unidad | 1 por circuito, con el calibre asignado en el cálculo |
| Interruptor diferencial | unidad | 1 por grupo resultante de la agrupación. Climatización y agua caliente generan siempre un grupo propio |
| Supresor de transiente | unidad | 1 por instalación |
| Protector de sobrevoltaje y corriente | unidad | 1 por instalación |

#### Accesorios: cajas y puntos

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Caja octogonal embutida | unidad | n = total de luminarias, una por centro de iluminación |
| Caja rectangular embutida | unidad | n = cajas troncales, una por grupo de interruptor, más cajas de mecanismo, más enchufes, más cajas adicionales, más cajas de climatización, agua caliente y conexiones fijas, más cajas de paso |
| Caja chuqui sobrepuesta | unidad | n = enchufes + cajas adicionales + interruptores + interruptores de pared + cajas de paso + conexiones fijas + climatización + agua caliente |
| Caja de paso | unidad | n = L / 20, sin decimales, y solo si el circuito tiene un tramo continuo de 20 m o más. Un circuito que suma 30 m en tramos cortos no lleva ninguna |
| Tapa ciega rectangular | unidad | n = cajas troncales + cajas adicionales + climatización + agua caliente + cajas de paso + conexiones fijas. Las cajas que ya quedan tapadas por su propio enchufe no se cuentan |
| Tapa ciega octogonal | unidad | Una por caja octogonal que queda como unión y no recibe luminaria |
| Enchufe | unidad | Uno por punto declarado, agrupado por amperaje y número de módulos |
| Punto de conexión directa | unidad | Se emite en lugar del enchufe cuando el termomagnético del circuito supera 16 A, porque a esa corriente la conexión debe ser fija |
| Portalámpara plafón E27 | unidad | n = total de ampolletas declaradas |

#### Conectores cónicos e insumos de empalme

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Conector cónico | unidad | Se acumula por combinación de color, número y rango de secciones. El color y el número salen de la sección del conductor y de cuántos cables entran al empalme |
| Tubo de estaño | unidad | n = n<sub>gruesas</sub> / 4 + n<sub>finas</sub> / 15, redondeado hacia arriba, mínimo 1. Las conexiones gruesas son las de más de 6 mm² y las de puesta a tierra, las finas son las de focos LED |
| Pasta para soldar | unidad | n = n<sub>tubos</sub> / 4, redondeado hacia arriba, mínimo 1 |
| Cinta autofundente de goma | unidad | n = n<sub>conexiones</sub> × 20 / 300, redondeado hacia arriba, mínimo 1. Se consideran 20 cm por conexión y rollos de 3 m |
| Cinta aislante PVC | unidad | n = cm<sub>total</sub> / 2000, redondeado hacia arriba, mínimo 1. Se consideran 35 cm por conexión de circuito y rollos de 20 m |

#### Iluminarias y componentes de tablero

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Luminaria o foco | unidad | Una por unidad declarada, agrupada por tipo, montaje, tecnología y potencia |
| Tablero de PVC | unidad | Se suman los puestos de todos los componentes, se agrega una reserva de n<sub>circ</sub> × 0,25, redondeado hacia arriba, por 3 puestos para ampliaciones futuras y el total se redondea al tamaño comercial siguiente entre 2, 4, 6, 8, 12, 16, 18, 24, 36, 42, 48, 54, 56 y 72 |
| Riel DIN | unidad | Tira de 1 m, o de 2 m cuando el tablero es de 56 o 72 puestos |
| Barra unipolar verde PE | unidad | Los polos salen del número de circuitos: 4 polos con 1 circuito, 6 con 2 o 3, 8 con 4 o 5, 10 con 6 o 7, 12 con 8 o 9, y 15 desde 10 |
| Barra repartidora bipolar | unidad | Los polos salen del número de diferenciales: 4 polos hasta 3 diferenciales, 7 polos de 4 a 6, y 11 polos de 7 a 10 |
| Barra repartidora tetrapolar | unidad | 1 unidad, más las barras adicionales que exija la cantidad de salidas |
| Luz piloto, portafusible y fusible | unidad | 1 de cada uno por tablero |
| Bornera de conexión | unidad | Una por circuito cuando el diferencial es exclusivo, y una adicional por circuito cuando la instalación supera los 8 circuitos, por organización del cableado de fase |

#### Terminales ferrul

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Terminal ferrul | unidad | Se acumula por sección normalizada a 1,5; 2,5; 4; 6; 10; 16; 25 o 35 mm², y el color queda determinado por esa sección: rojo, azul, naranjo, amarillo, rojo, azul, amarillo y gris respectivamente. Cada interruptor aporta una cantidad fija: el 9/12 aporta 2, el 9/15 aporta 5, el 9/32 aporta 8, y el conmutado 9/24 aporta 6 por grupo, sin importar cuántas luminarias comande |
| Terminal ferrul de acometida | unidad | 2 unidades, fase y neutro |
| Terminal ferrul de alimentador | unidad | 3 simples más 1 doble |
| Terminal de compresión tipo ojo | unidad | Uno por conexión de puesta a tierra a barra o carcasa |

#### Cableado interior del tablero

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Conductor interior de tablero | metro | La longitud base de cada tramo se estima con la geometría del tablero, contando el alto dos veces y el ancho una vez, y al total se le aplica un 10% de holgura |

#### Tornillería y tarugos

El total de tornillos se arma sumando los puntos de fijación reales de la instalación:

| Elemento a fijar | Tornillos que aporta |
|---|---|
| Cada abrazadera de conduit | 2 |
| Cada metro de canaleta sobrepuesta | 1 |
| Cada caja | 4 |
| Cada tapa ciega | 2 |
| Cada luminaria de montaje sobrepuesto | 2 |
| Riel DIN | 7 si la tira es de 1 m, 14 si es de 2 m |
| Tablero | 6, 8 o 10 según su cantidad de puestos |

Las luminarias embutidas no aportan tornillos, porque no se fijan de esa manera. El total se reparte después entre los distintos tipos de tornillo, ponderando por la cantidad real de cajas de cada ambiente, de modo que un ambiente con muchos puntos pesa más en el reparto que uno con pocos. El tipo de tornillo y la necesidad de tarugo dependen del material del tabique y del forrado declarados.

#### Sellos y aislación de panel SIP

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Espuma expansiva de poliuretano | unidad | n = n<sub>cajas</sub> / 10, redondeado hacia arriba, mínimo 1. Se estima un tubo de 750 ml cada 10 cajas embutidas. El tablero no requiere espuma |

#### Empalme, acometida y alimentador

| Material | Unidad | Cómo se calcula la cantidad |
|---|---|---|
| Unidad de medida monofásica | unidad | 1 |
| Caja de empalme metálica | unidad | 1 |
| Disyuntor del empalme | unidad | 1, con el calibre calculado |
| Portafusible de loza con fusible | unidad | 1 |
| Cable de acometida | metro | n = L<sub>transformador a empalme</sub>, redondeado al metro superior |
| Alimentador RV-K | metro | n = L<sub>empalme a tablero</sub>, redondeado al metro superior |
| Tubo conduit galvanizado | unidad | n = L / 3, redondeado hacia arriba, sumando el tramo de acometida y el de puesta a tierra |
| Conduit PVC del alimentador | unidad | n = L / 3, redondeado hacia arriba |
| Abrazadera del conduit del alimentador | unidad | n = L / d, redondeado hacia arriba, con d = 1,20 m hasta 25 mm y 1,50 m desde 32 mm |
| Abrazadera tipo caddy | unidad | Una por tramo de tubo galvanizado instalado |
| Cabeza de servicio, cáncamo y granpa | unidad | 1 de cada uno cuando la acometida es aérea |
| Mordaza para alimentador aéreo | unidad | 1 cuando el alimentador es aéreo |
| Conector HUB galvanizado | unidad | Uno por entrada de tubo a caja metálica, contando acometida y alimentador por separado |
| Terminal PVC conduit con 2 tuercas | unidad | Uno por entrada de conduit PVC, sumando acometida y alimentador |
| Sellador de roscas con teflón | unidad | Un envase cada cierta cantidad de uniones roscadas |
| Cámara tipo C con marco | unidad | n = 0 si el tramo subterráneo mide 20 m o menos, porque se resuelve en forma de U. Sobre 20 m, n = L / 90, redondeado hacia arriba |
| Boquilla PVC de cámara | unidad | 2 por cámara, entrada y salida |
| Caja de paso del alimentador | unidad | n = L / 20, sin decimales, solo cuando el alimentador va en ducto |
| Barra copperweld con conector | unidad | n = número de barras calculado para cada malla, PT1 y PT2 |
| Camarilla PVC naranjo | unidad | n = barras de PT1 + barras de PT2 |
| Conductor desnudo Cu 16 mm² | metro | L = (N − 1) × 2 × L<sub>barra</sub>, por cada malla |
| Poste de madera o pilar metálico | unidad | 1 cuando el empalme va en estructura independiente |

## Base normativa

El núcleo normativo del programa es la constante `BLOQUES_NORMATIVA`, una lista de tuplas con el material y sus artículos del RIC, con más de 130 entradas organizadas por reglamento de origen.

De esa lista se derivan dos cosas: la hoja Base Normativa completa, con una fila por material, y los hipervínculos de la columna Norma / RIC de la hoja Materiales.

La función `normalizar_ric_materiales()` limpia la referencia para mostrar solo el reglamento principal, de modo que "RIC 4.7.2" se muestra como "RIC 4". Cuando un material viene marcado como SEC sin un RIC explícito, la función infiere el reglamento aplicable a partir de palabras clave de la descripción.

Reglamentos cubiertos:

| RIC | Materia |
|---|---|
| RIC N°01 | Empalmes |
| RIC N°02 | Tableros eléctricos |
| RIC N°03 | Alimentadores y demanda |
| RIC N°04 | Canalizaciones, conductores y accesorios |
| RIC N°05 | Protecciones |
| RIC N°06 | Puesta a tierra |
| RIC N°07 | Equipos de climatización y agua caliente |
| RIC N°10 | Instalaciones en viviendas |
| RIC N°11 | Recintos con tina o ducha |

## Estructura del código

El programa es un único script de alrededor de 11.300 líneas, organizado de arriba hacia abajo en cinco capas. Este es el mapa del archivo, con las líneas aproximadas de cada bloque:

```
informe_electrico.py
|
+-- CAPA 1: BASE NORMATIVA                                    lineas 1 - 553
|   |
|   +-- Cabecera, comentario de orientacion e imports              1 - 35
|   +-- BLOQUES_NORMATIVA (catalogo de 130+ materiales)           42 - 138
|   +-- normalizar_ric_materiales()                              151 - 190
|   +-- aplicar_base_normativa_e_hipervinculos()                 192 - 553
|         escribe la hoja normativa y enlaza cada material
|
+-- CAPA 2: TABLAS Y FUNCIONES DE CALCULO                   lineas 554 - 1354
|   |
|   +-- conduit_por_tabla()  Tabla 4.17                          558 - 589
|   +-- canalizacion_recomendada_por_conductores()               591 - 604
|   +-- n_conductores_iluminacion_para_circuito()                605 - 660
|   +-- binpack_items()  reparto de cargas                       661 - 682
|   +-- CAPA 3: ENTRADA POR CONSOLA                              683 - 774
|   |     pedir_longitud_sub(), pedir_float_positivo(),
|   |     pedir_float_opcional(), limpiar_nombre_circuito(),
|   |     sugerir_nombre_circuito()
|   +-- canalizacion_recomendada()  canaleta por area            775 - 849
|   +-- resumen_items_por_ambiente(), parse_in_tm()              850 - 918
|   +-- descomponer_interruptores()  9/12, 9/15, 9/32            919 - 937
|   +-- _get_amb_row(), _lum_necesita_estano()                   938 - 960
|   +-- desc_luminaria_auto()                                    961 - 1028
|   +-- Tablas de ampacidad y factor de temperatura              1029 - 1155
|   +-- seleccionar_alimentador()                                1156 - 1206
|   +-- seleccionar_acometida()                                  1207 - 1246
|   +-- Tablas de ducto 4.19, 4.20, 4.29                         1247 - 1313
|   +-- ducto_nominal_tablas()                                   1314 - 1354
|
+-- CAPA 4: CUBICACION                                     lineas 1358 - 7650
|   |
|   +-- build_materiales_df()
|         diccionario de marcas                                       1410
|         bloques internos separados por
|         # ===== NOMBRE SECCION =====
|         ORDEN_SECCIONES y renumeracion final               7582 - 7650
|
+-- CAPA 5: SCRIPT PRINCIPAL                              lineas 7651 - 11310
    |
    +-- Parte 1  Ambientes y cargas                          7660 - 8213
    +-- Parte 2  Datos del sistema                           8214 - 8246
    +-- Helpers de circuitos                                 8247 - 8287
    |     normaliza_seleccion_ambientes(), add_circuito()
    +-- Climatizacion  RIC 7                                 8288 - 8610
    |     calcular_circuito_climatizacion()
    |     ingresar_equipo_climatizacion_inline()
    +-- Agua caliente  RIC 7 y 11                            8611 - 8804
    |     calcular_circuito_agua_caliente()
    |     ingresar_equipo_agua_caliente_inline()
    +-- Parte 2.1  Construccion de circuitos                 8805 - 9374
    +-- Parte 3  Tension, fp, temperatura y empalme          9387 - 9544
    +-- Parte 3.1  Puesta a tierra                           9545 - 9776
    |     _pedir_resistividad(), _calcular_pt()
    |     asignacion de termomagnetico por circuito
    +-- Diferenciales y empalme automatico                   9777 - 10145
    |     optimizar_agrupacion(), parametros_desde_empalme()
    |     factor de demanda y calibre de empalme
    +-- Parte 3.2  Seccion y caida de tension                10146 - 10277
    +-- Parte 3.3  Armado de DataFrames                      10278 - 10358
    +-- Parte 4  Escritura y formato del Excel              10359 - 11310
          cuadro de proteccion general, tabla de formulas,
          calculo de acometida y alimentador,
          llamada a build_materiales_df(),
          formato de las tres hojas y apertura del archivo
```

Las capas 2 y 3 aparecen entrelazadas en el archivo: las funciones que preguntan por consola están intercaladas entre las de cálculo, no en un bloque propio. Se listan por separado porque cumplen roles distintos.

**1. Base normativa.** La constante `BLOQUES_NORMATIVA`, la función `normalizar_ric_materiales()` y `aplicar_base_normativa_e_hipervinculos()`, que escribe la hoja y enlaza los materiales.

**2. Tablas y funciones de cálculo.** `conduit_por_tabla()`, `canalizacion_recomendada()`, `factor_temperatura_ft()`, `seleccionar_alimentador()`, `seleccionar_acometida()`, `ducto_nominal_tablas()`, `binpack_items()` y `descomponer_interruptores()`. Son funciones puras: reciben datos y devuelven un resultado, por lo que se pueden leer de forma aislada.

**3. Entrada por consola.** `pedir_float_positivo()`, `pedir_float_opcional()`, `pedir_longitud_sub()`, `limpiar_nombre_circuito()` y `sugerir_nombre_circuito()`, esta última con corrección difusa de errores de tipeo mediante `difflib`.

**4. Cubicación.** `build_materiales_df()`, la función más extensa del programa, dividida internamente en bloques separados por comentarios con el formato `# ===== NOMBRE SECCIÓN =====`.

**5. Script principal.** Código a nivel de módulo que se ejecuta de corrido, dividido en las partes descritas en el flujo de ejecución.

### Funciones principales

| Función | Responsabilidad |
|---|---|
| `build_materiales_df()` | Construye la cubicación completa |
| `seleccionar_alimentador()` | Sección del alimentador por caída de tensión, ampacidad y caída acumulada |
| `seleccionar_acometida()` | Sección de la acometida entre transformador y empalme |
| `factor_temperatura_ft()` | Factor de corrección por temperatura de la Tabla N°4.7, según método |
| `ducto_nominal_tablas()` | Diámetro de ducto según Tablas N°4.19, N°4.20 y N°4.29 |
| `conduit_por_tabla()` | Diámetro de conduit de circuitos según Tabla N°4.17 |
| `calcular_circuito_climatizacion()` | Termomagnético, curva, diferencial y conductor de un equipo de clima |
| `calcular_circuito_agua_caliente()` | Lo mismo para carga resistiva, con la lógica de volúmenes de baño |
| `optimizar_agrupacion()` | Backtracking que minimiza la cantidad de diferenciales |
| `_calcular_pt()` | Número de barras, resistencia y separación de la malla de tierra |
| `binpack_items()` | Distribución de cargas en circuitos con First-Fit Decreasing |
| `parametros_desde_empalme()` | Deriva diferencial, tope de TM y omnipolar del calibre del empalme |
| `descomponer_interruptores()` | Convierte N luminarias en una combinación de interruptores 9/12, 9/15 y 9/32 |
| `sugerir_nombre_circuito()` | Corrige nombres mal escritos |

### Convenciones internas

Los campos de los diccionarios de circuito que empiezan con guion bajo, como `_In_TM`, `_items` o `_es_climatizacion`, son auxiliares de cálculo y se eliminan antes de exportar a Excel.

En la cubicación, las filas con la columna Ítem vacía y la Descripción técnica llena son encabezados de sección, y se renderizan fusionadas y con fondo gris.

El orden final de las secciones lo fija la lista `ORDEN_SECCIONES`. Los ítems se renumeran después de reordenar, para que la numeración quede correlativa.

## Personalización

### Agregar un material nuevo

Primero se añade la entrada en `BLOQUES_NORMATIVA`:

```python
("Nombre exacto del material", "RIC 4 (5.12.4, 7.1.3)"),
```

Después se emite desde `build_materiales_df()`:

```python
add_row(
    desc="Nombre exacto del material 20mm",
    marcas_txt=marcas.get("Categoría", ""),
    norma="RIC 4",
    circuito=nombre_circuito,
    unidad="u",
    k=1,
    longitud_m=f"{cantidad_calculada} unid",
    cantidad=cantidad_calculada,
)
```

Los ocho argumentos son obligatorios, no tienen valor por omisión. `longitud_m` es la base de cálculo que se muestra antes de aplicar el factor, y admite texto libre, por ejemplo `"12 unid"` o `"25.40 m"`. El campo `norma` acepta el texto `"SEC"`, que es lo que hace que la columna Sello SEC quede marcada.

El texto de `desc` debe coincidir con la entrada de `BLOQUES_NORMATIVA` para que el hipervínculo se genere de forma correcta.

Si el material inaugura una partida nueva, primero hay que abrirla con `add_section("Nombre de la sección")` y agregar ese nombre a la lista `ORDEN_SECCIONES`. Una sección que quede sin filas no se escribe en el Excel.

### Cambiar las marcas sugeridas

Se edita el diccionario `marcas`, al inicio de `build_materiales_df()`:

```python
marcas = {
    "Protecciones": "Legrand, Schneider Electric, Bticino",
    "Conductores":  "Madeco, Cosesa, Revi",
}
```

### Ampliar las tablas

Las tablas de ampacidad y de ducto son diccionarios literales, así que se extienden agregando filas:

```python
DUCTO_N419_MM2[300.0] = {1: 75, 2: 100, 3: 125, 4: 150, 5: 175}
```

### Cambiar el orden de las secciones del Excel

Se reordena la lista `ORDEN_SECCIONES` dentro de `build_materiales_df()`.

### Subir el tope de circuitos

En la Parte 2 se modifica la validación `if cantidad_circuitos > 10`. Hay que tener presente que el dimensionamiento del tablero, riel DIN, barras y espacios, está calibrado para ese rango.

## Limitaciones

### Alcance de la instalación

- Solo monofásico 220 V. No contempla instalaciones trifásicas.
- Empalme máximo de 40 A, tipo A-9 o S-9. Si la demanda calculada supera ese valor, el programa emite un aviso y recomienda revisar los cálculos de forma manual.
- Hasta 10 circuitos.
- Todos los cálculos usan cobre, con rho igual a 0,0179 Ω·mm²/m. No contempla aluminio.

### Alcance de las tablas implementadas

- El diámetro del conduit PVC de los circuitos interiores está implementado hasta conductores de 16 mm² y 5 conductores por ducto. Para secciones mayores la función devuelve el máximo de la tabla.
- El diámetro del conduit PVC de la canalización subterránea está implementado hasta conductores de 25 mm² y un máximo de 3 conductores.

### Criterios de cálculo

- La selectividad está considerada según lo que indica la normativa, con curva B en iluminación, curva C en enchufes y curva D en el empalme. No está diseñada mediante curvas gráficas de disparo.
- Cuando el usuario no conoce la potencia de una luminaria, se usa un valor típico según tipo y montaje.
- La cubicación entrega cantidades, no valorización. No incluye precios.

### Climatización y corriente de arranque

El LRA, corriente de rotor bloqueado o *Locked Rotor Amps*, del compresor del aire acondicionado es lo que decide si hay que subir el calibre de la protección.

Los equipos inverter usan curva C, porque tienen arranque suave y disparan entre 5 y 10 veces la corriente nominal.

Los equipos on/off se prueban primero con curva C. Solo si el LRA no alcanza a cubrirse con ningún calibre se pasa a curva D, que dispara entre 10 y 20 veces la corriente nominal: el equipo tiene un arranque brusco y debe tolerar más corriente sin que el termomagnético se dispare antes de tiempo.

Este criterio no reemplaza un estudio de arranque completo para motores grandes.

### Operación

- Sin interfaz gráfica. Toda la interacción es por consola.
- Los datos se ingresan por consola en cada ejecución. No hay guardado ni recarga de proyectos, así que un error advertido al final obliga a reiniciar.

### Marco normativo cubierto

El programa implementa los RIC N°1, 2, 3, 4, 5, 6, 7, 10 y 11. Cualquier materia regulada por un RIC fuera de esa lista queda fuera del alcance de la herramienta.

## Problemas frecuentes

**ModuleNotFoundError: No module named 'openpyxl'**

Falta la dependencia de escritura de Excel. Se instala con `pip install openpyxl`.

**La consola muestra caracteres incorrectos**

Es un problema de codificación en Windows. Se ejecuta `chcp 65001` antes de correr el programa, o se usa Windows Terminal, que ya trabaja en UTF-8.

**ValueError: could not convert string to float**

Se ingresó texto donde se esperaba un número, o se usó coma decimal en un campo que espera punto. Hay que escribir 2.5 y no 2,5, salvo en los campos donde el programa acepta explícitamente ambos formatos.

**El Excel no se abre solo al terminar**

No es un error. El archivo ya quedó guardado en el directorio de trabajo. La apertura automática depende del sistema operativo y falla en silencio en entornos sin escritorio.

**PermissionError al escribir el archivo**

Hay un Excel generado antes que sigue abierto y bloqueado. Se cierra y se vuelve a ejecutar. Como el nombre incluye fecha y hora, esto normalmente no debería ocurrir.

**Advertencia de más de 6 barras de puesta a tierra**

La resistividad ingresada es muy alta para alcanzar los 20 Ω con barras verticales. Conviene revisar el valor de resistividad, considerar barras de 3 m si se usaron de 1,5 m, o encargar un estudio especial de puesta a tierra.

**El programa rechaza la combinación de empalme y alimentador**

No es un error del programa, es una validación normativa. La tabla de reglas está en [Parte 3: parámetros eléctricos y empalme](#parte-3-parámetros-eléctricos-y-empalme).

## Trabajo pendiente

- Guardar y recargar proyectos en JSON o YAML, para no reingresar todo
- Modo no interactivo, leyendo los datos desde un archivo de entrada
- Interfaz gráfica, con Tkinter o con Streamlit
- Soporte trifásico y empalmes sobre 63 A
- Valorización, integrando una lista de precios por proveedor
- Generación del diagrama unifilar
- Tests automatizados sobre las funciones de cálculo
- Separar el script en módulos: normativa, cálculo, cubicación y escritura del Excel
- Exportación del informe a PDF

## Contribuir

Las contribuciones son bienvenidas, en especial de instaladores autorizados que detecten discrepancias con el RIC vigente.

```bash
git checkout -b feature/nombre-descriptivo
git commit -m "Agrega Tabla N°4.21 para conductores en bandeja"
git push origin feature/nombre-descriptivo
```

Convenciones del proyecto:

- Los comentarios van en español y explican el porqué normativo, no solo lo que hace la línea
- Toda constante normativa lleva su referencia en un comentario, por ejemplo `# RIC 4, Tabla N°4.7`
- Los bloques de `build_materiales_df()` se separan con `# ===== NOMBRE SECCIÓN =====`
- Un material nuevo siempre viene acompañado de su entrada en `BLOQUES_NORMATIVA`
- No se rompe la compatibilidad de las columnas del Excel sin discutirlo antes en un issue

Para reportar un error de cálculo conviene incluir los datos de entrada que se usaron, el resultado que entregó el programa, el resultado esperado y la cita del artículo del RIC que lo respalda.

## Advertencia de uso

Este programa es una herramienta de apoyo al diseño y a la cubicación. No reemplaza el criterio, el cálculo ni la responsabilidad profesional de un instalador eléctrico autorizado por la SEC.

Los resultados deben ser revisados, validados y firmados por un profesional competente antes de ejecutar cualquier instalación o de presentar una declaración TE1 ante la Superintendencia de Electricidad y Combustibles.

La normativa RIC se actualiza de forma periódica. Conviene verificar siempre que las tablas y artículos implementados correspondan a la versión vigente al momento del proyecto.

Los autores y colaboradores no asumen responsabilidad por daños, pérdidas o incumplimientos normativos derivados del uso de esta herramienta.

## Autor

Justin Kuhl Abarca

## Licencia

Este proyecto está licenciado bajo la Licencia Apache, Versión 2.0. Puedes obtener una copia en:

http://www.apache.org/licenses/LICENSE-2.0

```
Copyright 2026 Justin Kuhl Abarca

Licencia bajo la Licencia Apache, Versión 2.0 (la "Licencia");
no puede utilizar este archivo salvo en cumplimiento de la Licencia.
Puede obtener una copia de la Licencia en:

    http://www.apache.org/licenses/LICENSE-2.0

Salvo que lo exija la legislación aplicable o se acuerde por escrito,
el software distribuido bajo la Licencia se distribuye "TAL CUAL",
SIN GARANTÍAS NI CONDICIONES DE NINGÚN TIPO, ya sean expresas o implícitas.
Consulte la Licencia para conocer el texto específico que rige los permisos
y limitaciones bajo la Licencia.
```

El texto completo se encuentra en el archivo [LICENSE](LICENSE) del repositorio.
