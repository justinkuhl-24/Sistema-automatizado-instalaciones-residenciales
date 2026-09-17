# Guía de uso - Informe Eléctrico

Esta guía tiene dos partes: primero cómo instalar todo lo necesario para correr el programa, y después todas las preguntas que hace, para que llegues con los datos listos antes de empezar.

## 1. Instalación

1. **Instalar Visual Studio Code:** desde `code.visualstudio.com`, con las opciones por defecto. Tutorial: https://www.youtube.com/watch?v=X_Z7d04x9-E
2. **(Opcional) Cambiar VS Code a español:** si se abrió en inglés, ve al ícono de extensiones (barra izquierda), busca "spanish" y selecciona "Spanish Language Pack for Visual Studio Code". Presiona Install. Abajo a la derecha va a aparecer una pestaña; presiona "Change Language and Restart" para que se reinicie en español.
3. **Crear y abrir la carpeta del proyecto:** en VS Code ve a Archivo → Abrir carpeta. En la ventana que se abre, elige dónde te acomode (Escritorio, Descargas, etc.), haz clic derecho → Nuevo → Carpeta, ponle un nombre y selecciónala.
4. **Instalar la extensión de Python en VS Code:** ve al ícono de extensiones (barra izquierda), busca "Python" (la de Microsoft) y dale clic a Instalar.
5. **Instalar Python:** entra a `python.org/downloads` (tutorial: https://www.youtube.com/watch?v=3Dcr-xCTvwg) y descarga la versión más reciente e instálala. En Windows, marca la casilla "Add Python to PATH" antes de instalar (si no, después la terminal no reconoce el comando python ni tampoco PIP).
6. **Seleccionar el intérprete de Python:** haz clic en el ícono del engranaje (Administrar, abajo a la izquierda), selecciona "Paleta de comandos", escribe "Select Interpreter" y presiona donde salga "Python: Select Interpreter". Elige el que dice "Python 3... (Recomendado)".
7. **Descargar el archivo del programa:** ve al repositorio de GitHub y descarga el archivo `.py`. Guárdalo dentro de la carpeta que creaste en el paso 3.
8. **Reiniciar el computador:** esto es necesario para que el sistema reconozca el PATH de Python y se pueda instalar pip.
9. **Abrir una terminal:** ya reiniciado el PC, abre VS Code, ve al Explorador (barra izquierda) y ubica la carpeta con el archivo `.py` que guardaste en el paso 7. Haz clic en el archivo para posicionarte en él, luego ve a los tres puntitos (···) al lado de "Ejecutar" → Terminal → Nueva terminal.
10. **Verificar Python e instalar las librerías:** en esa misma terminal escribe `py` y presiona Enter. Debería aparecer algo como "Python 3.... on win32 ...", eso significa que está bien instalado. Luego, en la misma terminal, escribe:

    ```
    pip install openpyxl
    ```

    y al terminar:

    ```
    pip install pandas openpyxl
    ```
11. **Listo:** el programa ya está instalado y listo para ser utilizado.

> **Mac / Linux:** el comando `pip install ...` es el mismo, pero si no se reconoce prueba con `pip3` (y `python3` en vez de `py`). En Mac y Linux normalmente no hace falta reiniciar el computador para que se reconozca el PATH.

## 2. Cómo ejecutar el programa

1. Para usar el código primero hay que posicionarse en el código y tener una terminal abierta.
2. Presiona el botón ▶ (Run) arriba a la derecha.
3. Se abre una terminal abajo: ahí el programa va preguntando y tú escribes la respuesta y presionas Enter.
4. Si te equivocas, el programa te avisa y te vuelve a preguntar lo mismo, no se cierra.
5. Al final arma un Excel y trata de abrirlo solo. Queda guardado en la misma carpeta del proyecto.

> Antes de correrlo, ten a mano un plano o croquis con las medidas de cada ambiente (m² y perímetro), y las fichas técnicas de los equipos grandes (aire acondicionado, calefón, etc.).

## 3. Todas las preguntas que hace el programa

El orden es: primero los ambientes de la casa, después datos generales, después arma los circuitos uno por uno, y al final el empalme, la tierra y los diferenciales.

### A. Por cada ambiente (living, cocina, dormitorios, etc.)

| Pregunta | Detalle |
|---|---|
| Cantidad de ambientes | Se pregunta una sola vez, antes de empezar a recorrerlos |
| Nombre del ambiente | Ej: Living, Cocina, Dormitorio 1, Pasillo |
| Dimensión en m² | Superficie del ambiente |
| Perímetro en metros | Define el mínimo de enchufes exigido |
| Material del tabique | Madera / Metalcon / Panel SIP |
| Forrado interior | Volcanita / Madera |
| Cantidad de luminarias | Mínimo 1 si el ambiente es ≤10 m², mínimo 2 si es mayor |
| Por cada luminaria: tipo, montaje y potencia | Tipo (foco LED, tubo, ampolleta, etc.), montaje (embutida/sobrepuesta), potencia en W (si no la sabes, se asignan 100W automático) |
| ¿Tiene conmutado (2 puntos)? | Si dices que sí, además pregunta cuántas luminarias son conmutadas y 4 longitudes de ese grupo conmutado (viajeros entre interruptores, retorno interruptor→lámpara, fase caja→primer interruptor, troncal→primera caja octogonal) |
| Por cada grupo de interruptor normal (no conmutado) | Las luminarias que no quedaron conmutadas (todas, si respondiste "no" en conmutado) se agrupan de a 1, 2 o 3 por interruptor. Por cada grupo pregunta: longitud troncal→interruptor; si el grupo controla 2 o más luminarias, también troncal→primera caja octogonal; si controla 3, además la distancia entre cada par de cajas octogonales seguidas |
| ¿Tiene componentes especiales enchufados? | Ej: horno, microondas, aire acondicionado (si es aire, pide BTU/h en vez de W) |
| Cantidad de enchufes comunes | Cocina mínimo 3, lavadero mínimo 1, dormitorio/living/comedor exige 1 doble o triple cada 8m de perímetro |
| Por cada enchufe: módulos y potencia | 1=simple, 2=doble, 3=triple. Si dices que conoces la potencia total, además pregunta la potencia de cada módulo por separado. Si no la sabes, se asignan 250W por norma |

> Si el ambiente es un pasillo, primero pregunta si tiene enchufes o no (puede no tener ninguno). Solo si dices que sí, pregunta cuántos.

### B. Datos generales del sistema (una sola vez)

| Pregunta | Detalle |
|---|---|
| Zona | "húmeda" o "seca" (cambia el tipo de cable) |
| Tipo de canalización | "embutida" o "sobrepuesta" |
| Material del forrado exterior de la casa | Fibrocemento / Madera / Siding PVC / Siding metálico |
| Cantidad de circuitos | Máximo 10, mínimo 2 o 3 según el área total de la casa |

### C. Por cada circuito

Primero se piden estos datos para cualquier circuito:

| Pregunta | Detalle |
|---|---|
| Nombre del circuito | Ej: iluminacion, enchufes generales, cocina-encimera, aire acondicionado, ducha eléctrica |
| Ambientes que componen el circuito | Elegidos de la lista de ambientes ya ingresados |
| Longitud real de la canalización en metros | Del tablero hasta el circuito (o hasta el equipo, si es agua caliente) |
| ¿Existe un tramo continuo ≥20m? | Solo se pregunta si la longitud ya es 20m o más |
| Longitud real de cocina o lavadero (si se mezclan con otros ambientes) | Si un circuito de enchufes junta cocina/lavadero con otros ambientes, pregunta la longitud real de cada cocina/lavadero por separado (van a su propio circuito dedicado), y también la del resto de ambientes |
| Longitud real de cada subcircuito (si la potencia es demasiado alta) | Si un circuito de enchufes/iluminación/especial junta demasiada potencia para un solo interruptor, el programa lo divide en subcircuitos y pregunta la longitud real de cada uno por separado |

> Si el nombre del circuito no se reconoce como iluminación, enchufes o climatización, el programa pregunta "¿quisiste decir...?" sugiriendo el más parecido.

Después, según el nombre que le pusiste, sigue una de estas rutas:

**Si es climatización (aire acondicionado, split):**

| Pregunta | Dónde sacar el dato |
|---|---|
| Tecnología: Inverter u On/Off | Ficha técnica |
| Corriente nominal y máxima en modo frío [A] | Placa del equipo |
| Potencia absorbida nominal [W] | Placa del equipo |
| Factor de potencia | Placa (si no aparece, usar 0.95) |
| Tipo de compresor (solo On/Off) | Rotativo o pistón |
| ¿La placa exterior indica LRA? | Si tienes el dato lo ingresas (si el valor no tiene sentido comparado con la corriente nominal, te pregunta si quieres reingresarlo o dejar que el programa lo estime); si no lo tienes, el programa lo estima |
| ¿Indica un MOCP (protección máxima)? | Placa o ficha técnica |
| ¿La temperatura del recinto supera 30°C? | Si sí, pide la temperatura máxima |

**Si es agua caliente (ducha, termo, calefón):**

| Pregunta | Detalle |
|---|---|
| Tipo de equipo | Ducha eléctrica / Termoelectro / Calefón eléctrico / Otro |
| Potencia del equipo [W] | Ej: 3300, 4400, 5500 |
| ¿Está dentro del "Volumen 1" del baño? | Solo si está en un baño y no es ducha |
| ¿El equipo trae interruptor incorporado? | Si no, se agrega un tablero externo |
| ¿La temperatura del recinto supera 30°C? | Si sí, pide la temperatura máxima |

> La ducha eléctrica solo se acepta si el ambiente elegido es un baño.

**Si no es clima ni agua caliente:**

| Pregunta | Detalle |
|---|---|
| ¿Este circuito es ESPECIAL? | Especial = horno, lavadora, encimera, etc. |
| Si es especial: qué componentes van en este circuito | Se elige por número, de los que ya ingresaste por ambiente |
| Si es iluminación: qué luminarias van en este circuito | Se elige por número |
| Si es de enchufes: qué enchufes van en este circuito | Se elige por número |

### D. Datos eléctricos generales (una sola vez)

| Pregunta | Ejemplo |
|---|---|
| Tensión nominal [V] | 220 |
| Factor de potencia | 0.92 |
| Temperatura ambiente típica de la zona [°C] | 25 |
| Longitud del alimentador (Empalme → Tablero) [m] | |
| Longitud de la acometida (Transformador → Empalme) [m] | |

### E. Empalme y acometida

Esta parte cambia bastante según lo que respondas. Ten a mano un croquis con las distancias entre el empalme, el poste (si hay), el tablero y las puestas a tierra.

| Pregunta | Detalle |
|---|---|
| Tipo de acometida | Aérea o subterránea |
| ¿El empalme está dentro de 15m del acceso a la propiedad? | Define si va en fachada o en estructura independiente |
| ¿Se instalará en fachada o en estructura independiente? | Debe calzar con la respuesta anterior |
| Si es independiente: material del poste | Madera o metálico |
| Si es independiente + acometida aérea: altura del tramo aéreo [m] | |
| Si es independiente + acometida subterránea: tramo enterrado hasta el medidor [m] | |
| ¿La acometida requiere mástil? (solo en fachada) | No se permite si la acometida es subterránea |
| Si requiere mástil: largo del mástil [m] | |
| Tipo de alimentador | Aéreo / Subterráneo / En ducto. En fachada solo se permite "en ducto" (salvo que ya haya mástil); en independiente no se permite "en ducto" |
| Si independiente + acometida aérea + alimentador subterráneo: tramo enterrado desde la salida del alimentador hasta el medidor [m] | |
| Si independiente + alimentador aéreo: tramo desde la llegada del alimentador aéreo en la casa hasta el TDA (tablero) [m] | Se preguntan las 2 longitudes seguidas |
| Si independiente + alimentador aéreo: tramo de subida del alimentador por el poste [m] | |
| Si independiente + alimentador subterráneo: tramo enterrado desde la salida del alimentador hasta el TDA [m] | |
| Si fachada + sin mástil + acometida subterránea + alimentador en ducto: tramo enterrado desde la salida de la acometida hasta el medidor [m] | |
| Si fachada + sin mástil + acometida aérea: distancia vertical entre la caja del empalme y el punto de llegada de la acometida [m] | |
| Distancia del empalme a la puesta a tierra N°1 | Metros hasta la primera camarilla |
| Distancia del tablero a la puesta a tierra N°2 | Metros hasta la segunda camarilla |
| Temperatura del suelo | Solo si hay tramo subterráneo (ENTER usa la misma temperatura ambiente) |

### F. Puesta a tierra (se pregunta 2 veces: PT del empalme y PT del tablero)

| Pregunta | Detalle |
|---|---|
| Largo de barra copperweld | 1) 3 metros (recomendado) o 2) 1,5 metros (ENTER = opción 1) |
| Resistividad del terreno | 1) Valor medido en Ohm·m, o 2) Estimar por tipo de terreno: fértil/húmedo (50), poco fértil (500), pedregoso/seco (3000) |

### G. Diferenciales (al final, una sola vez)

| Pregunta | Detalle |
|---|---|
| ¿1 diferencial por circuito, o agrupar hasta 3 circuitos por diferencial? | Escribes "1" o "3". Climatización y agua caliente siempre quedan con diferencial exclusivo |

## 4. Qué obtienes al final

Al responder la última pregunta, el programa guarda un Excel en la misma carpeta del proyecto, con tres hojas:

- **Informe:** resumen de la instalación y de cada circuito
- **Materiales:** lista completa de materiales a comprar
- **Base Normativa:** qué artículo del RIC respalda cada material
