# 02 · Recorrido: cómo se llegó a la solución

Relato por etapas de la sesión del 14/11/2026, de 00:37 a 02:52 aproximadamente. Los hechos y su
evidencia están en [bitacora.md](bitacora.md); aquí va qué se intentó, qué pasó y qué se decidió.
El resultado de cada prueba está en [04-resultados.md](04-resultados.md).

## Etapa 0. Plan sin probar

- Se partía de un plan con tres vías ([00-plan-inicial.md](00-plan-inicial.md)): COM en Windows,
  AppleScript en Mac y control de pantalla. Y una lista de pruebas en orden: M0 a M8 (Mac),
  W0 a W8 (Windows), C1 y C2 (pantalla).
- La duda principal era M4: ¿se puede **meter código nuevo** en un libro desde la terminal en Mac?
  Si no, en Mac solo se podrían ejecutar macros ya existentes.

## Etapa 1. Leer antes de tocar

- Se leyó el diccionario AppleScript de Excel (`Excel.sdef`, dentro de la aplicación). No hizo falta Xcode.
- Hallazgo: tiene `run VB macro` (ejecuta y devuelve texto, hasta 30 argumentos), pero **ningún objeto
  para módulos ni código VBA**.
- Decisión: AppleScript no puede escribir código; solo ejecutar lo que ya hay. El código tendrá que
  entrar **desde dentro**, con una macro que lo cargue.

## Etapa 2. AppleScript básico (M1, M2)

- La terminal habló con Excel, escribió 42 en A1 y lo leyó. Creó `prueba.xlsm` con `save workbook as`.
- macOS no pidió permiso de Automatización: ya estaba concedido.

## Etapa 3. La primera macro: el problema del huevo y la gallina

- Para que una macro cargue código, esa macro tiene que estar ya en el libro. La primera hay que
  meterla a mano o simulando a una persona.
- Intento: pulsar menús y teclas con AppleScript (System Events). El editor se abría, pero **las teclas
  sintéticas van a la ventana activa**, y delante estaba VS Code con el usuario escribiendo. Se paró.
- Plan B: escribir [vba/Cargadora.bas](../vba/Cargadora.bas) para pegarla a mano.

## Etapa 4. agent-desktop

- Propuesta del usuario: [agent-desktop](https://github.com/vercel-labs/agent-desktop) (Vercel Labs,
  Apache-2.0). Lee el árbol de accesibilidad de las aplicaciones y pulsa sus elementos sin robar el
  foco ni hacer capturas.
- Se instaló el binario de GitHub Releases con checksum verificado, sin npm (el paquete npm ejecuta un
  `postinstall`).
- Problemas que aparecieron:
  - Menú del editor desactivado: una celda estaba en edición. Escape enviado al proceso lo resolvió.
  - El área de código del editor **no está en el árbol de accesibilidad**. Se pegó con el portapapeles
    y Cmd+V (modo `--headed`, autorizado por el usuario).
- Resultado: la cargadora quedó en `prueba.xlsm`. M3 funcionó: `run VB macro "prueba.xlsm!Hola"`
  devolvió `hola desde VBA 00:47:46`.

## Etapa 5. El ajuste de confianza y M4

- La cargadora devolvió `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es
  de confianza`. Sin ventana, porque la función tiene `On Error`.
- Con agent-desktop se marcó en Preferencias > Seguridad la casilla "Acceso de confianza al modelo de
  objetos de proyectos de VBA". No se tocó la seguridad de macros.
- M4 funcionó: código escrito en la hoja `_codigo` por AppleScript, cargado con `AddFromString` y
  ejecutado. Recargado con cambios, también.
- Decisión: **la vía Mac es viable**. A partir de aquí el código entra siempre por la hoja `_codigo`.

## Etapa 6. Motor y vba.py

- Se escribió [vba/Motor.bas](../vba/Motor.bas): carga, exporta, lista y borra módulos de **cualquier**
  libro abierto. Se instaló con la cargadora, sin tocar el editor a mano.
- Hallazgo: al escribir código en celdas, Excel se come el apóstrofo inicial de los comentarios y trata
  `=` como fórmula. Decisión: cada línea va con una marca `|` delante que el motor quita.
- Se escribió [herramientas/vba.py](../herramientas/vba.py): órdenes cortas para la terminal, sin
  dependencias. Pasa los argumentos a `osascript` por `argv` para evitar problemas de comillas.
- Ida y vuelta (`cargar`, `exportar`, `diff`): idéntica, comentarios incluidos.
- Un error controlado con `On Error` volvió como texto: `ERROR 11: División por cero`.

## Etapa 7. Los tres cierres de Excel y las reglas que salieron de ellos

Aquí se aprendió lo más importante de la sesión.

**1er cierre (00:57).** Macro `SinControl` (1/0 sin `On Error`) lanzada a través de `Motor.Ejecutar`,
que usa `Application.Run` dentro de un `On Error`.

- El error **no subió** al `On Error` de `Ejecutar`. Salió la ventana "Se ha producido el error '11' en
  tiempo de ejecución" (Finalización, Depurar...).
- Mientras la ventana estaba abierta, agent-desktop no veía ninguna ventana de Excel. La llamada de
  AppleScript se quedó esperando.
- Se pulsó Intro con System Events. **Excel se cerró de golpe** (Microsoft Error Reporting). El
  `osascript` pendiente terminó con `Connection is invalid (-609)`.
- Se perdió lo no guardado (Motor y hoja `_codigo`). Se reinstaló desde los archivos `.bas`.
- Efecto secundario: al reabrir `prueba.xlsm` desde disco, `run VB macro` daba `Parameter error (-50)`.
  Excel preguntaba si deshabilitar las macros y el botón por defecto era "Deshabilitar macros".
  agent-desktop no ve ese diálogo; System Events sí. Decisión: orden `vba.py abrir`, que pulsa
  "Habilitar macros".

**2º cierre (00:59).** Misma prueba, pero pulsando "Finalización" por su nombre.

- Excel se cerró igual. Conclusión: no es la tecla; es **detener una ejecución lanzada desde AppleScript**.
- Decisión: no dejar que la macro se pare nunca. Para eso:
  - **Envoltorio** (`vba.py seguro`): un módulo temporal `ZZEnvoltorio` que llama a la macro
    **directamente** dentro de un `On Error`. Ahí el error sí sube. Resultado: `ERROR 11: División por
    cero`, mismo PID de Excel (no se reinició).
  - **SinVentanas**: `MsgBox` e `InputBox` propios del proyecto que apuntan en una hoja en vez de abrir
    ventana. Resultado: la macro siguió y el texto quedó en `_avisos`.

**3er cierre (01:09).** Error de compilación (`Option Explicit` y variable sin declarar) en una función
llamada desde AppleScript.

- Salió "Error de compilación". Se pulsó "Aceptar" sin cierre. El editor quedó en "[interrumpir]".
- Al pulsar **"Restablecer"**, Excel se cerró.
- Hallazgo útil: llamar a otra función del mismo módulo que sí compila funciona.
- Decisión: **compilar antes de ejecutar**, con ninguna macro en marcha. `vba.py compilar` pulsa
  Depurar > Compilar VBAProject con agent-desktop. Si hay error, cerrar ese aviso **no** tumba Excel,
  y el editor dice la línea y la columna.
- Una vez el menú Compilar estaba desactivado tras `cargar`. Decisión: la orden responde `SIN VERIFICAR`
  en ese caso, en vez de dar un falso OK.

De estos tres cierres salen las reglas de oro: orden `cargar` → `compilar` → `seguro`, nunca detener una
macro parada, y guardar tras cada carga. Están explicadas en
[05-buenas-practicas.md](05-buenas-practicas.md#a-reglas-de-oro).

## Etapa 8. La demo de valor liquidativo

- Un agente preparó [demo/](../demo/ESPECIFICACION.md) sin tocar Excel: herramienta inventada con un
  error de base 360/365.
- Hallazgo al copiar el CSV: con el Mac en castellano, escribir por `value` convirtió `100.0000000000`
  en `1,0E+12`. Decisión: escribir por `formula`, que Excel lee en notación inglesa (`vba.py hoja-csv`).
- Ciclo real: `FALLO 10 de 15` → `exportar Calculo` → línea 9 con `/ 360` → corregir → `cargar` →
  `OK 15/15`. Se repitió de principio a fin con las órdenes finales (abrir desde disco, cargar,
  compilar, seguro).
- M5 (`VBComponents.Import` desde archivo) funcionó, pero si el módulo existe crea `Importado1`.
  Corrección posterior: funcionó porque el usuario pulsó "Conceder acceso" (sandbox de Excel). Decisión:
  para reemplazar módulos se usa `vba.py cargar`, no `Import`.

## Etapa 9. Permisos

- Excel para Mac pidió "Conceder acceso" cada vez que creó o importó algo en una carpeta nueva.
- Pidió "Habilitar macros" cada vez que se abrió un `.xlsm` desde disco.
- Automatización y Accesibilidad de macOS ya estaban concedidas.
- Explicación de cada permiso: [05-buenas-practicas.md](05-buenas-practicas.md#b-permisos).

## Etapa 10. Trabajo en paralelo

- Excel en Mac es una sola instancia y atiende las órdenes de una en una. Dos agentes manejándolo se pisan.
- Decisión: agentes en paralelo solo para trabajo sin Excel (la demo y la vía Windows). Excel solo lo
  tocó la sesión principal.

## Etapa 11. olevba

- Se instaló `oletools==0.60.2` en un entorno aislado.
- `olevba --code` sacó los módulos de una copia del libro **sin abrir Excel** y encontró la línea con `/ 360`.
- `olevba -a` marcó falsos positivos ("Hex Strings", "Base64 Strings"). `olevba -j` da JSON.
- Decisión: sirve como herramienta de **revisión** de solo lectura, no como parte del ciclo.

## Etapa 12. Quitar avisos y desactivar el ajuste (01:18-01:27)

- **Complemento `.xlam`, para no tener que pulsar "Habilitar macros": no conseguido.** Ni AppleScript
  (`save workbook as ... file format add in file format`) ni VBA (`SaveAs FileFormat:=55`) crearon el
  archivo, y ninguno dio error. Decisión: seguir con `vba.py abrir` y no cerrar el lanzador durante la
  sesión. Guardarlo a mano queda como hipótesis.
- **Orden `vba.py confianza activar|desactivar`.** Hace con agent-desktop el recorrido de Preferencias >
  Seguridad y comprueba el valor de la casilla. Probada en los dos sentidos.
- Hallazgo: **sin el ajuste se siguen ejecutando las macros que ya existen** (`ejecutar EjecutarTests` →
  `FALLO 10 de 15`). El ajuste solo hace falta para leer o escribir código.
- Decisión: activarlo al empezar y desactivarlo al terminar. Al cerrar la sesión quedó desactivado.
- PowerShell en el Mac, para validar la sintaxis de `vba.ps1`: pendiente. A las 01:27 aún no estaba
  instalado. En Mac no hay COM, así que solo serviría para comprobar sintaxis.

## Etapa 13. Endurecer `vba.py` y limpiar los libros (01:45-02:35)

- Encargo: corregir los fallos que encontró una revisión independiente de `vba.py` (por ejemplo,
  `exportar` podía guardar un mensaje de error como si fuera código) y limpiar los `.xlsm` de datos
  personales antes de publicarlos. A mitad, el disco se llenó e iCloud vació la carpeta: se paró todo
  hasta recuperar espacio y comprobar que los libros estaban íntegros.
- Hallazgo principal: **`compilar` compilaba otro libro**, el proyecto activo del editor de VBA.
  Decisión: activar antes el proyecto pedido, comprobar el título del editor y responder
  `OK compila <libro>`. Las salidas `OK compila` anteriores pueden ser de otro libro.
- Vigilancia y foco: la vigilancia solo mira diálogos y solo pulsa Aceptar/OK en avisos de Excel. Con
  Excel en segundo plano, System Events no ve sus ventanas: `vba.py` pone Excel delante y, al terminar,
  devuelve el foco a la aplicación que estaba delante.

## Etapa 14. `confianza` con el editor delante y el libro "después" (02:40-02:52)

- `confianza` fallaba a veces (`ERROR no encuentro la casilla de confianza`). Una captura mostró la causa:
  con el editor de VBA delante, "Preferencias..." abre las Opciones del editor y no las de Excel. El clic
  funcionaba, pero en el sitio equivocado.
- Decisión: comprobar antes que la barra de menús no es la del editor y, si lo es, pulsar "Ver Microsoft
  Excel". Si la orden no termina en `OK`, sale con código 1. Probado también con el editor delante.
- Se creó `demo/vl-fondo-corregido.xlsm`, el después de la demo (`Calculo` corregido, `OK 15/15` e
  informe), para ver el resultado sin ejecutar nada. Al terminar, metadatos limpiados en los cuatro
  `.xlsm` y ajuste de confianza desactivado.

Hechos y salidas literales de las dos etapas: [bitacora.md](bitacora.md), secciones 13 y 14.

Siguiente: [03-solucion-mac.md](03-solucion-mac.md), la solución pieza a pieza.
