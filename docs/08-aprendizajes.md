# Aprendizajes: la manera adecuada y por qué se rompe si no se sigue

Si solo vas a leer un documento antes de usar esto, que sea este. Probamos muchas cosas y muchas fallaron.
Los fallos no fueron al azar: tienen tres causas, y la manera adecuada sale de entenderlas.
Detalle de cada prueba en [04-resultados.md](04-resultados.md) y de cada pieza en [03-solucion-mac.md](03-solucion-mac.md).
Las reglas para el día a día que salen de aquí, con la solución de problemas por síntoma, están en
[05-buenas-practicas.md](05-buenas-practicas.md).

## La respuesta corta

En Mac, la vía principal es **AppleScript + el motor VBA**, manejados con `herramientas/vba.py`, siguiendo
siempre este orden:

```
abrir  ->  cargar  ->  compilar  ->  seguro
```

agent-desktop **no** es la vía principal: es una ayuda para tres cosas puntuales.

| Pieza | Papel | ¿Imprescindible? |
|---|---|---|
| AppleScript (`osascript`, viene con macOS) | El canal con Excel: abrir, guardar, ejecutar macros y recibir lo que devuelven | Sí |
| Motor VBA (`vba/Motor.bas`, dentro del libro lanzador `prueba.xlsm`) | Mete, saca y reemplaza el código de cualquier libro abierto | Sí |
| `herramientas/vba.py` | Junta todo en órdenes cortas y aplica las precauciones de este documento | Sí, en la práctica |
| agent-desktop | Abrir el editor la primera vez, compilar (menú Depurar), cambiar el ajuste de confianza | Solo para compilar y para preparar |
| System Events (viene con macOS) | Pulsar "Habilitar macros" al abrir un libro (agent-desktop no ve esa ventana) | Solo al abrir |

## Por qué se rompe: las tres causas

### 1. Excel para Mac no está pensado para que lo maneje un programa desde fuera

- En **Windows** existe COM: un programa abre su propio Excel, invisible, sin molestar a nadie, y lo cierra al terminar.
- En **Mac** solo existe AppleScript, que habla con **el Excel que tienes abierto**. Hay una sola copia de Excel,
  compartida entre tú y la automatización.

Consecuencias que vimos:

| Lo que pasó | Por qué |
|---|---|
| Las teclas enviadas por AppleScript acabaron en VS Code | Las teclas sintéticas van a la ventana que está delante, no a Excel |
| No se puede poner a varios agentes a trabajar con Excel a la vez | Una sola copia de Excel que atiende las órdenes de una en una: se pisarían la hoja activa, las ventanas y el portapapeles |
| Los menús de Excel salían desactivados | Una celda estaba en modo edición; Excel bloquea los menús mientras tanto |
| `compilar` trae Excel al frente | Usa los menús del editor, que solo responden con Excel activo |
| La pantalla "salta" a Excel mientras trabajas en otra aplicación | `compilar`, `confianza`, `dialogos`, abrir o crear libros, la vigilancia y los avisos de Excel (`MsgBox`) ponen Excel delante. `cargar`, `seguro`, `exportar` y `leer` no, mientras no salga un aviso: van por detrás. `vba.py` devuelve el foco a la aplicación que estaba delante al terminar `compilar`, `confianza`, `abrir`, `dialogos` y la vigilancia, también si hay error |
| Con Excel en segundo plano, System Events ve 0 ventanas de Excel, ni siquiera un `MsgBox` abierto | macOS no publica las ventanas de una aplicación que no está delante (agent-desktop tampoco las ve). Al ponerla delante cambia de escritorio y tarda en mostrarlas: 0,3-0,5 s no bastaban. Por eso `vba.py` pone Excel delante sin teclas (`frontmost`) y espera hasta 3 s, y por eso los avisos traen Excel al frente |
| `activate` no siempre trae Excel delante | Con el usuario en VS Code, a veces no cambió y los menús de Excel salieron desactivados (`compilar` no encontraba Depurar > Compilar). `frontmost` por System Events sí |
| `confianza` no encontraba la casilla a veces | Con el editor de VBA delante (por ejemplo, después de `compilar`), los menús de arriba son los del editor: "Preferencias..." abre las Opciones del editor de VBA, no las de Excel. El clic funcionaba, pero en el sitio equivocado, y solo una captura lo mostró. Ahora `confianza` comprueba que la barra de menús no es la del editor y, si lo es, pulsa "Ver Microsoft Excel" |

**Qué hacer**: no escribir en otras ventanas mientras corre una orden que usa la interfaz (`compilar`,
`confianza`), usar un solo agente para Excel, y nombrar siempre el libro (`--libro`), nunca "el activo".

### 2. VBA espera que haya una persona delante

VBA nació para usarse a mano. Ante un error, un `MsgBox` o un `InputBox`, **abre una ventana y se para a
esperar** a que alguien pulse un botón. Si la macro la lanzó un programa, ese programa se queda esperando.

Y lo más grave, descubierto probando: **en Excel para Mac, detener una macro que se ha parado y que lanzó
AppleScript cierra Excel de golpe**. Pasó 3 de 3 veces el 14/11/2026:

1. Error de ejecución (división por cero), cerrado con Intro.
2. El mismo error, cerrado pulsando "Finalización" por su nombre (así que no era la tecla).
3. Error de compilación, y después "Restablecer" en el editor.

Se pierde todo lo no guardado y Excel reabre con libros "recuperados" (`prueba (version 1).xlsb`).

**Qué hacer**: que la ventana **nunca llegue a salir**. Desde que se aplicó esto, no hubo más cierres:

| Defensa | Contra qué | Cómo |
|---|---|---|
| `compilar` antes de ejecutar | Errores de compilación (variable no declarada, sintaxis) | Activa en el editor el proyecto del libro pedido, comprueba el título del editor y compila desde el menú sin ejecutar nada: cerrar ese aviso es seguro porque no hay macro en marcha. Devuelve módulo, línea y motivo, o `OK compila <libro>` |
| `seguro` en lugar de `ejecutar` | Errores de ejecución sin controlar | Genera un módulo temporal que llama a la macro dentro de un `On Error`: el error vuelve como texto |
| `SinVentanas.bas` en el libro | `MsgBox` e `InputBox` | Sustituye esas funciones por unas propias que apuntan el texto en la hoja `_avisos` |
| `cargar` guarda siempre | Perder trabajo si Excel se cae | Cada carga deja el libro guardado |
| Vigilancia de `vba.py` | Que la terminal se quede colgada | Si una macro no responde, mira solo los diálogos (`AXDialog`). Cierra con Aceptar/OK los avisos cuyo título o primer texto es "Microsoft Excel", hasta 10 seguidos, y devuelve lo que dio la macro. Ante un diálogo con "Visual Basic", "Finalización", "Depurar", "compilaci" o "tiempo de ejecuci" **no toca nada** y lo dice (`ERROR VENTANA VBA (no se ha tocado): ...`) |

Lo que no funcionó como defensa: atrapar el error con `Application.Run` dentro de un `On Error`. En Mac el
error no sube a quien llama; por eso el envoltorio llama a la macro directamente.

La vigilancia mira solo diálogos porque la ventana principal del editor de VBA también lleva "Visual
Basic": si contara, no cerraría nunca un `MsgBox` con el editor abierto. Detalle en
[03-solucion-mac.md](03-solucion-mac.md#herramientasvbapy).

### Un `OK` solo vale si se sabe de qué es

El 14/11 se descubrió que `compilar` **compilaba el proyecto equivocado**. Depurar > Compilar actúa sobre
el proyecto activo del editor de VBA (el de la ventana de código con el foco), no sobre el libro que se
activa en Excel. Con `--libro vl-fondo.xlsm` devolvió el error de compilación de `errores.xlsm`. Es decir:
**las salidas `OK compila` anteriores a la corrección pueden ser de otro libro**. La orden respondía `OK`
y nada indicaba de qué proyecto era.

Ahora `compilar` activa en el editor el proyecto pedido, lee el título del editor
("Microsoft Visual Basic - libro.xlsm"), no compila si no coincide y responde `OK compila <libro>`. La
lección general: una comprobación que no dice sobre qué actuó no demuestra nada. Por eso `compilar`
nombra ahora el libro en su salida. Lo mismo pasó con `confianza`: un clic en el menú equivocado parece
funcionar (ver la tabla de la causa 1). Mecanismo y corrección de `compilar`:
[03-solucion-mac.md](03-solucion-mac.md#compilar-vbapy-compilar).

### 3. Las protecciones de seguridad de Excel y macOS (están bien puestas)

No son fallos. Existen porque las macros son una vía clásica de virus. Hay que conocerlas y gestionarlas,
no quitarlas.

| Protección | Qué es y por qué existe | Qué vimos | Qué hacer |
|---|---|---|---|
| **"Acceso de confianza al modelo de objetos de proyectos de VBA"** | Permite que una macro lea y reescriba el código VBA de otros libros. Desactivado por defecto porque un virus de macro haría exactamente eso | Sin él, el motor devuelve `ERROR 1004`. Las macros que ya existen se pueden ejecutar igual | Activarlo al empezar a trabajar y desactivarlo al terminar: `vba.py confianza activar|desactivar` |
| **"Habilitar macros"** al abrir un `.xlsm` | Excel pregunta si fiarse de las macros del archivo; el botón por defecto es "Deshabilitar" | Abierto por AppleScript, quedaban deshabilitadas y `run VB macro` daba `Parameter error (-50)` | `vba.py abrir` (pulsa "Habilitar macros"); usarlo solo con libros propios. No cerrar el lanzador durante la sesión |
| **"Conceder acceso"** (sandbox de Excel) | Excel para Mac está aislado y necesita permiso para cada carpeta o archivo fuera de su contenedor | Salió al crear o importar en cada carpeta nueva; la importación de `.bas` funcionó porque se concedió | Trabajar siempre dentro de la misma carpeta ya concedida |
| **Automatización y Accesibilidad** (macOS) | Permisos para que un programa controle otra app o pulse sus botones | Ya estaban concedidos; no aparecieron | La primera vez en otro Mac, aceptarlos en Ajustes del Sistema > Privacidad y seguridad |

Intentamos evitar "Habilitar macros" guardando el motor como complemento (`.xlam`): no se consiguió desde
AppleScript ni desde VBA. Queda por probar a mano.

## Otros detalles que rompen cosas sin avisar

| Detalle | Qué pasa | Cómo lo resolvemos |
|---|---|---|
| Mac en castellano | Un número escrito como `100.5` se interpreta con el punto como separador de miles (`1,0E+12`) | `vba.py` escribe los datos por fórmula, que Excel lee siempre en notación inglesa |
| Código escrito en celdas | Excel se come el apóstrofo inicial de los comentarios y trata `=` como fórmula | Cada línea lleva una marca `\|` delante que el motor quita |
| Nombres de módulo | VBA no admite nombres que empiecen por `_` | El envoltorio se llama `ZZEnvoltorio` |
| Importar un `.bas` sobre un módulo que existe | Crea `Nombre1` en lugar de reemplazar | `vba.py cargar` reemplaza el código en su sitio |
| `Parameter error (-50)` | Casi siempre: la macro no existe o las macros están deshabilitadas | Comprobar con `vba.py modulos` y abrir con `vba.py abrir` |
| `Connection is invalid (-609)` | Excel se ha cerrado mientras esperábamos | Ver recuperación en [05-buenas-practicas.md](05-buenas-practicas.md) |
| Disco lleno con la carpeta en iCloud | iCloud vació la copia local de la carpeta y Excel se quedó sin autorrecuperación hasta reiniciarlo. Un fallo del sistema de archivos parece un fallo de la herramienta | Carpeta fuera de iCloud o "Mantener descargado", y espacio en disco. Síntomas y pasos en [05-buenas-practicas.md](05-buenas-practicas.md#c-solución-de-problemas-por-síntoma) |
| Excel guarda datos personales en el libro | En cada guardado reescribe el autor (`dc:creator`, `cp:lastModifiedBy`) y la ruta absoluta de la carpeta (`x15ac:absPath`). `cargar` guarda siempre | Limpiar con los libros cerrados justo antes de publicarlos ([05-buenas-practicas.md](05-buenas-practicas.md#antes-de-publicar-o-compartir-un-libro)). `vbaProject.bin` conserva una referencia de MSForms con la ruta del usuario: límite conocido |
| Nombres de libro con apóstrofo | En la referencia `'libro'!macro` el apóstrofo rompe la orden si no se duplica | `vba.py` pasa todos los valores por `argv` y duplica los apóstrofos. Probado con `prueba d'Enero.xlsm` |
| `save workbook as` deja de funcionar | Más tarde el 14/11 dio `Parameter error (-50)` al crear libros nuevos, cuando antes había funcionado. Causa sin confirmar | Crear el libro a mano o reiniciar Excel |

## ¿Hace falta instalar agent-desktop?

**No para el día a día; sí resulta útil.** El ciclo `cargar`, `seguro`, `exportar` y `abrir` funciona sin él.
agent-desktop hace en Mac lo que haría una persona con ratón y teclado, pero leyendo la estructura de las
ventanas en lugar de mirar la pantalla. Eso lo hace muy útil para lo que solo se puede hacer desde la interfaz:

| Tarea | Con agent-desktop | Sin agent-desktop |
|---|---|---|
| Primera instalación (pegar la cargadora en el editor) | Automática | 2 minutos a mano, una sola vez |
| Activar o desactivar el ajuste de confianza | `vba.py confianza ...` | A mano en Excel > Preferencias > Seguridad |
| Compilar antes de ejecutar | `vba.py compilar`, con módulo, línea y motivo | A mano: editor de VBA > Depurar > Compilar. System Events podría pulsar ese menú (sin probar) |
| Ciclo de trabajo (cargar, ejecutar, exportar) | No lo usa | Igual |

Por qué no es la vía principal aunque pueda "usar el ordenador como tú":

1. **No ve lo importante en Excel.** El área donde se escribe el código del editor de VBA no aparece en el
   árbol de accesibilidad, y las ventanas de aviso de Excel (error de VBA, "Habilitar macros") tampoco. Para
   escribir código tuvo que pegar con el portapapeles, y para esos avisos hizo falta System Events.
2. **Necesita la pantalla.** Varias acciones requieren Excel delante, y si alguien escribe a la vez en otra
   ventana se pisan. El motor trabaja sin interfaz: no depende de qué ventana esté activa.
3. **Es más frágil por naturaleza.** Depende de nombres de botones y menús (en castellano), del estado de la
   interfaz (una celda en edición desactiva los menús) y de que la ventana exista. Una orden directa a Excel
   no depende de nada de eso.
4. **Añade una herramienta externa** (binario sin firma de desarrollador, con permisos de Accesibilidad) que
   habría que instalar y aprobar en cada máquina, también en la de un usuario de empresa.

Regla práctica: **orden directa cuando existe (AppleScript, motor); interfaz solo cuando no hay otra vía**.
Para una persona que solo va a usar la solución, se puede prescindir de agent-desktop haciendo a mano la
preparación y la compilación. Para quien desarrolla con Claude, merece la pena tenerlo por `compilar`.

## Qué descartamos y por qué

| Camino | Por qué no |
|---|---|
| Pulsar teclas con AppleScript en el editor | Las teclas van a la ventana activa; frágil y peligroso si el usuario está escribiendo |
| agent-desktop para todo | No ve el área de código ni las ventanas de aviso de Excel, y necesita la pantalla. Útil solo para lo puntual |
| Control de pantalla por capturas | Lento y frágil; agent-desktop hace lo mismo por accesibilidad y de forma comprobable |
| `VBComponents.Import` como vía principal | Duplica módulos y pide permisos de carpeta; el motor reemplaza en su sitio |
| `xlwings` | En Mac usa AppleScript por debajo: no añade capacidad, solo una dependencia más |
| LibreOffice | No es Excel: no demuestra nada sobre Excel |

## Si la herramienta fuera nueva

Nada de esto haría falta: se haría en Python, con Excel solo como entrada y salida de datos. Ver
[07-vba-python-power-bi.md](07-vba-python-power-bi.md).
