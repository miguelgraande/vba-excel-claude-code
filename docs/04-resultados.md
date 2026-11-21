# 04 · Resultados

Pruebas del 14/11/2026 en macOS 26.6.2 con Excel 16.115. Plan original de las pruebas:
[00-plan-inicial.md](00-plan-inicial.md). Hechos con hora: [bitacora.md](bitacora.md).
Las nuevas pruebas (por ejemplo W0 a W8) se apuntan aquí, con fecha y salida real.

## Pruebas

### Mac

| ID | Prueba | Resultado | Evidencia |
|---|---|---|---|
| M0 | Preparación: Excel abierto, libro de prueba, ajuste de confianza | OK | El ajuste se activó durante M4 (estaba desmarcado) |
| M1 | La terminal habla con Excel | OK | `osascript -e 'tell application "Microsoft Excel" to return version'` → `16.115`. No pidió permiso de Automatización (ya concedido) |
| M2 | Escribir y leer celdas | OK | Escribe 42 en A1 y lee `42.0` |
| Extra | Crear un `.xlsm` desde la terminal | OK al principio; después falla | `save workbook as wb filename "<ruta>" file format macro enabled XML file format`. Más tarde, el mismo 14/11, la misma orden dio `Parameter error (-50)` sin crear archivo (con y sin apóstrofo en el nombre, también con ruta HFS). Causa sin confirmar |
| M3 | Ejecutar una macro existente y recibir su valor | OK | `run VB macro "prueba.xlsm!Hola"` → `hola desde VBA 00:47:46`. También con `"Hola"` a secas. Funcionó antes de activar el ajuste de confianza |
| M4 | Meter código nuevo sin archivos | OK | Sin ajuste: `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es de confianza`, sin ventana. Con ajuste: `OK Generado 3 lineas` y `Nueva` → `codigo nuevo, 2+3 = 5`. Recarga con cambio: `version 2, 10*7 = 70` |
| Extra | El motor se instala con la cargadora | OK | La cargadora cargó `Motor` sin tocar el editor a mano |
| Extra | Ida y vuelta de un módulo | Idéntica | `cargar` + `exportar` + `diff`, con comentarios `'` y líneas que empiezan por `=` |
| M5 | Importar un `.bas` desde archivo (`VBComponents.Import`) | OK con límite | Funcionó desde la carpeta del proyecto y desde la de Excel **porque el usuario pulsó "Conceder acceso"**. Si el módulo existe, crea `Importado1` al lado. Para reemplazar: `vba.py cargar` |
| M6a | Error controlado con `On Error` | OK | `ConControl` → `ERROR 11: División por cero` |
| M6b | Error sin controlar, sin envoltorio | Falla | Sale la ventana de error de VBA y, al cerrarla, Excel se cae (ver abajo) |
| M6c | Error sin controlar, con envoltorio | OK | `seguro SinControl` → `ERROR 11: División por cero`, mismo PID de Excel |
| M6d | `MsgBox` con SinVentanas | OK | `seguro ConMsgBox` → `despues del MsgBox`; el texto queda en la hoja `_avisos`. Mismo PID |
| Extra | Error de compilación en una función que no se llama | No molesta | `Bien` se ejecuta aunque `Mal`, en el mismo módulo, no compile |
| Extra | Detectar errores de compilación sin ejecutar | OK, con un fallo corregido después | `compilar` → `ERROR COMPILACION: Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`. Cerrar ese aviso no tumba Excel. Tras corregir: `OK compila`. Una vez de tres: `SIN VERIFICAR` (menú desactivado). **Fallo descubierto más tarde:** las salidas `OK compila` de esta fila y de M7 pueden ser de otro libro ([por qué](08-aprendizajes.md#un-ok-solo-vale-si-se-sabe-de-qué-es)). Ver las pruebas de `compilar` corregido abajo |
| Extra | Abrir un libro con macros sin intervención | OK | `vba.py abrir` pulsa "Habilitar macros" con System Events |
| M7 | Ciclo completo con referencia en Python | OK | Demo: `FALLO 10 de 15` → `exportar Calculo` → línea 9 `/ 360` → corregir → `cargar` → `OK 15/15` → `OK Informe 15 dias`. Repetido de principio a fin con las órdenes finales. Repetido otra vez con `vba.py` endurecido y con `demo/ejecutar_demo.py` (ver abajo) |
| M8 | Varios libros abiertos | OK | Lanzador y `vl-fondo.xlsm` abiertos; cada orden nombra su libro con `--libro` |
| Extra | agent-desktop: editor, menús, preferencias | OK | Abrió el editor, insertó un módulo, pegó, guardó y marcó el ajuste de confianza (verificado: valor 1) |
| Extra | olevba: leer el VBA sin Excel | OK | `olevba --code` sobre una copia encontró `TasaDiaria = ... / 360`. `-a` da falsos positivos; `-j` da JSON |
| Extra | `vba.py confianza activar\|desactivar` | OK | Probada en los dos sentidos. Desactivada: `vba.py modulos` → `ERROR 1004 ...` sin ventana. Activada: vuelve a funcionar. Al cerrar la sesión quedó desactivada |
| Extra | Ejecutar macros existentes sin el ajuste de confianza | OK | `ejecutar EjecutarTests` → `FALLO 10 de 15`; `ejecutar Hola` → `hola desde VBA 01:26:40` |
| Extra | Crear el motor como complemento `.xlam` desde la terminal | Falla | AppleScript `save workbook as ... file format add in file format`: sin error y sin archivo; después `The object you are trying to access does not exist (-1728)`. VBA `SaveAs FileFormat:=55`: devuelve vacío y sin archivo |

### Mac: `vba.py` endurecido (14/11, bitácora sección 13)

Pruebas tras corregir los hallazgos de una revisión independiente de `herramientas/vba.py`. Salida
literal. "Delante antes y después" es la aplicación que estaba delante al lanzar la orden y al terminar.

| Prueba | Orden | Resultado | Evidencia |
|---|---|---|---|
| Ajuste de confianza | `vba confianza activar` | OK | `OK confianza activada (antes: activada)` |
| Macro existente | `vba ejecutar Hola` | OK | `hola desde VBA 02:08:34` |
| Exportar | `vba exportar Calculo --libro vl-fondo.xlsm --salida /tmp/x.bas` | OK | `OK Calculo -> /tmp/x.bas`, código 0, idéntico a `demo/vba/Calculo.bas` sin las líneas `Attribute` |
| Exportar un módulo que no existe | `vba exportar NoExiste --libro vl-fondo.xlsm --salida /tmp/y.bas` | OK | `ERROR 9: Subíndice fuera de rango`, código 1, `/tmp/y.bas` no creado |
| Protección del lanzador | `vba cargar vba/Motor.bas` | OK | `ERROR no reemplazo Motor en el lanzador prueba.xlsm: es el codigo que hace la carga. ...`, código 1 |
| Protección del lanzador | `vba cargar vba/Cargadora.bas --modulo Módulo1` | OK | `ERROR no reemplazo Módulo1 en el lanzador prueba.xlsm: ...`, código 1 |
| Abrir un archivo que no existe | `vba abrir vba/pruebas/noexiste.xlsm` | OK | `ERROR no existe el archivo: .../vba/pruebas/noexiste.xlsm`, código 1 |
| Nombre con apóstrofo | `vba cargar vba/Cargadora.bas --libro "prueba d'Enero.xlsm" --modulo Saludo --sin-guardar` | OK | `OK Saludo 34 lineas en prueba d'Enero.xlsm` |
| Nombre con apóstrofo | `vba ejecutar Hola --libro "prueba d'Enero.xlsm"` | OK | `hola desde VBA 02:12:06` |
| Nombre con apóstrofo | `vba seguro Hola --libro "prueba d'Enero.xlsm"` | OK | `hola desde VBA 02:12:07`; `ZZEnvoltorio` borrado después |
| Nombre con apóstrofo | `vba abrir "vba/pruebas/prueba d'Enero.xlsm"` | OK | `OK abierto prueba d'Enero.xlsm`; después `ejecutar Bien` → `compila`. Cerrado con `close workbook` por argv |
| Vigilancia con un `MsgBox` real | `vba ejecutar ConAviso --libro errores.xlsm --timeout 10` | OK | `AVISO cerrado con Aceptar: Microsoft Excel / ventana de prueba (usa SinVentanas para evitarlo). Resultado: despues del MsgBox` |
| Vigilancia con dos `MsgBox` seguidos | `vba ejecutar DosAvisos --libro errores.xlsm --timeout 10` | OK | `AVISO cerrado con Aceptar: Microsoft Excel / primer aviso; Microsoft Excel / segundo aviso (usa SinVentanas para evitarlo). Resultado: despues de dos MsgBox`. Delante antes y después: VS Code |
| Salida grande | Función que devuelve 100 003 caracteres | OK | 100 004 bytes con el salto de línea, termina en `FIN`, sin bloqueo |
| `compilar`, antes de corregir la selección de proyecto | `vba compilar --libro vl-fondo.xlsm` | **Falla** | `ERROR COMPILACION: Microsoft Visual Basic / Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`: compiló otro libro |
| `compilar` corregido, libro que compila | `vba compilar --libro vl-fondo.xlsm` | OK | `OK compila vl-fondo.xlsm`. Delante antes y después: VS Code |
| `compilar` corregido, libro que no compila | `vba compilar --libro errores.xlsm` | OK | `ERROR COMPILACION: Microsoft Visual Basic / Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`. Delante antes y después: VS Code |
| Ciclo de la demo a mano | `cargar`, `compilar`, `seguro`, `cargar` corregido, `compilar`, `seguro` | OK | `OK Calculo 52 lineas en vl-fondo.xlsm`, `OK compila vl-fondo.xlsm`, `FALLO 10 de 15: ...`, `OK Calculo 53 lineas`, `OK compila vl-fondo.xlsm`, `OK 15/15`. Un `compilar` devolvió `ERROR no encuentro Depurar > Compilar en el editor`; repetido dos veces: `OK compila vl-fondo.xlsm`. Se añadió un reintento |
| Demo guiada | `python3 -I demo/ejecutar_demo.py` | OK | `OK compila vl-fondo.xlsm` (dos veces), `FALLO 10 de 15: ...`, `linea 9: TasaDiaria = ComisionPctDia(fila) / 100 / 360`, `OK 15/15`, `OK Informe 15 dias`, `DEMO CORRECTA`, código 0. Delante antes y después: VS Code |
| Demo guiada sin agent-desktop | `python3 -I demo/ejecutar_demo.py --sin-compilar` | OK | Mismos resultados sin compilar, `DEMO CORRECTA`, código 0 |
| Desactivar el ajuste | `vba confianza desactivar` | OK | `OK confianza desactivada (antes: activada)`; después `vba modulos` → `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es de confianza.` |

Sin probar: un error de ejecución real con la vigilancia (no se provoca a propósito: cerrarlo tumba
Excel; su clasificación se comprobó con los textos de las ventanas vistas antes), y el texto exacto del
aviso "Habilitar macros" en System Events (`abrir` lo sigue pulsando: los libros abren con macros).

Las pruebas con `ConAviso` y `DosAvisos` usaron un módulo de prueba (`Aviso`) que se borró de
`errores.xlsm` al limpiarlo. Para repetir la de un `MsgBox` con el módulo del repositorio (sin probar
así): `vba abrir vba/pruebas/errores.xlsm`, `vba cargar vba/pruebas/PruebaErrores.bas --libro
errores.xlsm` y `vba ejecutar ConMsgBox --libro errores.xlsm --timeout 10`. Debe responder
`AVISO cerrado con Aceptar: Microsoft Excel / ventana de prueba ... Resultado: despues del MsgBox`.
`cargar` guarda el libro: al terminar, `vba borrar PruebaErrores --libro errores.xlsm`. Para dos avisos
seguidos no hay módulo en el repositorio.

Después de estas pruebas se añadió a `demo/ejecutar_demo.py` un paso final que repite tests e informe con
la versión con el error, para que las hojas `Tests` e `Informe` del libro sean coherentes con su código.
Ese paso no estaba en las ejecuciones de esta tabla; se probó después (ver la tabla siguiente, última fila).

### Mac: `confianza` con el editor delante y libro "después" (14/11, bitácora sección 14)

| Prueba | Orden | Resultado | Evidencia |
|---|---|---|---|
| `confianza` antes del arreglo | `vba confianza activar\|desactivar` | **Falla a veces** | `ERROR no encuentro la casilla de confianza`. Una captura mostró que, con el editor de VBA delante, "Preferencias..." abría las Opciones del editor de VBA |
| Activar, con el ajuste desactivado | `vba confianza activar` | OK | `OK confianza activada (antes: desactivada)` |
| Activar otra vez | `vba confianza activar` | OK | `OK confianza activada (antes: activada)` |
| Desactivar con el editor de VBA delante, justo después de `compilar` | `vba confianza desactivar` | OK | `OK confianza desactivada (antes: activada)`; después `modulos` → `ERROR 1004` (desactivado de verdad) |
| Crear el libro "después" | `cp demo/vl-fondo.xlsm demo/vl-fondo-corregido.xlsm`; `vba abrir` del lanzador y de la copia; `vba cargar demo/solucion/Calculo_corregido.bas --libro vl-fondo-corregido.xlsm --modulo Calculo`; `vba compilar`, `vba seguro EjecutarTests`, `vba seguro GenerarInforme` y `vba guardar` con `--libro vl-fondo-corregido.xlsm` | OK | `OK Calculo 53 lineas en vl-fondo-corregido.xlsm`, `OK compila vl-fondo-corregido.xlsm`, `OK 15/15`, `OK Informe 15 dias`. Hoja `Tests` con las 15 filas en OK y `TasaDiaria = ComisionPctDia(fila) / 100 / 365`. Se abre en la hoja `Informe` |
| Cerrar Excel | `quit saving no` (AppleScript) | OK tras cerrar un diálogo | Primero `User cancelled (-128)`: seguía abierto el diálogo de Opciones del editor. Cerrado con "Cancelar", sin macro en marcha, Excel se cerró |
| Limpieza de metadatos | Reescritura del ZIP con los libros cerrados | OK | Los cuatro `.xlsm`: autor vacío, sin `absPath`, ZIP íntegros |
| Demo guiada con el paso final nuevo (ejecución final del 14/11) | `python3 -I demo/ejecutar_demo.py` | OK | Antes `FALLO 10 de 15`, `OK compila vl-fondo.xlsm` dos veces, después `OK 15/15`, `OK Informe 15 dias`. Paso 5: `OK Calculo 52 lineas en vl-fondo.xlsm`, `FALLO 10 de 15`, `OK Informe 15 dias`, `OK guardado vl-fondo.xlsm`. Final: `DEMO CORRECTA`. El libro queda con el código con el error y sus hojas coherentes |

### Control de pantalla

| ID | Prueba | Resultado | Evidencia |
|---|---|---|---|
| C1 | Hacerlo como una persona (editor, pegar, F5) | No probada como tal | No se usaron capturas. agent-desktop hizo los pasos de una persona por accesibilidad (abrir editor, insertar módulo, pegar, guardar), pero la ejecución se hizo con `run VB macro` |
| C2 | Cerrar una ventana que bloquea | Se puede, pero Excel se cae | System Events cerró la ventana de error (Intro y "Finalización"). Las dos veces Excel se cerró de golpe |

### Windows

| ID | Prueba | Resultado |
|---|---|---|
| W0 a W10 | Vía COM con `herramientas/windows/vba.ps1` | **Sin probar.** No había un Windows disponible. W9 y W10 (`compilar`) y el nuevo vigía de W6 se añadieron el 21/11, adaptados de MCP-Access ([06-windows.md §13](06-windows.md#13-técnicas-tomadas-de-mcp-access)). Plan en [06-windows.md](06-windows.md#11-plan-de-pruebas-en-windows) |
| Extra | Validar la sintaxis de `vba.ps1` con PowerShell en el Mac | **Pendiente.** Se estaba instalando; a las 01:27 `pwsh` aún no estaba. El 21/11 tampoco había `pwsh`. En Mac no hay COM: solo serviría para comprobar sintaxis |

## Qué no funcionó y por qué

| Intento | Qué pasó | Qué se hace en su lugar |
|---|---|---|
| Detener una macro lanzada desde AppleScript que se ha parado | **Excel se cierra de golpe** (Microsoft Error Reporting). 3 de 3 veces: Intro y "Finalización" en un error de ejecución, "Restablecer" en el editor tras un error de compilación | No dejar que se pare: `compilar` antes, `seguro` siempre, SinVentanas. Guardar tras cada carga |
| Atrapar el error con `Application.Run` dentro de un `On Error` (`Motor.Ejecutar`) | El error no sube al llamante; sale la ventana igual | Envoltorio que llama a la macro directamente |
| Pulsar teclas en el editor con AppleScript | Las teclas van a la ventana activa (VS Code) | agent-desktop, que actúa sin robar el foco |
| agent-desktop con ventanas modales de Excel | No las ve: Excel deja de publicar sus ventanas mientras hay una modal | System Events |
| agent-desktop para escribir en el área de código | No está en el árbol de accesibilidad | Portapapeles + Cmd+V una vez; después, el motor |
| Clic `--headed` de agent-desktop sobre un menú desactivado | Agotó el tiempo y agent-desktop se cerró (SIGABRT). Excel siguió bien | Escape antes, para salir del modo edición de la celda |
| Abrir un `.xlsm` desde disco por AppleScript | El botón por defecto es "Deshabilitar macros"; después `run VB macro` da `Parameter error (-50)` | `vba.py abrir` |
| Escribir números con `value` y el Mac en castellano | `100.0000000000` se leyó como `1,0E+12` | Escribir por `formula` (`vba.py hoja-csv`) |
| Escribir código en celdas tal cual | Excel se come el `'` inicial y trata `=` como fórmula | Marca `\|` delante de cada línea |
| `VBComponents.Import` para reemplazar un módulo | Crea `Nombre1` al lado y pide "Conceder acceso" | `vba.py cargar` |
| Crear un complemento `.xlam` desde la terminal (AppleScript o VBA) | No da error, pero no crea el archivo | `vba.py abrir` y no cerrar el lanzador. Guardarlo a mano: sin comprobar |
| Nombre de módulo que empieza por `_` | VBA no lo admite | `ZZEnvoltorio` |
| `compilar` activando solo el libro (`activate object workbook`) | Con `--libro vl-fondo.xlsm` compiló `errores.xlsm` | Activar antes en el editor el proyecto pedido y comprobar el título del editor ([mecanismo](03-solucion-mac.md#compilar-vbapy-compilar)) |
| Leer los avisos de Excel con Excel en segundo plano | System Events ve 0 ventanas de Excel, ni siquiera un `MsgBox`. Al ponerlo delante, una espera de 0,3-0,5 s no bastaba | Poner Excel delante (`frontmost`, sin teclas) y esperar hasta 3 s a que se vean sus ventanas |
| `tell application "Microsoft Excel" to activate` con el usuario en VS Code | A veces no trae Excel delante y los menús salen desactivados (`ERROR no encuentro Depurar > Compilar`) | `frontmost` por System Events |
| Vigilancia que mira todas las ventanas buscando "Visual Basic" | La ventana principal del editor también lo lleva: con el editor abierto, no cerraría nunca un `MsgBox` | Solo cuentan los diálogos (`AXDialog`) |
| `save workbook as` para crear un libro nuevo (más tarde el 14/11) | `Parameter error (-50)`, sin archivo. El libro en blanco quedó abierto y no se pudo cerrar (`close ... saving no` → -50; VBA `Close False` → `ERROR 1004`) | Causa sin confirmar. Crear el libro a mano o reiniciar Excel; ver [05-buenas-practicas.md](05-buenas-practicas.md#c-solución-de-problemas-por-síntoma) |

## Alternativas evaluadas

Veredictos: **sirve**, **sirve con límite**, **descartada**, **no probada**.

| Alternativa | Veredicto | Motivo |
|---|---|---|
| AppleScript (`osascript`) + motor + `vba.py` | Sirve (la elegida) | Ciclo completo demostrado. Nada que instalar salvo agent-desktop para `compilar` |
| `run VB macro` sobre macros existentes | Sirve | Devuelve el valor a la terminal. Funciona incluso sin el ajuste de confianza |
| Motor en un complemento `.xlam` para evitar "Habilitar macros" | Descartada por ahora | No se pudo crear desde la terminal. Guardarlo a mano y cargarlo desde la carpeta de inicio de Office: no probado |
| Cargadora con `AddFromString` | Sirve | Decisiva (M4). Sustituida en el día a día por el motor, que trabaja sobre cualquier libro |
| `VBComponents.Import` desde archivo | Sirve con límite | No reemplaza (crea `Nombre1`) y depende de "Conceder acceso" |
| agent-desktop | Sirve con límite | Bueno para menús y ajustes sin robar el foco. No ve ventanas modales ni el área de código |
| System Events para pulsar botones | Sirve con límite | Ve las modales. No sirve para escribir: las teclas van a la ventana activa |
| Teclas sintéticas con AppleScript en el editor | Descartada | Escriben en la ventana que esté delante |
| Control de pantalla por capturas (computer use) | Descartada como vía principal | Lenta y frágil. agent-desktop hace lo mismo por accesibilidad |
| Fabricar el `.xlsm` con el VBA desde fuera de Excel | Descartada sin probar | Exige librerías y generar el binario `vbaProject.bin`. El motor lo resuelve con Excel |
| LibreOffice | Descartada | No es Excel; no demuestra nada a la profesora |
| `olevba` (oletools 0.60.2) | Sirve con límite | Probada el 14/11. Lee el VBA sin abrir Excel, también fuera del Mac. Solo lee: no mete ni ejecuta código |
| `xlwings` | No probada | Hay que instalarla. En Mac usa AppleScript por debajo: añade comodidad, no capacidad. Para meter código seguiría necesitando el modelo de objetos de VBA y el ajuste de confianza |
| Complemento Claude para Excel | No probada | Según la profesora no ve el VBA. Sin comprobar |
| COM desde PowerShell (`vba.ps1`) o Python (`pywin32`) | No probada | Es la vía de Windows. Escrita, sin ejecutar. Ver [06-windows.md](06-windows.md) |

Siguiente: [05-buenas-practicas.md](05-buenas-practicas.md).
