# Bitácora de la sesión del 14/11/2026

Registro cronológico de hechos de la sesión de pruebas, en el Mac de quien las hacía (macOS 26.6.2, Excel 16.115,
Python 3.14, Node 20). Es la fuente de los demás documentos: aquí van los hechos con su evidencia;
la explicación ordenada está en los documentos de `docs/`. Horas aproximadas (hora peninsular española).

## 0. Punto de partida

- Clase de VBA del 13/11 (ver `contexto/`, carpeta local, fuera del repositorio; resumen en `docs/01-contexto.md`): la profesora de VBA dijo que Claude Code no puede ejecutar VBA
  porque el código vive dentro de Excel; su problema real es que no puede **validar** el código que Claude le da.
- Llegó un zip con `README.md` (plan con tres vías: COM en Windows, AppleScript en Mac, control de pantalla;
  pruebas M0-M8, W0-W8, C1-C2), `AGENTS.md` y `contexto/` (carpeta local, fuera del repositorio). Nada probado.
- La carpeta traía también material propio de organización, que se queda en local, fuera del repositorio.

## 1. Comprobaciones previas (solo lectura)

- El diccionario AppleScript de Excel se puede leer sin Xcode: `/Applications/Microsoft Excel.app/Contents/Resources/Excel.sdef`.
  Contiene `run VB macro` (devuelve texto, acepta hasta 30 argumentos) y `has vb project`, pero **ningún
  objeto para módulos o código VBA**: AppleScript no puede escribir código; solo ejecutar lo que ya hay.
- `xlwings` y `oletools` no estaban instalados.

## 2. AppleScript básico (00:37-00:38)

- M1: `osascript -e 'tell application "Microsoft Excel" to return version'` → `16.115`. macOS **no pidió**
  permiso de Automatización (ya estaba concedido a la terminal/VS Code de antes).
- M2: escribir 42 en A1 y leerlo → `42.0`.
- Crear `prueba.xlsm`: `save workbook as wb filename "<ruta>" file format macro enabled XML file format`.

## 3. Intento de pulsar teclas con AppleScript (00:39-00:41)

- System Events puede listar y pulsar menús de Excel (permiso de Accesibilidad ya concedido).
- Se pulsó Herramientas > Macro > Editor de Visual Basic. Una captura mostró que **delante estaba VS Code**
  con el usuario escribiendo: las teclas sintéticas van a la ventana activa. Se paró para no escribir en su chat.
- Se escribió `vba/Cargadora.bas` (funciones `Hola` y `Cargar`) para que el usuario la pegara a mano.

## 4. agent-desktop (00:42-00:48)

- Propuesta del usuario: https://github.com/vercel-labs/agent-desktop (Vercel Labs, Apache-2.0, Rust).
  Lee el árbol de accesibilidad de las apps y actúa sobre sus elementos sin capturas de pantalla.
- Instalación: repo clonado en `~/workspaces/agent-desktop` (v1.3.1, commit b0b7277b). Binario descargado de
  GitHub Releases (`agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz`), **checksum SHA-256 verificado** con
  `checksums.txt`, en `~/workspaces/agent-desktop-bin/v1.3.1/`. Sin npm ni instalación global (el paquete npm
  ejecuta un `postinstall`). Sin dependencias de red en el código. Firma ad hoc (sin Team ID).
  `agent-desktop status` → permisos accessibility, automation y screen_recording ya concedidos. Estado en `~/.agent-desktop`.
- Modo por defecto "headless": acciones semánticas (AXPress) sin robar foco ni mover el ratón. `--headed`
  permite teclado/ratón físico; el usuario lo autorizó expresamente.
- Problemas y soluciones:
  - Las ventanas de Excel salían "no visibles" hasta activar Excel.
  - El menú "Editor de Visual Basic" estaba **desactivado**: la barra de estado decía "Introducir" (la celda
    A1 estaba en modo edición). `agent-desktop press escape --app "Microsoft Excel"` (envía la tecla al
    proceso sin darle el foco) lo resolvió.
  - Un clic `--headed` sobre el menú desactivado agotó el tiempo; a las 00:45 hay un informe de cierre de
    **agent-desktop** (SIGABRT, `agent-desktop-2026-11-14-004520.ips`). Excel no se vio afectado.
- Con agent-desktop: se abrió el editor (clic headless), Insertar > Módulo, **el área de código no aparece en
  el árbol de accesibilidad** (solo los dos desplegables), así que se pegó con `clipboard-set` + `--headed press cmd+v`.
  El portapapeles del usuario quedó sustituido. Se guardó con el botón "Guardar prueba.xlsm".
- M3: `run VB macro "prueba.xlsm!Hola"` → `hola desde VBA 00:47:46`. También con `"Hola"` a secas.

## 5. Ajuste de confianza y M4 (00:48-00:50)

- La cargadora devolvió `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es de
  confianza.` (sin ventana, porque la función tiene `On Error`).
- Con agent-desktop: Excel > Preferencias... > Seguridad. Pantalla: "Seguridad de macros" con tres opciones
  (estaba "Deshabilitar todas las macros con notificación") y "Configuración de macro para programadores"
  con la casilla **"Acceso de confianza al modelo de objetos de proyectos de VBA"** (desmarcada). Se marcó
  con `toggle` (verificado: valor 1). No se tocó la seguridad de macros.
- M4: código escrito en la hoja `_codigo` por AppleScript, `run VB macro "prueba.xlsm!Cargar"` →
  `OK Generado 3 lineas`; `Nueva` → `codigo nuevo, 2+3 = 5`. Cambiado y recargado → `version 2, 10*7 = 70`.

## 6. Motor y vba.py (00:50-00:55)

- `vba/Motor.bas`: funciones `CargarEn(libro, modulo)`, `Exportar`, `Modulos`, `Borrar`, `Ejecutar`.
  Trabaja sobre cualquier libro abierto por nombre. Se instaló con la cargadora antigua (sin tocar a mano).
- Descubierto: al escribir código en celdas, Excel **se come el apóstrofo inicial** (comentarios) y trata
  `=` como fórmula. Solución: cada línea va con una marca `|` delante que `CargarEn` quita.
- `herramientas/vba.py` (Python estándar, sin dependencias) pasa argumentos a `osascript` por `argv`
  (sin problemas de comillas).
- Ida y vuelta (`cargar` + `exportar` + `diff`) idéntica, comentarios incluidos.
- M6a: `ConControl` (con `On Error`) → `ERROR 11: División por cero`.

## 7. Cierres de Excel (00:56-01:09)

- M6b: `SinControl` (1/0 sin `On Error`) llamado a través de `Motor.Ejecutar`, que usa `Application.Run`
  dentro de un `On Error`: **no se atrapó**. Salió la ventana "Se ha producido el error '11' en tiempo de
  ejecución: División por cero" con botones Ayuda, Continuar (gris), Finalización, Depurar.
  Mientras está abierta, agent-desktop **no ve ninguna ventana de Excel** (ni `alert`, ni `focused`).
  La llamada de AppleScript queda esperando.
- 1er cierre (00:57): Intro con System Events → Excel se cerró de golpe. Microsoft Error Reporting:
  "There was a problem with Microsoft Excel...". El `osascript` pendiente terminó con `Connection is invalid (-609)`.
  Al relanzar, Excel abrió "Libro1" y después ofreció `prueba (version 1).xlsb` (autorrecuperación, en
  `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/`).
  Se perdió lo no guardado (Motor, hoja `_codigo`); se reinstaló desde los archivos.
- Al reabrir `prueba.xlsm` desde disco por AppleScript, `run VB macro` daba `Parameter error (-50)`.
  Causa: Excel pregunta **"Este libro contiene macros. ¿Desea deshabilitar las macros antes de abrir el
  archivo?"** con "Deshabilitar macros" como botón por defecto. agent-desktop no ve ese diálogo; System Events sí:
  `click button "Habilitar macros" of window 1`. Automatizado en `vba.py abrir`.
- 2º cierre (00:59): misma prueba con vigilancia que pulsa **"Finalización" por su nombre** → Excel se cerró igual.
  Conclusión: no es la tecla; es detener la ejecución.
- Envoltorio: módulo temporal `ZZEnvoltorio` que llama a la macro **directamente** dentro de `On Error`.
  `seguro SinControl` → `ERROR 11: División por cero`, mismo PID de Excel. (Los nombres de módulo no pueden
  empezar por `_`.) El módulo se borra después.
- `SinVentanas.bas`: `MsgBox`/`InputBox` propios del proyecto. `seguro ConMsgBox` → `despues del MsgBox`;
  el texto quedó en la hoja `_avisos`. Mismo PID.
- 3er cierre (01:09): error de compilación (`Option Explicit` + variable no declarada) en una función llamada
  por AppleScript. Salió "Error de compilación"; la vigilancia pulsó "Aceptar" (sin cierre) y el editor quedó en
  "[interrumpir]" marcando la línea. Al pulsar **"Restablecer"** con agent-desktop, Excel se cerró.
- Llamar a una función correcta del mismo módulo (`Bien`) funcionó aunque otra (`Mal`) no compilara.
- `vba.py compilar`: con ninguna macro en marcha, Depurar > "Compilar VBAProject" (agent-desktop) →
  diálogo "Error de compilación: Variable no definida", "Aceptar" con System Events **sin cierre**, y el editor
  muestra "Lín 9, Col 20" (botón legible por accesibilidad). Tras corregir → `OK compila`.
  Una vez, tras `cargar`, el menú Compilar estaba desactivado (Excel lo daba por compilado): la orden ahora
  responde `SIN VERIFICAR` en ese caso.

## 8. Demo y datos (01:05-01:10)

- Un agente preparó `demo/` (sin tocar Excel): herramienta inventada de valor liquidativo, error de base 360/365.
- Al copiar CSV por `value` con el Mac en castellano, `100.0000000000` se convirtió en `1,0E+12`.
  Solución: escribir por `formula`, que Excel interpreta en notación inglesa.
- Ciclo: `FALLO 10 de 15` → `exportar Calculo` → línea 9 `/ 360` → corregir → `cargar` → `OK 15/15` → informe OK.
  Repetido de principio a fin con las órdenes de `GUIA.md`, hoy `docs/03-solucion-mac.md` (abrir desde disco, cargar, compilar, seguro).
- M5: `VBComponents.Import` funcionó desde la carpeta del proyecto y desde la carpeta de Excel, pero si el
  módulo existe crea `Importado1`. **Corrección posterior**: funcionó porque el usuario pulsó "Conceder acceso".

## 9. Permisos que pidió macOS / Excel

- **"Conceder acceso" (sandbox de Excel)**: Excel para Mac está aislado y no puede leer ni escribir fuera de
  su contenedor sin permiso. Cada vez que creó o importó algo en una carpeta nueva, pidió acceso y el usuario lo
  concedió. Constancia en `~/Library/Containers/com.microsoft.Excel/Data/Library/Preferences/com.microsoft.Excel.securebookmarks.plist`:
  entradas para la carpeta del proyecto, `demo`, `vba/pruebas`, los tres `.xlsm` e `Importado.bas`.
- **"Habilitar macros"**: cada vez que se abre un `.xlsm` desde disco con la seguridad en "con notificación".
- **Automatización y Accesibilidad**: ya concedidos antes; no aparecieron.

## 10. Trabajo en paralelo

- Excel en Mac es una sola instancia y atiende las órdenes de una en una: dos agentes manejando Excel se pisan.
  Se usaron agentes solo para trabajo sin Excel: demo (`demo/`) y vía Windows (`docs/windows.md`, hoy `docs/06-windows.md`;
  `herramientas/windows/vba.ps1`, sin probar).

## 11. olevba (01:15)

- `oletools==0.60.2` (PyPI, autor Philippe Lagadec, BSD, publicada 02/07/2024) en un entorno aislado
  `~/workspaces/venvs/oletools`; dependencias congeladas en `~/workspaces/venvs/oletools-freeze.txt` (hoy `herramientas/requirements-olevba.txt`).
- `olevba --code copia.xlsm` saca los módulos sin abrir Excel; encontró `TasaDiaria = ... / 360`.
  `olevba -a` marca "Hex Strings" y "Base64 Strings" como sospechosos: falsos positivos en este libro.
  `olevba -j` da JSON con nombre y código de cada módulo.

## 12. Complemento .xlam, ajuste de confianza por orden y PowerShell (01:18-01:27)

- **Complemento `.xlam` (para no tener que pulsar "Habilitar macros"): NO conseguido.**
  - AppleScript `save workbook as ... file format add in file format`: no da error pero **no crea ningún
    archivo** (buscado en todo el usuario, también en las carpetas de Office). El paso siguiente del script
    falló con `The object you are trying to access does not exist (-1728)`.
  - VBA `Workbooks.Add` + `SaveAs FileFormat:=55` (complemento) desde una función llamada por AppleScript:
    la función devuelve vacío, sin error ni ventana, y tampoco se crea el archivo.
  - Sin comprobar: guardarlo a mano con Archivo > Guardar como > "Complemento de Excel (.xlam)", ponerlo en
    la carpeta de inicio de Office y ver si Excel lo carga sin preguntar por las macros. Queda como hipótesis.
  - Mientras tanto, la vía probada es `vba.py abrir` (pulsa "Habilitar macros" con System Events) y no cerrar
    el libro lanzador durante la sesión.
- **`vba.py confianza activar|desactivar`**: abre Excel > Preferencias > Seguridad con agent-desktop, marca o
  desmarca la casilla y lee su valor para comprobarlo. Probado en los dos sentidos:
  - Desactivada: `vba.py modulos` → `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es de confianza.` (sin ventana).
  - Activada: vuelve a funcionar.
  - **Sin el ajuste se siguen ejecutando las macros que ya existen**: `ejecutar EjecutarTests` → `FALLO 10 de 15`,
    `ejecutar Hola` → `hola desde VBA 01:26:40`. El ajuste solo hace falta para leer o escribir código
    (motor, `cargar`, `exportar`, `modulos`, `seguro`).
  - Buena práctica adoptada: activarlo al empezar a trabajar y desactivarlo al terminar. **Al cerrar la sesión
    del 14/11 quedó desactivado.**
- **PowerShell en el Mac**: el usuario lo instaló durante la sesión para validar la sintaxis de `vba.ps1`.
  A las 01:27 `pwsh` todavía no estaba en el sistema; validación pendiente. En Mac no hay COM, así que PowerShell
  solo serviría para comprobar sintaxis, no para manejar Excel.

## 13. Endurecimiento de vba.py y limpieza de libros (01:45-02:35)

Encargo: corregir los hallazgos de una revisión independiente de `herramientas/vba.py`, probarlos en Excel,
limpiar los `.xlsm` de datos personales antes de subirlos a un repositorio. Un solo agente manejó Excel.
Copias previas de los tres `.xlsm` en el scratchpad de la sesión (`backup-xlsm/`, fuera del repositorio).

### Incidencias durante la sesión

- **Disco lleno (~01:55).** Quedaron unos 120 MB libres e iCloud vació la copia local de la carpeta (leer daba
  "Operation timed out", escribir ENOSPC). La primera reescritura de `vba.py` falló con `ETIMEDOUT` **sin tocar el
  archivo** (fecha y tamaño originales). Se paró todo hasta recuperar espacio (7,6 GB) y comprobar con `testzip`
  que los tres `.xlsm` estaban íntegros. Consecuencia en Excel: diálogo "Error inesperado. Se ha deshabilitado la
  Autorrecuperación para esta sesión de Excel." (un solo botón Aceptar). **En esta sesión de Excel no hay
  autorrecuperación**: hay que reiniciar Excel para recuperarla.
- **`save workbook as` desde AppleScript da `Parameter error (-50)`** al crear un libro nuevo en `vba/pruebas/`,
  con apóstrofo en el nombre y sin él (`prueba Enero.xlsm`), y también con la ruta en formato HFS. No se creó
  ningún archivo. Probable relación con el diálogo anterior o con el sandbox; sin confirmar. El libro en blanco
  `prueba Enero.xlsm` quedó abierto y **no se puede cerrar** (`close ... saving no` → -50; VBA
  `Workbooks(...).Close False` → `ERROR 1004: Fallo en el método 'Close'`). No tiene archivo en disco;
  desaparece al reiniciar Excel.
- **Alguien cerró a mano dos `MsgBox` de prueba** mientras la vigilancia esperaba. Esas pruebas no valen y se
  repitieron.

### Hallazgos nuevos (comprobados)

- **Cómo ve System Events un `MsgBox` real**: título de ventana vacío, subrol `AXDialog`, botón `Aceptar`,
  textos `Microsoft Excel` (el título del aviso, como primer texto) y el mensaje.
- **La ventana principal del editor de VBA** ("Microsoft Visual Basic - errores.xlsm [en ejecución]") también
  contiene "Visual Basic", pero es `AXStandardWindow`. Si la vigilancia mirara todas las ventanas, no cerraría
  nunca un `MsgBox` con el editor abierto. Por eso solo cuentan los diálogos (`AXDialog`).
- **Aviso de compilación real**: textos `Microsoft Visual Basic` / `Error de compilación: Variable no definida`.
  Confirma que los avisos de VBA llevan "Visual Basic".
- **Con Excel en segundo plano, System Events ve 0 ventanas de Excel**, ni siquiera un `MsgBox` abierto.
  agent-desktop tampoco. Al ponerlo delante, macOS cambia de escritorio y tarda en mostrar sus ventanas: una
  espera de 0,3-0,5 s no bastaba (la vigilancia devolvía `ERROR TIEMPO ... Dialogos: []` con el aviso abierto).
  `tell application "Microsoft Excel" to activate` a veces no lo trae delante con el usuario en VS Code, y
  entonces los menús de Excel salen desactivados (`compilar` respondía `ERROR no encuentro Depurar > Compilar`).
- **`compilar` compilaba el proyecto equivocado.** Depurar > Compilar actúa sobre el proyecto activo del
  editor (el de la ventana de código con el foco), no sobre el libro activado con `activate object workbook`.
  Con `--libro vl-fondo.xlsm` devolvió
  `ERROR COMPILACION: ... Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`.
  **Las salidas `OK compila` anteriores a este cambio pueden ser de otro libro.** El título del editor
  ("Microsoft Visual Basic - libro.xlsm") indica el proyecto activo. `Set Application.VBE.ActiveVBProject` no
  cambió el título; `VBComponents(x).CodeModule.CodePane.Show` sí.
- **agent-desktop**: los identificadores de `list-windows` (sin `--app`) dieron `WINDOW_NOT_FOUND` en
  `snapshot --window-id`, y `list-windows --app "Microsoft Excel"` devolvió a veces una lista vacía. Por eso la
  posición del error se lee ahora con `--surface focused`.
- **Excel vuelve a escribir el autor y la ruta al guardar.** Tras limpiar `vl-fondo.xlsm`, la demo (que guarda
  con `cargar`) dejó otra vez el nombre del autor en `core.xml` y `x15ac:absPath` en `workbook.xml`. La limpieza hay
  que repetirla con los libros cerrados justo antes de subirlos.
- **Saltos de pantalla ("latigazos")**: traen Excel al frente `compilar`, `confianza`, `abrir` (y crear libros),
  `dialogos`, la vigilancia y cualquier `MsgBox` que salga durante una macro.

### Cambios en `herramientas/vba.py`

Mismas órdenes y argumentos; comentarios en castellano.

| Punto | Cambio |
|---|---|
| a | `exportar`: si el resultado empieza por `ERROR` o `AVISO`, no escribe `--salida`, imprime el error y sale con código 1 |
| b | `macro()`: `communicate(input=..., timeout=...)` en lugar de `wait()`; tras el tiempo agotado, vigila y recoge con `communicate` (sin bloqueo con salidas grandes) |
| c | Vigilancia: lee título, subrol, botones y textos (`ventanas()`). Solo cuentan los diálogos `AXDialog`. Un diálogo con "Visual Basic", "Finalización", "Depurar", "compilaci", "tiempo de ejecuci" (o equivalentes en inglés) es de VBA: no se toca, se mata solo `osascript` y devuelve `ERROR VENTANA VBA (no se ha tocado): ...`. Solo pulsa Aceptar/OK en un diálogo cuyo título o primer texto es "Microsoft Excel" |
| d | `pulsar()` captura los fallos y devuelve si pulsó; tras pulsar espera con tiempo límite; si sale otro aviso, repite (máximo 10). La salida incluye lo que devolvió la macro |
| e | Por argv: libro y macro en `macro()` (la referencia `'libro'!macro` duplica los apóstrofos), `guardar`, `compilar` (`activate object workbook`), `pulsar` (botón, título y texto) y `abrir` (ruta) |
| f | `abrir`: comprueba que el archivo existe, código de salida de `osascript` y tiempo agotado; reconoce "Habilitar macros"/"Enable Macros" y, con `--sin-macros`, "Deshabilitar macros"/"Disable Macros" |
| g | `compilar`: reconoce "compilaci" y "Compile error"; sondea hasta 6 s; snapshot fallido sin `KeyError`; usa Aceptar u OK según exista |
| h | `ad()`: comprueba que el binario existe (mensaje con la ruta y `AGENT_DESKTOP`), código de salida y JSON válido |
| i | `cargar` se niega a reemplazar `Motor`, `Módulo1` o `Cargadora` en el lanzador |
| j | `seguro`: si el envoltorio devuelve `ERROR VENTANA` o `ERROR TIEMPO`, no intenta borrar y avisa de que `ZZEnvoltorio` queda en el libro; avisa también si el borrado falla |
| k | `--timeout SEGUNDOS` (60 por defecto) en `ejecutar`, `seguro`, `cargar`, `exportar` |
| l | `osa()` captura el tiempo agotado con mensaje claro |
| m | `confianza`: maneja la casilla ilegible tras el cambio; cierra con "Close" o "Cerrar" |
| Nuevo | Antes de leer o pulsar, pone Excel delante (`frontmost` por System Events, sin teclas) y espera hasta 3 s a que se vean sus ventanas |
| Nuevo | `compilar` activa primero en el editor el proyecto pedido (módulo temporal `ZZEnvoltorio` con `CodePane.Show`, que se borra) y comprueba el título del editor; si no coincide, no compila. Responde `OK compila <libro>`. Si el menú Compilar no aparece, lo reintenta una vez |
| Nuevo | Devolución del foco: `compilar`, `confianza`, `abrir`, `dialogos` y la vigilancia guardan la aplicación que estaba delante y la reactivan al terminar, también con error, salvo que fuera Excel. En la vigilancia, la aplicación se toma al lanzar la macro, porque el `MsgBox` ya ha traído Excel al frente |
| Nuevo | `seguro` usa una función común `temporal()` (meter, ejecutar y borrar `ZZEnvoltorio`) |

`demo/ejecutar_demo.py` (guion de otro agente): se añadió una parada si `compilar` devuelve `ERROR COMPILACION`
o `ERROR VENTANA`, para no ejecutar los tests sobre código que no compila.

### Pruebas (salida literal)

| Prueba | Salida |
|---|---|
| `vba confianza activar` | `OK confianza activada (antes: activada)` |
| `vba ejecutar Hola` | `hola desde VBA 02:08:34` |
| `vba modulos` | `ThisWorkbook 100 0`, `Hoja1 100 0`, `Módulo1 1 34`, `Hoja2 100 0`, `Motor 1 87`, `PruebaErrores 1 25`, `SinVentanas 1 35`, `Hoja3 100 0` |
| `vba exportar Calculo --libro vl-fondo.xlsm --salida /tmp/x.bas` | `OK Calculo -> /tmp/x.bas`, código 0, idéntico a `demo/vba/Calculo.bas` sin las líneas `Attribute` |
| `vba exportar NoExiste --libro vl-fondo.xlsm --salida /tmp/y.bas` | `ERROR 9: Subíndice fuera de rango`, código 1, `/tmp/y.bas` no creado |
| `vba cargar vba/Motor.bas` | `ERROR no reemplazo Motor en el lanzador prueba.xlsm: es el codigo que hace la carga. Reinstalalo con la cargadora o a mano.` (código 1) |
| `vba cargar vba/Cargadora.bas --modulo Módulo1` | `ERROR no reemplazo Módulo1 en el lanzador prueba.xlsm: ...` (código 1) |
| `vba abrir vba/pruebas/noexiste.xlsm` | `ERROR no existe el archivo: .../vba/pruebas/noexiste.xlsm` (código 1) |
| Apóstrofo: `vba cargar vba/Cargadora.bas --libro "prueba d'Enero.xlsm" --modulo Saludo --sin-guardar` | `OK Saludo 34 lineas en prueba d'Enero.xlsm` |
| `vba ejecutar Hola --libro "prueba d'Enero.xlsm"` | `hola desde VBA 02:12:06` |
| `vba seguro Hola --libro "prueba d'Enero.xlsm"` | `hola desde VBA 02:12:07` (`ZZEnvoltorio` borrado después) |
| `vba abrir "vba/pruebas/prueba d'Enero.xlsm"` (copia de `errores.xlsm`) | `OK abierto prueba d'Enero.xlsm`; `vba ejecutar Bien --libro "prueba d'Enero.xlsm"` → `compila` |
| `close workbook` por argv con apóstrofo | cerrado; el archivo de prueba se borró después |
| `MsgBox` real, `vba ejecutar ConAviso --libro errores.xlsm --timeout 10` | `AVISO cerrado con Aceptar: Microsoft Excel / ventana de prueba (usa SinVentanas para evitarlo). Resultado: despues del MsgBox` |
| Dos `MsgBox` seguidos, `vba ejecutar DosAvisos --libro errores.xlsm --timeout 10` | `AVISO cerrado con Aceptar: Microsoft Excel / primer aviso; Microsoft Excel / segundo aviso (usa SinVentanas para evitarlo). Resultado: despues de dos MsgBox`; delante antes y después: `Code` |
| Salida grande (función que devuelve 100 003 caracteres) | 100 004 bytes con salto de línea, termina en `FIN`, sin bloqueo |
| `vba compilar --libro vl-fondo.xlsm` (antes de corregir la selección de proyecto) | `ERROR COMPILACION: Microsoft Visual Basic / Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20` (proyecto equivocado) |
| `vba compilar --libro vl-fondo.xlsm` (corregido) | `OK compila vl-fondo.xlsm`; delante antes y después: `Code` |
| `vba compilar --libro errores.xlsm` (corregido) | `ERROR COMPILACION: Microsoft Visual Basic / Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`; delante antes y después: `Code` |
| Ciclo de la demo a mano | `cargar demo/vba/Calculo.bas` → `OK Calculo 52 lineas en vl-fondo.xlsm`; `compilar` → `OK compila vl-fondo.xlsm`; `seguro EjecutarTests` → `FALLO 10 de 15: 2025-03-10, ..., 2025-03-21`; `cargar demo/solucion/Calculo_corregido.bas --modulo Calculo` → `OK Calculo 53 lineas`; `compilar` → `ERROR no encuentro Depurar > Compilar en el editor` (intermitente; repetido dos veces después: `OK compila vl-fondo.xlsm`; se añadió un reintento); `seguro EjecutarTests` → `OK 15/15`; `cargar demo/vba/Calculo.bas` → `OK Calculo 52 lineas` |
| `python3 -I demo/ejecutar_demo.py` | `OK compila vl-fondo.xlsm` (x2), `FALLO 10 de 15: ...`, `linea 9: TasaDiaria = ComisionPctDia(fila) / 100 / 360`, `OK 15/15`, `OK Informe 15 dias`, `DEMO CORRECTA`, código 0; delante antes y después: `Code` |
| `python3 -I demo/ejecutar_demo.py --sin-compilar` | mismos resultados sin compilar, `DEMO CORRECTA`, código 0 |
| `vba confianza desactivar` | `OK confianza desactivada (antes: activada)`; después `vba modulos` → `ERROR 1004: El acceso mediante programación al proyecto de Visual Basic no es de confianza.` |

Sin probar: un error de ejecución real con la vigilancia (no se provoca: cerrarlo tumba Excel). Su
clasificación se comprobó con los textos de la sección 7. El texto exacto del aviso "Habilitar macros" en
System Events (`abrir` lo sigue pulsando: los libros abren con macros).

### Limpieza de los `.xlsm`

- `prueba.xlsm`: hojas `_codigo` y `_avisos` vaciadas y guardado. `errores.xlsm`: borrado el módulo de prueba
  `Aviso`; quedan `NoCompila` y las hojas.
- Con los libros cerrados, reescritura del ZIP con Python (`zipfile`): `docProps/core.xml` con `dc:creator` y
  `cp:lastModifiedBy` vacíos; en `xl/workbook.xml` eliminado `mc:AlternateContent` con `x15ac:absPath`
  (`/Users/<usuario>/.../`, `.../demo/`, `.../vba/pruebas/`). Resto de entradas comprobadas idénticas
  (contenido, orden), `testzip` correcto. `vl-fondo.xlsm` se limpió dos veces (la demo lo vuelve a guardar).
- Reabiertos con `vba abrir`: sin aviso de reparación, `vba ejecutar Hola` → `hola desde VBA 02:31:36`,
  `vba seguro EjecutarTests --libro vl-fondo.xlsm` → `FALLO 10 de 15: ...`. Comprobado después: ni el nombre del autor
  ni rutas de usuario en `core.xml` ni en `workbook.xml` de los tres.
- **Sin tocar:** `xl/vbaProject.bin` de los tres contiene la referencia de MSForms
  `{...-DE1B671906D3}#2.0#0#/Users/<usuario>/Library/Containers/...`.

### Estado al cerrar (02:35)

- Abiertos y guardados: `prueba.xlsm` y `vl-fondo.xlsm` (`Calculo` con el error `/ 360`). También
  `prueba Enero.xlsm`, en blanco y sin archivo, que no se puede cerrar.
- Ajuste de confianza **desactivado**.
- Pendiente: reiniciar Excel (autorrecuperación desactivada en esta sesión y libro `prueba Enero.xlsm`
  bloqueado). Al cerrarlo, descartar `prueba Enero.xlsm`. Si Excel guarda de nuevo los libros, repetir la
  limpieza antes de subirlos.

## 14. Copia "después" de la demo y fallo de `confianza` con el editor delante (02:40-02:52)

- **`vba.py confianza` fallaba de forma intermitente** con `ERROR no encuentro la casilla de confianza`.
  Causa encontrada con una captura: si delante está el **editor de VBA** (por ejemplo, justo después de
  `compilar`), el menú "Preferencias..." abre las **Opciones del editor de VBA**, no las Preferencias de
  Excel. Además, con Excel en segundo plano agent-desktop no veía la ventana con el foco.
  - Arreglo: `confianza` comprueba que la barra de menús no es la del editor (no tiene "Depurar"); si lo
    es, pulsa el botón "Ver Microsoft Excel" de la barra del editor (`activate object window 1` de
    AppleScript no bastó). La búsqueda de "Seguridad" y de la casilla trae Excel delante, espera hasta
    6 s y, si no ve la ventana con el foco, busca la ventana de Preferencias por su título.
  - Si falla, ahora sale con código 1 (antes devolvía 0 con el texto de error).
  - Probado: `OK confianza activada (antes: desactivada)`, `OK confianza activada (antes: activada)` y,
    con el editor delante tras compilar, `OK confianza desactivada (antes: activada)`; después,
    `modulos` → `ERROR 1004` (desactivado de verdad).
- **`demo/vl-fondo-corregido.xlsm`** (el "después" de la demo): copia de `vl-fondo.xlsm` con
  `Calculo` corregido. Salidas: `OK Calculo 53 lineas en vl-fondo-corregido.xlsm`,
  `OK compila vl-fondo-corregido.xlsm`, `OK 15/15`, `OK Informe 15 dias`; la hoja Tests tiene las
  15 filas en OK y `TasaDiaria = ComisionPctDia(fila) / 100 / 365`. Se abre en la hoja Informe.
- Excel cerrado con `quit saving no`; antes, la salida dio `User cancelled (-128)` porque el diálogo de
  Opciones del editor seguía abierto (se cerró con "Cancelar", sin macro en marcha).
- Metadatos limpiados con Excel cerrado en los cuatro `.xlsm` (autor vacío, sin `absPath`); ZIP íntegros.
- Ajuste de confianza: **desactivado** al terminar.

## 15. Revisión de coherencia y ejecución final de la demo (02:55-03:10)

- Revisión independiente de todo el repositorio (solo lectura) y corrección de la documentación: datos
  personales neutralizados, libro "después" explicado, arreglo de `confianza` documentado, `GUIA.md`
  borrado, repeticiones convertidas en enlaces.
- Hallazgo: `demo/vl-fondo.xlsm` tenía el código con el error pero sus hojas `Tests` e `Informe` mostraban
  el resultado bueno (15 OK) de la última ejecución, porque la demo volvía a cargar el código con el error
  sin repetir los tests. Arreglo en `demo/ejecutar_demo.py`: tras volver a la versión con el error, repite
  `seguro EjecutarTests` y `seguro GenerarInforme` y guarda.
- Ejecución final de `python3 -I demo/ejecutar_demo.py` (con `confianza activar` antes y `desactivar`
  después): `FALLO 10 de 15` -> `OK compila vl-fondo.xlsm` -> `OK 15/15` -> `OK Informe 15 dias`; paso 5:
  `OK Calculo 52 lineas en vl-fondo.xlsm`, `FALLO 10 de 15`, `OK Informe 15 dias`, `OK guardado vl-fondo.xlsm`;
  resultado `DEMO CORRECTA`.
- Órdenes con las que se creó `demo/vl-fondo-corregido.xlsm`: `cp demo/vl-fondo.xlsm demo/vl-fondo-corregido.xlsm`;
  `vba abrir prueba.xlsm`; `vba abrir demo/vl-fondo-corregido.xlsm`;
  `vba cargar demo/solucion/Calculo_corregido.bas --libro vl-fondo-corregido.xlsm --modulo Calculo`;
  `vba compilar`, `vba seguro EjecutarTests`, `vba seguro GenerarInforme` y `vba guardar` con
  `--libro vl-fondo-corregido.xlsm`; hoja Informe activada con AppleScript antes de guardar.
- Cierre: ajuste de confianza desactivado, Excel cerrado, metadatos de los cuatro `.xlsm` limpiados de
  nuevo (Excel los reescribe al guardar), sin archivos de bloqueo.
