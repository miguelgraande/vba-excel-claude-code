# 03 · La solución en Mac, pieza a pieza

Stack probado el 14/11/2026 en macOS 26.6.2 con Excel 16.115 (interfaz en castellano) y Python 3.14.
Cómo se llegó a él: [02-recorrido.md](02-recorrido.md). Evidencia de cada prueba: [04-resultados.md](04-resultados.md).

## Vista general

```text
Claude Code (terminal)
   |  escribe y lee archivos .bas
   v
herramientas/vba.py  ----------------------------+
   |  osascript (AppleScript)                    |  agent-desktop / System Events
   v                                             v  (solo menús y ventanas)
Excel para Mac
   |-- prueba.xlsm (libro lanzador)
   |     |-- Motor.bas: carga, exporta, lista y borra módulos
   |     '-- hoja _codigo: el código viaja aquí, una línea por celda
   '-- libro de trabajo (por ejemplo demo/vl-fondo.xlsm)
         |-- módulos del usuario
         |-- ZZEnvoltorio (temporal, lo pone y lo quita `seguro`)
         '-- SinVentanas (opcional)
```

La idea en tres frases:

1. AppleScript no puede escribir código VBA, pero sí escribir celdas y ejecutar macros.
2. Así que el código viaja por celdas y una macro ya instalada (el motor) lo mete en el módulo.
3. Todo lo que se ejecuta devuelve texto (`OK ...` o `ERROR ...`) y nunca abre ventanas, porque
   detener una macro parada cierra Excel.

## Las piezas

### AppleScript (`osascript`)

- **Qué es.** El lenguaje de automatización de macOS. `osascript` lo ejecuta desde la terminal. Viene con macOS.
- **Qué hace aquí.** Toda la conversación con Excel: escribir y leer celdas, crear hojas, abrir y
  guardar libros, y ejecutar macros con `run VB macro`.
- **Por qué así.** En Mac no existe COM. AppleScript es el canal que Excel para Mac ofrece hacia fuera
  y no hay que instalar nada.
- **Límite.** El diccionario de Excel (`Excel.sdef`) no tiene ningún objeto para módulos ni código.
  AppleScript no puede escribir VBA: solo ejecutar lo que ya hay. Microsoft no documenta `run VB macro`.
- **Descartado.** COM (solo Windows). `xlwings`: en Mac usa AppleScript por debajo, no añade capacidad
  y hay que instalarlo (no probado).

### `run VB macro`

- **Qué es.** Orden del diccionario AppleScript de Excel. Ejecuta una macro del libro y devuelve su
  valor como texto. Acepta hasta 30 argumentos.
- **Qué hace aquí.** Es la única puerta de entrada al VBA. `vba.py` la usa con la forma
  `run VB macro "'libro.xlsm'!Nombre" arg1 (...) arg2 (...)`. También funcionaron `"prueba.xlsm!Hola"`
  y `"Hola"` a secas.
- **Por qué así.** Devuelve el valor de una `Function`. Por eso todas las macros que se llaman desde la
  terminal son funciones que devuelven texto, y no `Sub`.
- **Señal de fallo.** `Parameter error (-50)`: casi siempre la macro no existe o las macros están
  desactivadas en ese libro.

### Hoja `_codigo` y la marca `|`

- **Qué es.** Una hoja del libro lanzador. En la columna A, una línea de código por celda, cada una con
  el carácter `|` delante.
- **Qué hace aquí.** Es el buzón por el que viaja el código: `vba.py` escribe las líneas con AppleScript
  y el motor las lee, quita el primer carácter y las mete en el módulo.
- **Por qué la marca.** Al escribir código en celdas, Excel **se come el apóstrofo inicial** (los
  comentarios desaparecen) y trata una línea que empieza por `=` como fórmula. Con `|` delante, ninguna
  línea empieza por `'` ni por `=`. El motor quita siempre el primer carácter (`Mid(..., 2)`).
- **Detalle.** `vba.py cargar` quita antes las líneas `Attribute VB_...`, que Excel añade al exportar
  módulos a archivo.
- **Descartado.** Escribir las líneas sin marca (pierde comentarios). Importar desde archivo
  (`VBComponents.Import`): pide "Conceder acceso" y, si el módulo existe, crea `Nombre1` al lado.
  Pulsar teclas en el editor: las teclas van a la ventana activa.

### Motor (`vba/Motor.bas`)

- **Qué es.** Un módulo VBA que vive en el libro lanzador `prueba.xlsm`. Código: [vba/Motor.bas](../vba/Motor.bas).
- **Qué hace aquí.** Cinco funciones que actúan sobre **cualquier libro abierto**, nombrado por su
  nombre (vacío = el propio lanzador):

  | Función | Hace |
  |---|---|
  | `CargarEn(libro, modulo)` | Lee `_codigo`, quita la marca y mete el código. Si el módulo existe, vacía sus líneas y escribe encima (no crea `Nombre1`). Si no existe, lo crea |
  | `Exportar(libro, modulo)` | Devuelve el código del módulo como texto |
  | `Modulos(libro)` | Lista nombre, tipo y número de líneas de cada componente |
  | `Borrar(libro, modulo)` | Elimina un módulo |
  | `Ejecutar(libro, macro)` | `Application.Run` dentro de un `On Error`. **No atrapa** los errores de la macro llamada (1er cierre de Excel). `vba.py` no la usa |

- **Por qué así.** Separa la herramienta del libro de trabajo: el libro que se corrige no necesita
  llevar el motor dentro. Todas las funciones devuelven `OK ...` o `ERROR n: ...`, sin ventanas.
- **Protección.** `vba.py cargar` se niega a reemplazar `Motor`, `Módulo1` o `Cargadora` en el libro
  lanzador: el motor se estaría reescribiendo a sí mismo mientras se ejecuta, y la cargadora es la vía
  para reinstalarlo. Probado el 14/11: `vba cargar vba/Motor.bas` → `ERROR no reemplazo Motor en el
  lanzador prueba.xlsm: ...`, código 1. Para actualizar el motor, ver "Si el
  libro lanzador se pierde o se estropea" en [09-instalacion.md](09-instalacion.md#si-el-libro-lanzador-se-pierde-o-se-estropea).
- **Requisito.** El ajuste "Acceso de confianza al modelo de objetos de proyectos de VBA". Sin él,
  devuelve `ERROR 1004`. Ver [05-buenas-practicas.md](05-buenas-practicas.md#acceso-de-confianza-al-modelo-de-objetos-de-proyectos-de-vba).
- **Descartado.** Usar solo la cargadora (trabaja solo sobre su propio libro y no quita la marca).

### Cargadora (`vba/Cargadora.bas`)

- **Qué es.** La primera macro del lanzador. Es la única que hay que meter a mano (o con agent-desktop),
  una sola vez. Código: [vba/Cargadora.bas](../vba/Cargadora.bas).
- **Qué hace aquí.** `Hola` devuelve un texto con la hora (prueba M3). `Cargar` lee `_codigo` (nombre
  del módulo en B1, código en la columna A), borra el módulo si existe y lo vuelve a crear con
  `AddFromString`. Sirvió para instalar el motor sin tocar el editor.
- **Por qué así.** Resuelve el huevo y la gallina: para cargar código hace falta una macro que ya esté
  dentro. Es mínima para que pegarla a mano sea fácil.
- **Límite.** Solo trabaja sobre su propio libro y no quita la marca `|`: un apóstrofo inicial se
  pierde. Probable motivo, no documentado, de que `Motor.bas` use `Rem` en vez de `'` en sus comentarios.

**Intento descartado por ahora: motor en un complemento `.xlam`.** La idea era guardar el motor en un
complemento para que Excel no preguntara por las macros al abrir el lanzador. No se consiguió crear el
archivo desde la terminal (14/11, 01:18-01:27):

- AppleScript `save workbook as ... file format add in file format` no da error pero no crea ningún
  archivo. El paso siguiente falló con `The object you are trying to access does not exist (-1728)`.
- VBA `Workbooks.Add` + `SaveAs FileFormat:=55` desde una función llamada por AppleScript devuelve vacío,
  sin error ni ventana, y tampoco crea el archivo.
- Sin comprobar: guardarlo a mano (Archivo > Guardar como > "Complemento de Excel (.xlam)"), ponerlo en la
  carpeta de inicio de Office y ver si Excel lo carga sin preguntar.

Mientras tanto: `vba.py abrir` y no cerrar el lanzador durante la sesión.

### `herramientas/vba.py`

- **Qué es.** Script de Python estándar, sin dependencias. Código y ayuda: [herramientas/vba.py](../herramientas/vba.py).
- **Qué hace aquí.** Junta todo en órdenes cortas para la terminal. Lista completa en
  [Referencia de órdenes](#referencia-de-órdenes).
- **Por qué así.**
  - Pasa los argumentos a `osascript` por `argv`, no dentro del texto del script. Así no hay problemas
    de comillas ni de caracteres especiales en el código. Los nombres de libro y de macro, las rutas y
    los botones también viajan como argumentos. En la referencia `'libro'!macro` los apóstrofos del
    nombre se duplican. Probado el 14/11 con `prueba d'Enero.xlsm`: `abrir`, `cargar`, `ejecutar` y
    `seguro` funcionaron.
  - Python lee y escribe los archivos `.bas`. Excel nunca abre esos archivos: el código entra por celdas
    y sale como valor devuelto. Por eso `cargar` y `exportar` no dependen del sandbox de archivos de Excel.
  - `--libro` nombra siempre el libro de trabajo. Sin él, la orden actúa sobre el lanzador. Nunca usa
    "el libro activo" (prueba M8).
- **Límites del código.**
  - Busca menús y botones por su texto en castellano: "Editor de Visual Basic", "Depurar", "Compilar",
    "Habilitar macros", "Aceptar", "Error de compilación", "(Código)", "Lín", "Preferencias",
    "Seguridad", "Acceso de confianza", "Ver Microsoft Excel" (botón del editor) y el título del editor
    "Microsoft Visual Basic - libro". Reconoce también en inglés "Enable Macros" y "Disable Macros"
    (`abrir`), "Compile error" (`compilar`), "OK" (avisos), "Security", "Trust access to the VBA project",
    "View Microsoft Excel" y "Close" o "Cerrar" (`confianza`). El resto, con Excel en otro idioma,
    necesitaría cambiar esos textos.
  - El nombre del lanzador está fijo en el código: `prueba.xlsm`.
  - La ruta de agent-desktop se puede cambiar con la variable de entorno `AGENT_DESKTOP`. Si no está,
    la orden lo dice con la ruta buscada.
  - Tiempo de espera: `ejecutar`, `seguro`, `cargar` y `exportar` aceptan `--timeout SEGUNDOS`
    (por defecto 60). Para una macro larga, subirlo antes de que salte la vigilancia. La espera ya no
    se bloquea con salidas grandes: probado con una función que devuelve 100 000 caracteres.
  - **Excel delante para ver sus ventanas.** Con Excel en segundo plano, System Events ve 0 ventanas
    de Excel, ni siquiera un `MsgBox` abierto (comprobado el 14/11; agent-desktop tampoco). Por eso,
    antes de leer o pulsar, `vba.py` pone Excel delante (`frontmost` por System Events, sin teclas) y
    espera hasta 3 s a que se vean sus ventanas: al cambiar de escritorio, macOS tarda en mostrarlas y
    0,3-0,5 s no bastaban. `activate` no se usa para esto: con el usuario en VS Code, a veces no traía
    Excel delante y los menús salían desactivados.
  - **Devolución del foco.** `compilar`, `confianza`, `abrir`, `dialogos` y la vigilancia guardan la
    aplicación que estaba delante y la vuelven a poner delante al terminar, también si hay error (salvo
    que fuera Excel). En la vigilancia se toma la aplicación al lanzar la macro, porque el `MsgBox` ya
    ha traído Excel al frente. Probado: con VS Code delante antes, VS Code delante después.
  - Vigilancia de ventanas: si una macro no responde en el tiempo de espera, lee las ventanas de Excel
    con System Events (título, subrol, botones y textos). **Solo mira los diálogos** (subrol
    `AXDialog`). La ventana principal del editor ("Microsoft Visual Basic - libro.xlsm [en ejecución]")
    también lleva "Visual Basic", pero es una ventana normal (`AXStandardWindow`): si contara, la
    vigilancia no cerraría nunca un `MsgBox` con el editor abierto.
    - Diálogo cuyo texto contiene "Visual Basic", "Finalización", "Depurar", "compilaci" o "tiempo de
      ejecuci" (o "Debug", "Compile error", "Run-time error" en inglés): es de VBA y **nunca pulsa
      nada**. Mata solo el proceso `osascript` y devuelve `ERROR VENTANA VBA (no se ha tocado): ...`.
      La macro queda parada en Excel.
    - Aviso de Excel: diálogo con "Aceptar" u "OK" cuyo título, o su primer texto, es "Microsoft Excel"
      (así ve System Events un `MsgBox`: título vacío y "Microsoft Excel" como primer texto). Es lo
      único que pulsa, y solo "Aceptar" u "OK". Después espera a que la macro termine; si sale otro
      aviso, lo cierra también, hasta 10 seguidos. Devuelve lo cerrado y lo que devolvió la macro:
      `AVISO cerrado con Aceptar: Microsoft Excel / ... (usa SinVentanas para evitarlo). Resultado: ...`.
    - Ningún diálogo conocido: devuelve `ERROR TIEMPO: sin respuesta en N s. Dialogos: [...]`.

    Hasta el 14/11 la vigilancia pulsaba "Finalización", y eso cerró Excel (2º cierre); se corrigió ese
    mismo día. Es un último recurso para no dejar la terminal colgada. La protección real es
    `compilar` + `seguro`.
- **Descartado.** `osascript -e` con el código metido en el texto del script (comillas). `xlwings`
  (no probado, no añade capacidad en Mac).

### Envoltorio `ZZEnvoltorio` (`vba.py seguro`)

- **Qué es.** Un módulo temporal que `seguro` genera, mete en el libro de trabajo, ejecuta y borra:

  ```vba
  Public Function Envoltorio() As String
      On Error GoTo Fallo
      Envoltorio = CStr(MiFuncion())      ' con --sub: Call MiSub y devuelve "OK"
      Exit Function
  Fallo:
      Envoltorio = "ERROR " & Err.Number & ": " & Err.Description
  End Function
  ```

- **Qué hace aquí.** Un error sin controlar en la macro sube hasta el `On Error` del envoltorio y vuelve
  a la terminal como texto, sin ventana de VBA.
- **Por qué así.** La llamada es **directa**, dentro del mismo proyecto. `Application.Run` dentro de un
  `On Error` (Motor.Ejecutar) no atrapó el error: salió la ventana igual. Con la llamada directa, sí:
  `seguro SinControl` → `ERROR 11: División por cero`, mismo PID de Excel.
- **Por qué ese nombre.** VBA no admite nombres de módulo que empiecen por `_`.
- **Límites.** No pasa argumentos. Llama a la macro por su nombre dentro del libro de trabajo. No
  protege de errores de **compilación**: para eso, `compilar` antes. Si la vigilancia devuelve
  `ERROR VENTANA` o `ERROR TIEMPO`, Excel está bloqueado: `seguro` no intenta borrar `ZZEnvoltorio` y
  avisa de que queda en el libro (`vba.py borrar ZZEnvoltorio --libro X` cuando Excel responda).
  También avisa si el borrado falla.
- **Descartado.** `Application.Run` con `On Error` (no atrapa). Cerrar la ventana de error después
  (cierra Excel).

### SinVentanas (`vba/SinVentanas.bas`)

- **Qué es.** Un módulo con `MsgBox` e `InputBox` propios. Código: [vba/SinVentanas.bas](../vba/SinVentanas.bas).
- **Qué hace aquí.** VBA busca los nombres primero en el propio proyecto, así que el código existente
  llama a estas versiones sin cambiar ni una línea. Cada llamada se apunta en la hoja `_avisos`
  (fecha, tipo, texto). `MsgBox` devuelve `vbOK`; `InputBox` devuelve el valor por defecto o vacío.
- **Por qué así.** Un `MsgBox` para la macro hasta que alguien pulse, y AppleScript se queda esperando.
  Un código antiguo tendrá muchos. Así no hay que reescribirlos.
- **Cómo se usa.** Cargarlo en el libro de trabajo: `vba.py cargar vba/SinVentanas.bas --libro X.xlsm`.
- **Límites.** No cubre `Application.InputBox`, formularios (`UserForm`) ni los diálogos de Excel.
  Sin probar con un caso así. `MsgBox` devuelve siempre `vbOK`: un `MsgBox` de Sí/No toma la rama de
  OK, que puede no ser la que se quería probar. Una llamada escrita como `VBA.MsgBox` se salta la
  sustitución y abre ventana. Además añade la hoja `_avisos` al libro de trabajo. Lista completa en
  [05-buenas-practicas.md](05-buenas-practicas.md#e-límites-conocidos).
- **Descartado.** `DisplayAlerts = False`: solo quita los avisos de Excel, no los `MsgBox` del código
  (documentación de Microsoft, ver [06-windows.md](06-windows.md#7-ventanas-modales-qué-se-puede-y-qué-no)).

### `compilar` (`vba.py compilar`)

- **Qué es.** Una orden que compila el proyecto VBA del libro **sin ejecutar nada**.
- **Qué hace aquí.**
  1. Si ya hay un diálogo de VBA abierto (macro parada), no hace nada y devuelve `ERROR VENTANA VBA (no
     se ha tocado): ...`.
  2. Activa en el editor de VBA el proyecto del libro pedido: mete un módulo temporal `ZZEnvoltorio`
     que muestra un módulo de ese libro (`CodePane.Show`) y lo borra.
  3. Activa el libro, pone Excel delante y, con agent-desktop, abre el editor de VBA.
  4. Lee el título del editor ("Microsoft Visual Basic - libro.xlsm"). Si el proyecto activo no es el
     pedido, **no compila** y devuelve `ERROR el editor tiene activo el proyecto de X, no el de Y`.
  5. Pulsa Depurar > Compilar VBAProject. Si no encuentra el menú, vuelve a abrir el editor y lo
     reintenta una vez.
  6. Espera hasta 6 s a que salga el aviso. Si sale "Error de compilación" (o "Compile error" con Excel
     en inglés), lee el mensaje con System Events, pulsa "Aceptar" (u "OK") y lee en el editor el
     módulo y la posición. Ejemplo real:
     `ERROR COMPILACION: Microsoft Visual Basic / Error de compilación: Variable no definida en errores.xlsm - NoCompila (Código), Lín 9, Col 20`.
  7. Si no hay error: `OK compila <libro>` (por ejemplo `OK compila vl-fondo.xlsm`).
  8. Devuelve el foco a la aplicación que estaba delante, también si hay error.
- **Por qué así.** Un error de compilación durante una macro lanzada desde AppleScript deja el editor en
  "[interrumpir]", y pulsar "Restablecer" cerró Excel (3er cierre). Con ninguna macro en marcha,
  cerrar ese aviso **no** tumba Excel.
- **Por qué el paso 2 (fallo corregido el 14/11).** Depurar > Compilar actúa sobre el proyecto activo
  del editor (el de la ventana de código con el foco), no sobre el libro activo de Excel. Antes de la
  corrección, `compilar --libro vl-fondo.xlsm` devolvió el error de compilación de `errores.xlsm`.
  **Las salidas `OK compila` anteriores a esa corrección pueden ser de otro libro.**
  `Set Application.VBE.ActiveVBProject` no cambió el título del editor; `CodePane.Show` sí.
- **Límites.**
  - Trae Excel al frente: mientras corre, no escribir en otra ventana.
  - Usa el módulo temporal `ZZEnvoltorio`, así que necesita el ajuste de confianza.
  - A veces, tras `cargar`, el menú Compilar está desactivado (Excel lo da por compilado). La orden
    responde `SIN VERIFICAR` y no garantiza nada. Visto una vez de tres; causa sin aclarar.
  - El menú Compilar no apareció una vez de cuatro (`ERROR no encuentro Depurar > Compilar en el
    editor`); por eso el reintento. Con el reintento no se ha vuelto a ver, pero son pocas pruebas.
  - Necesita agent-desktop.
- **Descartado.** Dejar que el error salga al ejecutar (cierra Excel al detenerlo).

### `abrir` (`vba.py abrir`)

- **Qué es.** Abre un libro desde disco y responde por el usuario al aviso de macros.
- **Qué hace aquí.** Si el archivo no existe, avisa y no hace nada más. Si existe, lanza
  `open POSIX file` en segundo plano. Durante 30 s mira cada segundo las ventanas de Excel con System
  Events y pulsa "Habilitar macros" (o "Enable Macros" con Excel en inglés) si aparece. Con
  `--sin-macros` pulsa "Deshabilitar macros" (o "Disable Macros"). Comprueba el código de salida de
  `osascript` y el tiempo agotado. Al terminar devuelve el foco a la aplicación que estaba delante.
- **Por qué así.** Con la seguridad en "Deshabilitar todas las macros con notificación", Excel pregunta
  y el botón por defecto es "Deshabilitar macros". Con las macros desactivadas, `run VB macro` da
  `Parameter error (-50)`. agent-desktop no ve ese diálogo; System Events sí.
- **Riesgo.** Habilita las macros sin mirar el libro. Usarla solo con libros propios o de prueba.
- **Alternativas.** Ver [Habilitar macros](05-buenas-practicas.md#habilitar-macros-al-abrir-un-libro).

### agent-desktop

- **Qué es.** Herramienta de Vercel Labs ([GitHub](https://github.com/vercel-labs/agent-desktop),
  Apache-2.0, Rust). Lee el árbol de accesibilidad de las aplicaciones y actúa sobre sus elementos por
  su nombre, sin capturas de pantalla. Por defecto ("headless") no roba el foco ni mueve el ratón.
- **Instalación usada.** v1.3.1, binario de GitHub Releases
  (`agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz`) con el SHA-256 comprobado contra `checksums.txt`,
  en `~/workspaces/agent-desktop-bin/v1.3.1/`. Sin npm ni instalación global (el paquete npm ejecuta
  un `postinstall`). Sin dependencias de red en el código. Firma ad hoc. Guarda estado en `~/.agent-desktop`.
- **Qué hace aquí.** Solo lo que AppleScript no alcanza: abrir el editor de VBA, pulsar menús, marcar o
  desmarcar el ajuste de confianza (`vba.py confianza`), compilar (`vba.py compilar`) y enviar Escape a
  Excel (`press escape --app "Microsoft Excel"`).
- **Por qué así.** Actúa sobre Excel aunque delante esté otra ventana. Es lo que paró el intento con
  teclas de AppleScript. Aun así, `compilar` y `confianza` ponen Excel delante antes de usar los menús:
  con Excel detrás, los menús salían a veces desactivados.
- **Límites.**
  - No ve las ventanas modales de Excel (error de VBA, "¿Habilitar macros?"). Mientras hay una, Excel
    no publica ninguna ventana.
  - El área de código del editor no está en el árbol de accesibilidad. Para escribir hubo que pegar con
    el portapapeles y Cmd+V en modo `--headed` (teclado físico, autorizado por el usuario). El portapapeles
    del usuario queda sustituido.
  - Un clic `--headed` sobre un menú desactivado agotó el tiempo y agent-desktop se cerró (SIGABRT).
    Excel no se vio afectado.
- **Descartado.** Control de pantalla por capturas (computer use): lento y frágil.

### System Events

- **Qué es.** El componente de macOS que permite a AppleScript manejar la interfaz de cualquier
  aplicación (ventanas, botones, menús) por accesibilidad.
- **Qué hace aquí.** Ver y pulsar las ventanas modales que agent-desktop no ve: `vba.py dialogos`,
  el aviso de macros en `abrir`, el "Aceptar" del error de compilación en `compilar` y la vigilancia de
  `ejecutar` y `seguro`.
- **Por qué así.** Sí ve los diálogos modales de Excel, pero **solo con Excel delante**: con Excel en
  segundo plano ve 0 ventanas, ni siquiera un `MsgBox`. Por eso `vba.py` pone Excel delante antes de
  mirar, y por eso los avisos traen Excel al frente.
- **Cómo ve un `MsgBox`.** Título de ventana vacío, subrol `AXDialog`, botón "Aceptar" y dos textos:
  "Microsoft Excel" (el título del aviso) y el mensaje. Un aviso de compilación lleva los textos
  "Microsoft Visual Basic" y "Error de compilación: ...".
- **Límite.** Las teclas que envía van a la ventana activa. No se usa para escribir.
- **Requisito.** Permiso de Accesibilidad para la terminal o el editor desde el que se lanza.

### olevba (oletools)

- **Qué es.** Herramienta de Python que extrae el código VBA de un archivo de Office **sin abrir Excel**.
  Solo lee.
- **Qué hace aquí.** Revisión, no forma parte del ciclo. Probado el 14/11: `olevba --code copia.xlsm`
  sacó los módulos y encontró `TasaDiaria = ... / 360`. `olevba -j` da JSON con nombre y código de cada
  módulo. `olevba -a` marcó "Hex Strings" y "Base64 Strings": falsos positivos en este libro.
- **Instalación usada.** `oletools==0.60.2` en un entorno aislado (`~/workspaces/venvs/oletools`), con
  las dependencias congeladas.
- **Para qué sirve.** Revisar un libro antes de abrirlo con macros, o leer el código en una máquina sin
  Excel.
- **Por qué no más.** No puede escribir código ni ejecutarlo. Con Excel abierto, `exportar` ya da lo mismo.

### Python como referencia independiente

- **Qué es.** Un script que calcula el resultado esperado a partir de la **regla de negocio** escrita,
  no a partir del VBA. En la demo: [demo/referencia.py](../demo/referencia.py), que lee `datos.csv` y
  escribe `esperado.csv`.
- **Qué hace aquí.** `esperado.csv` se copia a la hoja `Esperado`. El módulo `Tests` compara el valor
  de VBA con ese esperado, con tolerancia 1e-6, y escribe la hoja `Tests`.
- **Por qué así.** Es la respuesta al problema de la profesora. Si el valor esperado sale del propio VBA,
  el test no valida nada. Una implementación aparte, desde la especificación, sí.
- **Descartado.** Revisar el código "en teoría" sin ejecutarlo. Tests cuyo esperado lo calcula el mismo VBA.

## Preparación (una sola vez)

1. **Permisos de macOS.** Automatización (terminal → Excel y System Events) y Accesibilidad (terminal y
   agent-desktop). El 14/11 ya estaban concedidos. Ver [05-buenas-practicas.md](05-buenas-practicas.md#automatización-y-accesibilidad-de-macos).
2. **Ajuste de confianza.** Excel > Preferencias > Seguridad > "Acceso de confianza al modelo de objetos
   de proyectos de VBA". No es "una sola vez": se activa al empezar cada sesión y se desactiva al
   terminar, con `vba.py confianza activar` y `vba.py confianza desactivar` (ver
   [Referencia de órdenes](#referencia-de-órdenes)) o a mano.
3. **agent-desktop**, para `compilar` y `confianza`. Ver la instalación usada arriba.
4. **Libro lanzador.** El repositorio ya trae `prueba.xlsm` con la cargadora y el motor. Para hacerlo
   desde cero (pasos que se hicieron el 14/11; las órdenes exactas no quedaron registradas):
   1. Crear el libro con AppleScript: `save workbook as ... file format macro enabled XML file format`.
      Funcionó al principio de la sesión del 14/11; más tarde, ese mismo día, la misma orden dio
      `Parameter error (-50)` y no creó el archivo (causa sin confirmar, ver [05-buenas-practicas.md](05-buenas-practicas.md#c-solución-de-problemas-por-síntoma)).
      Si falla, crear el libro a mano con Archivo > Guardar como.
   2. Pegar [vba/Cargadora.bas](../vba/Cargadora.bas) en un módulo: Herramientas > Macro > Editor de
      Visual Basic > Insertar > Módulo. Guardar. Se hizo con agent-desktop; a mano también vale.
   3. Escribir `Motor.bas` en la hoja `_codigo` (columna A, sin marca), `Motor` en B1, y ejecutar
      `run VB macro "prueba.xlsm!Cargar"`. Guardar. No con `vba.py cargar`: escribe con la marca `|`,
      que `Cargar` no quita, y borra la columna B.
5. **Comprobar.** `python3 -I herramientas/vba.py modulos` debe listar `Motor`.

## Ciclo de trabajo

Es la versión de referencia del ciclo a mano, con el libro de la demo como ejemplo. Con Excel abierto y
desde la raíz del repositorio. Se usa una función de shell y no una variable: en zsh,
`V="python3 -I herramientas/vba.py"` seguido de `$V ...` falla porque la variable con espacios no se
parte en palabras.

```bash
vba() { python3 -I herramientas/vba.py "$@"; }
vba confianza activar                                   # ajuste de confianza, solo durante la sesión
vba abrir prueba.xlsm                                   # lanzador; pulsa "Habilitar macros"
vba abrir demo/vl-fondo.xlsm                            # libro de trabajo (llega con el error)
vba modulos --libro vl-fondo.xlsm                       # qué hay dentro
vba compilar --libro vl-fondo.xlsm                      # OK compila vl-fondo.xlsm
vba seguro EjecutarTests --libro vl-fondo.xlsm          # FALLO 10 de 15
vba exportar Calculo --libro vl-fondo.xlsm --salida /tmp/Calculo.bas
#   ... Claude lee y corrige /tmp/Calculo.bas ...
vba cargar /tmp/Calculo.bas --libro vl-fondo.xlsm --modulo Calculo   # 1. cargar: reemplaza y guarda
vba compilar --libro vl-fondo.xlsm                      # 2. compilar, siempre antes de ejecutar
vba seguro EjecutarTests --libro vl-fondo.xlsm          # 3. seguro, nunca ejecutar a secas: OK 15/15
vba seguro GenerarInforme --libro vl-fondo.xlsm         # OK Informe 15 dias
vba leer Tests A1:E16 --libro vl-fondo.xlsm             # detalle
vba confianza desactivar                                # al terminar
```

El archivo exportado va fuera del repositorio (`/tmp`) para no dejar copias sueltas en la raíz. Python
lo lee y lo escribe; Excel no lo abre, así que no pide "Conceder acceso".

Para dejar el libro como estaba, como hace la demo guiada: `vba cargar demo/vba/Calculo.bas --libro
vl-fondo.xlsm` y repetir `seguro EjecutarTests` y `seguro GenerarInforme`, para que las hojas `Tests` e
`Informe` cuenten lo mismo que el código (`FALLO 10 de 15`).

Orden obligatorio: **cargar → compilar → seguro**. Por qué: [05-buenas-practicas.md](05-buenas-practicas.md#a-reglas-de-oro).

## Referencia de órdenes

`--libro` vacío = el lanzador `prueba.xlsm`. Todas imprimen texto en la terminal. `--timeout S` fija los
segundos de espera antes de que actúe la vigilancia de ventanas (por defecto 60). Si la vigilancia
actúa, la salida es `AVISO cerrado con Aceptar: ... Resultado: ...`, `ERROR VENTANA VBA (no se ha
tocado): ...` o `ERROR TIEMPO: ...` (ver [`herramientas/vba.py`](#herramientasvbapy)).

| Orden | Qué hace | Salida típica |
|---|---|---|
| `abrir LIBRO.xlsm [--sin-macros]` | Abre el libro y pulsa "Habilitar macros" o "Enable Macros" (o "Deshabilitar"). Si el archivo no existe, avisa y no abre nada. Devuelve el foco | `OK abierto vl-fondo.xlsm` |
| `cargar ARCHIVO.bas [--libro L] [--modulo M] [--sin-guardar] [--timeout S]` | Mete el código en el módulo (por defecto, el nombre del archivo), lo reemplaza si existe y guarda el libro. Se niega a reemplazar `Motor`, `Módulo1` o `Cargadora` en el lanzador (código 1) | `OK Calculo 52 lineas en vl-fondo.xlsm` y `guardado ...` |
| `exportar MODULO [--libro L] [--salida F] [--timeout S]` | Saca el código del módulo a pantalla o a archivo. Si el resultado empieza por `ERROR` o `AVISO`, no escribe el archivo, imprime el mensaje y sale con código 1 | `OK Calculo -> /tmp/Calculo.bas` |
| `modulos [--libro L]` | Lista componentes: nombre, tipo y líneas | una línea por módulo |
| `borrar MODULO [--libro L]` | Elimina un módulo | `OK borrado X` |
| `compilar [--libro L]` | Activa en el editor el proyecto del libro, comprueba el título del editor y compila sin ejecutar (agent-desktop). Si el proyecto activo no es el pedido, no compila. Devuelve el foco | `OK compila vl-fondo.xlsm`, `ERROR COMPILACION: ...`, `ERROR el editor tiene activo el proyecto de ...` o `SIN VERIFICAR: ...` |
| `confianza activar\|desactivar` | Abre Excel > Preferencias > Seguridad con agent-desktop, marca o desmarca la casilla de confianza, lee su valor para comprobarlo y cierra. Si el editor de VBA está delante (la barra de menús tiene "Depurar"), pulsa antes su botón "Ver Microsoft Excel": con el editor delante, "Preferencias..." abre las Opciones del editor. Sale con código 1 si no termina en `OK`. Devuelve el foco | `OK confianza activada (antes: desactivada)`, `ERROR el editor de VBA sigue delante; ...` o `ERROR no encuentro la casilla de confianza` |
| `seguro MACRO [--sub] [--libro L] [--timeout S]` | Ejecuta la macro dentro del envoltorio. `--sub` si es un `Sub`. Si hay ventana o tiempo agotado, no borra `ZZEnvoltorio` y avisa | lo que devuelva la función, o `ERROR n: ...` |
| `ejecutar MACRO [ARG...] [--libro L] [--timeout S]` | Ejecuta con argumentos y sin envoltorio. Solo para macros con su propio `On Error` | lo que devuelva la función |
| `guardar [--libro L]` | Guarda el libro | `OK guardado ...` |
| `dialogos` | Pone Excel delante y lista sus ventanas que no son libros ni el editor, con subrol, título, botones y textos (System Events). Devuelve el foco | `[AXDialog]  \|\| Aceptar \|\| Microsoft Excel \| ...`, o `sin dialogos` |
| `hoja-csv ARCHIVO.csv HOJA [--libro L]` | Copia un CSV separado por `;` a una hoja desde A1. Escribe por `formula` para que los decimales con punto se lean bien. No respeta comillas y un texto que empiece por `=` se toma como fórmula | `OK 16 filas en Datos` |
| `leer HOJA RANGO [--libro L]` | Lee un rango y lo imprime separado por `;`. Los números salen con el formato regional (coma decimal en castellano) | filas de la hoja |

## Cómo hacer la demo

Guion completo y regla de negocio: [demo/ESPECIFICACION.md](../demo/ESPECIFICACION.md). Herramienta
inventada de valor liquidativo con módulos `Datos`, `Calculo`, `Informe` y `Tests`, y un error de base
360/365 en la comisión de gestión.

Libros de la demo:

- `demo/vl-fondo.xlsm`: el antes, con el error en `Calculo`.
- `demo/vl-fondo-corregido.xlsm`: el después de la demo, con `Calculo` corregido, tests 15/15 e informe.
  Se abre en la hoja `Informe`. Sirve para ver el resultado final sin ejecutar nada.

Este repositorio revela el error (README, `docs/` y `demo/solucion/`). Para probar que Claude lo encuentra
sin saber dónde está, ver [01-contexto.md](01-contexto.md#lo-que-todavía-no-está-demostrado).

Demo guiada, que hace todos los pasos de la tabla:

```bash
python3 -I demo/ejecutar_demo.py                  # con compilar (necesita agent-desktop)
python3 -I demo/ejecutar_demo.py --sin-compilar   # sin agent-desktop
```

Si `compilar` devuelve `ERROR COMPILACION` o `ERROR VENTANA`, el guion para y no ejecuta los tests.
Probado el 14/11: `DEMO CORRECTA` (código 0) con y sin `--sin-compilar`. Al terminar, el libro vuelve a
la versión con el error, con sus hojas `Tests` e `Informe` coherentes con ella (`FALLO 10 de 15`), salvo
con `--dejar-corregido`. Las mismas órdenes, a mano: [Ciclo de trabajo](#ciclo-de-trabajo).

Resultado real del 14/11/2026 (repetición a mano con `compilar` ya corregido; el informe, del guion):

| Paso | Orden | Resultado |
|---|---|---|
| Meter la versión con el error | `cargar demo/vba/Calculo.bas --libro vl-fondo.xlsm` | `OK Calculo 52 lineas en vl-fondo.xlsm` |
| Compilar | `compilar --libro vl-fondo.xlsm` | `OK compila vl-fondo.xlsm` |
| Tests con el código original | `seguro EjecutarTests --libro vl-fondo.xlsm` | `FALLO 10 de 15: 2025-03-10 … 2025-03-21` |
| Sacar el módulo | `exportar Calculo --libro vl-fondo.xlsm` | línea 9: `TasaDiaria = ComisionPctDia(fila) / 100 / 360` |
| Corregir y meter | `cargar demo/solucion/Calculo_corregido.bas --libro vl-fondo.xlsm --modulo Calculo` | `OK Calculo 53 lineas` (la corrección añade un comentario) |
| Compilar | `compilar --libro vl-fondo.xlsm` | `OK compila vl-fondo.xlsm` (una vez, `ERROR no encuentro Depurar > Compilar en el editor`; de ahí el reintento) |
| Repetir tests | `seguro EjecutarTests --libro vl-fondo.xlsm` | `OK 15/15` (incluye los 5 días que ya estaban bien) |
| Informe | `seguro GenerarInforme --libro vl-fondo.xlsm` | `OK Informe 15 dias` |

La corrección de referencia está en [demo/solucion/Calculo_corregido.bas](../demo/solucion/Calculo_corregido.bas)
y su explicación en [demo/solucion/SOLUCION.md](../demo/solucion/SOLUCION.md).

Siguiente: [04-resultados.md](04-resultados.md).
