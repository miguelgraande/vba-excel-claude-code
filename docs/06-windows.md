# Vía Windows (COM): preparar un portátil con Windows

Escrito el 14/11/2026 en un Mac, sin Windows a mano. **Nosotros no hemos probado nada de esta página en Windows.**

Hay dos scripts:

- **`herramientas/windows/vba.ps1`**: el del 14/11. **Informe de terceros (21/11/2026):** la profesora
  lo usó en Windows y le funcionó. No tenemos la orden exacta, la versión de Excel ni la salida, así que
  no cuenta como prueba W0 a W8 hasta tenerlas (regla de oro 8).
- **`herramientas/windows/vba-nuevo.ps1`**: copia de `vba.ps1` con `compilar` y un vigía que lee las
  ventanas, adaptados de MCP-Access el 21/11 (§13). Experimental: nadie lo ha ejecutado. Sí pasa el
  analizador de PowerShell 7.6.6 en el Mac, sin errores y con el C# del vigía compilando
  ([04-resultados.md](04-resultados.md)). Si pasa W0 a W11, puede sustituir a `vba.ps1`.

Cómo leer esta página:

- **[OFICIAL]**: lo dice la documentación de Microsoft (o de la librería), con enlace en §12.
- **[SUPOSICIÓN]**: deducción nuestra o práctica habitual. Hay que comprobarla en Windows.
- **[PROBADO]**: no hay ninguna todavía. Se marcará al pasar las pruebas W0 a W8 del [plan inicial](00-plan-inicial.md).

---

## 1. La idea en una frase

En Windows, cualquier programa puede manejar Excel por COM. Excel expone también el editor de VBA
(`VBProject`), así que desde fuera se puede sacar el código a texto, meter código nuevo, ejecutar
una macro y leer lo que devuelve. Es lo que hace Microsoft en la KB 219905 **[OFICIAL]**.

PowerShell viene con Windows y sabe hablar COM (`New-Object -ComObject`) **[OFICIAL]**. Por eso el
script no necesita instalar nada.

---

## 2. Requisitos

| Requisito | Cómo comprobarlo (solo lectura) | Nota |
|---|---|---|
| Windows con Excel de escritorio (2016, 2019, 2021 o Microsoft 365) | Abrir Excel > Archivo > Cuenta | Excel en la web no sirve: no tiene VBA |
| Windows PowerShell 5.1 | `$PSVersionTable.PSVersion` | Viene con Windows 10 y 11 |
| PowerShell en modo completo | `$ExecutionContext.SessionState.LanguageMode` debe decir `FullLanguage` | Si dice `ConstrainedLanguage`, la empresa usa AppLocker o WDAC y **no se puede usar COM** desde PowerShell **[OFICIAL]**. Ver §8 |
| Permiso para ejecutar scripts | `Get-ExecutionPolicy -List` | Ver §2.1 |
| Ajuste de confianza VBA activado | `reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM` | Ver §3 |
| Libro de prueba `.xlsm` con datos inventados | - | Nunca un libro real ni de cliente |

### 2.1 Política de ejecución de scripts

- En Windows de escritorio, si nadie la ha cambiado, la política es `Restricted`: no deja ejecutar
  ningún `.ps1` **[OFICIAL]**.
- Se puede saltar **solo para ese proceso**, sin cambiar nada del equipo:
  `powershell -NoProfile -ExecutionPolicy Bypass -File .\vba.ps1 ...` **[OFICIAL]** (ámbito `Process`).
- Si la empresa fija la política por directiva de grupo (`MachinePolicy` o `UserPolicy` en
  `Get-ExecutionPolicy -List`), esa manda y no se puede cambiar **[OFICIAL]**. Entonces hay que pedirlo
  a informática. No buscar atajos.
- Si el `.ps1` llega por correo, Teams o descarga, Windows le pone la "marca de la web". Se quita con
  `Unblock-File .\vba.ps1` (o Propiedades > Desbloquear) **[OFICIAL]** para el caso de libros; para
  scripts es el mismo mecanismo **[SUPOSICIÓN]**.

---

## 3. Activar el acceso de confianza al modelo de objetos de VBA

Sin este ajuste, `exportar`, `importar` y `modulos` fallan con "El acceso mediante programación al
proyecto de Visual Basic no es de confianza". `ejecutar` no lo necesita: solo usa `Application.Run`
**[SUPOSICIÓN]**, coherente con lo visto en Mac el 14/11.

### 3.1 Por la interfaz (la vía recomendada)

Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza >
Configuración de macros > casilla **"Confiar en el acceso al modelo de objetos de proyectos de VBA"**
**[OFICIAL]** (en inglés: *Trust access to the VBA project object model*).

Microsoft avisa de que es un riesgo: con la casilla marcada, las macros de **cualquier** libro que se
abra pueden tocar el VBA. Recomienda marcarla solo mientras se trabaja y quitarla después **[OFICIAL]**.

### 3.2 Por el registro (lo que hace la casilla por debajo)

- Clave del usuario: `HKCU\Software\Microsoft\Office\16.0\Excel\Security`, valor DWORD `AccessVBOM`
  (1 = activado, 0 = desactivado). Fuente: blog de soporte de Microsoft archivado en Learn **[OFICIAL, archivo]**.
- `16.0` es la versión interna de Excel 2016, 2019, 2021 y Microsoft 365 **[SUPOSICIÓN]**: la plantilla
  de directivas de Office usa `16.0` para todas, pero no hemos visto una página de Microsoft que lo diga
  así de claro.
- El mismo blog avisa de que cambiarlo con Excel abierto puede no tener efecto. Cerrar todos los Excel antes.
- Leer el valor (no cambia nada):
  `reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM`
- Preferir la casilla. Tocar el registro a mano solo si lo deciden quien prepara la prueba y la usuaria del portátil, y nunca la rama de directivas.

### 3.3 Si lo bloquea una directiva de grupo (GPO)

Existe una directiva de Office llamada **"Trust access to Visual Basic Project"**
(Microsoft Excel 2016 > Opciones de Excel > Seguridad > Centro de confianza). Su texto oficial
(plantillas ADMX de Office **[OFICIAL]**, leído en una copia pública de esas plantillas):

- **Activada**: el acceso queda permitido y el usuario no puede cambiar la casilla.
- **Desactivada**: no hay acceso, la casilla aparece desmarcada y **bloqueada** (gris).
- **Sin configurar**: no hay acceso por defecto, pero el usuario puede marcar la casilla.
- Se guarda en `HKCU\Software\Policies\Microsoft\Office\16.0\Excel\Security\AccessVBOM`
  (1 activada, 0 desactivada).

Qué significa para un portátil de empresa:

- Las guías de seguridad CIS para Office recomiendan poner esta directiva en **Desactivada**. Un portátil
  de empresa endurecido según CIS tendrá la casilla en gris **[SUPOSICIÓN]** sobre un equipo concreto.
- Si la casilla está en gris: **solo informática puede abrirla**. El script lo detecta y lo dice: lee
  las dos claves (directiva y usuario) y muestra sus valores. No escribe nada.
- Sin el ajuste, lo único que queda es `ejecutar` (macros que ya estén en el libro) y la vía de
  control de pantalla (C1 y C2 del [plan inicial](00-plan-inicial.md)).
- Además de GPO, Microsoft gestiona estas directivas desde Cloud Policy e Intune **[OFICIAL]**: el
  efecto en el equipo es el mismo.

---

## 4. Cómo abre Excel el script

| Pieza | Qué hace | Fuente |
|---|---|---|
| `New-Object -ComObject Excel.Application` | Crea una instancia **nueva** de Excel (equivale a `CreateObject`). `GetObject` en cambio se engancha a la primera que se abrió | **[OFICIAL]** KB 288902, que además recomienda usar una instancia nueva y no la del usuario |
| PID propio | Compara los `EXCEL.EXE` antes y después de crearla para saber cuál es la nuestra | **[SUPOSICIÓN]**: si otro Excel arranca en el mismo segundo, no se identifica y se avisa |
| `Visible = $false` | Excel abierto por COM arranca invisible | **[OFICIAL]** KB 288902 (tabla) |
| `DisplayAlerts = $false` | Quita los avisos de Excel y elige la respuesta por defecto. En código entre procesos no se reactiva solo | **[OFICIAL]** |
| `AutomationSecurity` | Al abrir por COM, las macros quedan **activadas** por defecto (`msoAutomationSecurityLow`). El script usa `ForceDisable` (3) para `modulos`, `exportar`, `importar` y `compilar`, y `Low` (1) solo para `ejecutar` | **[OFICIAL]** para los valores. **[SUPOSICIÓN]** que con macros apagadas se puede seguir leyendo y escribiendo el VBA |
| `EnableEvents = $false` | Evita que `Workbook_Open` y otros eventos se disparen al abrir. `-Eventos` los activa | **[OFICIAL]** la propiedad. **[SUPOSICIÓN]** que impide `Workbook_Open` al abrir por COM |
| `Workbooks.Open(ruta, 0, soloLectura)` | Sin actualizar vínculos. Solo lectura para `modulos` y `exportar` | **[OFICIAL]** |
| `Close($false)` + `Quit()` | Cierra sin guardar y sale. Con `DisplayAlerts = $false`, `Quit` no pregunta | **[OFICIAL]** |
| `ReleaseComObject` + `GC.Collect()` | Suelta las referencias COM para que `EXCEL.EXE` termine | **[OFICIAL]** el método. Microsoft avisa de que mal usado puede romper cosas; aquí se llama solo al final |
| Último recurso | Si nuestro `EXCEL.EXE` sigue vivo 10 s después, se mata **solo ese** PID | **[SUPOSICIÓN]** |

---

## 5. Uso del script

```powershell
cd C:\ruta\vba-excel-claude-code
$vba = ".\herramientas\windows\vba.ps1"
# una sola vez por ventana de PowerShell:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Desde Claude Code (que en Windows suele lanzar los comandos en Git Bash) **[SUPOSICIÓN]**:

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File herramientas/windows/vba.ps1 modulos prueba.xlsm
```

### 5.1 Comandos

| Comando | Qué hace | Ejemplo |
|---|---|---|
| `modulos <libro>` | Lista `nombre;tipo;líneas` | `& $vba modulos .\prueba.xlsm` |
| `exportar <libro> <carpeta>` | Saca cada módulo a `.bas`, `.cls` o `.frm` (+ `.frx`). Sobrescribe. Omite hojas sin código. Convierte a UTF-8 para VS Code y git (`-Ansi` lo deja como lo saca Excel) | `& $vba exportar .\prueba.xlsm .\src` |
| `importar <libro> <archivo>` | Mete el módulo. Si ya hay uno con ese nombre, lo reemplaza. Guarda el libro | `& $vba importar .\prueba.xlsm .\src\Calculo.bas` |
| `ejecutar <libro> <Macro> [args]` | Ejecuta con `Application.Run` y escribe lo que devuelve. Cierra **sin guardar** | `& $vba ejecutar .\prueba.xlsm Suma 2 3` |
| `compilar <libro>` | **Solo en `vba-nuevo.ps1`.** Depurar > Compilar sin ejecutar nada. Responde `OK compila <libro>`, `ERROR COMPILACION: ...` con módulo, línea y columna, o `SIN VERIFICAR: ...`. Cierra **sin guardar**. Solo `OK` cuenta como comprobado (§13) | `& $vba compilar .\prueba.xlsm` |

Opciones:

| Opción | Para | Efecto |
|---|---|---|
| `-Leer "Tests"` o `-Leer "Tests!A1:D20"` | ejecutar | Tras la macro, imprime la hoja (rango usado) o el rango, separado por `;` |
| `-Guardar` | ejecutar | Guarda el libro al terminar |
| `-Tiempo 120` | ejecutar, compilar | Segundos máximos. Si la macro no vuelve, mata nuestro Excel. `0` = sin límite. En `compilar`, máximo 60 |
| `-Eventos` | todos | Deja activos los eventos (`Workbook_Open`, `Worksheet_Change`...) |
| `-Visible` | todos | Muestra Excel. Útil para ver qué ventana se ha quedado atascada |
| `-Ansi` | exportar | No convierte a UTF-8 |

Salida: lo útil va a la salida normal; avisos y errores, a la de errores. Código de salida 0 (bien),
1 (uso incorrecto) o 2 (fallo). Si la macro devuelve el texto "ERROR ...", el código sigue siendo 0:
hay que leer el texto.

### 5.2 Detalles que importan

- **Argumentos**: `2` llega como entero, `2.5` como decimal (siempre con **punto**), `true`/`false` como
  booleano. Lo demás, texto. Para forzar texto: `s:007`. Un argumento que empiece por `-` lo
  interpreta PowerShell como opción: ponerlo entre comillas o con `s:`.
- **Números de vuelta**: se escriben con punto decimal (`2.5`), no con coma, aunque Windows esté en
  castellano. Así se comparan sin líos con Python.
- **Nombre de la macro**: si no lleva `!`, el script añade `'prueba.xlsm'!` para no ejecutar por error
  una macro de otro libro. Vale `Modulo.Macro`. Máximo 30 argumentos y solo por posición **[OFICIAL]**.
- **Mejor `Function` que `Sub`**: `Application.Run` devuelve lo que devuelva la macro **[OFICIAL]**. Una
  `Sub` no devuelve nada.
- **Parámetros de las funciones de prueba**: usar `ByVal` o `Variant` **[SUPOSICIÓN]**: no está claro
  cómo convierte `Run` un entero a un parámetro `ByRef x As Double`.
- **Módulos de documento** (`ThisWorkbook`, `Hoja1`): no se pueden borrar. `importar` vacía su código con
  `DeleteLines` y lo mete con `AddFromString` **[OFICIAL]** los métodos. Los demás se borran con
  `Remove` y se vuelven a meter con `Import` **[OFICIAL]**.
- **Nombre del módulo**: sale de la línea `Attribute VB_Name = "..."`. Si falta en un `.bas`, se usa el
  nombre del archivo. Si Excel lo renombra a `Modulo11` al reimportar, el script lo corrige.
- **Tildes**: Excel exporta en la codificación ANSI de Windows (en castellano, Windows-1252)
  **[SUPOSICIÓN]**. VS Code abre en UTF-8 y, si guarda un archivo ANSI, rompe las tildes para siempre.
  Por eso `exportar` pasa a UTF-8 e `importar` vuelve a ANSI. Un carácter que no exista en ANSI
  (por ejemplo una flecha) llega a Excel como `?`.
- **Libro abierto en otro Excel**: nuestro Excel lo abre en solo lectura. `importar` y `-Guardar` se
  niegan a seguir y avisan. Cerrar el libro en el otro Excel y repetir.
- **Proyecto con contraseña**: el script lo detecta y para. Desbloquearlo a mano (Alt+F11).

---

## 6. Ciclo completo

Objetivo: que la profesora vea que Claude escribe, mete, ejecuta, comprueba contra una referencia
independiente y corrige, sin copiar y pegar.

1. **Ver qué hay**
   `& $vba modulos .\prueba.xlsm`
2. **Sacar el código a texto** (y guardarlo en git para ver cada cambio)
   `& $vba exportar .\prueba.xlsm .\src`
   `git add src; git commit -m "Código original"`
3. **Claude edita en VS Code** `src\Calculo.bas`. No hace falta abrir el editor de VBA.
4. **Meterlo en el libro**
   `& $vba importar .\prueba.xlsm .\src\Calculo.bas`
   Salida esperada: `OK Calculo 42 lineas`.
5. **Ejecutar los tests y leer la hoja Tests**
   `& $vba ejecutar .\prueba.xlsm Tests.EjecutarTodo -Leer "Tests"`
   Salida esperada, por ejemplo:
   ```
   FALLO 1 de 5
   --- Tests!A1:D6
   Caso;Esperado;Obtenido;Estado
   sin comisiones;100.25;100.25;OK
   con comisiones;99.75;100.25;FALLO
   ...
   ```
6. **Comparar con Python**: un script en Python hace el mismo cálculo con los mismos datos y Claude
   compara columna a columna. Es la referencia independiente que pedía el problema.
7. **Corregir y repetir** pasos 3 a 6 hasta `OK 5 de 5`. Los casos que ya iban bien siguen en la hoja,
   así que también se comprueba que no se ha roto nada.
8. **Revisar el cambio**: `git diff src\Calculo.bas` enseña exactamente qué ha cambiado.

Patrón para `Tests.EjecutarTodo`, sin ventanas:

```vba
Public Function EjecutarTodo() As String
    Dim ws As Worksheet, fila As Long, fallos As Long
    On Error GoTo Fallo
    Set ws = ThisWorkbook.Worksheets("Tests")
    ws.Cells.ClearContents
    ws.Range("A1:D1").Value = Array("Caso", "Esperado", "Obtenido", "Estado")
    fila = 2
    ' Probar (no incluida aqui) llama a la funcion real, escribe una fila,
    ' avanza fila y suma 1 a fallos si no cuadra
    Probar ws, fila, fallos, "sin comisiones", 100.25, ValorLiquidativo(1000, 0)
    Probar ws, fila, fallos, "con comisiones", 99.75, ValorLiquidativo(1000, 0.005)
    If fallos = 0 Then EjecutarTodo = "OK " & (fila - 2) & " de " & (fila - 2) _
    Else EjecutarTodo = "FALLO " & fallos & " de " & (fila - 2)
    Exit Function
Fallo:
    EjecutarTodo = "ERROR " & Err.Number & ": " & Err.Description
End Function
```

---

## 7. Ventanas modales: qué se puede y qué no

El problema: si una macro abre una ventana (MsgBox, InputBox, error sin controlar, error de compilación),
Excel se queda esperando a que alguien pulse. Por COM, la llamada a `Run` no vuelve nunca.

Lo que dice Microsoft **[OFICIAL]**:

- `DisplayAlerts = $false` quita los **avisos de Excel** (guardar cambios, sobrescribir...) y elige la
  respuesta por defecto. Nada más.
- `DisplayAlerts` **no** se aplica a los avisos de seguridad.
- Ante un error inesperado, Office está pensado para mostrar una ventana modal. Si nadie la ve, el
  proceso se cuelga. "Ciertas prácticas reducen el riesgo, pero no lo evitan del todo" (KB 257757).
- Microsoft no recomienda ni da soporte a automatizar Office **desatendido** (servicios, tareas sin
  sesión). Nuestro caso es otro: la usuaria sentada delante, con su sesión abierta. Pero el riesgo de
  cuelgue es el mismo.

Lo que **no** se puede **[SUPOSICIÓN]**:

- Quitar por configuración los `MsgBox` o `InputBox` del propio código VBA: `DisplayAlerts` no los toca.
- Quitar la ventana de "Error en tiempo de ejecución" (Finalizar / Depurar) ni la de error de compilación.
- Cerrar la ventana desde COM: mientras hay una ventana modal, Excel rechaza las llamadas. Sí se puede
  por fuera de COM, con la API de ventanas de Windows (Win32): es lo que hace el vigía (§13).

Lo que **sí** se puede:

| Medida | Dónde | Efecto |
|---|---|---|
| `On Error GoTo` en cada función de prueba y devolver `"ERROR ..."` | VBA | El error no llega a abrir ventana. Ya probado así en Mac (M4) |
| Sin `MsgBox` en el código nuevo. En código antiguo, cambiar `MsgBox` por una función `Aviso` que en modo prueba escribe en una hoja | VBA (copia de prueba) | El código de 2015 tendrá muchas |
| `EnableEvents = $false` y macros apagadas al exportar o importar | script | Evita que `Workbook_Open` abra ventanas al abrir el libro |
| Vigía de ventanas | `vba-nuevo.ps1` | Cada 0,5 s lee las ventanas de **nuestro** Excel. Aviso simple (solo "Aceptar"): lo pulsa y lo cuenta (`AVISO cerrado con Aceptar: ...`). Otra ventana ("Finalizar"/"Depurar", InputBox, Sí/No): no pulsa nada, la describe (`ERROR VENTANA VBA ...`) y mata nuestro Excel sin esperar al tiempo (§13) |
| `compilar` antes de `ejecutar` | `vba-nuevo.ps1` | El error de compilación sale sin ninguna macro en marcha, así que cerrar su aviso es seguro |
| Vigía `-Tiempo` | script | Si `Run` no vuelve en N segundos, mata **solo** nuestro Excel. El script dice "Tiempo agotado, probable ventana modal" y no se queda colgado |
| `-Visible` | script | Para ver qué ventana es y cerrarla a mano |
| Control de pantalla (C2) | otra herramienta | Plan B para cerrar la ventana sin matar Excel |

Al matar Excel se pierde lo no guardado. Por eso `ejecutar` no guarda por defecto y se trabaja
siempre sobre una copia.

---

## 8. Riesgos

| Riesgo | Qué pasa | Qué hacer |
|---|---|---|
| **GPO bloquea el ajuste de confianza** | Casilla en gris. Sin `exportar`/`importar` | Pedirlo a informática solo para un perfil o equipo de pruebas. No intentar saltarlo |
| **GPO fija la política de scripts** | No se puede ejecutar el `.ps1` | Pedirlo a informática, o usar Python (§10) si está permitido |
| **AppLocker/WDAC** (PowerShell en `ConstrainedLanguage`) | No se puede crear Excel por COM desde PowerShell **[OFICIAL]** | El script lo detecta al empezar. Solo informática |
| **Libro con marca de la web** (llegó por correo o Teams) | Excel bloquea sus macros **[OFICIAL]** | `Unblock-File .\prueba.xlsm` o guardarlo en una ubicación de confianza. Cómo afecta a un libro abierto por COM: **[SUPOSICIÓN]**, probar |
| **Ventanas modales** | Cuelgue | §7 |
| **Varias instancias de Excel** | En clase se comentó que el método actual era inestable con varios Excel abiertos. `GetObject` se engancha a la primera instancia abierta **[OFICIAL]**, que puede no ser la buena (posible causa, **[SUPOSICIÓN]**) | El script crea su propia instancia y no toca las demás. El vigía y el cierre final solo matan nuestro PID |
| **El mismo libro abierto dos veces** | Nuestro Excel lo abre en solo lectura | `importar` y `-Guardar` paran y avisan |
| **Datos de clientes** | Lo que lee `-Leer` y el código exportado pasan a Claude, es decir, salen del portátil | Solo libros con datos inventados. Antes de tocar una herramienta real, permiso expreso de la organización propietaria. El código de la herramienta también puede ser confidencial |
| **Ajuste de confianza activado de más** | Cualquier libro con macros podría reescribir VBA | Activarlo solo durante la sesión y quitarlo al terminar **[OFICIAL]** |
| **Se pierde trabajo** | `importar` guarda el libro; el vigía mata Excel | Trabajar sobre una copia y con el código en git |

---

## 9. En qué se diferencia de lo que hacemos en Mac

| | Mac (probado el 14/11/2026) | Windows (sin probar) |
|---|---|---|
| Canal | AppleScript (`osascript`) | COM desde PowerShell o Python |
| Instancia de Excel | La que esté abierta. Hay una sola por usuario **[SUPOSICIÓN]** | Una propia y nueva por cada orden |
| Meter código | Desde **dentro**: `vba.py cargar` escribe el código en la hoja `_codigo` del lanzador, con una marca `\|` delante de cada línea, y el motor (`CargarEn`) quita la marca y reemplaza el módulo en cualquier libro abierto | Desde **fuera**: `VBComponents.Import` del `.bas`, sin cargadora ni hoja `_codigo` |
| Sacar código | Desde VBA: el motor (`Exportar`) lo devuelve como texto y `vba.py exportar` lo guarda | `VBComponent.Export` directo |
| Ejecutar | `run VB macro "libro.xlsm!Funcion"` | `Application.Run "'libro.xlsm'!Funcion"` |
| Documentación | Ninguna oficial para `run VB macro` ni la inyección | Microsoft documenta todo el modelo (KB 219905 y referencia VBE) |
| Ajuste de confianza | Excel > Preferencias > Seguridad | Centro de confianza, o GPO que puede bloquearlo |
| Sin el ajuste | "ERROR 1004" devuelto como texto, sin ventana (M4) | Se espera un error COM que el script atrapa y explica **[SUPOSICIÓN]** |
| Libro abierto | Excel debe tener el libro abierto | El script lo abre y lo cierra cada vez |
| Ventanas modales | Mismo riesgo | Mismo riesgo, más el vigía que mata solo la instancia propia |
| Compilar | Menú Depurar > Compilar pulsado con agent-desktop | Botón Compilar (ID 578) pulsado por COM y aviso leído por Win32. Sin agent-desktop |

En Mac la herramienta equivalente es [herramientas/vba.py](../herramientas/vba.py).

---

## 10. Alternativas con Python (breve)

Útiles si PowerShell está bloqueado pero Python está permitido. **Exigen instalar** paquetes: preguntar antes.

**pywin32** (`pip install pywin32`). `DispatchEx` crea una instancia nueva aunque ya haya otra abierta.
La ayuda de pywin32 apenas documenta `DispatchEx`: está en su código fuente y en la lista de correo
de pywin32 **[OFICIAL, poco documentado]**.

```python
import win32com.client as win32

xl = win32.DispatchEx("Excel.Application")   # instancia propia
xl.Visible = False
xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(r"C:\pruebas\prueba.xlsm")
    for c in wb.VBProject.VBComponents:         # exige el ajuste de confianza
        print(c.Name, c.Type, c.CodeModule.CountOfLines)
    print(xl.Run("'prueba.xlsm'!Hola"))         # valor devuelto
finally:
    if wb is not None:
        wb.Close(False)
    xl.Quit()
    del wb, xl
```

**xlwings** (`pip install xlwings`). Envuelve lo anterior. `xw.App()` como `with` abre una instancia
propia y la cierra al salir para no dejar procesos zombi **[OFICIAL]**. `wb.macro("Hola")()` ejecuta y
devuelve el valor **[OFICIAL]**. Para meter código sigue haciendo falta `wb.api.VBProject...` y el
ajuste de confianza.

```python
import xlwings as xw
with xw.App(visible=False) as app:
    app.display_alerts = False
    wb = app.books.open(r"C:\pruebas\prueba.xlsm")
    print(wb.macro("Hola")())
```

---

## 11. Plan de pruebas en Windows

**Para ejecutarlas, usa [10-probar-en-windows.md](10-probar-en-windows.md)**: órdenes exactas y salidas
esperadas con los libros del repositorio. Esta tabla es el resumen original; si no coinciden, manda la 10
(por ejemplo, aquí W3 usa una función `Suma` que no existe en ningún libro del repositorio).

Las mismas W0 a W8 del [plan inicial](00-plan-inicial.md), con el script:

| ID | Comando | Qué confirma |
|---|---|---|
| W0 | `$ExecutionContext.SessionState.LanguageMode`, `Get-ExecutionPolicy -List`, `reg query ... AccessVBOM` | Que el portátil lo permite |
| W1 | `& $vba modulos .\prueba.xlsm -Visible` | Abre Excel propio, lee el proyecto, cierra. Mirar en el Administrador de tareas que no queda `EXCEL.EXE` |
| W2 | `& $vba ejecutar .\prueba.xlsm Hola -Leer "Hoja1!A1"` | Lee celdas |
| W3 | `& $vba ejecutar .\prueba.xlsm Suma 2 3` (con una función `Suma` de prueba en el libro) | Devuelve `5` |
| W4 | `& $vba exportar .\prueba.xlsm .\src` | Archivos en `src`, tildes bien en VS Code |
| W5 | Editar `src\Nueva.bas`, `importar`, `ejecutar` | Código nuevo entra y corre |
| W6 | Macro con `MsgBox` y otra con error sin controlar, `-Tiempo 15` | El vigía corta a los 15 s y no queda Excel colgado |
| W7 | Ciclo de §6 | Ciclo completo |
| W8 | Con un Excel propio de la usuaria abierto, repetir W1 a W3 | Su Excel no se toca |
| W9 | `vba-nuevo.ps1 compilar .\prueba.xlsm` con un libro que compila | `OK compila prueba.xlsm` |
| W10 | `vba-nuevo.ps1 compilar` con un libro con una variable sin declarar y `Option Explicit` | `ERROR COMPILACION: ...` con módulo, línea y columna correctos. Repetirlo dos veces seguidas: la segunda no debe decir `SIN VERIFICAR` |
| W11 | W6 con `vba-nuevo.ps1` | La del `MsgBox` termina con `AVISO cerrado con Aceptar: ...` y su resultado. La del error devuelve `ERROR VENTANA VBA ...` con el texto del error en segundos, sin esperar a los 15 s. No queda Excel colgado |

W0 a W8 valen para los dos scripts; W9 a W11, solo para `vba-nuevo.ps1`.

Además, sin ajuste de confianza: `modulos` debe fallar con el diagnóstico del registro y sin ventana.

Apuntar cada resultado en [04-resultados.md](04-resultados.md), con fecha y salida real.

---

## 12. Fuentes

Microsoft (oficiales):

- [KB 219905: añadir y ejecutar una macro VBA desde fuera de Excel](https://support.microsoft.com/en-us/help/219905/how-to-dynamically-add-and-run-a-vba-macro-from-visual-basic)
- [Security notes for Office solution developers: casilla "Trust access to the VBA project object model" y su riesgo](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/security-notes-for-microsoft-office-solution-developers)
- [Blog de soporte de Microsoft archivado: clave AccessVBOM en HKCU](https://learn.microsoft.com/en-us/archive/blogs/cristib/vba-how-to-programmatically-enable-access-to-the-vba-object-model-using-macros)
- KB 282830 ("Programmatic access to Office VBA project is denied"): el enlace https://support.microsoft.com/kb/282830 hoy redirige a una página general de macros. No se ha podido leer el original.
- [Plantillas ADMX de Office (descarga oficial)](https://www.microsoft.com/download/details.aspx?id=49030)
- [Macros de Internet bloqueadas por defecto: marca de la web, Unblock-File, ubicaciones de confianza, GPO/Intune/Cloud Policy](https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked)
- [Habilitar o deshabilitar macros en Microsoft 365](https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
- [KB 288902: GetObject y CreateObject en Office (instancia nueva frente a existente)](https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/getobject-createobject-behavior)
- [KB 257757: automatización de Office desatendida y ventanas modales](https://support.microsoft.com/en-us/help/257757)
- [Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
- [Application.DisplayAlerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts)
- [Application.EnableEvents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableevents)
- [Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run)
- [Application.Quit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit)
- [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)
- [Workbook.Close](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.close)
- [VBComponents.Import](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/import-method-vba-add-in-object-model)
- [VBComponent.Export](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/export-method-vba-add-in-object-model) (falla si el archivo ya existe)
- [VBComponents.Remove](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/remove-method-vba-add-in-object-model)
- [CodeModule.AddFromString](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/addfromstring-method-vba-add-in-object-model)
- [CodeModule.DeleteLines](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletelines-method-vba-add-in-object-model)
- [Marshal.ReleaseComObject](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject)
- [New-Object -ComObject](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/new-object)
- [about_Execution_Policies](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies)
- [about_Language_Modes (ConstrainedLanguage, AppLocker, WDAC)](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_language_modes)

Otras:

- [Copia pública del texto ADMX "Trust access to Visual Basic Project" (Excel 2016) y su clave](https://admx-offsec-nl.pages.dev/policy/microsoftoffice/microsoft-excel-2016/excel-options/security/trust-center/trust-access-to-visual-basic-project/) (no es de Microsoft; reproduce la plantilla oficial)
- [CIS Microsoft Office Enterprise: recomienda "Trust access to Visual Basic Project" en Disabled](https://tenable.com/audits/items/CIS_Microsoft_Office_Enterprise_v1.2.0_L1.audit:f88b9ac37f182f89b54b3a874dffd508)
- [pywin32: código de DispatchEx](https://github.com/mhammond/pywin32/blob/main/com/win32com/client/__init__.py)
- [Lista de correo python-win32: DispatchEx abre una instancia nueva de Excel](https://mail.python.org/pipermail/python-win32/2008-February/007014.html)
- [xlwings: App (instancia propia, display_alerts, kill, macro)](https://docs.xlwings.org/en/stable/api/app.html)
- [xlwings: Book.macro](https://xlwings.readthedocs.io/en/stable/api/book.html)
- [MCP-Access, commit 3e0fe05: `mcp_access/compile.py` y `mcp_access/vba_exec.py`](https://github.com/unmateria/MCP-Access) (técnicas de §13)
- [SendMessageTimeout y WM_GETTEXT: leer el texto de controles de otro proceso](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtextw) (GetWindowText no lo hace)

## Lo aprendido en Mac que conviene llevar a Windows

Añadido el 14/11/2026 tras las pruebas en Mac ([03-solucion-mac.md](03-solucion-mac.md)). Es VBA, así que aplica igual;
lo que no está comprobado en Windows es cómo reacciona Excel por COM.

- **Envoltorio con llamada directa.** `Application.Run` dentro de un `On Error` no atrapó el error de la
  macro llamada (Mac). Un módulo temporal que llama a la macro **directamente** sí lo atrapa y lo
  devuelve como texto. Es lo que hace `vba.py seguro`; se puede replicar en `vba.ps1` importando un
  módulo `ZZEnvoltorio` antes de `ejecutar` y borrándolo después.
- **SinVentanas.** `MsgBox` e `InputBox` propios en el proyecto ([vba/SinVentanas.bas](../vba/SinVentanas.bas))
  sustituyen a los de VBA y apuntan en la hoja `_avisos`. Probado en Mac.
- **No detener una macro parada.** En Mac, detener con "Finalización" o "Restablecer" una macro lanzada
  desde fuera tumbó Excel 3 de 3 veces. En Windows, con una instancia propia, el script puede matar ese
  Excel por tiempo, pero se pierde lo no guardado. Mismo remedio: compilar antes y usar el envoltorio.
- **Compilar antes de ejecutar.** En Mac se hace con el menú Depurar > Compilar (agent-desktop). En
  Windows lo hace `vba-nuevo.ps1 compilar` (§13) [SUPOSICIÓN, sin probar].
- **Separador decimal.** Con el sistema en castellano, escribir números como texto con punto falló en
  Mac. Pasar los datos como números o por la propiedad `Formula`, que usa notación inglesa.

## 13. Técnicas tomadas de MCP-Access

Añadido el 21/11/2026. [MCP-Access](https://github.com/unmateria/MCP-Access) es un servidor MCP para
Access en Windows (Python y pywin32, MIT según su `pyproject.toml`), revisado en el commit `3e0fe05` con
`security-review` (ClawScan y SkillSpector, más revisión manual: sin malware). No se ha copiado su
código. Se han adaptado a Excel y a PowerShell sus técnicas de `compile.py` y `vba_exec.py`, y en esas
técnicas tropezó con lo mismo que nosotros en Mac el 14/11.
**Todo está sin probar en Windows [SUPOSICIÓN]**. Por eso está en `vba-nuevo.ps1` y no en `vba.ps1`,
que es el que le funcionó a la profesora.

| Técnica | Qué resuelve | Cómo queda en `vba-nuevo.ps1` |
|---|---|---|
| Activar un panel de código del libro antes de compilar (`_ensure_code_pane`) | Compilar actúa sobre el proyecto **activo** del editor, no sobre el pedido. Es nuestro fallo del 14/11 en Mac ([03-solucion-mac.md](03-solucion-mac.md#compilar-vbapy-compilar)) | `CodePane.Show()` y comprobar `VBE.ActiveVBProject.FileName`. Si no es el libro: `SIN VERIFICAR`, no compila |
| Marcar el proyecto como cambiado antes de compilar (insertar y borrar una línea) | Si Excel da el proyecto por compilado, el menú Compilar sale desactivado (visto en Mac una vez de tres) | `InsertLines` y `DeleteLines` en memoria. El libro se abre en solo lectura y no se guarda |
| Pulsar Compilar por el menú del editor | Compila todos los módulos | `VBE.CommandBars.FindControl(1, 578).Execute()`. El ID no depende del idioma |
| Vigía que lee el texto **antes** de cerrar (`_compile_dialog_watchdog`) | Saber qué decía el error, no solo que lo hubo | Trabajo aparte con Win32: busca ventanas `#32770` del PID propio y lee su texto con `WM_GETTEXT` |
| Lugar del error (`ActiveCodePane.GetSelection`) | Módulo, línea y columna | `ERROR COMPILACION: ... - Modulo, Lin 9, Col 20 : linea` |
| Distinguir "no se pudo pulsar Compilar" de "el código no compila" | Evitar culpar al código de un fallo de la herramienta | `SIN VERIFICAR: no se pudo pulsar Compilar (...)` |
| Comprobación final propia (no de MCP-Access) | Excel no tiene `IsCompiled` como Access | `OK` solo si no hubo aviso y el botón Compilar queda desactivado tras pulsarlo |

Lo que **no** se ha tomado:

- **Pulsar "Finalizar" o "Cancelar" en un error en tiempo de ejecución.** MCP-Access lo hace. En Mac,
  detener una macro parada cerró Excel 3 de 3 veces (regla de oro 1). El vigía no pulsa nada en esas
  ventanas: las describe y mata solo nuestro Excel, que es propio e invisible.
- **Descompilar y compactar**, `RunCommand` y `taskkill` de `msaccess.exe`: son de Access.
- **Python y pywin32**: `vba-nuevo.ps1` sigue sin dependencias.

Diferencia con Mac: el vigía de Mac solo cierra avisos titulados "Microsoft Excel", porque allí la
instancia de Excel es compartida. En Windows filtra por el PID de nuestro Excel, así que cierra
cualquier aviso con un solo botón "Aceptar" de esa instancia.

Suposiciones que hay que comprobar (W6, W9, W10):

- Que el `MsgBox` y la ventana de error de VBA son de clase `#32770` y visibles aunque Excel esté invisible.
  Si no lo fueran, queda el límite `-Tiempo` de antes.
- Que `VBE.MainWindow.Visible` y `CodePane.Show()` funcionan con Excel invisible.
- Que `Execute()` no se bloquea para siempre mientras el aviso está abierto (el vigía lo cierra).
- Que el botón Compilar se desactiva tras compilar bien. Si no, la orden dirá `SIN VERIFICAR`, nunca un
  `OK` falso.
- Que `compilar` funciona con las macros apagadas (`AutomationSecurity` ForceDisable), como `modulos`,
  `exportar` e `importar`. Se apagan para no ejecutar código del libro al abrirlo (`Workbook_Open` o
  funciones VBA en celdas al recalcular). Si el editor no compilara así, la orden dirá `SIN VERIFICAR`.
