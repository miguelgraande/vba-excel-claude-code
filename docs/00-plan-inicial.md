# Ejecutar VBA desde Claude Code: opciones y pruebas

> Plan del 13/11/2026, conservado como histórico. Resultados en [04-resultados.md](04-resultados.md).
> Se han corregido los enlaces para que funcionen desde `docs/` y se ha anonimizado: sin nombres, citas
> ni enlaces a la grabación de la clase. El problema planteado está en [01-contexto.md](01-contexto.md).

Origen: debate con una profesora de VBA en una clase del 13/11/2026.
Objetivo: demostrar en la próxima sesión que Claude Code puede escribir VBA, meterlo en un libro, ejecutarlo, leer el resultado y corregirlo, sin copiar y pegar.

Estado: **probado en Mac el 14/11/2026.** Cómo lo hacemos, qué funcionó y qué no: [03-solucion-mac.md](03-solucion-mac.md). Windows: [06-windows.md](06-windows.md), sin probar.

---

## 1. La idea clave

VBA no tiene un programa propio que lo ejecute. Solo funciona **dentro** de una aplicación de Office, en nuestro caso Excel. No se puede "compilar fuera" y ver si va.

Por eso Claude siempre necesita a **Excel como intermediario**. La pregunta no es si Claude puede ejecutar VBA, sino **cómo le da órdenes a Excel**.

El problema planteado: abrir un Excel desde fuera sí se puede; lo que se daba por imposible es ejecutar el código desde el entorno de Visual Basic dentro del Excel. Y el problema de fondo: **no hay forma de validar** que el código que propone Claude es correcto.

### Nombres que se mezclan

| Nombre | Qué es |
|---|---|
| VBA (Visual Basic for Applications) | El lenguaje. Vive dentro de Excel, Word, PowerPoint |
| Editor de VBA (VBE) | La ventana que abre Alt+F11 dentro de Excel |
| VS Code (Visual Studio Code) | El editor donde corre Claude Code. No ejecuta VBA |
| COM | El "mando a distancia" de Office en **Windows** |
| AppleScript | El equivalente en **Mac**, más limitado |

---

## 2. Esquema

```
                    Claude Code (terminal / VS Code)
                                 |
          escribe el código VBA como texto (.bas)
                                 |
        +------------------------+------------------------+
        |                        |                        |
   WINDOWS: COM            MAC: AppleScript      CUALQUIERA: control
   (Python + pywin32)      (osascript)           de pantalla (computer use)
        |                        |                        |
        +------------------------+------------------------+
                                 |
                               EXCEL
              mete el código, ejecuta la macro, guarda
                                 |
               resultado en celdas de una hoja "Tests"
                                 |
             Claude lee las celdas y compara con Python
                                 |
                  OK -> fin     FALLO -> corrige y repite
```

---

## 3. Las tres vías

| | Windows: COM | Mac: AppleScript | Control de pantalla |
|---|---|---|---|
| Cómo funciona | Un script de Python habla con Excel por COM | `osascript` manda órdenes a Excel | Claude ve la pantalla, pulsa Alt+F11, pega, F5 |
| Meter código en el libro | Sí, documentado (`VBComponents.Import`, `AddFromString`) | **Dudoso.** Hay informes contradictorios y no hay documentación de Microsoft para Mac | Sí, pegando en el editor |
| Ejecutar macro | Sí (`xl.Run`) | Sí, solo macros que ya estén en el libro (`run VB macro`). Sintaxis sin documentación oficial | Sí (F5) |
| Leer resultado | Sí | Sí | Por captura de pantalla |
| Hay que instalar | Python + `pywin32` | Nada (viene con macOS). `xlwings` opcional | Herramienta de computer use |
| Velocidad y fiabilidad | Alta | Media | Baja: lenta y frágil |
| Para quién | **La profesora** (su ordenador parece Windows: usa Alt+F11 y F12) | Quien prepara la demo (Mac) | Plan B o para cerrar ventanas que bloquean |

Ajuste obligatorio en las dos primeras vías: **"Confiar en el acceso al modelo de objetos de proyectos de VBA"**.
- Windows: Archivo > Opciones > Centro de confianza > Configuración > Configuración de macros.
- Mac: Excel > Preferencias > Seguridad.

---

## 4. Lista de pruebas para ir descartando

Siempre con un **libro de prueba con datos inventados** (`prueba.xlsm`). Nunca con un libro de cliente.
Cada prueba depende de la anterior. Si una falla, se apunta por qué y se pasa a la alternativa indicada.

### Mac (AppleScript): se puede probar ya en tu Mac

| ID | Pregunta | Cómo probarla | Si falla |
|---|---|---|---|
| M0 | ¿Está todo preparado? | Excel abierto con `prueba.xlsm` y el ajuste de confianza activado | - |
| M1 | ¿La terminal puede hablar con Excel? | `osascript -e 'tell application "Microsoft Excel" to get name of active workbook'`. La primera vez macOS pide permiso de automatización | Ajustes del Sistema > Privacidad y seguridad > Automatización |
| M2 | ¿Puede escribir y leer celdas? | `osascript -e 'tell application "Microsoft Excel" to set value of range "A1" of active sheet to 42'` y leerla después | Revisar la sintaxis del diccionario de Excel |
| M3 | ¿Puede ejecutar una macro que ya existe? | Macro `Hola` que escribe en una celda. `run VB macro "Hola"` dentro de `tell application "Microsoft Excel"` | Probar con el nombre del libro delante (`"prueba.xlsm!Hola"`) |
| M4 | ¿Puede meter código nuevo **sin archivos**? | Claude escribe el código en una celda de una hoja oculta `_codigo`. Una macro `Cargadora` lo lee y lo mete con `CodeModule.AddFromString`. Se lanza con M3 | Si falla aquí, en Mac solo se pueden ejecutar macros existentes: pasar a Windows |
| M5 | ¿Puede importar un `.bas` desde archivo? | `VBComponents.Import` desde una carpeta. Ojo: Office en Mac está aislado (sandbox) y puede pedir permiso de acceso al archivo | Usar la carpeta propia de Excel (`~/Library/Containers/com.microsoft.Excel/Data/`) o quedarse con M4 |
| M6 | ¿Qué pasa si la macro da error? | Macro con un error a propósito, con y sin `On Error`. Ver si la ventana de error bloquea `osascript` | Obligar a que todo el código escriba sus errores en la hoja |
| M7 | ¿Funciona el ciclo completo? | Función de prueba + tests en VBA + mismo cálculo en Python. Claude corrige hasta que todo da OK | - |
| M8 | ¿Aguanta con varios libros abiertos? | Abrir otro Excel y repetir M7 nombrando el libro, no "el libro activo" | Es la inestabilidad que se comentó en clase |

### Windows (COM): necesita un Windows con Excel

| ID | Pregunta | Cómo probarla | Si falla |
|---|---|---|---|
| W0 | ¿Está todo preparado? | Python, `pip install pywin32`, ajuste de confianza activado | En un portátil de empresa, informática puede bloquear el ajuste |
| W1 | ¿Python puede abrir Excel? | `xl = win32com.client.DispatchEx("Excel.Application")` y `wb = xl.Workbooks.Open(ruta)` | Revisar la instalación de Office |
| W2 | ¿Lee y escribe celdas? | `wb.Sheets("Tests").Range("A1").Value` | - |
| W3 | ¿Ejecuta una macro? | `xl.Run("prueba.xlsm!Hola")` | - |
| W4 | ¿Saca los módulos a texto? | `for c in wb.VBProject.VBComponents: c.Export(...)`. Así se pueden editar en VS Code y guardar en git | Error de acceso: falta el ajuste de confianza |
| W5 | ¿Mete código nuevo y lo ejecuta? | Borrar el módulo, `VBComponents.Import(ruta_bas)`, ejecutar | - |
| W6 | ¿Qué pasa con los errores? | `xl.DisplayAlerts = False` y una macro con un error sin controlar | Igual que M6 |
| W7 | ¿Funciona el ciclo completo? | Como M7 | - |
| W8 | ¿Aguanta con varios libros? | `DispatchEx` abre una instancia de Excel propia y separada | - |

### Control de pantalla: solo para comparar o desbloquear

| ID | Pregunta | Cómo probarla |
|---|---|---|
| C1 | ¿Claude puede hacerlo como una persona? | Que abra el editor, pegue el código, pulse F5 y lea la celda. Medir tiempo y fallos frente a M7 o W7 |
| C2 | ¿Puede cerrar una ventana que bloquea? | Provocar el error de M6 y que la cierre |

### Herramientas de apoyo

- **olevba** (oletools): lee el VBA de un `.xlsm` sin abrir Excel. Solo lee. Funciona en cualquier sistema, también en Linux o en un servidor. Útil para que Claude analice una herramienta antes de tocarla.
- **xlwings**: librería de Python que funciona en Windows y Mac (en Mac usa AppleScript por debajo). Más cómoda, pero hay que instalarla.
- **agent-desktop** (https://github.com/vercel-labs/agent-desktop, Apache-2.0): control de apps de escritorio por el árbol de accesibilidad, sin capturas ni mover el ratón. Candidata para meter código en el editor de VBA sin robar el foco. Hay que instalarla (npm). Sin probar.
- **Claude en Excel** (complemento): según la profesora no ve el VBA. Sin comprobar.

---

## 5. Qué habría que ver para darlo por bueno

Una versión pequeña de su caso real (herramienta de valor liquidativo de 2015, muchos módulos):

1. Libro inventado con tres módulos (datos, cálculo, informe) y un error escondido.
2. Una incidencia escrita como las reales: "el valor liquidativo sale mal cuando hay comisiones".
3. Claude saca los módulos, encuentra el error, lo corrige, lo mete en el libro y ejecuta.
4. Compara con el mismo cálculo en Python, que hace de referencia independiente.
5. Vuelve a pasar los casos que ya funcionaban para comprobar que no ha roto nada.

Responde directamente al problema de fondo: no había forma de validar que el código era correcto.

---

## 6. Riesgos y límites

- **Ventanas emergentes.** Un `MsgBox` o un error sin controlar para todo hasta que alguien pulse. Un código de 2015 tendrá muchas.
- **Mac.** `run VB macro` y la inyección de código no tienen documentación oficial. Hay que probarlo, no suponerlo.
- **Portátil de empresa.** Informática puede tener bloqueado el ajuste de confianza.
- **Datos de clientes.** Antes de conectar Claude a una herramienta real, que la organización propietaria lo autorice.

---

## 7. Contexto

- El problema planteado y las vías propuestas, resumidos: [01-contexto.md](01-contexto.md).

---

## 8. Resultados de las pruebas

| ID | Fecha | Resultado | Comando o código usado | Notas y errores |
|---|---|---|---|---|
| M1 | 14/11/2026 | OK | `osascript -e 'tell application "Microsoft Excel" to return version'` devolvió 16.115 | No pidió permiso de automatización |
| M2 | 14/11/2026 | OK | `set value of range "A1" of active sheet to 42` y lectura: 42.0 | También creó `prueba.xlsm` con `save workbook as ... file format macro enabled XML file format` |
| M3 | 14/11/2026 | OK | `tell application "Microsoft Excel" to run VB macro "prueba.xlsm!Hola"` | Devuelve el valor de una Function a la terminal |
| M4 | 14/11/2026 | OK | Código en hoja `_codigo`, `run VB macro "prueba.xlsm!Cargar"` y después `"prueba.xlsm!Nueva"` | «OK Generado 3 lineas» y «codigo nuevo, 2+3 = 5». Recarga con cambio: «version 2, 10*7 = 70». Sin ajuste de confianza: «ERROR 1004», sin ventana |
| M5 | 14/11/2026 | OK con límite | Función con `ThisWorkbook.VBProject.VBComponents.Import ruta` | Lee desde el proyecto y desde la carpeta de Excel. Si el módulo existe crea `Importado1`: para reemplazar, mejor `vba.py cargar` |
| M6 | 14/11/2026 | OK con envoltorio | `vba.py seguro SinControl`; `SinVentanas.bas` para `MsgBox` | Sin envoltorio sale la ventana y, al cerrarla, **Excel se cae** (3/3). Con envoltorio: `ERROR 11: División por cero`. Ver [03-solucion-mac.md](03-solucion-mac.md) |
| M7 | 14/11/2026 | OK | Demo `demo/`: `seguro EjecutarTests`, `exportar Calculo`, corregir, `cargar`, `seguro EjecutarTests` | `FALLO 10 de 15` → `OK 15/15` |
| M8 | 14/11/2026 | OK | Lanzador `prueba.xlsm` + `vl-fondo.xlsm` abiertos, `--libro` en cada orden | Nombrar siempre el libro, nunca "el activo" |

---

## 9. Fuentes

- [Microsoft KB 219905: añadir y ejecutar macros VBA desde fuera (Windows)](https://support.microsoft.com/en-us/help/219905/how-to-dynamically-add-and-run-a-vba-macro-from-visual-basic)
- [Microsoft: AppleScriptTask y el aislamiento de Office en Mac](https://learn.microsoft.com/en-us/office/vba/office-mac/applescripttask)
- [SumProduct: ajustes de seguridad de Excel para Mac](https://sumproduct.com/blog/article/excel-for-mac-security-settings-in-excel-for-mac)
- [Apple Discussions: lanzar macros de Excel con AppleScript](https://discussions.apple.com/thread/353150)
- [xlwings: Book.macro](https://xlwings.readthedocs.io/en/stable/api/book.html)
