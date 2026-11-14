# 01 · Contexto: el problema y lo que se ha demostrado

## El problema planteado

En una clase de VBA del 13/11/2026, una profesora de VBA planteó lo siguiente:

- El código VBA vive **dentro** de Excel. Se daba por hecho que una IA puede generarlo y analizarlo,
  pero **no ejecutarlo**.
- Quien mantiene una herramienta antigua de **valor liquidativo** (hecha hacia 2015, con muchos módulos
  y cuyo autor ya no está) recibe de Claude correcciones que copia y pega a mano.
- No tiene forma de **validar** que esas correcciones son correctas: Claude no puede ejecutarlas y
  comprobar el resultado.
- Su alternativa actual, sacar el código de VBA a ficheros para trabajarlo fuera de Excel, es inestable
  cuando hay varios Excel abiertos.

## Qué es VBA y por qué una IA no lo ejecuta sola

- VBA (Visual Basic for Applications) es el lenguaje propio de Office. Se usa sobre todo en Excel.
- El código se guarda dentro del libro (`.xlsm`) o de un complemento (`.xlam`). Si el libro se guarda
  como `.xlsx`, el código se pierde.
- Se edita en el editor de VBA (VBE), que se abre desde Excel. Hay dos tipos de código: funciones
  (`Function`, devuelven un valor) y macros (`Sub`, ejecutan un proceso).
- VBA no tiene un programa propio que lo ejecute. Solo corre dentro de una aplicación de Office.
  No se puede "compilar fuera" y ver si funciona.
- Claude Code corre en la terminal. Puede escribir VBA como texto, pero para ejecutarlo necesita que
  **Excel haga de intermediario**. La pregunta real es **cómo le da órdenes a Excel**.

## VBA no es VB.NET: una confusión frecuente

Si se pregunta a un buscador o a un asistente "cómo usar Claude Code con Visual Basic", la respuesta
suele hablar de otra cosa. Conviene separarlas:

| | VB.NET y VB6 ("Visual Basic" a secas) | VBA (este repositorio) |
|---|---|---|
| Qué es | Lenguajes para programas independientes (escritorio, web) | Lenguaje de macros que vive **dentro** de un libro de Excel |
| Dónde está el código | Archivos de texto en disco (`.vb`, `.frm`, `.bas`, `.cls`) | Dentro del archivo `.xlsm`, un ZIP con el VBA en binario |
| Cómo se compila y ejecuta | Desde la terminal: `dotnet build`, `msbuild`, `vb6.exe /make` (Windows) | Solo dentro de Excel |
| Qué hace falta para usarlo con Claude Code | Nada especial: lee, edita, compila y prueba los archivos como en cualquier proyecto. Extensiones de Visual Studio o VS Code, o la terminal, con un `CLAUDE.md` que diga versión y forma de compilar | Un puente con Excel para meter, ejecutar y sacar el código. Es lo que resuelve este repositorio |

Las respuestas habituales sobre VBA se quedan en **sacar las macros a archivos `.bas` (con Rubberduck,
solo Windows, o con scripts), editarlas y volver a meterlas**. Eso era lo que ya se hacía y no
resolvía el problema: falta **ejecutar** el código y **validar** el resultado sin copiar y pegar.

Dos ideas de esas respuestas que no son exactas:

- "La conexión directa con Excel desde fuera no funciona de forma nativa": sí funciona, por COM en
  Windows y por AppleScript en Mac, con el ajuste "Acceso de confianza al modelo de objetos de proyectos
  de VBA" activado. Es lo que se demuestra aquí.
- "Los `.xlsm` están comprimidos y protegidos": están comprimidos (son un ZIP), pero no protegidos por
  defecto; `olevba` lee su VBA sin abrir Excel.

Lo que sí compartimos con esas recomendaciones: Claude Code en VS Code y un archivo de instrucciones en
la raíz del proyecto ([AGENTS.md](../AGENTS.md), que `CLAUDE.md` importa).

## La meta

No se discute si "lo ejecuta Claude o Excel": lo ejecuta Excel. Hay que demostrar el ciclo completo
sin copiar y pegar:

1. Claude escribe o corrige el VBA como texto.
2. Lo mete en el libro.
3. Lo ejecuta y recibe el resultado en la terminal.
4. Lo compara con una **referencia independiente** (Python), calculada desde la regla de negocio y no
   desde el propio VBA.
5. Corrige y repite, y comprueba que lo que ya funcionaba sigue funcionando.

El paso 4 es el que responde al problema de validación.

## Las vías que propusimos

El plan inicial ([00-plan-inicial.md](00-plan-inicial.md)) proponía tres vías para dar órdenes a Excel:

| Vía | Cómo | Estado |
|---|---|---|
| COM (Windows) | PowerShell o Python manejan Excel desde fuera; Microsoft documenta cómo meter y sacar módulos | Preparada, sin probar: [06-windows.md](06-windows.md) |
| AppleScript (Mac) | `osascript` escribe celdas y ejecuta macros con `run VB macro` | Probada: es la base de la solución en Mac |
| Control de pantalla | Una herramienta que mira la pantalla y pulsa como una persona | Descartada como vía principal: lenta y frágil |

Lo que se añadió al probar, porque AppleScript no puede escribir código VBA:

- **Motor VBA**: un módulo dentro de un libro lanzador que mete, saca y borra el código de cualquier
  libro abierto. El código viaja por celdas.
- **agent-desktop**: maneja los menús de Excel por accesibilidad, sin robar el foco. Se usa para abrir
  el editor, compilar y cambiar el ajuste de confianza.
- **olevba**: lee el VBA de un libro sin abrir Excel. Solo para revisar.

Pieza a pieza en [03-solucion-mac.md](03-solucion-mac.md). Pruebas y alternativas descartadas en
[04-resultados.md](04-resultados.md).

## Lo que pedía el problema y lo que se ha demostrado

Demostrado en Mac el 14/11/2026 con una herramienta inventada de valor liquidativo
([demo/ESPECIFICACION.md](../demo/ESPECIFICACION.md)).

| Lo que pedía el problema | ¿Demostrado? | Cómo |
|---|---|---|
| Modificar el VBA sin copiar y pegar | Sí | `vba.py exportar` saca el módulo a texto; `vba.py cargar` lo vuelve a meter en el libro |
| Ejecutarlo desde Claude Code | Sí | Lo ejecuta Excel; lo lanza Claude desde la terminal y recibe el resultado como texto |
| Validar que el código es correcto | Sí | Tests en VBA contra una referencia independiente en Python: `FALLO 10 de 15` → corrección de una línea → `OK 15/15` |
| No romper lo que funcionaba | Sí | Los 5 días sin comisión salen OK antes y después de la corrección |
| Trabajar con varios Excel abiertos | Sí, en Mac | Cada orden nombra su libro (`--libro`), nunca "el libro activo" |

## Lo que todavía no está demostrado

> **1. Que Claude encuentre el error sin saber dónde está.**
> En la prueba del 14/11, la sesión ya sabía que el error era la base 360/365, porque lo había diseñado
> el agente que preparó la demo. La prueba válida es una sesión limpia de Claude que solo reciba la
> incidencia ("el valor liquidativo sale mal cuando hay comisiones") y la regla de negocio, y que llegue
> sola a `OK 15/15`. Para hacerla, esa sesión debe recibir, en una carpeta aparte, solo
> `demo/vl-fondo.xlsm` y los apartados 2 y 4 de [demo/ESPECIFICACION.md](../demo/ESPECIFICACION.md)
> (más `herramientas/vba.py` y el lanzador para manejar Excel). No sirve abrirla en este repositorio: el
> README, `docs/` y `demo/solucion/` revelan el error.
>
> **2. Windows.**
> La persona que mantiene la herramienta probablemente usa Windows (usa atajos como Alt+F11 y F12).
> La vía COM está preparada ([06-windows.md](06-windows.md) y `herramientas/windows/vba.ps1`), pero no
> se ha ejecutado nunca. En un portátil de empresa, informática puede bloquear por directiva el ajuste
> de confianza que hace falta para meter y sacar código.
>
> **3. La herramienta real.**
> Es más grande y de 2015. Si tiene muchos `MsgBox`, formularios (`UserForm`), `Application.InputBox` o
> errores sin controlar, hay que aplicar `compilar`, `seguro` y SinVentanas
> ([05-buenas-practicas.md](05-buenas-practicas.md#a-reglas-de-oro)). SinVentanas no cubre formularios ni
> `Application.InputBox`. Además, antes de tocarla hace falta la autorización de la organización
> propietaria: solo se ha trabajado con datos inventados.

## Condiciones de las pruebas

- Mac con macOS 26.6.2 y Excel 16.115 en castellano. La vía es AppleScript.
- Solo libros con datos inventados.

Siguiente: [02-recorrido.md](02-recorrido.md), cómo se llegó a la solución.
