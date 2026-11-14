# Ejecutar y validar VBA desde Claude Code (Excel para Mac)

Entrada para una IA (Claude Code o Codex) que abre esta carpeta sin saber nada del proyecto. Léela entera
antes de tocar Excel. Repositorio: <https://github.com/miguelgraande/vba-excel-claude-code>.

## Lo primero: ofrecer el ejemplo

En la primera conversación en esta carpeta, después de presentarte en dos líneas, pregunta al usuario:

> Este repositorio trae un ejemplo listo: una herramienta VBA inventada de valor liquidativo con un error
> en las comisiones. Puedo abrirla en tu Excel, pasar los tests (fallan 10 de 15), sacar el código,
> corregirlo, volver a meterlo y comprobar que pasan los 15. Tarda unos dos minutos. ¿Lo ejecuto?

Si responde que sí:

1. Comprueba lo previo sin cambiar nada: que es un Mac, que Excel está instalado y abierto
   (`osascript -e 'application "Microsoft Excel" is running'`) y que estás en la raíz del repositorio.
   Si falta algo, explícalo y remite a [docs/09-instalacion.md](docs/09-instalacion.md).
2. Avisa de dos cosas antes de empezar: Excel puede ponerse delante unos segundos (al abrir libros y al
   compilar; al terminar, el foco vuelve a la aplicación que estaba delante) y conviene no usar Excel
   mientras corre. Si Excel pide "Conceder acceso" a la carpeta o
   "Habilitar macros", es normal (ver [docs/05-buenas-practicas.md](docs/05-buenas-practicas.md)).
3. Si el ajuste de confianza de VBA no está activo, pide permiso para activarlo
   (`python3 -I herramientas/vba.py confianza activar`, o a mano en Excel > Preferencias > Seguridad).
4. Ejecuta la demo guiada y enseña la salida tal cual:

   ```bash
   python3 -I demo/ejecutar_demo.py
   ```

   Si agent-desktop no está instalado, añade `--sin-compilar`. Debe terminar con `DEMO CORRECTA`
   (`FALLO 10 de 15` antes y `OK 15/15` después). El libro vuelve a la versión con el error, con sus
   hojas `Tests` e `Informe` coherentes con ella (`FALLO 10 de 15`), para repetirla.
   Si `compilar` devuelve `ERROR COMPILACION` o `ERROR VENTANA`, el guion para sin ejecutar los tests.
   Probado el 14/11 con y sin `--sin-compilar`.
5. Al terminar, explica en pocas líneas qué ha pasado (qué es cada paso y por qué el test es fiable: el valor
   esperado sale de `demo/referencia.py`, no del VBA) y ofrece desactivar el ajuste de confianza
   (`vba.py confianza desactivar`). Dile también que puede abrir `demo/vl-fondo-corregido.xlsm` para ver
   el después: `Calculo` corregido, tests 15/15 e informe (se abre en la hoja `Informe`).

Si algo falla, no improvises con Excel: busca el mensaje en la tabla de síntomas de
[docs/05-buenas-practicas.md](docs/05-buenas-practicas.md) y respeta las reglas de oro de abajo (sobre todo:
nunca detener una macro parada). Detalle de la demo: [demo/ESPECIFICACION.md](demo/ESPECIFICACION.md); la
solución está aparte, en [demo/solucion/SOLUCION.md](demo/solucion/SOLUCION.md) y
[demo/solucion/Calculo_corregido.bas](demo/solucion/Calculo_corregido.bas).

## Qué es y qué problema resuelve

- Una IA no podía ejecutar VBA porque vive dentro de Excel, y sin ejecutarlo no hay forma de validar las correcciones que propone.
- Aquí la IA escribe el VBA como texto, lo mete en un libro abierto, lo compila, lo ejecuta sin ventanas y recibe el resultado en la terminal.
- El resultado se compara con una referencia independiente en Python, calculada desde la regla de negocio y no desde el VBA.

## Estado

- **Mac: probado.** Ciclo completo el 14/11/2026 (macOS 26.6.2, Excel 16.115 en castellano, Python 3.14). Pruebas M1 a M8 en `docs/04-resultados.md`.
- **Windows: sin probar.** Hay un script (`herramientas/windows/vba.ps1`) y un plan, pero nunca se ha ejecutado.

## Orden de lectura

| # | Documento | Para qué |
|---|---|---|
| 1 | [docs/08-aprendizajes.md](docs/08-aprendizajes.md) | La manera adecuada de trabajar y las tres causas por las que se rompe si no se sigue |
| 2 | [docs/09-instalacion.md](docs/09-instalacion.md) | Instalación en un Mac nuevo paso a paso y cómo comprobar que funciona |
| 3 | [docs/03-solucion-mac.md](docs/03-solucion-mac.md) | Cada pieza, el ciclo de trabajo y la referencia de órdenes de `vba.py` |
| 4 | [docs/05-buenas-practicas.md](docs/05-buenas-practicas.md) | Reglas de oro, permisos, tabla de síntomas y recuperación si Excel se cae |
| 5 | [docs/04-resultados.md](docs/04-resultados.md) | Qué está probado, con qué orden exacta y qué falló |
| 6 | [docs/06-windows.md](docs/06-windows.md) | Vía Windows por COM. Sin probar |
| 7 | [docs/07-vba-python-power-bi.md](docs/07-vba-python-power-bi.md) | Opinión: cuándo VBA, Python o Power BI. No probado |
| 8 | [docs/01-contexto.md](docs/01-contexto.md) | El problema planteado y lo que falta por demostrar |
| 9 | [docs/02-recorrido.md](docs/02-recorrido.md) | Cómo se llegó a la solución, por etapas |
| 10 | [docs/bitacora.md](docs/bitacora.md) | Hechos de la sesión del 14/11 en orden, con evidencia. Fuente de los demás |

`docs/00-plan-inicial.md` es histórico: solo para consultar la lista original de pruebas. La demo se
describe en [demo/ESPECIFICACION.md](demo/ESPECIFICACION.md). Este repositorio revela el error de la demo
(README, `docs/` y `demo/solucion/`). Para probar que una IA lo encuentra sola, hace falta una sesión
limpia en una carpeta aparte, solo con `demo/vl-fondo.xlsm` y los apartados 2 y 4 de
`demo/ESPECIFICACION.md` ([docs/01-contexto.md](docs/01-contexto.md#lo-que-todavía-no-está-demostrado)).

## Herramientas y versiones

| Herramienta | Versión | Uso | Enlace |
|---|---|---|---|
| Excel para Mac | Probado 16.115, **interfaz en castellano** | Ejecuta el VBA. `compilar`, `confianza` y `abrir` buscan botones por su nombre en castellano | Microsoft 365 |
| Python 3 | Probado 3.14. Solo biblioteca estándar | `herramientas/vba.py` y la referencia de la demo | Viene en macOS o `brew install python` |
| AppleScript (`osascript`) y System Events | Los de macOS | Canal con Excel y pulsar avisos | Vienen con macOS |
| agent-desktop | v1.3.1, binario verificado por SHA-256 | `vba.py compilar` y `vba.py confianza` | <https://github.com/vercel-labs/agent-desktop> |
| oletools (`olevba`) | 0.60.2, dependencias en `herramientas/requirements-olevba.txt` | Opcional: leer el VBA de un libro sin abrir Excel | <https://pypi.org/project/oletools/> |
| Este repositorio | | Motor VBA, libro lanzador `prueba.xlsm`, script, demo | <https://github.com/miguelgraande/vba-excel-claude-code> |

No hace falta `xlwings`, Node, npm ni ninguna librería de Python.

## Preparación (resumen)

Detalle y comprobaciones en [docs/09-instalacion.md](docs/09-instalacion.md).

1. Clonar el repositorio y trabajar siempre dentro de su carpeta.
2. agent-desktop v1.3.1 en `~/workspaces/agent-desktop-bin/v1.3.1/` (o la ruta en `AGENT_DESKTOP`), con su suma comprobada.
3. Permisos de macOS de Automatización y Accesibilidad para la terminal o el editor, y Accesibilidad para agent-desktop.
4. Abrir Excel, activar el ajuste de confianza y abrir el lanzador `prueba.xlsm`.
5. Comprobar: `vba ejecutar Hola`, `vba modulos`, `vba compilar`.

## Ciclo obligatorio

Define la función una vez por terminal (en zsh, una variable con espacios no sirve como orden):

```bash
vba() { python3 -I herramientas/vba.py "$@"; }
```

El ciclo completo, orden a orden y con la salida esperada de cada paso, está en
[docs/03-solucion-mac.md, ciclo de trabajo](docs/03-solucion-mac.md#ciclo-de-trabajo). La versión
automática es la demo guiada: `python3 -I demo/ejecutar_demo.py`.

Orden: **abrir -> cargar -> compilar -> seguro**, con el ajuste de confianza activado al empezar y
desactivado al terminar. Solo se sigue adelante si el paso anterior devuelve `OK`. `compilar` responde
`OK compila <libro>`: comprueba que nombra el libro pedido
([por qué](docs/08-aprendizajes.md#un-ok-solo-vale-si-se-sabe-de-qué-es)). Si responde `SIN VERIFICAR`,
trátalo como no comprobado. Todas las órdenes y sus salidas:
[docs/03-solucion-mac.md](docs/03-solucion-mac.md#referencia-de-órdenes).

## Reglas de oro

1. **Nunca detener una macro parada lanzada desde la terminal.** Cerrar esa ventana (Intro, "Finalización",
   "Restablecer") cerró Excel 3 de 3 veces. Si `vba.py` devuelve `ERROR VENTANA VBA (no se ha tocado)`,
   no pulses nada y avisa al usuario. La vigilancia de `vba.py` solo cierra con Aceptar/OK los avisos de
   Excel (`MsgBox` con título "Microsoft Excel") y lo indica con `AVISO cerrado con Aceptar: ... Resultado: ...`.
2. **Un solo agente maneja Excel.** Excel para Mac es una sola instancia compartida. Los demás agentes
   trabajan solo con archivos (`.bas`, documentos, referencia Python).
3. **Nombrar siempre el libro** con `--libro`. Nunca "el libro activo".
4. **Ajuste de confianza activado solo durante la sesión:** `vba confianza activar` al empezar y
   `vba confianza desactivar` al terminar.
5. **Trabajar dentro de la carpeta del repositorio.** Excel para Mac está aislado (sandbox) y pide
   "Conceder acceso" en cada carpeta nueva.
6. **Datos inventados, nunca libros reales.** Una herramienta real necesita antes el permiso de la
   organización propietaria.
7. **No instalar nada sin preguntar al usuario** (tampoco `xlwings` ni paquetes de Python).
8. **No dar nada por probado sin ejecutarlo** y apuntar resultado, orden exacta y error, si lo hubo, en
   `docs/04-resultados.md`.
9. El VBA que se llama desde la terminal es una `Function` que devuelve texto, con `On Error` y sin
   `MsgBox`. Para código heredado con ventanas, cargar `vba/SinVentanas.bas`.
10. **Carpeta fuera de iCloud, o marcada como "Mantener descargado", y con espacio en disco**
    ([por qué](docs/05-buenas-practicas.md#a-reglas-de-oro), regla 13).
11. **Antes de publicar libros, limpiar sus metadatos** con los libros cerrados: Excel reescribe el autor
    y la ruta de la carpeta en cada guardado. Ver
    [docs/05-buenas-practicas.md](docs/05-buenas-practicas.md#antes-de-publicar-o-compartir-un-libro).
12. Castellano directo, frases cortas, sin emojis ni raya larga.

## Si Excel se cae

Síntomas: `Connection is invalid (-609)`, ventana de Microsoft Error Reporting o libros "recuperados" como
`prueba (version 1).xlsb`. Sigue
[docs/05-buenas-practicas.md, sección D](docs/05-buenas-practicas.md#d-recuperarse-si-excel-se-cae): cerrar
sin guardar los recuperados, reabrir con `vba abrir`, revisar con `vba modulos`, recargar desde los `.bas`
y volver a `compilar` antes de `seguro`.
