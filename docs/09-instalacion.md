# Instalación en un Mac nuevo, paso a paso

Para una persona o una IA que parte de cero. Todo probado el 14/11/2026 en macOS 26.6.2 con Excel 16.115
(interfaz en castellano), salvo lo marcado "sin comprobar". Por qué se usa cada pieza:
[08-aprendizajes.md](08-aprendizajes.md) y [03-solucion-mac.md](03-solucion-mac.md).

## Qué se usa y qué es obligatorio

| Herramienta | Para qué | ¿Obligatoria? | Dónde |
|---|---|---|---|
| Microsoft Excel para Mac | Ejecuta el VBA | Sí | Microsoft 365 |
| Python 3 | `herramientas/vba.py` (solo biblioteca estándar) y la referencia de la demo | Sí | Ya viene en macOS o `brew install python` |
| AppleScript (`osascript`) y System Events | Hablar con Excel y pulsar "Habilitar macros" | Sí | Vienen con macOS |
| Este repositorio | Motor VBA, libro lanzador, script, demo | Sí | <https://github.com/miguelgraande/vba-excel-claude-code> |
| agent-desktop v1.3.1 | `vba.py compilar` y `vba.py confianza` | Recomendada (sin ella, compilar y el ajuste se hacen a mano) | https://github.com/vercel-labs/agent-desktop |
| oletools 0.60.2 (`olevba`) | Leer el VBA de un libro sin abrir Excel. Dependencias fijadas en `herramientas/requirements-olevba.txt` | Opcional | https://pypi.org/project/oletools/ |
| PowerShell (`pwsh`) | Solo comprobar la sintaxis del script de Windows; en Mac no hay COM | Opcional | https://github.com/PowerShell/PowerShell |

No hace falta: `xlwings`, Node, npm ni ninguna librería de Python.

## Paso 1. Clonar el repositorio

```bash
mkdir -p ~/workspaces && cd ~/workspaces
git clone https://github.com/miguelgraande/vba-excel-claude-code.git
cd vba-excel-claude-code
```

Trabaja **siempre dentro de esta carpeta**: Excel pedirá "Conceder acceso" a ella la primera vez (paso 4) y
cada carpeta nueva vuelve a pedirlo.

**Dónde clonarlo.** Fuera de carpetas sincronizadas con iCloud (Escritorio y Documentos lo están si
"Escritorio y Documentos" está activado en iCloud Drive), o con la carpeta del proyecto marcada como
"Mantener descargado" en el Finder. Deja también espacio libre en disco
([por qué](05-buenas-practicas.md#a-reglas-de-oro), regla 13).

Define una función para no repetir la ruta del script (en zsh, una variable con espacios no funciona como orden):

```bash
vba() { python3 -I herramientas/vba.py "$@"; }
```

## Paso 2. agent-desktop (recomendado)

Se usa el binario publicado, verificado por su suma SHA-256, sin npm ni instalación global (el paquete npm
ejecuta un script de instalación). Comandos para Apple Silicon (arm64):

```bash
B=~/workspaces/agent-desktop-bin/v1.3.1 && mkdir -p "$B" && cd "$B"
gh release download v1.3.1 -R vercel-labs/agent-desktop \
  -p 'agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz' -p 'checksums.txt'
grep 'agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz' checksums.txt | shasum -a 256 -c   # debe decir OK
tar -xzf agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz
./agent-desktop status          # permisos: accessibility, automation, screen_recording
cd -
```

- SHA-256 esperado del archivo para Apple Silicon (línea de `checksums.txt` de la release v1.3.1):

  ```text
  9b7b7cfc74892eab46c2c85d6cb7dc4863bf8dab0a222cfcfb8ae4604c69c86c  agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz
  ```

  Si `shasum -a 256 agent-desktop-v1.3.1-aarch64-apple-darwin.tar.gz` da otro valor, no lo descomprimas.
- El archivo trae dos ejecutables: `agent-desktop` y `agent-desktop-macos-helper`. El segundo debe quedar
  **en la misma carpeta** que el primero; no los separes al moverlos.
- En un Mac Intel, el archivo es `agent-desktop-v1.3.1-x86_64-apple-darwin.tar.gz` (SHA-256 en la misma
  `checksums.txt`: `469da10cfb68928c58caff88b92b30526e0ca895ac06bc86b7f533b1d9b309b3`). Sin probar en Intel.
- Sin `gh`, descárgalo de la página de releases de GitHub y comprueba la suma igual.
- `vba.py` lo busca en `~/workspaces/agent-desktop-bin/v1.3.1/agent-desktop`. Si está en otro sitio:
  `export AGENT_DESKTOP=/ruta/al/agent-desktop`.
- Si `status` dice que falta Accesibilidad: `./agent-desktop permissions --guide`.

## Paso 3. Permisos de macOS

La primera vez que la terminal (o VS Code) controle Excel, macOS pedirá permiso. Acéptalo, o actívalo en
Ajustes del Sistema > Privacidad y seguridad:

- **Automatización**: la terminal o VS Code puede controlar "Microsoft Excel" y "System Events".
- **Accesibilidad**: para la terminal o VS Code (System Events) y para agent-desktop.

## Paso 4. Excel

1. Abre Excel.
2. Activa el ajuste de confianza (necesario para leer y escribir código VBA):
   - con agent-desktop: `vba confianza activar`
   - a mano: Excel > Preferencias > Seguridad > marcar "Acceso de confianza al modelo de objetos de proyectos de VBA".
   No cambies la seguridad de macros: deja "Deshabilitar todas las macros con notificación".
3. Abre el libro lanzador, que ya trae el motor dentro:

   ```bash
   vba abrir prueba.xlsm
   ```

   Si Excel muestra "Conceder acceso" a la carpeta, pulsa **Conceder acceso** (es el aislamiento de Excel
   para Mac; ver [05-buenas-practicas.md](05-buenas-practicas.md)).

## Paso 5. Comprobar que todo funciona

```bash
vba ejecutar Hola               # -> hola desde VBA hh:mm:ss    (AppleScript y macros habilitadas)
vba modulos                     # -> lista con Motor            (ajuste de confianza activo)
vba compilar                    # -> OK compila prueba.xlsm     (agent-desktop)
```

`compilar` con el lanzador no se ha repetido desde el cambio del 14/11 (activar el proyecto pedido en el
editor); con `--libro vl-fondo.xlsm` sí: `OK compila vl-fondo.xlsm`. Mientras corre, Excel se pone
delante; al terminar, el foco vuelve a la aplicación que estaba delante.

Si algo falla, busca el mensaje en la tabla de síntomas de [05-buenas-practicas.md](05-buenas-practicas.md).

## Paso 6. La demo, de principio a fin

La forma corta es el guion, que hace todos los pasos:

```bash
python3 -I demo/ejecutar_demo.py                 # termina en DEMO CORRECTA
python3 -I demo/ejecutar_demo.py --sin-compilar  # si no tienes agent-desktop
```

Si `compilar` devuelve `ERROR COMPILACION` o `ERROR VENTANA`, el guion para sin ejecutar los tests.
Probado el 14/11 con y sin `--sin-compilar`: `DEMO CORRECTA`. Al terminar, el libro vuelve a la versión
con el error, con sus hojas `Tests` e `Informe` coherentes con ella (`FALLO 10 de 15`). El resultado
final, ya hecho, está en `demo/vl-fondo-corregido.xlsm` (se abre en la hoja `Informe`).

A mano, paso a paso: [03-solucion-mac.md, ciclo de trabajo](03-solucion-mac.md#ciclo-de-trabajo).

Detalle de la demo: [../demo/ESPECIFICACION.md](../demo/ESPECIFICACION.md). La solución está aparte, en
[../demo/solucion/SOLUCION.md](../demo/solucion/SOLUCION.md) y
[../demo/solucion/Calculo_corregido.bas](../demo/solucion/Calculo_corregido.bas). Para probar que una IA
encuentra el error sola no basta con ocultarle `demo/solucion/`: ver
[01-contexto.md](01-contexto.md#lo-que-todavía-no-está-demostrado).

## Paso 7. Al terminar

```bash
vba confianza desactivar
```

Guarda el trabajo (`cargar` ya guarda; si no, `vba guardar --libro <libro>`).

**Si vas a publicar o compartir los libros:** Excel escribe en cada guardado el autor y la ruta absoluta
de la carpeta dentro del `.xlsm`. Límpialos con los libros cerrados justo antes de publicarlos: ver
[05-buenas-practicas.md](05-buenas-practicas.md#antes-de-publicar-o-compartir-un-libro).

## Opcional: olevba

En un entorno aislado de Python, con las versiones fijadas en
[herramientas/requirements-olevba.txt](../herramientas/requirements-olevba.txt) (oletools 0.60.2 y sus
dependencias, congeladas tal como se probaron el 14/11/2026). Desde la carpeta del repositorio:

```bash
python3 -m venv ~/workspaces/venvs/oletools
~/workspaces/venvs/oletools/bin/pip install -r herramientas/requirements-olevba.txt
~/workspaces/venvs/oletools/bin/olevba --code copia-del-libro.xlsm      # código de todos los módulos
~/workspaces/venvs/oletools/bin/olevba -j copia-del-libro.xlsm          # lo mismo en JSON
```

Trabaja sobre una copia del libro. Los avisos "Hex Strings" y "Base64 Strings" de `olevba -a` fueron
falsos positivos en nuestros libros.

## Si el libro lanzador se pierde o se estropea

`prueba.xlsm` ya lleva el motor. `vba cargar` se niega a reemplazar `Motor`, `Módulo1` o `Cargadora` en
el lanzador (el motor se reescribiría a sí mismo, y la cargadora es la vía para reinstalarlo). Para **actualizar el motor**, rehaz el lanzador en un libro nuevo con los pasos de abajo y el
`vba/Motor.bas` nuevo: es el camino que ya se siguió el 14/11. Hipótesis sin comprobar: usar la
cargadora (`Cargar`) del lanzador existente para reemplazar `Motor`, ya que no es el motor quien se
reescribe. Se haría a mano, como en el paso 4.3 de la preparación de
[03-solucion-mac.md](03-solucion-mac.md#preparación-una-sola-vez): código de `Motor.bas` en la columna A
de la hoja `_codigo`, **sin marca**, `Motor` en B1, y `run VB macro "prueba.xlsm!Cargar"`. `vba.py` no
sirve para esto: escribe las líneas con la marca `|` (que `Cargar` no quita) y borra la columna B. El
lanzador debe llamarse `prueba.xlsm`: el nombre está fijo en `vba.py`. Guarda una copia del lanzador
antiguo antes de sustituirlo.

Para rehacerlo en un libro nuevo:

1. Crea un libro y guárdalo como `.xlsm` en la carpeta del repositorio. A mano es lo seguro: el 14/11,
   `save workbook as` desde AppleScript llegó a dar `Parameter error (-50)` al crear libros nuevos.
2. Editor de VBA (Herramientas > Macro > Editor de Visual Basic) > Insertar > Módulo, y pega
   `vba/Cargadora.bas`. Guarda.
3. Con la cargadora dentro, el motor se instala solo: ver el paso 4 de "Preparación (una sola vez)" en
   [03-solucion-mac.md](03-solucion-mac.md#preparación-una-sola-vez).
