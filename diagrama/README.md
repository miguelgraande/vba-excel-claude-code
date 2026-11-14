# Diagramas

Dos diagramas hechos con [stack-diagram](https://github.com/caru-ini/stack-diagram) v0.3.0.

| Diagrama | Qué muestra | Fuente editable | Página | Imágenes |
|---|---|---|---|---|
| Arquitectura en Mac | Claude Code y la terminal llaman a `vba.py`. `vba.py` habla con Excel por AppleScript, y con System Events y agent-desktop solo para avisos y menús. Las protecciones de macOS y Excel forman una capa intermedia. Dentro de Excel, el lanzador `prueba.xlsm` (hoja `_codigo` y Motor) mete y saca código de `demo/vl-fondo.xlsm`. Fuera de Excel, la referencia en Python y olevba. Incluye el ciclo `abrir -> cargar -> compilar -> seguro` | `fuente.html` | `arquitectura.html` | `arquitectura.png`, `arquitectura-dark.png` |
| Descartado y Windows | Lo descartado o no usado, atenuado y tachado, con su motivo, y la vía COM de Windows (escrita, sin probar) | `fuente-otros.html` | `descartado-windows.html` | `descartado-windows.png`, `descartado-windows-dark.png` |

Datos sacados de `docs/03-solucion-mac.md`, `docs/04-resultados.md`, `docs/06-windows.md` y
`docs/08-aprendizajes.md` (pruebas del 14/11/2026). Los iconos están en `icons/`, con su origen en
`icons/manifest.json` (devicon, svgl y Lucide). Los logos son marcas de sus dueños.

## Reconstruir

Desde la raíz del repositorio, con `SD` apuntando a la carpeta de la skill stack-diagram:

```bash
SD=<ruta>/stack-diagram

# página autocontenida (iconos y fuentes incrustados)
python3 -I "$SD/scripts/build.py" diagrama/fuente.html diagrama/icons diagrama/arquitectura.html --fonts "$SD/assets/fonts"
python3 -I "$SD/scripts/build.py" diagrama/fuente-otros.html diagrama/icons diagrama/descartado-windows.html --fonts "$SD/assets/fonts"

# PNG claro y oscuro con Google Chrome sin ventana (sin Playwright)
python3 -I diagrama/captura_chrome.py diagrama/arquitectura.html diagrama/arquitectura.png 1110 1240
python3 -I diagrama/captura_chrome.py diagrama/descartado-windows.html diagrama/descartado-windows.png 1110 540
```

`captura_chrome.py` hace lo mismo que `scripts/render_png.py` de la skill: captura solo el lienzo,
a escala 2 y en los dos temas. Se usa porque en el Mac donde se hicieron no está Playwright. Con Playwright
instalado vale también:
`python3 "$SD/scripts/render_png.py" diagrama/arquitectura.html diagrama/arquitectura.png --scale 2`.

Las coordenadas se colocan a mano (reglas en `references/geometry.md` de la skill). Si se cambia
el tamaño del lienzo, hay que cambiarlo en `.canvas`, en `svg.edges` y en su `viewBox`, y pasar el
nuevo ancho y alto a `captura_chrome.py`. Después de cada cambio, mirar los PNG antes de publicar.
