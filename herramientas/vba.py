#!/usr/bin/env python3
"""Puente terminal <-> VBA de Excel para Mac, por AppleScript. Sin dependencias.

Requisitos: Excel abierto con el libro lanzador (prueba.xlsm) que contiene el modulo Motor,
y el ajuste "Acceso de confianza al modelo de objetos de proyectos de VBA" activado.

Uso:
  vba.py abrir LIBRO.xlsm [--sin-macros]               abre el libro y pulsa Habilitar macros
  vba.py cargar ARCHIVO.bas [--libro L] [--modulo M] [--sin-guardar]
                                                       mete el codigo (reemplaza el modulo) y guarda
  vba.py guardar [--libro L]                           guarda el libro
  vba.py compilar [--libro L]                          compila sin ejecutar (menu del editor, via agent-desktop)
  vba.py confianza activar|desactivar                  ajuste "Acceso de confianza al modelo de objetos VBA" (agent-desktop)
  vba.py dialogos                                      lista los dialogos de Excel (subrol, titulo, botones, textos)
  vba.py ejecutar MACRO [ARG...] [--libro L]           ejecuta y muestra lo que devuelve
  vba.py seguro MACRO [--sub] [--libro L]              ejecuta dentro de un On Error generado (sin ventanas de error)
  vba.py exportar MODULO [--libro L] [--salida F]      saca el codigo del modulo a texto
  vba.py modulos [--libro L]                           lista modulos, tipo y lineas
  vba.py borrar MODULO [--libro L]                     elimina un modulo
  vba.py hoja-csv ARCHIVO.csv HOJA [--libro L]         copia un CSV (;) a una hoja desde A1
  vba.py leer HOJA RANGO [--libro L]                   lee un rango y lo imprime separado por ;

--libro vacio = el libro lanzador. El codigo pasa por la hoja _codigo del lanzador.
--timeout SEGUNDOS (ejecutar, seguro, cargar, exportar; por defecto 60): espera antes de mirar las ventanas.
Todos los valores (libros, macros, rutas, botones) pasan a AppleScript por argv, nunca dentro del texto del script.
"""
import argparse
import functools
import json
import os
import time
import pathlib
import subprocess
import sys

LANZADOR = "prueba.xlsm"
MARCA = "|"
# Modulos del lanzador que "cargar" no reemplaza: el motor se reescribiria a si mismo mientras se
# ejecuta, y la cargadora (en Modulo1) es la via para reinstalarlo.
PROTEGIDOS = {"motor", "módulo1", "modulo1", "cargadora"}
MAX_AVISOS = 10


def osa(script, *args, timeout=120):
    try:
        r = subprocess.run(["osascript", "-", *map(str, args)], input=script,
                           capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise SystemExit(f"ERROR TIEMPO: AppleScript sin respuesta en {timeout} s "
                         "(Excel puede estar esperando en una ventana: mira `vba.py dialogos`)")
    if r.returncode:
        raise SystemExit("AppleScript: " + r.stderr.strip())
    return r.stdout.rstrip("\n")

# Foco: compilar, confianza, abrir, la vigilancia y dialogos traen Excel al frente (menus del editor, avisos,
# ventanas que System Events solo ve con Excel delante). Al terminar, incluso con error, se devuelve el foco
# a la aplicacion que estaba delante, salvo que fuera el propio Excel. Asi no hay "saltos" que dejen al
# usuario en Excel mientras trabaja en otra ventana.
def app_delante():
    try:
        return osa('tell application "System Events" to return name of first application process whose frontmost is true',
                   timeout=10)
    except SystemExit:
        return ""


def devolver_foco(app):
    if app and app != "Microsoft Excel":
        try:
            osa('on run argv\ntell application "System Events" to set frontmost of application process (item 1 of argv) to true\n'
                'end run\n', app, timeout=10)
        except SystemExit:
            pass


def con_foco(funcion):
    @functools.wraps(funcion)
    def envuelta(*args, **kwargs):
        app = app_delante()
        try:
            return funcion(*args, **kwargs)
        finally:
            devolver_foco(app)
    return envuelta


TRAER = """
on traerExcel()
    -- Al ponerse Excel delante, macOS puede cambiar de escritorio: esperar a que se vean sus ventanas.
    tell application "System Events" to tell process "Microsoft Excel"
        set frontmost to true
        repeat 15 times
            if (count of windows) > 0 then exit repeat
            delay 0.2
        end repeat
    end tell
    delay 0.3
end traerExcel
"""


# Lista las ventanas de Excel vistas por System Events: titulo, subrol, botones y textos.
# Separadores de control (29, 30, 31) para que ningun texto de la ventana los confunda.
# Comprobado el 14/11 con un MsgBox real: titulo vacio, subrol AXDialog, boton "Aceptar" y el titulo del
# aviso ("Microsoft Excel") como primer texto. La ventana principal del editor de VBA ("Microsoft Visual
# Basic - libro [en ejecucion]") es AXStandardWindow, igual que las de los libros.
VENTANAS = TRAER + """
on lista(xs, sep)
    set r to {}
    repeat with x in xs
        set v to contents of x
        if v is not missing value then set end of r to (v as text)
    end repeat
    set AppleScript's text item delimiters to sep
    set s to r as text
    return s
end lista

on run argv
    set FS to character id 30
    set US to character id 31
    set RS to character id 29
    set salida to ""
    tell application "System Events" to tell process "Microsoft Excel"
        if (count of argv) > 0 and item 1 of argv is "delante" then my traerExcel()
        repeat with w in windows
            try
                set t to name of w
                if t is missing value then set t to ""
                set sr to ""
                try
                    set sr to subrole of w
                    if sr is missing value then set sr to ""
                end try
                set bs to my lista(name of every button of w, US)
                set ts to my lista(value of every static text of w, US)
                set salida to salida & t & FS & sr & FS & bs & FS & ts & RS
            end try
        end repeat
    end tell
    return salida
end run
"""


def traer_excel():
    """Pone Excel delante (frontmost por System Events, sin teclas) y espera a que se vean sus ventanas.
    `activate` no basta: con el usuario en otra aplicacion, macOS a veces no cambia y los menus de Excel
    salen desactivados (visto el 14/11 con compilar)."""
    osa(TRAER + "\non run\nmy traerExcel()\nend run\n", timeout=20)


def ventanas(delante=True):
    """Ventanas de Excel (avisos, errores de VBA, MsgBox, libros, editor) vistas por System Events.
    Devuelve una lista de dict con titulo, subrol, botones y textos. Si falla, lista vacia.
    Comprobado el 14/11: con Excel en segundo plano, System Events no ve ninguna de sus ventanas (0), ni
    siquiera un MsgBox abierto. Por eso, por defecto, pone Excel al frente (frontmost, sin teclas) antes de leer."""
    try:
        bruto = osa(VENTANAS, *(["delante"] if delante else []), timeout=20)
    except SystemExit:
        return []
    salida = []
    for reg in bruto.split("\x1d"):
        if not reg.strip():
            continue
        partes = (reg.split("\x1e") + ["", "", "", ""])[:4]
        salida.append({
            "titulo": partes[0].strip(),
            "subrol": partes[1].strip(),
            "botones": [b.strip() for b in partes[2].split("\x1f") if b.strip()],
            "textos": [" ".join(t.split()) for t in partes[3].split("\x1f") if t.strip()],
            "crudos": [t for t in partes[3].split("\x1f") if t.strip()],
        })
    return salida


def es_dialogo(v):
    """Solo los dialogos cuentan como avisos. Las ventanas normales (libros, editor de VBA) no."""
    return v["subrol"] == "AXDialog"


def describir(v):
    return f"{v['titulo'] or '(sin titulo)'} | botones: {', '.join(v['botones'])} | {' / '.join(v['textos'])}"


def _texto(v):
    return " ".join([v["titulo"], *v["botones"], *v["textos"]])


# Marcas de los avisos de error de VBA (bitacora 7): botones Finalizacion y Depurar, textos
# "Se ha producido el error '11' en tiempo de ejecucion" y "Error de compilacion".
MARCAS_VBA = ("Visual Basic", "Finalizaci", "Depurar", "Debug", "compilaci", "Compile error",
              "tiempo de ejecuci", "Run-time error")


def es_ventana_vba(v):
    """Aviso de error de VBA (ejecucion o compilacion): nunca se pulsa nada en el durante una macro.
    Llevan "Visual Basic" en el titulo en cualquier idioma. La ventana principal del editor tambien lo lleva,
    pero no es un dialogo (AXStandardWindow): ahi solo cuentan los botones de error."""
    if es_dialogo(v):
        return any(m in _texto(v) for m in MARCAS_VBA)
    return bool({"Finalización", "Depurar", "End", "Debug"} & set(v["botones"]))


def es_compilacion(v):
    texto = _texto(v)
    return es_dialogo(v) and ("compilaci" in texto or "Compile error" in texto)


def es_aviso_excel(v):
    """MsgBox o aviso de Excel: dialogo con Aceptar u OK cuyo titulo (o primer texto, que es donde
    System Events deja el titulo) es "Microsoft Excel". Nunca un aviso de VBA."""
    if not es_dialogo(v) or es_ventana_vba(v):
        return False
    if not ({"Aceptar", "OK"} & set(v["botones"])):
        return False
    primero = v["textos"][0] if v["textos"] else ""
    return v["titulo"].startswith("Microsoft Excel") or primero.startswith("Microsoft Excel")


@con_foco
def dialogos():
    """Ventanas de Excel que no son libros ni el editor, en texto, para mostrarlas."""
    return [f"[{v['subrol']}] {v['titulo']} || {' | '.join(v['botones'])} || {' | '.join(v['textos'])}"
            for v in ventanas() if v["subrol"] != "AXStandardWindow"]


PULSAR = TRAER + """
on run argv
set boton to item 1 of argv
set titulo to item 2 of argv
set marca to item 3 of argv
my traerExcel()
tell application "System Events" to tell process "Microsoft Excel"
    repeat with w in windows
        try
            set t to name of w
            if t is missing value then set t to ""
            set sr to ""
            try
                set sr to subrole of w
                if sr is missing value then set sr to ""
            end try
            if sr is not "AXStandardWindow" and t is titulo and ((name of every button of w) contains boton) ¬
                and (marca is "" or ((value of every static text of w) contains marca)) then
                click button boton of w
                return "ok"
            end if
        end try
    end repeat
end tell
return "no"
end run
"""


def pulsar(boton, titulo="", marca=""):
    """Pulsa un boton de una ventana de Excel que no sea un libro ni el editor, con ese titulo exacto y, si se
    da, con ese texto (con dos avisos a la vez, distingue cual). Quien llama decide antes si es seguro
    (es_aviso_excel, es_compilacion sin macro, Habilitar macros). Devuelve True si lo pulso."""
    try:
        return osa(PULSAR, boton, titulo, marca, timeout=20) == "ok"
    except SystemExit:
        return False


def _matar(proc):
    """Mata solo el osascript que espera. No toca Excel ni sus ventanas."""
    proc.kill()
    try:
        proc.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        pass


def _vigilar(proc, timeout, app=""):
    """La macro no ha respondido a tiempo. Regla aprendida el 14/11: detener una macro parada (Finalizacion,
    Restablecer, o Aceptar en un error de compilacion, que la deja en pausa) cierra Excel de golpe. Solo se
    pulsa Aceptar/OK en un aviso de Excel (MsgBox), que deja seguir a la macro. Ante cualquier ventana de VBA
    no se toca nada: se mata solo osascript y se informa. Si tras cerrar un aviso sale otro, se repite.
    Al terminar devuelve el foco a app (la que estaba delante al lanzar la macro: un MsgBox trae Excel al frente)."""
    try:
        return _vigilar_avisos(proc, timeout)
    finally:
        devolver_foco(app)


def _vigilar_avisos(proc, timeout):
    cerrados = []

    def previos():
        return f" Antes se cerraron: {'; '.join(cerrados)}." if cerrados else ""

    for _ in range(MAX_AVISOS):
        vistas = ventanas()
        de_vba = [v for v in vistas if es_ventana_vba(v)]
        if de_vba:
            _matar(proc)
            return (f"ERROR VENTANA VBA (no se ha tocado): {describir(de_vba[0])}. La macro esta parada. "
                    "Cerrarla tumba Excel: guarda lo que puedas desde otra via o acepta perder lo no guardado."
                    + previos())
        aviso = next((v for v in vistas if es_aviso_excel(v)), None)
        if not aviso:
            _matar(proc)
            dial = [describir(v) for v in vistas if v["subrol"] != "AXStandardWindow"]
            return f"ERROR TIEMPO: sin respuesta en {timeout} s. Dialogos: {dial}" + previos()
        boton = "Aceptar" if "Aceptar" in aviso["botones"] else "OK"
        if not pulsar(boton, aviso["titulo"], aviso["crudos"][-1] if aviso["crudos"] else ""):
            _matar(proc)
            return f"ERROR no he podido pulsar {boton} en: {describir(aviso)}" + previos()
        cerrados.append(" / ".join(aviso["textos"]))
        # Esperar a que la macro termine, mirando cada 2 s si sale otra ventana.
        fin = time.time() + timeout
        while True:
            try:
                out, err = proc.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                if time.time() > fin or any(es_ventana_vba(v) or es_aviso_excel(v) for v in ventanas()):
                    break
                continue
            resultado = out.rstrip("\n") if proc.returncode == 0 else "AppleScript: " + err.strip()
            return (f"AVISO cerrado con {boton}: {'; '.join(cerrados)} (usa SinVentanas para evitarlo). "
                    f"Resultado: {resultado}")
    _matar(proc)
    return f"ERROR demasiados avisos seguidos ({MAX_AVISOS}); la macro sigue en Excel." + previos()


def macro(nombre, *args, libro=LANZADOR, timeout=60):
    """Ejecuta una macro. Si se queda colgada por una ventana, la vigila (ver _vigilar) y lo devuelve como texto."""
    # Referencia VBA 'libro'!macro: los apostrofos del nombre del libro se duplican. Va por argv (item 1).
    ref = "'" + libro.replace("'", "''") + "'!" + nombre
    params = " ".join(f"arg{i + 1} (item {i + 2} of argv)" for i in range(len(args)))
    script = (
        "on run argv\n"
        f"with timeout of {timeout + 600} seconds\n"
        f'tell application "Microsoft Excel" to return (run VB macro (item 1 of argv) {params})\n'
        "end timeout\nend run\n"
    )
    app = app_delante()
    proc = subprocess.Popen(["osascript", "-", ref, *map(str, args)], stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    try:
        out, err = proc.communicate(input=script, timeout=timeout)
    except subprocess.TimeoutExpired:
        return _vigilar(proc, timeout, app)
    if proc.returncode:
        raise SystemExit("AppleScript: " + err.strip())
    return out.rstrip("\n")


AD = os.environ.get("AGENT_DESKTOP", str(pathlib.Path.home() / "workspaces/agent-desktop-bin/v1.3.1/agent-desktop"))


def ad(*args):
    """Llama a agent-desktop y devuelve su JSON. Si la orden falla, el dict lleva ok=False."""
    if not (os.path.isfile(AD) and os.access(AD, os.X_OK)):
        raise SystemExit(f"ERROR no encuentro agent-desktop en {AD}. Indica su ruta con la variable AGENT_DESKTOP.")
    try:
        r = subprocess.run([AD, *args], capture_output=True, text=True, timeout=90)
    except subprocess.TimeoutExpired:
        raise SystemExit(f"ERROR TIEMPO: agent-desktop {args[0]} sin respuesta en 90 s")
    try:
        d = json.loads(r.stdout)
    except json.JSONDecodeError:
        detalle = (r.stderr or r.stdout).strip()[:300]
        raise SystemExit(f"ERROR agent-desktop {args[0]} (codigo {r.returncode}) no devolvio JSON: {detalle}")
    if not isinstance(d, dict):
        raise SystemExit(f"ERROR agent-desktop {args[0]} devolvio un JSON inesperado")
    if r.returncode:
        d["ok"] = False
    return d


def _arbol(d):
    """Arbol de un snapshot de agent-desktop, o {} si fallo."""
    if not d.get("ok"):
        return {}
    data = d.get("data")
    return (data.get("tree") or {}) if isinstance(data, dict) else {}


def _buscar(n, prueba):
    if not isinstance(n, dict):
        return None
    if prueba(n):
        return n
    for c in n.get("children") or []:
        r = _buscar(c, prueba)
        if r:
            return r
    return None


def item_menu(prefijo, espera=0):
    """Busca un elemento de menu de Excel por el principio de su nombre. Con espera, reintenta hasta esos
    segundos: tras poner Excel delante, sus menus tardan en estar disponibles."""
    limite = time.time() + espera
    while True:
        arbol = _arbol(ad("snapshot", "--app", "Microsoft Excel", "--surface", "menubar", "-i", "--compact"))
        item = _buscar(arbol, lambda n: n.get("role") == "menuitem" and str(n.get("name", "")).startswith(prefijo))
        if item or time.time() >= limite:
            return item
        time.sleep(1)


def _donde_compila():
    """Lee en el editor de VBA (ventana con el foco tras el aviso) el modulo y la posicion del cursor."""
    arbol = _arbol(ad("snapshot", "--app", "Microsoft Excel", "--surface", "focused", "--compact"))
    if not arbol:
        return " (no he podido leer el editor)"
    modulo = _buscar(arbol, lambda n: "(C\u00f3digo)" in str(n.get("name", "")) and "focused" in (n.get("states") or []))
    modulo = modulo or _buscar(arbol, lambda n: "(C\u00f3digo)" in str(n.get("name", "")))
    pos = _buscar(arbol, lambda n: str(n.get("name", "")).startswith("L\u00edn"))
    return f" en {modulo['name'] if modulo else '?'}, {pos['name'] if pos else '?'}"


MOSTRAR_EN_EDITOR = """Option Explicit
Public Function Envoltorio() As String
    Dim c As Object, elegido As Object
    On Error GoTo Fallo
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Name <> "ZZEnvoltorio" Then
            If elegido Is Nothing Then Set elegido = c
            If c.Type = 1 Then
                Set elegido = c
                Exit For
            End If
        End If
    Next c
    elegido.CodeModule.CodePane.Show
    Envoltorio = "OK " & elegido.Name
    Exit Function
Fallo:
    Envoltorio = "ERROR " & Err.Number & ": " & Err.Description
End Function
"""


def _proyecto_editor():
    """Libro del proyecto activo en el editor de VBA, leido del titulo "Microsoft Visual Basic - libro.xlsm"."""
    for v in ventanas():
        if v["titulo"].startswith("Microsoft Visual Basic - "):
            return v["titulo"][len("Microsoft Visual Basic - "):].replace(" [en ejecuci\u00f3n]", "").replace(
                " [interrumpir]", "").strip()
    return ""


@con_foco
def compilar(libro):
    """Compila el proyecto VBA del libro desde el menu del editor (agent-desktop).
    Depurar > Compilar compila el proyecto activo del editor (el de la ventana de codigo con el foco), no el
    libro activo de Excel: visto el 14/11, con --libro vl-fondo.xlsm compilo errores.xlsm. Por eso antes un
    modulo temporal (ZZEnvoltorio, como en seguro) muestra en el editor un modulo del libro (CodePane.Show),
    se borra, y se comprueba en el titulo del editor que el proyecto activo es el pedido; si no, no se compila.
    Si hay error de compilacion, cierra el aviso con Aceptar/OK (seguro: no hay macro en marcha, comprobado)
    y dice modulo, linea y columna. Cualquier otra ventana de VBA (macro parada) no se toca."""
    parada = [v for v in ventanas() if es_ventana_vba(v)]
    if parada:
        return f"ERROR VENTANA VBA (no se ha tocado): {describir(parada[0])}. Hay una macro parada; no compilo."
    r = temporal(MOSTRAR_EN_EDITOR, libro)
    if not r.startswith("OK"):
        return f"ERROR no he podido mostrar {libro} en el editor: {r}"
    osa('on run argv\ntell application "Microsoft Excel" to activate object workbook (item 1 of argv)\nend run\n', libro)
    traer_excel()
    editor = item_menu("Editor de Visual Basic", espera=5)
    if editor:
        ad("click", editor["ref_id"])
        time.sleep(1)
    activo = _proyecto_editor()
    if activo != libro:
        return (f"ERROR el editor tiene activo el proyecto de {activo or '?'}, no el de {libro}. "
                "No compilo para no dar el resultado de otro libro.")
    item = item_menu("Compilar", espera=5)
    if not item:
        # A veces (1 de 4 el 14/11) el editor no queda al frente a la primera: se pulsa otra vez su menu.
        traer_excel()
        editor = item_menu("Editor de Visual Basic", espera=5)
        if editor:
            ad("click", editor["ref_id"])
            time.sleep(1)
        item = item_menu("Compilar", espera=5)
    if not item:
        return "ERROR no encuentro Depurar > Compilar en el editor"
    if "disabled" in (item.get("states") or []):
        return "SIN VERIFICAR: el menu Compilar estaba desactivado (Excel lo da por compilado); no garantiza que compile"
    if not ad("click", item["ref_id"]).get("ok"):
        return "ERROR agent-desktop no pudo pulsar Depurar > Compilar"
    # Sondear hasta 6 s por si sale el aviso de compilacion.
    fin = time.time() + 6
    while time.time() < fin:
        time.sleep(0.5)
        for v in ventanas():
            if es_compilacion(v):
                mensaje = " / ".join(v["textos"])
                boton = next((b for b in ("Aceptar", "OK") if b in v["botones"]), None)
                if not boton or not pulsar(boton, v["titulo"], v["crudos"][-1] if v["crudos"] else ""):
                    return f"ERROR COMPILACION: {mensaje} (no he podido cerrar el aviso: {describir(v)})"
                time.sleep(1)
                return f"ERROR COMPILACION: {mensaje}{_donde_compila()}"
            if es_ventana_vba(v):
                return f"ERROR VENTANA VBA (no se ha tocado): {describir(v)}"
    return f"OK compila {libro}"


TITULOS_PREFERENCIAS = ("Preferencias", "Seguridad", "Preferences", "Security")


def _foco():
    """Arbol de la ventana con el foco en Excel. Si Excel no esta delante, agent-desktop no ve sus ventanas
    (WINDOW_NOT_FOUND, visto el 14/11): se trae Excel delante y, si aun asi no hay arbol, se busca la
    ventana de Preferencias por su titulo."""
    traer_excel()
    arbol = _arbol(ad("snapshot", "--app", "Microsoft Excel", "--surface", "focused", "--compact"))
    if arbol:
        return arbol
    lista = ad("list-windows", "--app", "Microsoft Excel")
    for v in (lista.get("data") or []) if isinstance(lista.get("data"), list) else []:
        if any(t in (v.get("title") or "") for t in TITULOS_PREFERENCIAS):
            return _arbol(ad("snapshot", "--app", "Microsoft Excel", "--window-id", v["id"], "--compact"))
    return {}


def _esperar(prueba, espera=6):
    """Busca en la ventana de Preferencias hasta `espera` segundos: el panel tarda en cargar."""
    limite = time.time() + espera
    while True:
        n = _buscar(_foco(), prueba)
        if n or time.time() >= limite:
            return n
        time.sleep(1)


def _casilla(espera=6):
    return _esperar(lambda n: "Acceso de confianza" in str(n.get("name", ""))
                    or "Trust access to the VBA project" in str(n.get("name", "")), espera)


@con_foco
def confianza(activar):
    """Marca o desmarca en Excel > Preferencias > Seguridad la casilla
    "Acceso de confianza al modelo de objetos de proyectos de VBA" (agent-desktop). Devuelve el estado final."""
    traer_excel()
    # Si delante esta el editor de VBA (por ejemplo, despues de compilar), "Preferencias..." abre las
    # Opciones del editor y no las de Excel (visto el 14/11). Se pone delante una ventana normal de Excel
    # y se comprueba que la barra de menus ya no es la del editor (que tiene "Depurar").
    for _ in range(3):
        if not item_menu("Depurar"):
            break
        # El boton "Ver Microsoft Excel" de la barra del editor es lo que pulsaria una persona;
        # `activate object window 1` de AppleScript no saca al editor de delante.
        lista = ad("list-windows", "--app", "Microsoft Excel")
        vbe = next((v["id"] for v in (lista.get("data") or []) if isinstance(v, dict)
                    and "Visual Basic" in (v.get("title") or "")), None)
        boton = None
        if vbe:
            arbol = _arbol(ad("snapshot", "--app", "Microsoft Excel", "--window-id", vbe, "-i", "--compact"))
            boton = _buscar(arbol, lambda n: n.get("role") == "button"
                            and n.get("name") in ("Ver Microsoft Excel", "View Microsoft Excel"))
        if boton:
            ad("click", boton["ref_id"])
        else:
            osa('tell application "Microsoft Excel" to activate object window 1', timeout=20)
        time.sleep(1.5)
    else:
        return "ERROR el editor de VBA sigue delante; cierralo o pasa a una ventana de Excel y repite"
    pref = item_menu("Preferencias")
    if not pref:
        return "ERROR no encuentro Excel > Preferencias"
    ad("click", pref["ref_id"])
    time.sleep(1)
    casilla = _casilla(espera=1)  # si Preferencias ya estaba abierta en Seguridad
    if not casilla:
        seg = _esperar(lambda n: n.get("role") == "button" and n.get("name") in ("Seguridad", "Security"))
        if seg:
            ad("click", seg["ref_id"])
            time.sleep(1)
        casilla = _casilla()
    if not casilla:
        return "ERROR no encuentro la casilla de confianza"
    antes = str(casilla.get("value")) == "1"
    if antes != activar:
        ad("toggle", casilla["ref_id"])
        time.sleep(1)
        casilla = _casilla()
    despues = (str(casilla.get("value")) == "1") if casilla else None
    cerrar = _buscar(_foco(), lambda n: n.get("role") == "button" and n.get("name") in ("Close", "Cerrar"))
    if cerrar:
        ad("click", cerrar["ref_id"])
    txt_antes = "activada" if antes else "desactivada"
    if despues is None:
        return f"ERROR no he podido leer la casilla tras cambiarla (antes: {txt_antes}); compruebalo a mano"
    estado = "activada" if despues else "desactivada"
    ok = "OK" if despues == activar else "ERROR"
    return f"{ok} confianza {estado} (antes: {txt_antes})"


def guardar(libro=LANZADOR):
    osa('on run argv\ntell application "Microsoft Excel" to save workbook (item 1 of argv)\nend run\n', libro)


ABRIR = """on run argv
set f to POSIX file (item 1 of argv)
tell application "Microsoft Excel" to open f
end run
"""


@con_foco
def abrir(ruta, habilitar=True, espera=30):
    """Abre un libro y pulsa Habilitar macros (o Deshabilitar) si Excel lo pregunta."""
    p = pathlib.Path(ruta).expanduser().resolve()
    if not p.is_file():
        raise SystemExit(f"ERROR no existe el archivo: {p}")
    botones = ("Habilitar macros", "Enable Macros") if habilitar else ("Deshabilitar macros", "Disable Macros")
    proc = subprocess.Popen(["osascript", "-", str(p)], stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    proc.stdin.write(ABRIR)
    proc.stdin.close()
    pulsado = False
    fin = time.time() + espera
    while time.time() < fin and proc.poll() is None:
        if not pulsado:
            for v in ventanas():
                boton = next((b for b in botones if b in v["botones"]), None)
                if boton and v["subrol"] != "AXStandardWindow" and pulsar(boton, v["titulo"]):
                    pulsado = True
                    break
        time.sleep(1)
    try:
        _, err = proc.communicate(timeout=30)
    except subprocess.TimeoutExpired:
        _matar(proc)
        raise SystemExit(f"ERROR TIEMPO abriendo {p.name}: mira `vba.py dialogos`")
    if proc.returncode:
        raise SystemExit("AppleScript: " + err.strip())
    return p.name


def escribir_columna(lineas, hoja="_codigo", libro=LANZADOR):
    script = (
        "on run argv\n"
        "set libro to item 1 of argv\nset nomHoja to item 2 of argv\n"
        "set filas to {}\n"
        "repeat with i from 3 to count of argv\nset end of filas to {item i of argv}\nend repeat\n"
        'tell application "Microsoft Excel"\n'
        "set wb to workbook libro\n"
        "if (name of every worksheet of wb) does not contain nomHoja then\n"
        "set ws to make new worksheet at end of wb\nset name of ws to nomHoja\nend if\n"
        "set ws to worksheet nomHoja of wb\n"
        'clear contents range "A1:B20000" of ws\n'
        'set value of range ("A1:A" & (count of filas)) of ws to filas\n'
        "end tell\nend run\n"
    )
    osa(script, libro, hoja, *lineas)


def cmd_cargar(a):
    ruta = pathlib.Path(a.archivo)
    modulo = a.modulo or ruta.stem
    destino = a.libro or LANZADOR
    if destino == LANZADOR and modulo.casefold() in PROTEGIDOS:
        raise SystemExit(f"ERROR no reemplazo {modulo} en el lanzador {LANZADOR}: es el codigo que hace la carga. "
                         "Reinstalalo con la cargadora o a mano.")
    texto = ruta.read_text(encoding="utf-8")
    lineas = [l for l in texto.splitlines() if not l.startswith("Attribute VB_")]
    escribir_columna([MARCA + l for l in lineas])
    r = macro("CargarEn", a.libro or "", modulo, timeout=a.timeout)
    print(r)
    if r.startswith("OK") and not a.sin_guardar:
        guardar(destino)
        print("guardado " + destino)


def cmd_ejecutar(a):
    print(macro(a.macro, *a.args, libro=a.libro or LANZADOR, timeout=a.timeout))


ENVOLTORIO = """Option Explicit
Public Function Envoltorio() As String
    On Error GoTo Fallo
    {llamada}
    Exit Function
Fallo:
    Envoltorio = "ERROR " & Err.Number & ": " & Err.Description
End Function
"""


def seguro(nombre, libro=LANZADOR, es_sub=False, timeout=60):
    """Ejecuta una macro dentro de un On Error generado al vuelo (modulo ZZEnvoltorio del libro destino).
    Un error sin controlar sube hasta el envoltorio y vuelve como texto, sin ventana de VBA."""
    llamada = f"Call {nombre}\n    Envoltorio = \"OK\"" if es_sub else f"Envoltorio = CStr({nombre}())"
    return temporal(ENVOLTORIO.format(llamada=llamada), libro, timeout)


def temporal(codigo, libro=LANZADOR, timeout=60):
    """Mete el codigo (con una funcion Envoltorio) en el modulo ZZEnvoltorio del libro, la ejecuta y borra
    el modulo. Devuelve lo que devuelva Envoltorio."""
    escribir_columna([MARCA + l for l in codigo.splitlines()])
    destino = "" if libro == LANZADOR else libro
    r = macro("CargarEn", destino, "ZZEnvoltorio", timeout=timeout)
    if not r.startswith("OK"):
        return r
    resultado = macro("Envoltorio", libro=libro, timeout=timeout)
    if resultado.startswith(("ERROR VENTANA", "ERROR TIEMPO")):
        # Excel esta bloqueado por una ventana o una macro sin terminar: no se intenta borrar nada.
        return (resultado + f"\nAVISO: ZZEnvoltorio queda en {libro}; borralo cuando Excel responda "
                f"(vba.py borrar ZZEnvoltorio --libro {libro}).")
    b = macro("Borrar", destino, "ZZEnvoltorio", timeout=timeout)
    if not b.startswith("OK"):
        resultado += f"\nAVISO: no se pudo borrar ZZEnvoltorio de {libro}: {b}"
    return resultado


def cmd_seguro(a):
    print(seguro(a.macro, a.libro or LANZADOR, a.sub, timeout=a.timeout))


def cmd_exportar(a):
    texto = macro("Exportar", a.libro or "", a.modulo, timeout=a.timeout)
    if texto.startswith(("ERROR", "AVISO")):
        # No se escribe --salida: un error no debe acabar guardado como si fuera el codigo del modulo.
        raise SystemExit(texto)
    texto = texto.replace("\r\n", "\n").replace("\r", "\n")
    if a.salida:
        pathlib.Path(a.salida).write_text(texto + "\n", encoding="utf-8")
        print(f"OK {a.modulo} -> {a.salida}")
    else:
        print(texto)


def cmd_abrir(a):
    print("OK abierto " + abrir(a.archivo, habilitar=not a.sin_macros))


def cmd_guardar(a):
    guardar(a.libro or LANZADOR)
    print("OK guardado " + (a.libro or LANZADOR))


def cmd_confianza(a):
    r = confianza(a.accion == "activar")
    print(r)
    if not r.startswith("OK"):
        sys.exit(1)


def cmd_compilar(a):
    print(compilar(a.libro or LANZADOR))


def cmd_dialogos(a):
    print("\n".join(dialogos()) or "sin dialogos")


def cmd_modulos(a):
    print(macro("Modulos", a.libro or "").replace("\r", "\n"))


def cmd_borrar(a):
    print(macro("Borrar", a.libro or "", a.modulo))


def cmd_hoja_csv(a):
    # Se escribe por "formula" y no por "value": Excel interpreta las formulas en notacion inglesa
    # (punto decimal) sea cual sea el idioma del Mac. Con "value" y el Mac en castellano,
    # AppleScript lee 100.5 como 1005 o 1,0E+12.
    filas = [l.split(";") for l in pathlib.Path(a.archivo).read_text(encoding="utf-8").splitlines() if l.strip()]
    ancho = max(len(f) for f in filas)
    plano = []
    for f in filas:
        plano += (f + [""] * (ancho - len(f)))
    script = (
        "on run argv\n"
        "set libro to item 1 of argv\nset nomHoja to item 2 of argv\nset ancho to (item 3 of argv) as integer\n"
        "set filas to {}\nset fila to {}\n"
        "repeat with i from 4 to count of argv\n"
        "set v to item i of argv\n"
        "set end of fila to v\n"
        "if (count of fila) = ancho then\nset end of filas to fila\nset fila to {}\nend if\n"
        "end repeat\n"
        'tell application "Microsoft Excel"\n'
        "set wb to workbook libro\n"
        "if (name of every worksheet of wb) does not contain nomHoja then\n"
        "set ws to make new worksheet at end of wb\nset name of ws to nomHoja\nend if\n"
        "set ws to worksheet nomHoja of wb\n"
        "clear contents used range of ws\n"
        "set formula of (get resize (range \"A1\" of ws) row size (count of filas) column size ancho) to filas\n"
        "end tell\nend run\n"
    )
    osa(script, a.libro or LANZADOR, a.hoja, ancho, *plano)
    print(f"OK {len(filas)} filas en {a.hoja}")


def cmd_leer(a):
    script = (
        "on run argv\n"
        'tell application "Microsoft Excel" to set v to value of range (item 3 of argv) of worksheet (item 2 of argv) of workbook (item 1 of argv)\n'
        "set salida to \"\"\n"
        "repeat with fila in v\n"
        "set AppleScript's text item delimiters to \";\"\n"
        "set salida to salida & (fila as text) & linefeed\n"
        "end repeat\nreturn salida\nend run\n"
    )
    print(osa(script, a.libro or LANZADOR, a.hoja, a.rango))


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    s = p.add_subparsers(dest="cmd", required=True)
    T = dict(type=int, default=60, metavar="SEGUNDOS", help="espera antes de mirar las ventanas (60)")
    c = s.add_parser("cargar"); c.add_argument("archivo"); c.add_argument("--libro"); c.add_argument("--modulo"); c.add_argument("--sin-guardar", action="store_true"); c.add_argument("--timeout", **T); c.set_defaults(f=cmd_cargar)
    c = s.add_parser("abrir"); c.add_argument("archivo"); c.add_argument("--sin-macros", action="store_true"); c.set_defaults(f=cmd_abrir)
    c = s.add_parser("guardar"); c.add_argument("--libro"); c.set_defaults(f=cmd_guardar)
    c = s.add_parser("dialogos"); c.set_defaults(f=cmd_dialogos)
    c = s.add_parser("compilar"); c.add_argument("--libro"); c.set_defaults(f=cmd_compilar)
    c = s.add_parser("confianza"); c.add_argument("accion", choices=["activar", "desactivar"]); c.set_defaults(f=cmd_confianza)
    c = s.add_parser("ejecutar"); c.add_argument("macro"); c.add_argument("args", nargs="*"); c.add_argument("--libro"); c.add_argument("--timeout", **T); c.set_defaults(f=cmd_ejecutar)
    c = s.add_parser("seguro"); c.add_argument("macro"); c.add_argument("--libro"); c.add_argument("--sub", action="store_true"); c.add_argument("--timeout", **T); c.set_defaults(f=cmd_seguro)
    c = s.add_parser("exportar"); c.add_argument("modulo"); c.add_argument("--libro"); c.add_argument("--salida"); c.add_argument("--timeout", **T); c.set_defaults(f=cmd_exportar)
    c = s.add_parser("modulos"); c.add_argument("--libro"); c.set_defaults(f=cmd_modulos)
    c = s.add_parser("borrar"); c.add_argument("modulo"); c.add_argument("--libro"); c.set_defaults(f=cmd_borrar)
    c = s.add_parser("hoja-csv"); c.add_argument("archivo"); c.add_argument("hoja"); c.add_argument("--libro"); c.set_defaults(f=cmd_hoja_csv)
    c = s.add_parser("leer"); c.add_argument("hoja"); c.add_argument("rango"); c.add_argument("--libro"); c.set_defaults(f=cmd_leer)
    a = p.parse_args()
    a.f(a)


if __name__ == "__main__":
    main()
