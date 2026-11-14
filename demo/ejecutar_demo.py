#!/usr/bin/env python3
"""Demo guiada: Claude Code corrige y valida VBA dentro de Excel (Mac).

Recorre el ciclo completo sobre la herramienta inventada de valor liquidativo (demo/vl-fondo.xlsm):
  1. comprueba que todo esta listo
  2. carga la version con el error y pasa los tests      -> FALLO 10 de 15
  3. saca el modulo Calculo a texto y muestra la linea culpable
  4. carga la version corregida, compila y repite tests  -> OK 15/15
  5. genera el informe
  6. deja el libro como estaba (con el error), salvo --dejar-corregido

Uso, desde la raiz del repositorio y con Excel abierto:
  python3 -I demo/ejecutar_demo.py [--dejar-corregido] [--sin-compilar]

--sin-compilar  no usa agent-desktop (ni trae Excel al frente); la compilacion previa se omite.
Solo usa la biblioteca estandar de Python. Llama a herramientas/vba.py para cada paso.
"""
import argparse
import pathlib
import subprocess
import sys
import tempfile

RAIZ = pathlib.Path(__file__).resolve().parent.parent
VBA = [sys.executable, "-I", str(RAIZ / "herramientas" / "vba.py")]
LIBRO = "vl-fondo.xlsm"


def paso(titulo):
    print(f"\n=== {titulo}")


def vba(*args, permitir_error=False):
    r = subprocess.run([*VBA, *args], cwd=RAIZ, capture_output=True, text=True)
    salida = (r.stdout + r.stderr).strip()
    print("   $ vba " + " ".join(args))
    for linea in salida.splitlines():
        print("     " + linea)
    if r.returncode and not permitir_error:
        sys.exit(f"\nParado: la orden anterior fallo. Mira docs/05-buenas-practicas.md (tabla de sintomas).")
    return salida


def compilar(a):
    """Compila salvo --sin-compilar. Si no compila, para: ejecutar codigo que no compila deja una ventana de VBA
    que, al cerrarla, tumba Excel. Otros fallos de compilar (menu no encontrado...) solo avisan."""
    if a.sin_compilar:
        return
    r = vba("compilar", "--libro", LIBRO, permitir_error=True)
    if "ERROR COMPILACION" in r or "ERROR VENTANA" in r:
        sys.exit("\nParado: el codigo no compila o hay una ventana de VBA abierta. No se ejecutan los tests.")


def excel_abierto():
    r = subprocess.run(["osascript", "-e", 'application "Microsoft Excel" is running'],
                       capture_output=True, text=True)
    return r.stdout.strip() == "true"


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dejar-corregido", action="store_true", help="no volver a poner la version con el error")
    p.add_argument("--sin-compilar", action="store_true", help="omitir la compilacion (no usa agent-desktop)")
    a = p.parse_args()

    paso("0. Comprobaciones")
    if not excel_abierto():
        sys.exit("Excel no esta abierto. Abrelo y vuelve a ejecutar la demo.")
    vba("abrir", "prueba.xlsm")
    vba("abrir", "demo/vl-fondo.xlsm")
    modulos = vba("modulos", "--libro", LIBRO, permitir_error=True)
    if "1004" in modulos:
        sys.exit("\nFalta el ajuste de confianza de VBA. Ejecuta: python3 -I herramientas/vba.py confianza activar\n"
                 "(o Excel > Preferencias > Seguridad) y vuelve a lanzar la demo.")
    if "Calculo" not in modulos:
        sys.exit("\nEl libro de la demo no tiene el modulo Calculo. Revisa demo/vl-fondo.xlsm.")

    paso("1. La herramienta tal como llega, con la incidencia: 'el valor liquidativo sale mal cuando hay comisiones'")
    vba("cargar", "demo/vba/Calculo.bas", "--libro", LIBRO)
    compilar(a)
    antes = vba("seguro", "EjecutarTests", "--libro", LIBRO)
    print("   -> Los tests comparan el VBA con el resultado de demo/referencia.py (Python, calculado aparte).")

    paso("2. Sacar el modulo Calculo del libro a texto y buscar la causa")
    with tempfile.TemporaryDirectory() as tmp:
        salida = pathlib.Path(tmp) / "Calculo.bas"
        vba("exportar", "Calculo", "--libro", LIBRO, "--salida", str(salida))
        for n, linea in enumerate(salida.read_text(encoding="utf-8").splitlines(), 1):
            if "TasaDiaria =" in linea:
                print(f"     linea {n}: {linea.strip()}")
    print("   -> La comision anual se reparte entre 360 dias; la regla de negocio dice 365.")

    paso("3. Meter la correccion en el libro, compilar y repetir los tests")
    vba("cargar", "demo/solucion/Calculo_corregido.bas", "--libro", LIBRO, "--modulo", "Calculo")
    compilar(a)
    despues = vba("seguro", "EjecutarTests", "--libro", LIBRO)

    paso("4. Informe")
    vba("seguro", "GenerarInforme", "--libro", LIBRO)

    if not a.dejar_corregido:
        paso("5. Dejar el libro como estaba (version con el error) para repetir la demo")
        vba("cargar", "demo/vba/Calculo.bas", "--libro", LIBRO)
        # Repetir tests e informe con la version con el error: asi las hojas Tests e Informe del libro
        # cuentan lo mismo que su codigo (FALLO 10 de 15) y nadie ve "todo OK" sobre codigo con el error.
        vba("seguro", "EjecutarTests", "--libro", LIBRO)
        vba("seguro", "GenerarInforme", "--libro", LIBRO)
        vba("guardar", "--libro", LIBRO)

    paso("Resultado")
    print(f"   antes:   {antes.splitlines()[0] if antes else '?'}")
    print(f"   despues: {despues.splitlines()[0] if despues else '?'}")
    ok = antes.startswith("FALLO") and despues.startswith("OK")
    print("   DEMO CORRECTA" if ok else "   RESULTADO INESPERADO: revisa la salida de cada paso")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
