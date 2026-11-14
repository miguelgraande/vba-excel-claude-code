"""Referencia independiente del valor liquidativo (VL) del fondo de demo.

Implementa la regla de negocio de ESPECIFICACION.md, no el codigo VBA.
Lee datos.csv y escribe esperado.csv (fecha;vl) en esta misma carpeta.

Uso: python3 -I demo/referencia.py
Solo biblioteca estandar.
"""

import csv
from datetime import date
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
ENTRADA = CARPETA / "datos.csv"
SALIDA = CARPETA / "esperado.csv"

# Regla de negocio (ESPECIFICACION.md, apartado 2).
DIAS_ANIO = 365  # la comision de gestion se devenga por dias naturales, base 365


def leer_datos(ruta):
    filas = []
    with open(ruta, newline="", encoding="utf-8") as f:
        for registro in csv.DictReader(f, delimiter=";"):
            filas.append(
                {
                    "fecha": date.fromisoformat(registro["fecha"].strip()),
                    "patrimonio": float(registro["patrimonio"]),
                    "participaciones": float(registro["participaciones"]),
                    "comision_pct": float(registro["comision_gestion_pct"]),
                }
            )
    if not filas:
        raise ValueError("datos.csv no tiene filas")
    return filas


def calcular_vl(filas):
    """Devuelve [(fecha, vl)] aplicando la regla dia a dia."""
    resultado = []
    fecha_anterior = None
    for dia in filas:
        if fecha_anterior is None:
            dias = 1  # primer dia de valoracion
        else:
            dias = (dia["fecha"] - fecha_anterior).days
            if dias <= 0:
                raise ValueError(f"fechas no ordenadas en {dia['fecha']}")
        if dia["participaciones"] <= 0:
            raise ValueError(f"participaciones no validas en {dia['fecha']}")

        comision_anual = dia["patrimonio"] * dia["comision_pct"] / 100
        comision_periodo = comision_anual * dias / DIAS_ANIO
        vl = (dia["patrimonio"] - comision_periodo) / dia["participaciones"]

        resultado.append((dia["fecha"], vl))
        fecha_anterior = dia["fecha"]
    return resultado


def escribir_esperado(ruta, resultado):
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        escritor = csv.writer(f, delimiter=";", lineterminator="\n")
        escritor.writerow(["fecha", "vl"])
        for fecha, vl in resultado:
            escritor.writerow([fecha.isoformat(), f"{vl:.10f}"])


def main():
    resultado = calcular_vl(leer_datos(ENTRADA))
    escribir_esperado(SALIDA, resultado)
    print(f"OK {len(resultado)} dias escritos en {SALIDA.name}")
    for fecha, vl in resultado:
        print(f"{fecha.isoformat()}  {vl:.6f}")


if __name__ == "__main__":
    main()
