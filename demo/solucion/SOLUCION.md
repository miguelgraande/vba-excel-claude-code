# Solución de la demo

Especificación de la demo: [../ESPECIFICACION.md](../ESPECIFICACION.md). Para probar que Claude encuentra el
error sin saber dónde está: [docs/01-contexto.md](../../docs/01-contexto.md#lo-que-todavía-no-está-demostrado).

## El error

En `vba/Calculo.bas`, función `TasaDiaria`:

```vb
TasaDiaria = ComisionPctDia(fila) / 100 / 360
```

Divide la comisión anual entre **360** en vez de **365**. Es la base de los depósitos y del mercado monetario, aplicada por error a la comisión de un fondo, que según la regla se devenga por días naturales sobre 365. El arreglo en `Calculo_corregido.bas` cambia solo ese número (y añade un comentario).

Por qué es realista y sutil: el código "parece" correcto, la diferencia es de unos 0,00006 euros por participación (quinto decimal), y solo aparece cuando la comisión no es cero. Los lunes la diferencia se triplica porque devengan 3 días.

## Qué debe detectar el test

Comprobado el 14/11/2026 en Python (simulación de la lógica de `Calculo.bas` contra `esperado.csv`, tolerancia 1e-6) y **ejecutado en Excel el 14/11/2026**: con el error, `FALLO 10 de 15: 2025-03-10 ... 2025-03-21`; tras corregir, `OK 15/15`; el informe, `OK Informe 15 dias`. Se repitió de principio a fin abriendo el libro desde disco (abrir, cargar, compilar, seguro). Evidencia en [docs/bitacora.md](../../docs/bitacora.md), §8 y §13 (pruebas repetidas tras corregir `compilar`).

| Fecha | Comisión % | Días | VL correcto (365) | VL con error (360) | Diferencia | Test |
|---|---|---|---|---|---|---|
| 2025-03-03 | 0.00 | 1 | 100.000000 | 100.000000 | +0.000000 | OK |
| 2025-03-04 | 0.00 | 1 | 99.891449 | 99.891449 | +0.000000 | OK |
| 2025-03-05 | 0.00 | 1 | 99.846839 | 99.846839 | +0.000000 | OK |
| 2025-03-06 | 0.00 | 1 | 99.516672 | 99.516672 | +0.000000 | OK |
| 2025-03-07 | 0.00 | 1 | 99.206739 | 99.206739 | +0.000000 | OK |
| 2025-03-10 | 1.50 | 3 | 98.853665 | 98.853496 | -0.000169 | FALLO |
| 2025-03-11 | 1.50 | 1 | 98.657373 | 98.657317 | -0.000056 | FALLO |
| 2025-03-12 | 1.50 | 1 | 98.647785 | 98.647728 | -0.000056 | FALLO |
| 2025-03-13 | 1.50 | 1 | 98.469015 | 98.468959 | -0.000056 | FALLO |
| 2025-03-14 | 1.50 | 1 | 98.451407 | 98.451351 | -0.000056 | FALLO |
| 2025-03-17 | 1.50 | 3 | 98.550528 | 98.550359 | -0.000169 | FALLO |
| 2025-03-18 | 1.50 | 1 | 98.721840 | 98.721783 | -0.000056 | FALLO |
| 2025-03-19 | 1.50 | 1 | 99.163803 | 99.163747 | -0.000057 | FALLO |
| 2025-03-20 | 1.50 | 1 | 99.289729 | 99.289672 | -0.000057 | FALLO |
| 2025-03-21 | 1.50 | 1 | 99.764959 | 99.764902 | -0.000057 | FALLO |

Resultado esperado en Excel (coincide con lo obtenido el 14/11):

- Con `Calculo.bas`: `FALLO 10 de 15: 2025-03-10, 2025-03-11, 2025-03-12, 2025-03-13, 2025-03-14, 2025-03-17, 2025-03-18, 2025-03-19, 2025-03-20, 2025-03-21`
- Con `Calculo_corregido.bas`: `OK 15/15` (diferencia máxima en la simulación: 4,9e-11, por los 10 decimales de `esperado.csv`).

La diferencia más pequeña con el error (0,000056) es 56 veces la tolerancia, así que el test no puede pasarla por alto.

## Si algo sale distinto en Excel

- `ERROR 9` al ejecutar: falta una hoja (`Datos` o `Esperado`) o tiene otro nombre.
- `FALLO 15 de 15` con diferencias enormes: las fechas o los números se pegaron como texto con coma decimal. El VBA acepta texto con punto decimal y fechas `aaaa-mm-dd` o `dd/mm/aaaa`, pero no números con coma escritos como texto.
- `obtenido = sin datos`: la fecha de `Esperado` no está en `Datos`.
