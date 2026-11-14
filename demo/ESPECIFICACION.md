# Demo: valor liquidativo de un fondo ficticio

Caso pequeño para enseñar a la profesora de VBA que Claude Code puede encontrar un error en una herramienta VBA, corregirlo dentro del libro y validar la corrección contra una referencia independiente en Python ([docs/00-plan-inicial.md](../docs/00-plan-inicial.md) §5).

Fondo, cifras y fechas inventados. No corresponden a ninguna gestora ni cliente.

La solución está en [solucion/](solucion/SOLUCION.md).

---

## 1. Archivos

| Archivo | Qué es |
|---|---|
| `datos.csv` | 15 días de valoración. Se copia a la hoja `Datos` desde A1 |
| `vba/Datos.bas` | Lee la hoja `Datos` y utilidades comunes (fechas, números, hoja de salida) |
| `vba/Calculo.bas` | Cálculo del VL |
| `vba/Informe.bas` | Escribe la hoja `Informe` con el detalle de cada día |
| `vba/Tests.bas` | Compara el VL de VBA con la hoja `Esperado` y escribe la hoja `Tests` |
| `vl-fondo.xlsm` | El libro de la demo, el antes: hojas `Datos` y `Esperado` ya copiadas, los cuatro módulos cargados y hojas `Tests` e `Informe` coherentes con ese código (`FALLO 10 de 15`) |
| `vl-fondo-corregido.xlsm` | El después de la demo, con `Calculo` corregido, tests 15/15 e informe. Se abre en la hoja `Informe` |
| `ejecutar_demo.py` | Demo guiada de principio a fin: tests, exportar, corregir, compilar, tests e informe. Al terminar vuelve a dejar `vl-fondo.xlsm` con la versión con el error |
| `solucion/` | La solución: `SOLUCION.md` (explicación) y `Calculo_corregido.bas` (el mismo módulo `Calculo` con el arreglo, para comparar al final) |
| `referencia.py` | Referencia independiente en Python. Lee `datos.csv` y escribe `esperado.csv` |
| `esperado.csv` | VL correcto por día (`fecha;vl`). Se copia a la hoja `Esperado` desde A1 |

Orden de carga de módulos: `Datos`, `Calculo`, `Informe`, `Tests`. `Calculo` e `Informe` usan funciones de `Datos`; `Tests` usa `Datos` y `Calculo`.

Funciones que se llaman desde la terminal (todas devuelven texto y nunca abren ventanas; si algo falla devuelven `ERROR n: descripción`):

| Función | Devuelve |
|---|---|
| `ContarDias()` | `OK 15 dias, de 2025-03-03 a 2025-03-21` |
| `VLTodos()` | Una línea por día: `aaaa-mm-dd VL` |
| `VLDeFecha("2025-03-10")` | VL de ese día |
| `GenerarInforme()` | `OK Informe 15 dias` y la hoja `Informe` |
| `EjecutarTests()` | `OK n/n` o `FALLO k de n: <fechas>` y la hoja `Tests` |

---

## 2. Regla de negocio del valor liquidativo

Fuente de verdad común para el VBA y para Python.

Para cada día de valoración:

1. **Patrimonio**: valor de la cartera a precios de cierre, en euros, antes de descontar la comisión de gestión de ese día. Viene en los datos.
2. **Días de devengo**: días naturales transcurridos desde la valoración anterior. Un lunes después de un viernes cuenta 3. El primer día de la serie cuenta 1.
3. **Comisión de gestión devengada** = patrimonio × (comisión de gestión anual en % / 100) × días de devengo / **365**. Base 365 días naturales, siempre.
4. **Patrimonio neto** = patrimonio − comisión de gestión devengada.
5. **Valor liquidativo (VL)** = patrimonio neto / participaciones en circulación.

Sin redondeos intermedios. El VL se compara con una tolerancia de 0,000001 (1e-6) euros por participación.

Errores de datos que deben avisarse en vez de calcular: fechas desordenadas o repetidas, participaciones cero o negativas, celdas vacías.

---

## 3. Datos

`datos.csv`: separador `;`, punto decimal, cabecera en la fila 1.

| Columna | Hoja `Datos` | Contenido |
|---|---|---|
| `fecha` | A | Día de valoración, `aaaa-mm-dd` |
| `patrimonio` | B | Euros, antes de la comisión del día |
| `participaciones` | C | Participaciones en circulación |
| `comision_gestion_pct` | D | Comisión de gestión anual, en % |

Escenario: fondo lanzado el 03/03/2025 con VL 100. Durante la semana de lanzamiento (03/03 a 07/03) la gestora bonifica la comisión de gestión (0,00 %). Desde el 10/03 se cobra el 1,50 % anual. Hay suscripciones algunos días, por eso cambian las participaciones. Los lunes devengan 3 días.

Así hay días que no dependen de la comisión (los 5 primeros) y días que sí (los 10 siguientes, dos de ellos lunes).

---

## 4. La incidencia

Tal como la escribiría la persona que mantiene la herramienta:

> **Incidencia:** el valor liquidativo sale mal cuando hay comisiones.
> Desde que empezamos a cobrar la comisión de gestión, el VL que da la herramienta no cuadra con el que nos manda el depositario. La diferencia es pequeña, en el quinto decimal, pero está todos los días. La semana de lanzamiento cuadraba bien.

---

## 5. Guion de la demo ([docs/00-plan-inicial.md](../docs/00-plan-inicial.md) §5)

Preparación: ya está hecha en `vl-fondo.xlsm`, con las hojas `Datos` (copiado `datos.csv`) y `Esperado` (copiado `esperado.csv`) y los módulos `Datos`, `Calculo`, `Informe` y `Tests` cargados. Si se cambian los datos, ejecutar `python3 -I demo/referencia.py` y volver a copiar los CSV con `vba.py hoja-csv`.

Los cinco pasos:

1. **El libro.** Enseñar el libro inventado con cuatro módulos (`Datos`, `Calculo`, `Informe` y `Tests`) y el error escondido. Ejecutar `GenerarInforme()` desde la terminal: la hoja `Informe` aparece sin tocar Excel.
2. **La incidencia.** Dar a Claude el texto del apartado 4 y la regla del apartado 2. Nada más.
3. **Encontrar y corregir.** Claude saca los módulos del libro, lee el código, localiza el error, lo corrige en `Calculo`, lo vuelve a meter en el libro y ejecuta `EjecutarTests()`.
4. **Validar contra Python.** Antes de corregir, `EjecutarTests()` debe dar `FALLO 10 de 15` con las fechas del 10/03 al 21/03. Después del arreglo debe dar `OK 15/15`. Python es la referencia: calcula desde la regla, no copia el VBA.
5. **No romper lo que funcionaba.** Los 5 días de la semana de lanzamiento salen OK antes y después del arreglo. Enseñar la hoja `Tests` (esperado, obtenido, diferencia) y comparar `vba/Calculo.bas` con `solucion/Calculo_corregido.bas`.

Idea de cierre: responde al problema que planteó la profesora, que no tenía forma de comprobar si el código que le daba la IA era correcto. La validación ya no depende de leer el código: la hace una referencia independiente y el resultado se ve en la hoja `Tests`.

---

## 6. Solución

Está aparte, en [solucion/SOLUCION.md](solucion/SOLUCION.md), junto con `solucion/Calculo_corregido.bas`.
El resultado final, ya hecho, está en `vl-fondo-corregido.xlsm`. Para probar que Claude encuentra el error
sin saber dónde está, ver [docs/01-contexto.md](../docs/01-contexto.md#lo-que-todavía-no-está-demostrado):
este repositorio lo revela en varios sitios.
