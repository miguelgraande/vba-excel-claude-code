# Probar en Windows: lo que falta y cómo hacerlo

Para quien abra este repositorio en un PC con Windows y Excel, sea una persona o una IA. Todo lo que
hace falta ya está en el repositorio. Solo faltan las pruebas: nadie del proyecto ha tenido un Windows
a mano. Lee antes las reglas de oro de [AGENTS.md](../AGENTS.md). La demo guiada de Mac
(`demo/ejecutar_demo.py`) **no funciona en Windows**: usa AppleScript.

## Estado a 21/11/2026

| Pieza | Qué es | Estado |
|---|---|---|
| [herramientas/windows/vba.ps1](../herramientas/windows/vba.ps1) | Saca y mete código, lista módulos y ejecuta macros por COM, con una instancia propia e invisible de Excel | Una profesora lo usó en Windows y le funcionó (nos lo contaron; no tenemos la orden ni la salida). Nosotros no lo hemos ejecutado |
| [herramientas/windows/vba-nuevo.ps1](../herramientas/windows/vba-nuevo.ps1) | `vba.ps1` más `compilar` y un vigía que lee las ventanas | Nunca ejecutado. Pasa el analizador de PowerShell 7.6.6 sin errores (en un Mac) |
| Demo `demo/vl-fondo.xlsm` y referencia en Python | Herramienta inventada con un error en las comisiones | Probada en Mac. El VBA no depende del Mac |
| Ajuste de confianza, directivas de empresa, ventanas | Riesgos de Windows | Explicados en [06-windows.md](06-windows.md), sin comprobar |

Si `vba-nuevo.ps1` pasa todas las pruebas de abajo, puede sustituir a `vba.ps1`. Si falla alguna de
W9 a W11, se queda `vba.ps1`.

## Antes de empezar

1. Windows 10 u 11 con Excel de escritorio. Comprueba los requisitos de [06-windows.md §2](06-windows.md#2-requisitos).
2. Activa la casilla "Confiar en el acceso al modelo de objetos de proyectos de VBA"
   ([06-windows.md §3](06-windows.md#3-activar-el-acceso-de-confianza-al-modelo-de-objetos-de-vba)) y
   desactívala al terminar.
3. Abre PowerShell en la carpeta del repositorio y prepara copias, para no tocar los originales:

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   $vba   = ".\herramientas\windows\vba.ps1"
   $nuevo = ".\herramientas\windows\vba-nuevo.ps1"
   New-Item -ItemType Directory -Force pruebas-win | Out-Null
   Copy-Item demo\vl-fondo.xlsm pruebas-win\copia.xlsm
   Copy-Item demo\vl-fondo.xlsm pruebas-win\copia-errores.xlsm
   # Calculo_corregido.bas no lleva el nombre del modulo dentro: con este nombre de archivo reemplaza a Calculo
   Copy-Item demo\solucion\Calculo_corregido.bas pruebas-win\Calculo.bas
   ```

   Si los archivos llegaron por correo, Teams o descarga: `Unblock-File` sobre los `.ps1` y los `.xlsm`.
   `pruebas-win` no se sube a GitHub: los libros guardan el nombre de quien los guarda (regla de oro 11).

## Las pruebas

Primero W0 a W8 con `$vba`, luego **las mismas** con `$nuevo`, y por último W9 a W11 con `$nuevo`.
Rehaz las copias del paso 3 antes de cada vuelta. Después de cada orden, en el Administrador de tareas no
debe quedar ningún `EXCEL.EXE` que no fuera tuyo.

| ID | Orden | Resultado esperado |
|---|---|---|
| W0 | `$ExecutionContext.SessionState.LanguageMode`; `Get-ExecutionPolicy -List`; `reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Security" /v AccessVBOM` | `FullLanguage`; ninguna directiva (`MachinePolicy`, `UserPolicy`) en `Restricted`; `AccessVBOM` a `0x1` |
| W1 | `& $vba modulos pruebas-win\copia.xlsm` | Líneas `Datos;Modulo;152`, `Calculo;Modulo;52`, `Informe;Modulo;42`, `Tests;Modulo;69`, más hojas y `ThisWorkbook` con 0 |
| W2 | `& $vba ejecutar pruebas-win\copia.xlsm EjecutarTests -Leer "Tests!A1:E3"` | `FALLO 10 de 15: 2025-03-10, ...` y debajo las filas de la hoja `Tests` separadas por `;` |
| W3 | `& $vba ejecutar pruebas-win\copia.xlsm VLDeFecha 2025-03-03` | `2025-03-03 100.000000` (pasa un argumento y vuelve el texto) |
| W4 | `& $vba exportar pruebas-win\copia.xlsm pruebas-win\src` | Rutas de `Calculo.bas`, `Datos.bas`, `Informe.bas` y `Tests.bas`. En `Calculo.bas` está la línea `TasaDiaria = ComisionPctDia(fila) / 100 / 360` (el error: la regla dice 365) |
| W5 | `& $vba importar pruebas-win\copia.xlsm pruebas-win\Calculo.bas` | `OK Calculo N lineas` (en Mac, 53) y ningún módulo `Calculo1` ni `Calculo_corregido` en `modulos` |
| W6 | Ver la nota de W6 debajo | `ConControl` devuelve `ERROR 11: ...`. `ConMsgBox` y `SinControl`, con `-Tiempo 15`: a los 15 s, `Tiempo agotado (15 s)...`, sin Excel colgado |
| W7 | `& $vba ejecutar pruebas-win\copia.xlsm EjecutarTests` y luego `VLDeFecha 2025-03-10` | `OK 15/15` y `2025-03-10 98.853665` (el valor de `demo/esperado.csv`, calculado en Python aparte) |
| W8 | Con un Excel tuyo abierto y un libro tuyo sin guardar, repite W1 a W3 | Tu Excel no se cierra ni cambia |
| W9 | `& $nuevo compilar pruebas-win\copia.xlsm` | `OK compila copia.xlsm` |
| W10 | `& $nuevo importar pruebas-win\copia-errores.xlsm vba\pruebas\NoCompila.bas` y dos veces `& $nuevo compilar pruebas-win\copia-errores.xlsm` | Las dos veces, `ERROR COMPILACION: ... en copia-errores.xlsm - NoCompila, Lin 9, ...`. Nunca `OK` ni `SIN VERIFICAR` |
| W11 | Como W6, con `$nuevo` | `ConMsgBox`: `despues del MsgBox` y, en la salida de errores, `AVISO cerrado con Aceptar: ... ventana de prueba ...`. `SinControl`: `ERROR VENTANA VBA (no se ha pulsado nada): ...` en pocos segundos, sin esperar a los 15 |

**Nota de W6 y W11.** Mete primero el módulo de pruebas en la copia y luego lanza las tres funciones:

```powershell
& $vba importar pruebas-win\copia.xlsm vba\pruebas\PruebaErrores.bas
& $vba ejecutar pruebas-win\copia.xlsm ConControl
& $vba ejecutar pruebas-win\copia.xlsm ConMsgBox  -Tiempo 15
& $vba ejecutar pruebas-win\copia.xlsm SinControl -Tiempo 15
```

**Sin el ajuste de confianza**, `modulos` debe fallar con un diagnóstico del registro y sin ventanas.

## Si algo no sale como se espera

- No pulses nada en una ventana de Excel con "Finalizar" o "Depurar" (regla de oro 1). Para ver qué
  ventana es, repite la orden con `-Visible`.
- Copia la salida completa, también la de errores. No la resumas.
- Las suposiciones que más pueden fallar están en [06-windows.md §13](06-windows.md#13-técnicas-tomadas-de-mcp-access)
  (vigía y `compilar`) y en [06-windows.md §4](06-windows.md#4-cómo-abre-excel-el-script) (instancia propia).

## Cómo dejar el resultado

Apunta cada prueba en [04-resultados.md](04-resultados.md), en la tabla de Windows: fecha, versión de
Windows y de Excel (Archivo > Cuenta), script (`vba.ps1` o `vba-nuevo.ps1`), orden exacta, salida literal
y si coincide con lo esperado. Si no puedes subir cambios al repositorio, abre una incidencia (issue) en
GitHub con esos mismos datos o envíaselos al autor.

Quedan por probar después, con el ciclo ya funcionando en Windows: una sesión limpia de IA que encuentre
el error sola ([01-contexto.md](01-contexto.md#lo-que-todavía-no-está-demostrado)) y la herramienta real,
solo con permiso de la organización propietaria.
