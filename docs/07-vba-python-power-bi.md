# VBA, Python o Power BI: qué usaríamos y por qué

Este repositorio demuestra que Claude Code puede mantener y validar una herramienta VBA dentro de Excel.
Eso no significa que VBA sea la mejor tecnología para algo nuevo. Aquí está nuestra opinión, separada de lo
probado: **nada de este documento está construido ni medido**; es una recomendación.

## Resumen

- **Para algo nuevo, no usaríamos VBA.** Usaríamos Python, con Excel solo como entrada y salida de datos,
  y un panel encima (aplicación a medida o Power BI).
- **Para una herramienta VBA que ya existe y funciona**, lo de este repositorio tiene sentido: permite
  corregirla con pruebas sin reescribirla de golpe.
- **El camino recomendado es migrar por partes**: mantener el VBA protegido por tests mientras se pasa la
  lógica a Python. La referencia Python de los tests ya es el primer trozo de esa migración. Si esa
  referencia acaba siendo el código de producción, hay que conservar aparte casos esperados calculados
  desde la especificación, para no quedarse sin una comprobación independiente.

## Qué es cada cosa

| Tecnología | Qué es | Para qué es buena |
|---|---|---|
| VBA | Lenguaje de macros de Office (de los años 90), vive dentro del archivo de Excel | Automatizar acciones dentro de Excel: botones, formatos, mover hojas |
| Python | Lenguaje de programación general, el más usado en análisis de datos | Cálculo, datos, pruebas automáticas, programas y aplicaciones |
| Power BI | Herramienta de Microsoft para cuadros de mando e informes | Visualizar datos, refrescarlos de forma programada y compartirlos en la empresa |
| Aplicación a medida | Programa hecho para un proceso concreto (por ejemplo, en Python) | Hacer exactamente el flujo que necesita el usuario, con su panel |

## Por qué Python es mejor para algo nuevo

| Aspecto | VBA | Python |
|---|---|---|
| Pruebas automáticas | No hay herramienta estándar; hay que montarlas (como en este repositorio). En Windows existe Rubberduck, un complemento libre para el editor de VBA con tests unitarios; en Mac no | Estándar (`pytest`). No corren solas en cada cambio: hace falta CI (por ejemplo, GitHub Actions) o un hook de git que las lance |
| Control de versiones | El código vive dentro de un binario `.xlsm`; hay que exportarlo para usar git | Archivos de texto: git funciona directamente |
| Revisión por una IA | Hay que sacar el código del libro y volver a meterlo (todo este repositorio) | Claude lee, ejecuta y prueba el código directamente |
| Ejecución | Necesita Excel. Lo de "Excel en primer plano y una sola instancia" es de la vía Mac de este repositorio (AppleScript sobre el Excel abierto), no de VBA: en Windows, por COM, se pueden abrir varias instancias de Excel | Sin Excel, en cualquier máquina o servidor, en paralelo |
| Datos grandes y matrices | Lento sobre todo al leer y escribir celda a celda. Pasando el rango a una matriz en memoria (`Range.Value` a un `Variant`) y escribiéndolo de una vez, va razonable | `pandas` y `numpy` están hechos para eso |
| Ventanas que bloquean | Un `MsgBox` o un error sin controlar abren una ventana y esperan a una persona. Que Excel se cierre al detener esa macro es lo observado aquí, en Mac, con macros lanzadas desde la terminal | Una excepción sin controlar también detiene el programa, pero sin ventana: deja el error escrito y se puede capturar y registrar |
| Futuro | No recibe funciones nuevas desde VBA7 (Office 2010), aunque se mantiene. Excel en la web no lo ejecuta | Activo, con un ecosistema enorme |

Microsoft ofrece dos vías nuevas que conviene conocer, aunque no las hemos probado:

- **Office Scripts**: automatización en TypeScript pensada para Excel en la web. Sustituye a VBA en ese
  entorno, pero no ejecuta VBA ni lo convierte.
- **Python en Excel**: fórmula `=PY()` dentro de la celda. El código se ejecuta en la nube de Microsoft,
  no en el ordenador, con un conjunto de librerías fijado por Microsoft. Depende de la licencia de
  Microsoft 365. Disponibilidad en Excel para Mac: sin comprobar. Que los datos salgan a la nube puede
  ser un problema en una gestora.

## Cuándo sigue teniendo sentido VBA

1. **Herramientas heredadas que funcionan** y de las que depende un proceso mensual, como la de valor
   liquidativo del problema. Reescribirlas tiene coste y riesgo: hay que revalidar cada resultado.
2. **Entornos corporativos cerrados**: en banca o gestoras es habitual que solo se pueda usar Excel y que
   informática no permita instalar Python ni ejecutar programas propios.
3. **Automatizaciones pequeñas pegadas a la interfaz de Excel**: un botón que da formato, una macro que
   ordena hojas. Ahí Python no aporta.

## Cuándo Power BI

- Cuando lo que se necesita es **ver y compartir** resultados: cuadros de mando, informes periódicos,
  filtros, permisos por usuario dentro de la empresa (Microsoft 365).
- Se puede automatizar el **refresco de datos** programado y conectarlo a Excel, bases de datos o archivos.
- No sustituye a la **lógica de cálculo** compleja: la complementa. El cálculo (por ejemplo, el valor
  liquidativo) conviene hacerlo en Python o en la base de datos y que Power BI lo muestre.
- **Power BI Desktop**, la aplicación donde se diseñan los informes, solo existe para Windows. En Mac
  hay que usar una máquina virtual o el servicio web, que tiene menos funciones de diseño.
- **Compartir y permisos** requieren licencia Power BI Pro o Premium por usuario (PPU) para quien publica
  y para quien consume, o que el espacio de trabajo esté en una capacidad Premium o Fabric.
- **Refrescar un Excel local** de forma programada exige una puerta de enlace de datos local (gateway)
  instalada en un equipo encendido, o guardar el Excel en OneDrive o SharePoint.

## La alternativa que proponemos: aplicación a medida en Python

En lugar de mantener VBA dentro de Excel:

```
Excel con los datos (el usuario sigue trabajando ahí)
        |
        v
Aplicación Python  ->  lee el Excel (pandas / openpyxl), sin abrir Excel
        |               calcula (valor liquidativo, comisiones, controles)
        |               valida con tests automáticos
        v
Panel (Streamlit o Dash): tablas, gráficos, alertas, con refresco programado
        |
        v
Salidas: Excel o PDF de informe, histórico, registro de cambios
```

- **Leer el Excel con openpyxl**: no recalcula fórmulas. Con `data_only=True` lee el último valor que
  Excel guardó en el archivo; si el libro se generó sin abrirlo en Excel, esa celda puede estar vacía.
  Si la aplicación escribe en un `.xlsm`, hay que abrirlo con `keep_vba=True` o el VBA se pierde al
  guardar.
- **Panel**: Streamlit o Dash dan una pantalla web con tablas, gráficos y botones. El "tiempo real" no es
  automático: necesitan un proceso servidor en marcha y un refresco programado (`run_every` en un
  fragmento de Streamlit, `dcc.Interval` en Dash) que vuelva a leer el Excel de entrada.
- **"Compilado"**: PyInstaller no compila. Empaqueta el intérprete de Python, el bytecode y las librerías
  en un ejecutable para que el usuario no tenga que instalar Python. Hay que generarlo en cada sistema
  operativo (uno para Windows, otro para Mac). Con `pandas` pasa de 100 MB. Los antivirus dan a menudo
  falsos positivos con estos ejecutables, y en Mac hay que firmarlo y notarizarlo con una cuenta de
  desarrollador de Apple para que Gatekeeper lo deje abrir. En empresas, informática puede bloquear
  ejecutables no firmados: hay que contar con ellos.
- **Ventajas**: pruebas automáticas reales, versiones en git, sin ventanas que bloqueen, sin depender de
  Excel en primer plano, y Claude puede trabajar sobre el código directamente.
- **Límites**: hay que reescribir y revalidar la lógica; hay que acordar con informática dónde se ejecuta;
  el usuario tiene que acostumbrarse a una herramienta nueva.

## Qué haríamos nosotros, en una frase

**VBA en 2027, solo para mantener lo que ya existe mientras se migra.** Lo construido en este repositorio
no es una apuesta por VBA: es el puente seguro para salir de él sin romper nada.

| Situación | Qué haríamos |
|---|---|
| Algo nuevo (cálculo, informe, automatización) | Python. Excel solo como entrada y salida de datos; la lógica en Python, con pruebas y en git; encima un panel (Streamlit o Dash) o un informe automático |
| Cuadros de mando para compartir en una empresa que ya usa Microsoft 365 | Power BI mostrando resultados calculados en Python o en la base de datos (recuerda: Desktop solo en Windows y licencias para compartir) |
| Herramienta VBA heredada de la que depende un proceso | No reescribirla de golpe: el plan por fases de abajo |
| Entorno donde informática solo permite Excel | Quedarse en la fase 1: VBA protegido con tests ya es una mejora enorme |

## Plan por fases para una herramienta VBA heredada

Pensado para un caso como el del problema planteado: una herramienta de valor liquidativo hecha hacia 2015,
con muchos módulos, que se usa cada mes y cuyo autor ya no está. Las duraciones son orientativas.

### Fase 0. Preguntar qué permite el entorno (antes de todo)

- En una gestora o un banco es probable que informática no deje instalar Python, ejecutar programas
  propios ni activar el ajuste de confianza de VBA. Hay que preguntarlo primero.
- Si la respuesta es "solo Excel", el plan termina en la fase 1. Si permite Python, se sigue.

### Fase 1. Proteger lo que hay (semanas)

- Exactamente lo que hace este repositorio: tests del VBA contra una referencia Python independiente,
  calculada desde la regla de negocio y no desde el propio VBA.
- Desde ese momento, cada corrección se valida y se repasan los casos que ya funcionaban.
- El proceso mensual sigue funcionando igual; nadie nota el cambio salvo en que hay menos errores.

### Fase 2. Pasar el cálculo a Python por partes (meses)

- Se migra una función o un bloque cada vez (por ejemplo, primero el devengo de comisiones).
- Cada parte migrada se ejecuta **en paralelo** con el VBA, con los mismos datos, durante varios cierres
  mensuales. Cuando coinciden siempre, el VBA de esa parte se retira.
- La referencia Python de la fase 1 ya es el primer trozo de esta migración.
- Ojo: si la referencia Python pasa a ser el código de producción, **deja de ser un oráculo independiente**
  (un error en ella ya no lo detecta nadie). Hay que conservar casos esperados calculados aparte, desde la
  especificación (a mano o con otra implementación), y probar contra ellos.

### Fase 3. Poner el panel encima

- Con el cálculo ya en Python: panel con tablas, gráficos, alertas e histórico, o un informe PDF/Excel
  automático. Power BI si la empresa ya lo tiene y lo que se busca es compartir y filtrar.

### Fase 4. Dejar Excel solo para la entrada de datos

- La gente sigue introduciendo o revisando datos en Excel, que es donde trabaja.
- Del VBA queda nada o casi nada (si acaso, algo pegado a la interfaz, como un botón).

### Por qué este orden y no reescribir todo de golpe

- **Es un cálculo con consecuencias reales.** Migrar por partes, comparando siempre con lo que ya funciona,
  evita el "big bang": reescribir todo y descubrir al final que la versión nueva no da lo mismo.
- **El proceso nunca se para.** En cada fase hay una versión que funciona y está validada.
- **Cada fase aporta valor por sí sola.** Si el proyecto se detiene en la fase 1 o 2, lo hecho sigue sirviendo.

## Qué no hemos comprobado

- No hemos construido ni la aplicación Python ni un panel para este caso.
- No sabemos qué permite instalar o ejecutar el entorno de la persona que mantiene la herramienta.
- Las ventajas de la tabla son generales del lenguaje, no medidas en este proyecto. Tampoco hemos medido
  la velocidad de VBA con matrices frente a celda a celda.
- No hemos probado Rubberduck, Office Scripts, Python en Excel ni Power BI. Lo dicho sobre ellos (licencias,
  plataformas, gateway, ejecución en la nube) sale de su documentación y puede cambiar: comprobarlo antes
  de decidir.
- No sabemos si Python en Excel está disponible en el Excel para Mac de este proyecto.
- No hemos empaquetado nada con PyInstaller ni medido su tamaño en este caso.
