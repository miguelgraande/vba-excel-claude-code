# =====================================================================
# SIN PROBAR: escrito el 14/11/2026 en un Mac (compilar y vigia: 21/11); probar en Windows antes de usar.
# =====================================================================
#
# vba.ps1: puente terminal <-> VBA de Excel en Windows, por COM.
# No instala nada: usa Windows PowerShell 5.1 (viene con Windows) y el Excel instalado.
#
# Uso (desde la carpeta del script):
#   powershell -NoProfile -ExecutionPolicy Bypass -File .\vba.ps1 <comando> <libro> [...]
#
#   modulos  <libro>                       lista modulos: nombre, tipo, lineas
#   exportar <libro> <carpeta>             saca cada modulo a .bas/.cls/.frm (UTF-8)
#   importar <libro> <archivo.bas|.cls>    mete el modulo; reemplaza el del mismo nombre y guarda
#   ejecutar <libro> <Macro> [args...]     ejecuta y escribe en pantalla lo que devuelve
#   compilar <libro>                       Depurar > Compilar sin ejecutar nada: OK compila / ERROR COMPILACION
#
# Opciones:
#   -Leer "Hoja" o "Hoja!A1:C20"   (ejecutar) tras la macro, imprime ese rango separado por ;
#   -Guardar                       (ejecutar) guarda el libro al terminar (por defecto NO guarda)
#   -Tiempo 120                    (ejecutar, compilar) segundos maximos; despues mata SOLO nuestro Excel.
#                                  0 = sin limite
#
# Vigia de ventanas (ejecutar y compilar): mira cada 0,5 s las ventanas de NUESTRO Excel y lee su texto.
#   - Aviso simple (un solo boton Aceptar/OK): pulsa Aceptar y lo cuenta (AVISO cerrado con Aceptar).
#   - En compilar, el aviso de error de compilacion: lo lee y pulsa Aceptar.
#   - Cualquier otra (Finalizar/Depurar, InputBox, Si/No): NO pulsa nada, la describe y mata nuestro Excel.
# Tecnica adaptada de MCP-Access (github.com/unmateria/MCP-Access, commit 3e0fe05, MIT):
# compile.py (_ensure_code_pane, _compile_dialog_watchdog) y vba_exec.py (_find_dialog_hwnds_by_pid).
#   -Eventos                       activa eventos (Workbook_Open, etc.). Por defecto apagados
#   -Visible                       muestra la ventana de Excel (para ver una ventana atascada)
#   -Ansi                          (exportar) deja los archivos en ANSI, como los saca Excel
#
# Codigos de salida: 0 bien, 1 uso incorrecto, 2 fallo.
#
# Reglas que sigue:
#   - Abre una instancia de Excel PROPIA y nueva. No toca los Excel que ya tengas abiertos.
#   - DisplayAlerts = $false, Visible = $false salvo -Visible.
#   - try/finally: siempre cierra el libro sin guardar (salvo importar o -Guardar),
#     sale de Excel y libera los objetos COM para no dejar EXCEL.EXE colgado.
#   - No escribe en el registro. Solo lo LEE para explicar por que falla el acceso al VBA.
#
# Comentarios sin tildes a proposito: Windows PowerShell 5.1 lee los .ps1 sin BOM como ANSI.

[CmdletBinding()]
param(
    [Parameter(Position = 0)] [string] $Comando,
    [Parameter(Position = 1)] [string] $Libro,
    [Parameter(Position = 2, ValueFromRemainingArguments = $true)] [string[]] $Resto,
    [string] $Leer,
    [switch] $Guardar,
    [int] $Tiempo = 120,
    [switch] $Eventos,
    [switch] $Visible,
    [switch] $Ansi
)

# Primero de todo: en modo restringido (AppLocker/WDAC de empresa) PowerShell no deja usar COM
if ($ExecutionContext.SessionState.LanguageMode -ne 'FullLanguage') {
    Write-Error ("PowerShell esta en modo $($ExecutionContext.SessionState.LanguageMode). " +
        'Ese modo no deja crear Excel por COM. Hay que hablarlo con informatica.')
    exit 2
}

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Constantes (documentacion oficial de Microsoft)
$MSO_SEGURIDAD_BAJA        = 1    # msoAutomationSecurityLow: macros activadas
$MSO_SEGURIDAD_DESACTIVADA = 3    # msoAutomationSecurityForceDisable: macros apagadas sin avisos
$TIPOS = @{ 1 = 'Modulo'; 2 = 'Clase'; 3 = 'Formulario'; 11 = 'ActiveX'; 100 = 'Documento' }
$EXTENSIONES = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.cls' }
$INV = [System.Globalization.CultureInfo]::InvariantCulture

# Salida en UTF-8 para que las tildes lleguen bien a Claude Code
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

function Mostrar-Uso {
    Write-Host @'
Uso:
  vba.ps1 modulos  <libro>
  vba.ps1 exportar <libro> <carpeta> [-Ansi]
  vba.ps1 importar <libro> <archivo.bas|.cls|.frm>
  vba.ps1 ejecutar <libro> <Macro> [args...] [-Leer "Hoja!A1:C20"] [-Guardar] [-Tiempo 120]
  vba.ps1 compilar <libro> [-Tiempo 60]
Opciones comunes: -Visible -Eventos
'@
}

# ---------------------------------------------------------------------
# Vigia de ventanas (Win32). Corre en un trabajo aparte porque Run y Execute
# no vuelven mientras haya una ventana modal abierta.
# ---------------------------------------------------------------------
$CODIGO_VENTANAS = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VentanasVba {
    delegate bool Recorrer(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(Recorrer f, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr padre, Recorrer f, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, StringBuilder l, uint f, uint t, out IntPtr r);
    [DllImport("user32.dll")]
    static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint f, uint t, out IntPtr r);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    static string Clase(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
    // WM_GETTEXT con tiempo maximo: GetWindowText no lee controles de otro proceso
    public static string Texto(IntPtr h) {
        IntPtr n;
        if (SendMessageTimeout(h, 0x000E, IntPtr.Zero, IntPtr.Zero, 2, 1000, out n) == IntPtr.Zero) return "";
        var s = new StringBuilder(n.ToInt32() + 1);
        IntPtr r;
        SendMessageTimeout(h, 0x000D, (IntPtr)s.Capacity, s, 2, 1000, out r);
        return s.ToString();
    }
    // Ventanas modales (clase #32770) visibles del proceso pid
    public static IntPtr[] Dialogos(int pid) {
        var lista = new List<IntPtr>();
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p == (uint)pid && IsWindowVisible(h) && Clase(h) == "#32770") lista.Add(h);
            return true;
        }, IntPtr.Zero);
        return lista.ToArray();
    }
    public static IntPtr[] Hijos(IntPtr padre, string clase) {
        var lista = new List<IntPtr>();
        EnumChildWindows(padre, (h, l) => { if (Clase(h) == clase) lista.Add(h); return true; }, IntPtr.Zero);
        return lista.ToArray();
    }
    public static void Pulsar(IntPtr boton) { PostMessage(boton, 0x00F5, IntPtr.Zero, IntPtr.Zero); } // BM_CLICK
}
'@

# Devuelve lineas TIPO<tab>TITULO<tab>TEXTO<tab>BOTONES, con TIPO = AVISO (cerrado con Aceptar),
# VENTANA (no se toca; mata nuestro Excel) o AGOTADO (tiempo; mata nuestro Excel).
$VIGIA = {
    param([int] $id, [int] $seg, [string] $codigo, [string] $modo)
    Add-Type -TypeDefinition $codigo
    $fin = if ($seg -gt 0) { (Get-Date).AddSeconds($seg) } else { [datetime]::MaxValue }
    $vistas = @{}
    while ((Get-Date) -lt $fin) {
        if (-not (Get-Process -Id $id -ErrorAction SilentlyContinue)) { return }
        foreach ($d in [VentanasVba]::Dialogos($id)) {
            if ($vistas.ContainsKey([string] $d)) { continue }
            $vistas[[string] $d] = $true
            $titulo = [VentanasVba]::Texto($d)
            $texto = @([VentanasVba]::Hijos($d, 'Static') | ForEach-Object { [VentanasVba]::Texto($_).Trim() } | Where-Object { $_ }) -join ' / '
            $botones = @([VentanasVba]::Hijos($d, 'Button') | ForEach-Object { @{ h = $_; t = [VentanasVba]::Texto($_).Replace('&', '').Trim() } } | Where-Object { $_.t })
            $aceptar = @($botones | Where-Object { $_.t -in @('Aceptar', 'OK') })
            $conEdicion = @([VentanasVba]::Hijos($d, 'Edit')).Count -gt 0
            $linea = "$titulo`t$texto`t" + (@($botones | ForEach-Object { $_.t }) -join ', ')
            # Aviso simple: un solo boton Aceptar. En compilar vale tambien Aceptar + Ayuda (error de compilacion)
            $cerrable = $aceptar.Count -eq 1 -and -not $conEdicion -and ($modo -eq 'compilar' -or $botones.Count -eq 1)
            if ($cerrable) {
                [VentanasVba]::Pulsar($aceptar[0].h)
                "AVISO`t$linea"
            } else {
                "VENTANA`t$linea"
                Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
                return
            }
        }
        Start-Sleep -Milliseconds 500
    }
    Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
    'AGOTADO'
}

function Iniciar-Vigia([int] $id, [int] $seg, [string] $modo) {
    return Start-Job -ScriptBlock $VIGIA -ArgumentList $id, $seg, $CODIGO_VENTANAS, $modo
}

# Para el vigia y devuelve sus lineas ya separadas: @{ Tipo; Titulo; Texto; Botones }
function Parar-Vigia($vigia, [int] $esperar = 0) {
    if (-not $vigia) { return @() }
    if ($esperar -gt 0) { $null = Wait-Job $vigia -Timeout $esperar }
    Stop-Job $vigia -ErrorAction SilentlyContinue
    $lineas = @(Receive-Job $vigia -ErrorAction SilentlyContinue)
    Remove-Job $vigia -Force -ErrorAction SilentlyContinue
    return @($lineas | ForEach-Object {
        $p = (([string] $_) + "`t`t`t") -split "`t"
        @{ Tipo = $p[0]; Titulo = $p[1]; Texto = $p[2]; Botones = $p[3] }
    })
}

function Describir-Ventana($v) { return "$($v.Titulo) / $($v.Texto) [botones: $($v.Botones)]" }

# Tras un error de compilacion el editor deja el cursor en la linea del fallo (ActiveCodePane)
function Lugar-Error($excel) {
    try {
        $panel = $excel.VBE.ActiveCodePane
        $l1 = 0; $c1 = 0; $l2 = 0; $c2 = 0
        $panel.GetSelection([ref] $l1, [ref] $c1, [ref] $l2, [ref] $c2)
        $modulo = $panel.CodeModule
        return " - $($modulo.Parent.Name), Lin $l1, Col $c1 : $($modulo.Lines($l1, 1).Trim())"
    } catch { return '' }
}

function Salir-Con([int] $codigo, [string] $mensaje) {
    if ($mensaje) { [Console]::Error.WriteLine($mensaje) }
    exit $codigo
}

# Codificacion ANSI del sistema (la que usa el editor de VBA al exportar e importar).
# Suposicion sin probar: en Windows en castellano es Windows-1252.
function Codificacion-Ansi {
    $cp = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
    return [System.Text.Encoding]::GetEncoding($cp)
}

# Lee un archivo de texto: si es UTF-8 valido lo trata como UTF-8; si no, como ANSI.
function Leer-Texto([string] $ruta) {
    $bytes = [System.IO.File]::ReadAllBytes($ruta)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    try {
        $estricto = New-Object System.Text.UTF8Encoding($false, $true)
        return $estricto.GetString($bytes)
    } catch {
        return (Codificacion-Ansi).GetString($bytes)
    }
}

# Convierte un argumento de texto en numero o booleano si lo parece.
# Los decimales van con punto (2.5), nunca con coma. Prefijo s: fuerza texto (s:007).
function Convertir-Argumento([string] $s) {
    if ($s.StartsWith('s:')) { return $s.Substring(2) }
    $entero = 0L
    $doble = 0.0
    if ([long]::TryParse($s, [System.Globalization.NumberStyles]::Integer, $INV, [ref] $entero)) {
        if ($entero -ge [int]::MinValue -and $entero -le [int]::MaxValue) { return [int] $entero }
        return $entero
    }
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $INV, [ref] $doble)) { return $doble }
    if ($s -eq 'true' -or $s -eq 'verdadero') { return $true }
    if ($s -eq 'false' -or $s -eq 'falso') { return $false }
    return $s
}

# Pasa un valor de Excel a texto. Numeros con punto decimal para compararlos con Python.
function Texto-De($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { return $v.ToString('R', $INV).Replace('E+', 'E') }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm:ss', $INV) }
    if ($v -is [System.DBNull]) { return '' }
    return [string] $v
}

# Escribe un valor que puede ser simple o una matriz de Excel (1..n, 1..m).
function Escribir-Valor($v) {
    if ($v -is [System.Array]) {
        if ($v.Rank -eq 2) {
            for ($f = $v.GetLowerBound(0); $f -le $v.GetUpperBound(0); $f++) {
                $celdas = @()
                for ($c = $v.GetLowerBound(1); $c -le $v.GetUpperBound(1); $c++) { $celdas += (Texto-De $v[$f, $c]) }
                Write-Output ($celdas -join ';')
            }
        } else {
            foreach ($x in $v) { Write-Output (Texto-De $x) }
        }
    } else {
        Write-Output (Texto-De $v)
    }
}

# Explica por que no hay acceso al modelo de objetos de VBA. Solo lee el registro.
function Diagnostico-AccessVBOM {
    $lineas = @()
    $ver = '16.0'
    $directiva = "HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security"
    $usuario   = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    foreach ($par in @(@('Directiva de grupo (GPO)', $directiva), @('Ajuste del usuario', $usuario))) {
        $valor = $null
        try { $valor = (Get-ItemProperty -Path $par[1] -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch { }
        if ($null -eq $valor) { $lineas += "  $($par[0]): sin valor ($($par[1]))" }
        else { $lineas += "  $($par[0]): AccessVBOM = $valor ($($par[1]))" }
    }
    $lineas += '  Si la directiva vale 0, la casilla esta bloqueada: solo informatica puede cambiarla.'
    $lineas += '  Si no hay directiva: Archivo > Opciones > Centro de confianza > Configuracion del Centro de confianza >'
    $lineas += '  Configuracion de macros > "Confiar en el acceso al modelo de objetos de proyectos de VBA".'
    return ($lineas -join [Environment]::NewLine)
}

function Obtener-Proyecto($wb) {
    try {
        $proyecto = $wb.VBProject
        $null = $proyecto.VBComponents.Count
    } catch {
        throw ("Sin acceso al proyecto VBA: $($_.Exception.Message)" + [Environment]::NewLine + (Diagnostico-AccessVBOM))
    }
    if ($proyecto.Protection -eq 1) {
        throw 'El proyecto VBA esta protegido con contrasena. Hay que desbloquearlo a mano en el editor (Alt+F11).'
    }
    return $proyecto
}

# Quita la cabecera de un .cls exportado (VERSION, BEGIN..END, Attribute VB_*)
# para meter solo el codigo en un modulo de documento (ThisWorkbook, Hoja1...).
function Quitar-Cabecera([string[]] $lineas) {
    $i = 0
    $dentroBegin = $false
    while ($i -lt $lineas.Length) {
        $l = $lineas[$i]
        if ($l -match '^VERSION\s') { $i++; continue }
        if ($l -match '^BEGIN\b') { $dentroBegin = $true; $i++; continue }
        if ($dentroBegin) { if ($l -match '^END\s*$') { $dentroBegin = $false }; $i++; continue }
        if ($l -match '^Attribute\s+VB_') { $i++; continue }
        break
    }
    if ($i -ge $lineas.Length) { return @() }
    return $lineas[$i..($lineas.Length - 1)]
}

# ---------------------------------------------------------------------
# Comprobaciones previas (no abren Excel)
# ---------------------------------------------------------------------
if (-not $Comando -or $Comando -in @('-h', '--help', 'ayuda', 'help')) { Mostrar-Uso; exit 1 }
if ($Comando -notin @('modulos', 'exportar', 'importar', 'ejecutar', 'compilar')) { Mostrar-Uso; Salir-Con 1 "Comando desconocido: $Comando" }
if (-not $Libro) { Mostrar-Uso; Salir-Con 1 'Falta el libro.' }
if (-not $Resto) { $Resto = @() }

try { $rutaLibro = (Resolve-Path -LiteralPath $Libro).ProviderPath }
catch { Salir-Con 1 "No encuentro el libro: $Libro" }

$extLibro = [System.IO.Path]::GetExtension($rutaLibro).ToLowerInvariant()
if ($Comando -eq 'importar' -and $extLibro -notin @('.xlsm', '.xlsb', '.xls', '.xlam', '.xltm')) {
    Salir-Con 1 "El libro $extLibro no admite macros. Usa .xlsm o .xlsb."
}

switch ($Comando) {
    'exportar' { if ($Resto.Count -lt 1) { Salir-Con 1 'Falta la carpeta de destino.' } }
    'importar' {
        if ($Resto.Count -lt 1) { Salir-Con 1 'Falta el archivo .bas/.cls/.frm.' }
        try { $rutaModulo = (Resolve-Path -LiteralPath $Resto[0]).ProviderPath }
        catch { Salir-Con 1 "No encuentro el archivo: $($Resto[0])" }
    }
    'ejecutar' { if ($Resto.Count -lt 1) { Salir-Con 1 'Falta el nombre de la macro.' } }
}

# ---------------------------------------------------------------------
# Excel propio
# ---------------------------------------------------------------------
$excel = $null
$wb = $null
$pidExcel = $null
$vigia = $null
$guardarAlFinal = $false
$codigoSalida = 0

try {
    # PID de los Excel que ya existian, para reconocer el nuestro y no tocar los demas
    $antes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

    # CreateObject / New-Object -ComObject abre una instancia nueva (KB 288902)
    $excel = New-Object -ComObject Excel.Application

    $nuevos = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $antes -notcontains $_.Id } | ForEach-Object { $_.Id })
    if ($nuevos.Count -eq 1) { $pidExcel = $nuevos[0] }
    else { [Console]::Error.WriteLine('Aviso: no se pudo identificar el proceso de Excel propio; no habra limite de tiempo.') }

    $excel.Visible = [bool] $Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = [bool] $Eventos
    $excel.ScreenUpdating = [bool] $Visible

    # Macros apagadas salvo para ejecutar y compilar: asi exportar/modulos/importar no disparan codigo al abrir.
    # compilar las deja activas por si con el proyecto desactivado el editor no compila (sin comprobar).
    if ($Comando -in @('ejecutar', 'compilar')) { $excel.AutomationSecurity = $MSO_SEGURIDAD_BAJA }
    else { $excel.AutomationSecurity = $MSO_SEGURIDAD_DESACTIVADA }

    # Open(FileName, UpdateLinks=0, ReadOnly). compilar toca el codigo solo en memoria y no guarda
    $soloLectura = ($Comando -in @('modulos', 'exportar', 'compilar'))
    $wb = $excel.Workbooks.Open($rutaLibro, 0, $soloLectura)

    if (($Comando -eq 'importar' -or $Guardar) -and $wb.ReadOnly) {
        throw 'El libro se ha abierto en solo lectura (probablemente esta abierto en otro Excel). Cierralo y repite.'
    }

    switch ($Comando) {

        'modulos' {
            $proyecto = Obtener-Proyecto $wb
            foreach ($comp in $proyecto.VBComponents) {
                $tipo = $TIPOS[[int] $comp.Type]
                if (-not $tipo) { $tipo = "Tipo $($comp.Type)" }
                Write-Output ('{0};{1};{2}' -f $comp.Name, $tipo, $comp.CodeModule.CountOfLines)
            }
        }

        'exportar' {
            $proyecto = Obtener-Proyecto $wb
            $carpeta = $Resto[0]
            if (-not (Test-Path -LiteralPath $carpeta)) { $null = New-Item -ItemType Directory -Path $carpeta }
            $carpeta = (Resolve-Path -LiteralPath $carpeta).ProviderPath
            $ansi = Codificacion-Ansi
            foreach ($comp in $proyecto.VBComponents) {
                $t = [int] $comp.Type
                # Hojas y ThisWorkbook vacios no aportan nada
                if ($t -eq 100 -and $comp.CodeModule.CountOfLines -eq 0) { continue }
                $ext = $EXTENSIONES[$t]
                if (-not $ext) { [Console]::Error.WriteLine("Omitido $($comp.Name): tipo $t"); continue }
                $destino = Join-Path $carpeta ($comp.Name + $ext)
                # Export falla si el archivo ya existe (documentacion de Microsoft): se borra antes
                if (Test-Path -LiteralPath $destino) { Remove-Item -LiteralPath $destino -Force }
                if ($t -eq 3) {
                    $frx = [System.IO.Path]::ChangeExtension($destino, '.frx')
                    if (Test-Path -LiteralPath $frx) { Remove-Item -LiteralPath $frx -Force }
                }
                $comp.Export($destino)
                if (-not $Ansi) {
                    # Excel exporta en ANSI. Se pasa a UTF-8 sin BOM para VS Code y git.
                    $texto = [System.IO.File]::ReadAllText($destino, $ansi)
                    [System.IO.File]::WriteAllText($destino, $texto, (New-Object System.Text.UTF8Encoding($false)))
                }
                Write-Output $destino
            }
        }

        'importar' {
            $proyecto = Obtener-Proyecto $wb
            $ext = [System.IO.Path]::GetExtension($rutaModulo).ToLowerInvariant()
            if ($ext -notin @('.bas', '.cls', '.frm')) { throw "Extension no admitida: $ext" }

            # Texto normalizado: saltos de linea CRLF, como espera el editor de VBA
            $texto = (Leer-Texto $rutaModulo).TrimEnd("`r", "`n")
            [string[]] $lineas = @($texto -split "`r?`n")

            # Nombre del modulo: Attribute VB_Name o, si falta, el nombre del archivo
            $nombre = $null
            foreach ($l in $lineas) {
                if ($l -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') { $nombre = $Matches[1]; break }
            }
            if (-not $nombre) {
                $nombre = [System.IO.Path]::GetFileNameWithoutExtension($rutaModulo)
                if ($ext -eq '.bas') { [string[]] $lineas = @("Attribute VB_Name = `"$nombre`"") + $lineas }
                else { [Console]::Error.WriteLine("Aviso: $ext sin Attribute VB_Name; se usara $nombre") }
            }

            $existente = $null
            foreach ($c in $proyecto.VBComponents) { if ($c.Name -eq $nombre) { $existente = $c; break } }

            if ($null -ne $existente -and [int] $existente.Type -eq 100) {
                # Modulo de documento (ThisWorkbook, Hoja1): no se puede borrar.
                # Se vacia con DeleteLines y se mete el codigo con AddFromString.
                $codigo = (Quitar-Cabecera $lineas) -join "`r`n"
                $cm = $existente.CodeModule
                if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
                if ($codigo) { $cm.AddFromString($codigo) }
                Write-Output "OK $nombre (documento) $($cm.CountOfLines) lineas"
            } else {
                if ($null -ne $existente) { $proyecto.VBComponents.Remove($existente) }

                # Copia temporal en ANSI + CRLF, que es lo que lee Import
                $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('vba_' + [guid]::NewGuid().ToString('N'))
                $null = New-Item -ItemType Directory -Path $tmpDir
                try {
                    $tmp = Join-Path $tmpDir ($nombre + $ext)
                    [System.IO.File]::WriteAllText($tmp, (($lineas -join "`r`n") + "`r`n"), (Codificacion-Ansi))
                    if ($ext -eq '.frm') {
                        $frx = [System.IO.Path]::ChangeExtension($rutaModulo, '.frx')
                        if (Test-Path -LiteralPath $frx) { Copy-Item -LiteralPath $frx -Destination ([System.IO.Path]::ChangeExtension($tmp, '.frx')) }
                    }
                    $nuevo = $proyecto.VBComponents.Import($tmp)
                } finally {
                    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
                }
                # Si Excel lo renombra (Modulo11), se devuelve a su nombre
                if ($nuevo.Name -ne $nombre) {
                    try { $nuevo.Name = $nombre } catch { [Console]::Error.WriteLine("Aviso: el modulo quedo como $($nuevo.Name)") }
                }
                Write-Output "OK $($nuevo.Name) $($nuevo.CodeModule.CountOfLines) lineas"
            }
            $guardarAlFinal = $true
        }

        'ejecutar' {
            $macro = $Resto[0]
            $argumentos = @()
            if ($Resto.Count -gt 1) { foreach ($a in $Resto[1..($Resto.Count - 1)]) { $argumentos += , (Convertir-Argumento $a) } }
            if ($argumentos.Count -gt 30) { throw 'Application.Run admite como maximo 30 argumentos.' }

            # Nombre completo: 'Libro.xlsm'!Modulo.Macro (evita ejecutar una macro de otro libro)
            if ($macro -notmatch '!') { $macro = "'" + $wb.Name.Replace("'", "''") + "'!" + $macro }

            # Vigia: cierra avisos simples con Aceptar; ante otra ventana o sin respuesta a tiempo, mata NUESTRO Excel
            if ($pidExcel) { $vigia = Iniciar-Vigia $pidExcel $Tiempo 'ejecutar' }

            $resultado = $null
            try {
                # Application.Run devuelve lo que devuelva la macro (documentacion de Microsoft).
                # Argumentos solo por posicion. .Invoke() permite pasar un numero variable.
                [object[]] $parametros = @($macro) + $argumentos
                $resultado = $excel.Run.Invoke($parametros)
            } catch {
                $vistas = Parar-Vigia $vigia 5; $vigia = $null
                $ventana = @($vistas | Where-Object { $_.Tipo -eq 'VENTANA' })
                if ($ventana.Count -gt 0) {
                    $excel = $null; $wb = $null
                    throw ("ERROR VENTANA VBA (no se ha pulsado nada): $(Describir-Ventana $ventana[0]). " +
                        'Excel propio cerrado a la fuerza; se pierde lo no guardado. Compila antes y usa SinVentanas.')
                }
                if (@($vistas | Where-Object { $_.Tipo -eq 'AGOTADO' }).Count -gt 0) {
                    $excel = $null; $wb = $null
                    throw "Tiempo agotado ($Tiempo s). Probable ventana modal (MsgBox, InputBox o error sin controlar). Excel propio cerrado a la fuerza."
                }
                $msg = $_.Exception.Message
                if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
                throw "Fallo al ejecutar $macro : $msg"
            }
            $vistas = Parar-Vigia $vigia; $vigia = $null
            foreach ($v in @($vistas | Where-Object { $_.Tipo -eq 'AVISO' })) {
                [Console]::Error.WriteLine("AVISO cerrado con Aceptar: $(Describir-Ventana $v) (usa SinVentanas para evitarlo)")
            }

            Escribir-Valor $resultado

            if ($Leer) {
                if ($Leer -match '^(.+)!(.+)$') { $hoja = $Matches[1].Trim("'"); $rango = $Matches[2] }
                else { $hoja = $Leer.Trim("'"); $rango = $null }
                $ws = $wb.Worksheets.Item($hoja)
                if ($rango) { $r = $ws.Range($rango) } else { $r = $ws.UsedRange }
                Write-Output "--- $hoja!$($r.Address($false, $false))"
                Escribir-Valor $r.Value2
            }
            if ($Guardar) { $guardarAlFinal = $true }
        }

        'compilar' {
            # Mismo formato de salida que vba.py compilar en Mac. Solo OK vale como comprobado.
            $proyecto = Obtener-Proyecto $wb
            $nombreLibro = $wb.Name
            if (-not $pidExcel) { Write-Output 'SIN VERIFICAR: sin PID de Excel propio no hay vigia; no compilo'; $codigoSalida = 2; break }

            # 1. Un componente con codigo, primero modulos normales
            $comp = $null
            foreach ($soloModulos in @($true, $false)) {
                foreach ($c in $proyecto.VBComponents) {
                    if ($soloModulos -and [int] $c.Type -ne 1) { continue }
                    if ($c.CodeModule.CountOfLines -gt 0) { $comp = $c; break }
                }
                if ($comp) { break }
            }
            if (-not $comp) { Write-Output "SIN VERIFICAR: $nombreLibro no tiene codigo que compilar"; $codigoSalida = 2; break }

            # 2. Marcar el proyecto como cambiado (solo en memoria): si Excel lo da por compilado, Compilar sale desactivado
            $cm = $comp.CodeModule
            $cm.InsertLines($cm.CountOfLines + 1, "' marca temporal de vba.ps1 compilar")
            $cm.DeleteLines($cm.CountOfLines, 1)

            # 3. Compilar actua sobre el proyecto ACTIVO del editor: se activa un panel de ESTE libro y se comprueba
            $excel.VBE.MainWindow.Visible = $true
            $cm.CodePane.Show()
            $activo = $null
            try { $activo = $excel.VBE.ActiveVBProject.FileName } catch { }
            if (-not $activo -or [System.IO.Path]::GetFullPath($activo) -ne [System.IO.Path]::GetFullPath($rutaLibro)) {
                Write-Output "SIN VERIFICAR: el editor tiene activo el proyecto de '$activo', no el de $nombreLibro; no compilo"
                $codigoSalida = 2; break
            }

            # 4. Boton Compilar por su ID (578), que no depende del idioma de Excel
            $boton = $null
            try { $boton = $excel.VBE.CommandBars.FindControl(1, 578) } catch { }
            if ($null -eq $boton) { Write-Output 'SIN VERIFICAR: no encuentro el boton Compilar (ID 578) en el editor'; $codigoSalida = 2; break }
            if (-not $boton.Enabled) {
                Write-Output 'SIN VERIFICAR: el boton Compilar esta desactivado (Excel lo da por compilado); no garantiza que compile'
                $codigoSalida = 2; break
            }

            # 5. Pulsar con el vigia en marcha: lee el aviso de error y pulsa Aceptar
            $limite = if ($Tiempo -gt 0) { [Math]::Min($Tiempo, 60) } else { 60 }
            $vigia = Iniciar-Vigia $pidExcel $limite 'compilar'
            Start-Sleep -Seconds 3   # que el vigia este listo antes de que salga un aviso
            $falloBoton = $null
            try { $boton.Execute() } catch { $falloBoton = $_.Exception.Message }
            Start-Sleep -Seconds 2   # avisos que salen tarde
            $vistas = Parar-Vigia $vigia; $vigia = $null

            $ventana = @($vistas | Where-Object { $_.Tipo -in @('VENTANA', 'AGOTADO') })
            if ($ventana.Count -gt 0) {
                $excel = $null; $wb = $null
                if ($ventana[0].Tipo -eq 'AGOTADO') { throw "Tiempo agotado ($limite s) al compilar. Excel propio cerrado a la fuerza." }
                throw "ERROR VENTANA VBA (no se ha pulsado nada): $(Describir-Ventana $ventana[0]). Excel propio cerrado a la fuerza."
            }
            $avisos = @($vistas | Where-Object { $_.Tipo -eq 'AVISO' })
            if ($avisos.Count -gt 0) {
                # El ultimo aviso es el del error; los anteriores suelen ser otros
                $a = $avisos[$avisos.Count - 1]
                Write-Output "ERROR COMPILACION: $($a.Titulo) / $($a.Texto) en $nombreLibro$(Lugar-Error $excel)"
                $codigoSalida = 2
            } elseif ($falloBoton) {
                Write-Output "SIN VERIFICAR: no se pudo pulsar Compilar ($falloBoton). Eso no quiere decir que el codigo tenga errores"
                $codigoSalida = 2
            } elseif ($boton.Enabled) {
                Write-Output 'SIN VERIFICAR: Compilar sigue activo tras pulsarlo; no se sabe si ha compilado'
                $codigoSalida = 2
            } else {
                Write-Output "OK compila $nombreLibro"
            }
        }
    }

    if ($guardarAlFinal) { $wb.Save() }
}
catch {
    [Console]::Error.WriteLine("ERROR: $($_.Exception.Message)")
    $codigoSalida = 2
}
finally {
    # 1. Cerrar sin guardar (lo que habia que guardar ya se guardo arriba)
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    # 2. Salir de Excel
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    # 3. Liberar COM (ReleaseComObject) y recoger basura para soltar las referencias intermedias
    foreach ($o in @($wb, $excel)) {
        if ($null -ne $o) { try { $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    $wb = $null; $excel = $null; $proyecto = $null; $comp = $null; $existente = $null; $nuevo = $null; $cm = $null; $ws = $null; $r = $null; $boton = $null; $c = $null
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    # 4. Vigia fuera
    if ($vigia) { Stop-Job $vigia -ErrorAction SilentlyContinue; Remove-Job $vigia -Force -ErrorAction SilentlyContinue }
    # 5. Ultimo recurso: si NUESTRO Excel sigue vivo a los 10 s, se cierra a la fuerza
    if ($pidExcel) {
        $p = Get-Process -Id $pidExcel -ErrorAction SilentlyContinue
        if ($p) {
            if (-not $p.WaitForExit(10000)) {
                Stop-Process -Id $pidExcel -Force -ErrorAction SilentlyContinue
                [Console]::Error.WriteLine("Aviso: EXCEL.EXE $pidExcel no se cerro solo; cerrado a la fuerza.")
            }
        }
    }
}
exit $codigoSalida
