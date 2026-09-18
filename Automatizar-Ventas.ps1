<#
.SINOPSIS
    Verifica en AC (Administracion de Clientes) que ventas quedaron activas,
    cuales no, y el motivo de las que no lo estan.

.DESCRIPCION
    Automatiza el paso a paso de "PASO A PASO AC.docx":

      1. Toma los numeros de la columna E del Excel de ventas
      2. Ingresa a AC con el usuario de red de la sesion de Windows
      3. Consulta cada numero por criterio MIN / MSISDN
      4. Selecciona la linea y lee la grilla de resultados
      5. Obtiene el ESTADO (activa / suspendida / otro)
      6. Para las que NO estan activas: abre la ficha, pulsa Ctrl+Shift+H
         y lee el HISTORIAL para saber el motivo

    La automatizacion es de SOLO LECTURA. AC obliga a guardar un
    "Solicitud Tickler" para cerrar una ficha de cliente; guardarlo dejaria
    un registro en produccion, por lo que el script nunca lo hace: cuando
    necesita cerrar una ficha, reinicia AC.

.EJEMPLOS
    .\Automatizar-Ventas.ps1 -Numeros 3001234567
    .\Automatizar-Ventas.ps1 -Archivo ".\entrada\ventas.xlsx"
    .\Automatizar-Ventas.ps1 -Archivo ".\entrada\ventas.xlsx" -HistorialSiempre
    .\Automatizar-Ventas.ps1 -GuardarClave        (guarda la clave cifrada y sale)
#>

[CmdletBinding()]
param(
    [string]   $Archivo,
    [string[]] $Numeros,
    [string]   $Columna    = 'E',
    [int]      $FilaInicio = 2,
    [string]   $Hoja,
    [string]   $Password,
    [switch]   $GuardarClave,
    [string]   $BaseDatos  = 'AC_PRODUCCION',
    [switch]   $HistorialSiempre,
    [string]   $Salida,
    [int]      $MaxReintentos = 2
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path

. "$raiz\lib\Win32.ps1"
. "$raiz\lib\Remote.ps1"
. "$raiz\lib\AC.ps1"
. "$raiz\lib\Excel.ps1"

$dirEntrada = Join-Path $raiz 'entrada'
$dirSalida  = Join-Path $raiz 'salida'
$dirConfig  = Join-Path $raiz 'config'
foreach ($d in @($dirEntrada, $dirSalida, $dirConfig)) {
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
}
$archivoClave = Join-Path $dirConfig 'clave.dat'

function Write-Titulo {
    param([string]$Texto)
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor DarkCyan
    Write-Host "  $Texto" -ForegroundColor Cyan
    Write-Host ("=" * 72) -ForegroundColor DarkCyan
}

# ---------------------------------------------------------------------
#  Clave de red (nunca se guarda en texto plano)
# ---------------------------------------------------------------------

function Save-Clave {
    $sec = Read-Host "Clave de red de $env:USERDOMAIN\$env:USERNAME" -AsSecureString
    $sec | ConvertFrom-SecureString | Set-Content -LiteralPath $archivoClave -Encoding UTF8
    Write-Host "Clave guardada cifrada en $archivoClave" -ForegroundColor Green
    Write-Host "Solo se puede descifrar con tu usuario de Windows en este equipo." -ForegroundColor Gray
}

function Get-Clave {
    param([string]$Explicita)
    if ($Explicita) { return $Explicita }
    if (Test-Path -LiteralPath $archivoClave) {
        try {
            $sec = Get-Content -LiteralPath $archivoClave -Raw | ConvertTo-SecureString
            $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
            try   { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        } catch {
            Write-Host "No se pudo leer la clave guardada: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    $sec = Read-Host "Clave de red de $env:USERDOMAIN\$env:USERNAME" -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    try   { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

if ($GuardarClave) { Save-Clave; return }

# ---------------------------------------------------------------------
#  Numeros a verificar
# ---------------------------------------------------------------------

Write-Titulo "VERIFICACION DE VENTAS EN AC"

$listaNumeros = @()
$origen = ""

if ($Numeros -and $Numeros.Count -gt 0) {
    foreach ($n in $Numeros) {
        $limpio = ConvertTo-NumeroLimpio -Valor $n
        if ($limpio) { $listaNumeros += $limpio }
        else { Write-Host "Numero ignorado (no son 10 digitos): '$n'" -ForegroundColor Yellow }
    }
    $origen = "parametro -Numeros"
} else {
    if (-not $Archivo) {
        $cand = Get-ChildItem -LiteralPath $dirEntrada -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in '.xlsx', '.xlsm', '.csv', '.txt' } |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $cand) {
            throw "No se indico -Archivo ni -Numeros, y la carpeta 'entrada' esta vacia.`n" +
                  "Copia alli el Excel de ventas (se usa la columna $Columna) y vuelve a ejecutar."
        }
        $Archivo = $cand.FullName
    }
    if (-not [System.IO.Path]::IsPathRooted($Archivo)) { $Archivo = Join-Path $raiz $Archivo }
    $lec = Get-NumerosDesdeArchivo -Path $Archivo -Columna $Columna -FilaInicio $FilaInicio -Hoja $Hoja
    $listaNumeros = $lec.Numeros
    $origen = $Archivo
    if ($lec.Descartados -gt 0) {
        Write-Host "Se descartaron $($lec.Descartados) celdas que no son numeros de 10 digitos." -ForegroundColor Yellow
    }
}

if ($listaNumeros.Count -eq 0) { throw "No hay numeros que verificar." }
Write-Host "Origen : $origen"
Write-Host "Numeros: $($listaNumeros.Count)"
Write-Host "Modo   : $(if ($HistorialSiempre) { 'historial para TODOS los numeros' } else { 'historial solo para las lineas NO activas' })"

$clave = Get-Clave -Explicita $Password

# ---------------------------------------------------------------------
#  Estructura de resultados
# ---------------------------------------------------------------------

$resultados = New-Object System.Collections.ArrayList
$sello = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $Salida) { $Salida = Join-Path $dirSalida "ventas_$sello.csv" }
$dirCapturas = Join-Path $dirSalida "capturas_$sello"

function New-Registro {
    param([string]$Numero)
    return [ordered]@{
        NUMERO              = $Numero
        ENCONTRADO          = ''
        ACTIVA              = ''
        ESTADO              = ''
        MOTIVO              = ''
        FECHA_ESTADO        = ''
        USUARIO_ESTADO      = ''
        NOMBRE              = ''
        CUSTCODE            = ''
        PLAN                = ''
        TECNOLOGIA          = ''
        TIPO_CLIENTE        = ''
        CENTRO_COSTOS       = ''
        HISTORIAL_COMPLETO  = ''
        OBSERVACION         = ''
        CONSULTADO          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

# ---------------------------------------------------------------------
#  PASADA 1 - estado de cada numero (sin abrir fichas, sin Ticklers)
# ---------------------------------------------------------------------

Write-Titulo "PASADA 1 - ESTADO DE CADA NUMERO"
Write-Host "No mover el raton ni el teclado mientras corre: AC se maneja con clics reales." -ForegroundColor Yellow
Write-Host ""

$ctx = Start-ACSession -Password $clave -BaseDatos $BaseDatos
$i = 0
foreach ($num in $listaNumeros) {
    $i++
    Write-Host ("[{0}/{1}] {2}" -f $i, $listaNumeros.Count, $num) -ForegroundColor White
    $reg = New-Registro -Numero $num
    $intento = 0
    $listo = $false

    while (-not $listo -and $intento -le $MaxReintentos) {
        $intento++
        try {
            if (-not (Get-ACProcess)) {
                Write-Paso "AC no esta en ejecucion; reiniciando sesion..." "WARN"
                $ctx = Start-ACSession -Password $clave -BaseDatos $BaseDatos
            }
            $r = Invoke-ACBusqueda -Numero $num

            if (-not $r.Encontrado) {
                $reg.ENCONTRADO = 'NO'
                $reg.ACTIVA     = 'NO'
                $reg.ESTADO     = 'NO ENCONTRADO'
                $reg.MOTIVO     = $r.Mensaje
                $reg.OBSERVACION= 'AC no devolvio ninguna linea para este numero.'
                Write-Paso "No encontrado. $($r.Mensaje)" "WARN"
            } else {
                # si el cliente tiene varias lineas, se toma la del numero consultado
                $fila = $r.Filas | Where-Object { ($_.Campos.MIN -replace '[^0-9]','') -eq $num } | Select-Object -First 1
                if (-not $fila) { $fila = $r.Filas[0] }
                $c = $fila.Campos

                $reg.ENCONTRADO    = 'SI'
                $reg.ESTADO        = $c.ESTADO
                $reg.NOMBRE        = $c.NOMBRE
                $reg.CUSTCODE      = $c.CUSTCODE
                $reg.PLAN          = $c.PLAN
                $reg.TECNOLOGIA    = $c.TECNOLOGIA
                $reg.TIPO_CLIENTE  = $c.'TIPO DE CLIENTE'
                $reg.CENTRO_COSTOS = $c.'CENTRO DE COSTOS'
                $reg.ACTIVA        = if ($c.ESTADO -and $c.ESTADO.Trim().ToUpper() -eq 'ACTIVO') { 'SI' } else { 'NO' }

                if ($r.Filas.Count -gt 1) {
                    $reg.OBSERVACION = "El cliente tiene $($r.Filas.Count) lineas; se reporta la del numero consultado."
                }
                $col = if ($reg.ACTIVA -eq 'SI') { "OK" } else { "WARN" }
                Write-Paso "$($c.NOMBRE) - ESTADO: $($c.ESTADO)" $col
            }
            $listo = $true
        } catch {
            $msg = $_.Exception.Message
            Write-Paso "Error (intento $intento): $msg" "ERROR"
            if ($intento -gt $MaxReintentos) {
                $reg.ENCONTRADO  = 'ERROR'
                $reg.ACTIVA      = 'REVISAR'
                $reg.ESTADO      = 'ERROR'
                $reg.OBSERVACION = $msg
                $c = Get-ACContext
                if ($c -and $c.Principal) {
                    [void](Save-WindowShot -Handle $c.Principal.Handle -Path (Join-Path $dirCapturas "error_$num.png"))
                }
            } else {
                # reiniciar AC deja el aplicativo en un estado limpio
                Stop-ACSession
                Start-Sleep -Seconds 2
                $ctx = Start-ACSession -Password $clave -BaseDatos $BaseDatos
            }
        }
    }

    [void]$resultados.Add([pscustomobject]$reg)
    try { Close-ACResultados -Contexto $ctx } catch { }
}

# ---------------------------------------------------------------------
#  PASADA 2 - motivo de las lineas que no estan activas
# ---------------------------------------------------------------------

$pendientes = @($resultados | Where-Object {
    $_.ENCONTRADO -eq 'SI' -and ($HistorialSiempre -or $_.ACTIVA -ne 'SI')
})

if ($pendientes.Count -gt 0) {
    Write-Titulo "PASADA 2 - MOTIVO (HISTORIAL, Ctrl+Shift+H)"
    Write-Host "Lineas por revisar: $($pendientes.Count)"
    Write-Host "AC se reinicia despues de cada ficha: es la unica forma de cerrarla sin guardar un Tickler." -ForegroundColor Gray
    Write-Host ""

    $j = 0
    foreach ($reg in $pendientes) {
        $j++
        Write-Host ("[{0}/{1}] {2} ({3})" -f $j, $pendientes.Count, $reg.NUMERO, $reg.ESTADO) -ForegroundColor White
        try {
            Stop-ACSession
            $ctx = Start-ACSession -Password $clave -BaseDatos $BaseDatos

            $r = Invoke-ACBusqueda -Numero $reg.NUMERO
            if (-not $r.Encontrado) {
                $reg.OBSERVACION = (("$($reg.OBSERVACION) No se pudo reabrir para el historial: $($r.Mensaje)").Trim())
                continue
            }

            [void](Open-ACFicha)
            $h = Get-ACHistorial

            $filas = @($h.EstadoContrato.Rows)
            if ($filas.Count -gt 0) {
                $ultima = $filas[0].Campos
                $reg.MOTIVO         = $ultima.MOTIVO
                $reg.FECHA_ESTADO   = $ultima.'VALIDO DESDE'
                $reg.USUARIO_ESTADO = $ultima.USUARIO
                $reg.HISTORIAL_COMPLETO = (($filas | ForEach-Object {
                    "$($_.Campos.ESTADO) / $($_.Campos.MOTIVO) / $($_.Campos.'VALIDO DESDE') / $($_.Campos.USUARIO)"
                }) -join ' || ')

                $estadoHist = "$($ultima.ESTADO)".Trim()
                if ($estadoHist -and $estadoHist.ToUpper() -ne "$($reg.ESTADO)".Trim().ToUpper()) {
                    $reg.OBSERVACION = (("$($reg.OBSERVACION) Ultimo movimiento del historial: '$estadoHist' (la grilla reporta '$($reg.ESTADO)').").Trim())
                }
                Write-Paso "Motivo: $($ultima.MOTIVO) ($($ultima.ESTADO), $($ultima.'VALIDO DESDE'))" "OK"
            } else {
                $reg.OBSERVACION = (("$($reg.OBSERVACION) El historial no devolvio movimientos.").Trim())
                Write-Paso "El historial no devolvio movimientos." "WARN"
            }

            [void](Save-WindowShot -Handle $h.VentanaHandle -Path (Join-Path $dirCapturas "historial_$($reg.NUMERO).png"))
        } catch {
            $reg.OBSERVACION = (("$($reg.OBSERVACION) Error al consultar el historial: $($_.Exception.Message)").Trim())
            Write-Paso "Error: $($_.Exception.Message)" "ERROR"
        }
    }
}

# ---------------------------------------------------------------------
#  Salida
# ---------------------------------------------------------------------

Write-Titulo "RESUMEN"

$activas    = @($resultados | Where-Object { $_.ACTIVA -eq 'SI' })
$noActivas  = @($resultados | Where-Object { $_.ACTIVA -eq 'NO' })
$conError   = @($resultados | Where-Object { $_.ACTIVA -eq 'REVISAR' })

Write-Host ("Activas        : {0}" -f $activas.Count)   -ForegroundColor Green
Write-Host ("No activas     : {0}" -f $noActivas.Count) -ForegroundColor Yellow
Write-Host ("Con error      : {0}" -f $conError.Count)  -ForegroundColor Red
Write-Host ("Total          : {0}" -f $resultados.Count)

if ($noActivas.Count -gt 0) {
    Write-Host ""
    Write-Host "Lineas que NO quedaron activas:" -ForegroundColor Yellow
    $noActivas | Select-Object NUMERO, NOMBRE, ESTADO, MOTIVO, FECHA_ESTADO | Format-Table -AutoSize | Out-String | Write-Host
}

$resultados | Export-Csv -LiteralPath $Salida -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Reporte: $Salida" -ForegroundColor Cyan
if (Test-Path -LiteralPath $dirCapturas) { Write-Host "Capturas: $dirCapturas" -ForegroundColor Cyan }

try { Stop-ACSession } catch { }
