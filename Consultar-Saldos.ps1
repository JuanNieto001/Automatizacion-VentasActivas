<#
.SINOPSIS
    Para las lineas suspendidas, averigua cuanto debe el cliente.

.DESCRIPCION
    Una linea suspendida casi siempre lo esta por falta de pago, y a diferencia
    de una desactivada por portabilidad SI se recupera: si el cliente se pone
    al dia antes de que la venta cumpla sus tres meses, la tercera cuota se
    salva. Para poder gestionarlo hay que saber cuanto debe.

    La ficha del cliente tiene un boton "Saldo" que abre tres importes: Saldo
    Equipo, Saldo Servicios y Saldo Total. Ademas se leen cuatro casillas de
    cobranza; "En Demanda" y "Castigada" avisan que la deuda ya paso a cobro
    juridico o que el operador la dio por perdida, y en esos casos no vale la
    pena que un asesor llame.

    Va de a un numero por vez: abrir la ficha obliga a reiniciar AC despues,
    porque cerrarla exigiria guardar un Tickler y eso escribiria en produccion.

.EJEMPLO
    .\Consultar-Saldos.ps1 -Numeros 3044760292
    .\Consultar-Saldos.ps1 -DesdeCsv .\salida\ventas_JULIO.csv
#>

[CmdletBinding()]
param(
    [string[]] $Numeros,
    [string[]] $DesdeCsv,
    [string]   $Salida,
    [string]   $Password,
    [string]   $BaseDatos = 'AC_PRODUCCION',
    [int]      $Limite = 0
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$raiz\lib\Win32.ps1"
. "$raiz\lib\Remote.ps1"
. "$raiz\lib\AC.ps1"

$dirSalida = Join-Path $raiz 'salida'
if (-not $Salida) { $Salida = Join-Path $dirSalida ("saldos_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv") }

# --- numeros a consultar: los suspendidos de los CSV, o los que se indiquen ---
$lista = @()
if ($Numeros) { $lista = @($Numeros) }
else {
    if (-not $DesdeCsv) {
        $DesdeCsv = @(Get-ChildItem -LiteralPath $dirSalida -Filter 'ventas_*.csv' -ErrorAction SilentlyContinue |
                      ForEach-Object { $_.FullName })
    }
    $vistos = @{}
    foreach ($f in $DesdeCsv) {
        foreach ($r in (Import-Csv -LiteralPath $f -Encoding UTF8)) {
            if ($r.ESTADO -match 'suspension' -and -not $vistos.ContainsKey($r.NUMERO)) {
                $vistos[$r.NUMERO] = $true
                $lista += $r.NUMERO
            }
        }
    }
}
if ($Limite -gt 0 -and $lista.Count -gt $Limite) { $lista = @($lista | Select-Object -First $Limite) }
if (-not $lista.Count) { throw "No hay numeros que consultar." }

Write-Host ("Lineas por consultar: {0}" -f $lista.Count)
Write-Host "Cada una obliga a reiniciar AC; calcula ~30 s por linea." -ForegroundColor Yellow
Write-Host "Usa el raton: no toques el equipo mientras corre." -ForegroundColor Yellow
Write-Host ""

# --- clave ---
$clave = $Password
if (-not $clave) {
    $arch = Join-Path $raiz 'config\clave.dat'
    if (-not (Test-Path -LiteralPath $arch)) { throw "No hay clave guardada. Usa -Password o Automatizar-Ventas.ps1 -GuardarClave" }
    $sec = Get-Content -LiteralPath $arch | ConvertTo-SecureString
    $b = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    $clave = [Runtime.InteropServices.Marshal]::PtrToStringAuto($b)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b)
}

$resultados = New-Object System.Collections.ArrayList
$sw = [Diagnostics.Stopwatch]::StartNew()
$i = 0

foreach ($n in $lista) {
    $i++
    $reg = [ordered]@{
        NUMERO = $n; NOMBRE = ''; ESTADO = ''
        SALDO_EQUIPO = ''; SALDO_SERVICIOS = ''; SALDO_TOTAL = ''; DEBE = ''
        RESPON_PAGO = ''; FLAG_NO_COBRAR = ''; EN_DEMANDA = ''; CASTIGADA = ''
        OBSERVACION = ''; CONSULTADO = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
    try {
        Stop-ACSession -Todas
        $null = Start-ACSession -Password $clave -BaseDatos $BaseDatos

        $r = $null
        for ($t = 1; $t -le 3; $t++) { $r = Invoke-ACBusqueda -Numero $n; if ($r.Encontrado) { break }; Start-Sleep 2 }
        if (-not $r.Encontrado) { $reg.OBSERVACION = "No se encontro en AC: $($r.Mensaje)" }
        else {
            if ($r.Filas -and $r.Filas.Count) {
                $c = $r.Filas[0].Campos
                if ($c) { $reg.NOMBRE = $c.NOMBRE; $reg.ESTADO = $c.ESTADO }
            }
            $null = Open-ACFicha
            Start-Sleep -Seconds 2

            try {
                $ind = Get-ACIndicadoresCobranza
                $reg.RESPON_PAGO    = $ind.'Respon. Pago'
                $reg.FLAG_NO_COBRAR = $ind.'Flag No Cobrar'
                $reg.EN_DEMANDA     = $ind.'En Demanda'
                $reg.CASTIGADA      = $ind.'Castigada'
            } catch { $reg.OBSERVACION = ("$($reg.OBSERVACION) indicadores: $($_.Exception.Message)").Trim() }

            $s = Get-ACSaldo
            $reg.SALDO_EQUIPO    = $s.SaldoEquipo
            $reg.SALDO_SERVICIOS = $s.SaldoServicios
            $reg.SALDO_TOTAL     = $s.SaldoTotal
            $reg.DEBE            = if ($null -ne $s.TotalNum) { [string]$s.TotalNum } else { '' }
        }
    } catch {
        $reg.OBSERVACION = ("$($reg.OBSERVACION) $($_.Exception.Message)").Trim()
    }

    [void]$resultados.Add([pscustomobject]$reg)
    $col = if ($reg.DEBE) { 'Yellow' } else { 'Gray' }
    Write-Host ("  [{0}/{1}] {2}  {3}  debe {4}  {5}" -f $i, $lista.Count, $n, $reg.ESTADO, `
                $(if ($reg.SALDO_TOTAL) { $reg.SALDO_TOTAL } else { '?' }), `
                $(if ($reg.EN_DEMANDA -eq 'SI') { '[EN DEMANDA]' } elseif ($reg.CASTIGADA -eq 'SI') { '[CASTIGADA]' } else { '' })) -ForegroundColor $col

    if (($i % 10) -eq 0) { $resultados | Export-Csv -LiteralPath $Salida -NoTypeInformation -Encoding UTF8 }
}

Stop-ACSession -Todas
$sw.Stop()
$resultados | Export-Csv -LiteralPath $Salida -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("{0} lineas en {1:N0} s  ({2:N1} s por linea)" -f $resultados.Count, $sw.Elapsed.TotalSeconds, `
            ($sw.Elapsed.TotalSeconds / [Math]::Max(1,$resultados.Count))) -ForegroundColor Cyan
$con = @($resultados | Where-Object { $_.DEBE })
if ($con.Count) {
    $suma = ($con | ForEach-Object { [double]$_.DEBE } | Measure-Object -Sum).Sum
    Write-Host ("Con saldo leido: {0}   deuda total: `${1:N0}   promedio: `${2:N0}" -f `
                $con.Count, $suma, ($suma/$con.Count))
}
Write-Host ("CSV: {0}" -f $Salida) -ForegroundColor Cyan
