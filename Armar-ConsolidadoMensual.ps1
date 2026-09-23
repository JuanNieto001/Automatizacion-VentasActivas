<#
.SINOPSIS
    Arma el entregable unico del mes: un solo .xlsx con el consolidado
    verificado mas las pestanas que alimentan el tablero.

.DESCRIPCION
    Toma el consolidado ya verificado por Automatizar-Ventas.ps1 y le agrega
    tres hojas, sin tocar las que traia el archivo original:

      NO ACTIVAS       lineas que existen en AC pero se cayeron, con que hacer
                       con cada una y cuanta urgencia hay segun su vencimiento
      NO ENCONTRADAS   numeros que no aparecen en AC: revision distinta, hay
                       que averiguar si la venta llego a existir
      CUOTA 3          seguimiento del pago: a Claro le paga cada venta en tres
                       cuotas, asi que la linea tiene que seguir activa tres
                       meses despues de vendida

    Van TODAS las ventas en la hoja CUOTA 3, no solo las que ya cumplieron los
    tres meses: lo que el tablero necesita ver es cuanto esta por definirse y
    cuanto esta en riesgo.

.EJEMPLO
    .\Armar-ConsolidadoMensual.ps1 -Mes JULIO `
        -Verificado ".\salida\Consolidado de ventas portabilidado JULIO - VERIFICADO 2026-09-22.xlsx" `
        -Resultados ".\salida\ventas_JULIO.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Mes,
    [Parameter(Mandatory=$true)][string]$Verificado,
    [Parameter(Mandatory=$true)][string]$Resultados,
    [string]$Hoja = 'CONSOLIDADO',
    [string]$Columna = 'E',
    [string]$ColumnaFecha = 'B',
    # Columna con el ESTADO DIME. No todos los consolidados la traen: el de
    # septiembre del 23-sep, por ejemplo, usa la AF para "VENTANA DE CAMBIO".
    # Vacio = el archivo no la tiene y la comparacion con DIME se omite.
    [string]$ColumnaDime = 'AF',
    [string]$ColumnaAsesor = 'D',
    [string]$ColumnaCcAsesor = 'C',
    [string]$Salida,
    [datetime]$Hoy = (Get-Date).Date
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$raiz\lib\Excel.ps1"
. "$raiz\lib\ExcelHojas.ps1"

function Digitos { param($s) [regex]::Replace("$s", '[^0-9]', [string]::Empty) }

if (-not [System.IO.Path]::IsPathRooted($Verificado)) { $Verificado = Join-Path $raiz $Verificado }
if (-not [System.IO.Path]::IsPathRooted($Resultados)) { $Resultados = Join-Path $raiz $Resultados }
if (-not $Salida) { $Salida = Join-Path (Join-Path $raiz 'salida') "Consolidado $Mes.xlsx" }

Write-Host "Origen : $Verificado"
Write-Host "Datos  : $Resultados"
Write-Host ""

$res = @{}
foreach ($r in (Import-Csv -LiteralPath $Resultados -Encoding UTF8)) { $res[$r.NUMERO] = $r }
Write-Host ("Resultados de verificacion: {0}" -f $res.Count)

$c = Get-XlsxCeldas -Path $Verificado -Hoja $Hoja
$ultima = 0
foreach ($k in $c.Keys) { if ($k -match '^[A-Z]{1,2}(\d+)$') { $f=[int]$Matches[1]; if ($f -gt $ultima) { $ultima = $f } } }
Write-Host ("Filas en la hoja $Hoja : {0}" -f ($ultima - 1))

$base = [datetime]'1899-12-30'
$noActivas = @(); $noEncontradas = @(); $cuota3 = @(); $sinResultado = 0

for ($f = 2; $f -le $ultima; $f++) {
    $n = Digitos $c["$Columna$f"]; if ($n.Length -ne 10) { continue }
    $r = $res[$n]
    if (-not $r) { $sinResultado++; continue }

    $s = "$($c["$ColumnaFecha$f"])".Trim(); $x = 0
    $vta   = if ([double]::TryParse($s, [ref]$x) -and $x -gt 1000) { $base.AddDays([Math]::Floor($x)) } else { $null }
    $vence = if ($vta) { $vta.AddMonths(3) } else { $null }
    $af    = if ($ColumnaDime) { "$($c["$ColumnaDime$f"])".Trim().ToUpper() } else { '' }
    $activa = ($r.ENCONTRADO -eq 'SI' -and $r.ACTIVA -eq 'SI')

    $comun = [ordered]@{
        'NUMERO'            = $n
        'FECHA VENTA'       = if ($vta)   { $vta.ToString('yyyy-MM-dd') }   else { '' }
        'CUMPLE 3 MESES'    = if ($vence) { $vence.ToString('yyyy-MM-dd') } else { '' }
        'DIAS PARA 3 MESES' = if ($vence) { [string][int](($vence - $Hoy).TotalDays) } else { '' }
        'ESTADO AC'         = if ($r.ESTADO) { $r.ESTADO } else { 'no existe en AC' }
        'ESTADO DIME'       = if ($ColumnaDime) { $af } else { '(no viene en el archivo)' }
        'ASESOR'            = "$($c["$ColumnaAsesor$f"])".Trim()
        'CC ASESOR'         = "$($c["$ColumnaCcAsesor$f"])".Trim()
        'TEAM LEADER'       = "$($c["Z$f"])".Trim()
        'CLIENTE'           = "$($c["M$f"])".Trim()
        'CEDULA CLIENTE'    = "$($c["K$f"])".Trim()
        'CIUDAD'            = "$($c["I$f"])".Trim()
        'DEPARTAMENTO'      = "$($c["J$f"])".Trim()
        'PLAN'              = "$($c["G$f"])".Trim()
        'CARGO FIJO'        = "$($c["T$f"])".Trim()
        'FILA CONSOLIDADO'  = [string]$f
        'VERIFICADO EL'     = $Hoy.ToString('yyyy-MM-dd')
    }

    # --- CUOTA 3: van todas, con el estado del pago ---
    $cuando = if (-not $vence) { 'sin fecha de venta' }
              elseif ($vence -le $Hoy) { '1 - ya cumplio' }
              elseif ($vence -le $Hoy.AddDays(30)) { '2 - cumple en 30 dias o menos' }
              else { '3 - cumple mas adelante' }
    $estadoCuota = if (-not $vence) { 'SIN FECHA' }
                   elseif ($vence -le $Hoy) { if ($activa) { 'COBRADA' } else { 'PERDIDA' } }
                   else { if ($activa) { 'EN CAMINO' } else { 'EN RIESGO' } }
    $o = [ordered]@{ 'ESTADO 3a CUOTA' = $estadoCuota; 'CUANDO CUMPLE' = $cuando
                     'ACTIVA HOY' = $(if ($activa) { 'SI' } else { 'NO' }) }
    foreach ($k in $comun.Keys) { $o[$k] = $comun[$k] }
    $cuota3 += [pscustomobject]$o

    if ($activa) { continue }

    # --- NO ENCONTRADAS ---
    # Una venta reciente que no aparece casi nunca esta caida: la portabilidad
    # tarda dias habiles en aprovisionarse. Medido en este mismo consolidado el
    # 23-sep: a 0-3 dias solo el 24% figura activa, a mas de 15 dias el 97%.
    $diasVenta = if ($vta) { [int](($Hoy - $vta).TotalDays) } else { $null }
    if ($r.ENCONTRADO -ne 'SI') {
        $o = [ordered]@{ 'QUE HACER' = $(
            if ($null -ne $diasVenta -and $diasVenta -le 7) {
                "Venta de hace $diasVenta dias: casi seguro sigue en tramite, volver a consultar"
            } elseif ($af -eq 'EXITOSO') {
                'DIME la da por exitosa pero no existe en AC: buscar la orden de portabilidad'
            } elseif ($af) {
                "DIME ya la marcaba $af : confirmar que no se activo"
            } else {
                'No existe en AC pese a tener dias: verificar si la venta se concreto'
            }) }
        foreach ($k in $comun.Keys) { $o[$k] = $comun[$k] }
        $noEncontradas += [pscustomobject]$o
        continue
    }

    # --- NO ACTIVAS ---
    $vencida = ($vence -and $vence -le $Hoy)
    $suspend = ($r.ESTADO -match 'suspension')
    $urg = if ($null -ne $diasVenta -and $diasVenta -le 7) { '0 - VENTA RECIENTE: aun en tramite, no perseguir' }
           elseif ($vencida) { '1 - YA VENCIO: cuota perdida' }
           elseif ($vence -and $vence -le $Hoy.AddDays(30)) { '2 - VENCE EN 30 DIAS' }
           else { '3 - Vence mas adelante' }
    $que = if ($null -ne $diasVenta -and $diasVenta -le 7) {
               "Venta de hace $diasVenta dias: la portabilidad tarda dias habiles, esperar"
           } elseif ($suspend) {
               if ($vencida) { 'Suspendida y vencida: la cuota ya no se recupera' }
               else { 'Suspendida: si el cliente se pone al dia antes del vencimiento, se salva la cuota' }
           } else {
               if ($vencida) { 'Desactivada y vencida: analizar la causa' }
               else { 'Contactar al cliente antes de la fecha de vencimiento' }
           }
    $o = [ordered]@{ 'URGENCIA' = $urg; 'QUE HACER' = $que }
    foreach ($k in $comun.Keys) { $o[$k] = $comun[$k] }
    $noActivas += [pscustomobject]$o
}

if ($sinResultado -gt 0) {
    Write-Host ("AVISO: {0} filas del Excel no tienen resultado de verificacion." -f $sinResultado) -ForegroundColor Yellow
}

$noActivas     = @($noActivas     | Sort-Object URGENCIA, 'CUMPLE 3 MESES')
$noEncontradas = @($noEncontradas | Sort-Object 'CUMPLE 3 MESES')
$cuota3        = @($cuota3        | Sort-Object 'CUANDO CUMPLE', 'ESTADO 3a CUOTA', 'FECHA VENTA')

Write-Host ""
Write-Host ("  NO ACTIVAS     : {0}" -f $noActivas.Count)
Write-Host ("  NO ENCONTRADAS : {0}" -f $noEncontradas.Count)
Write-Host ("  CUOTA 3        : {0}" -f $cuota3.Count)
$cuota3 | Group-Object 'ESTADO 3a CUOTA' | Sort-Object Name | ForEach-Object {
    Write-Host ("      {0,-12} {1,5}" -f $_.Name, $_.Count)
}

$hojas = @()
if ($noActivas.Count)     { $hojas += @{ Nombre='NO ACTIVAS';     Filas=$noActivas;     Columnas=@($noActivas[0].PSObject.Properties.Name) } }
if ($noEncontradas.Count) { $hojas += @{ Nombre='NO ENCONTRADAS'; Filas=$noEncontradas; Columnas=@($noEncontradas[0].PSObject.Properties.Name) } }
if ($cuota3.Count)        { $hojas += @{ Nombre='CUOTA 3';        Filas=$cuota3;        Columnas=@($cuota3[0].PSObject.Properties.Name) } }
if (-not $hojas.Count) { throw "No hay nada que agregar." }

$r2 = Add-HojasAXlsx -Origen $Verificado -Destino $Salida -Hojas $hojas
Write-Host ""
Write-Host ("Entregable: {0}" -f $r2.Archivo) -ForegroundColor Cyan
Write-Host ("  hojas agregadas: {0}" -f ($r2.Agregadas -join ', '))
if ($r2.Reemplazadas) { Write-Host ("  hojas reemplazadas: {0}" -f ($r2.Reemplazadas -join ', ')) }
