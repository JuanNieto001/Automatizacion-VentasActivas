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
    # Una venta con menos de estos dias que no aparece activa NO se da por
    # caida: la portabilidad tarda dias habiles en aprovisionarse, asi que
    # todavia no se sabe. Esas van a la hoja EN TRAMITE, aparte de las que de
    # verdad se cayeron. Medido el 05-oct: a 0-4 dias solo el 40% figura activa,
    # a 5-7 dias el 87%, a mas de 15 dias el 97%.
    [int]$DiasTramite = 7,
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
$noActivas = @(); $noEncontradas = @(); $cuota3 = @(); $enTramite = @(); $sinResultado = 0

for ($f = 2; $f -le $ultima; $f++) {
    $n = Digitos $c["$Columna$f"]; if ($n.Length -ne 10) { continue }
    $r = $res[$n]
    if (-not $r) { $sinResultado++; continue }

    $s = "$($c["$ColumnaFecha$f"])".Trim(); $x = 0
    $vta   = if ([double]::TryParse($s, [ref]$x) -and $x -gt 1000) { $base.AddDays([Math]::Floor($x)) } else { $null }
    $vence = if ($vta) { $vta.AddMonths(3) } else { $null }
    $af    = if ($ColumnaDime) { "$($c["$ColumnaDime$f"])".Trim().ToUpper() } else { '' }
    $activa = ($r.ENCONTRADO -eq 'SI' -and $r.ACTIVA -eq 'SI')
    # "Fresca" se mide contra la fecha en que SE CONSULTO, no contra hoy: una
    # venta que estaba en tramite cuando la miramos sigue siendo "no se sabe",
    # aunque hoy ya tenga dias. Si el CSV no trae CONSULTADO, se usa Hoy.
    $refFecha = $Hoy
    if ($r.PSObject.Properties.Name -contains 'CONSULTADO' -and $r.CONSULTADO) {
        try { $refFecha = ([datetime]$r.CONSULTADO).Date } catch { }
    }
    $diasVenta = if ($vta) { [int](($refFecha - $vta).TotalDays) } else { $null }
    $fresca = ($null -ne $diasVenta -and $diasVenta -le $DiasTramite)

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
    # Una venta fresca no activa no esta "en riesgo": todavia no se sabe, puede
    # aparecer en cualquier momento. Se marca EN TRAMITE para no asustar al
    # tablero con ventas que van bien.
    $estadoCuota = if (-not $vence) { 'SIN FECHA' }
                   elseif ($activa) { if ($vence -le $Hoy) { 'COBRADA' } else { 'EN CAMINO' } }
                   elseif ($fresca) { 'EN TRAMITE' }
                   elseif ($vence -le $Hoy) { 'PERDIDA' }
                   else { 'EN RIESGO' }
    $o = [ordered]@{ 'ESTADO 3a CUOTA' = $estadoCuota; 'CUANDO CUMPLE' = $cuando
                     'ACTIVA HOY' = $(if ($activa) { 'SI' } else { 'NO' }) }
    foreach ($k in $comun.Keys) { $o[$k] = $comun[$k] }
    $cuota3 += [pscustomobject]$o

    if ($activa) { continue }

    # --- EN TRAMITE ---
    # Una venta reciente que no aparece activa NO esta caida: la portabilidad
    # tarda dias habiles, asi que todavia no se sabe. Se saca de NO ACTIVAS y de
    # NO ENCONTRADAS, donde parecia una venta perdida, y se deja aqui con lo
    # unico cierto: hay que volver a consultarla.
    if ($fresca) {
        $o = [ordered]@{
            'ESTADO HOY' = if ($r.ENCONTRADO -ne 'SI') { 'todavia no aparece en AC' } else { "en AC: $($r.ESTADO)" }
            'QUE HACER'  = "Venta de hace $diasVenta dias: aun en tramite, volver a consultar en unos dias"
        }
        foreach ($k in $comun.Keys) { $o[$k] = $comun[$k] }
        $enTramite += [pscustomobject]$o
        continue
    }

    # --- NO ENCONTRADAS (ya con dias suficientes y aun no existe) ---
    if ($r.ENCONTRADO -ne 'SI') {
        $o = [ordered]@{ 'QUE HACER' = $(
            if ($af -eq 'EXITOSO') {
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

    # --- NO ACTIVAS (existe en AC pero caida, con dias suficientes) ---
    $vencida = ($vence -and $vence -le $Hoy)
    $suspend = ($r.ESTADO -match 'suspension')
    $urg = if ($vencida) { '1 - YA VENCIO: cuota perdida' }
           elseif ($vence -and $vence -le $Hoy.AddDays(30)) { '2 - VENCE EN 30 DIAS' }
           else { '3 - Vence mas adelante' }
    $que = if ($suspend) {
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
$enTramite     = @($enTramite     | Sort-Object 'FECHA VENTA' -Descending)
$cuota3        = @($cuota3        | Sort-Object 'CUANDO CUMPLE', 'ESTADO 3a CUOTA', 'FECHA VENTA')

Write-Host ""
Write-Host ("  NO ACTIVAS (de verdad caidas) : {0}" -f $noActivas.Count)
Write-Host ("  NO ENCONTRADAS (con dias)     : {0}" -f $noEncontradas.Count)
Write-Host ("  EN TRAMITE (no se sabe aun)   : {0}" -f $enTramite.Count)
Write-Host ("  CUOTA 3                       : {0}" -f $cuota3.Count)
$cuota3 | Group-Object 'ESTADO 3a CUOTA' | Sort-Object Name | ForEach-Object {
    Write-Host ("      {0,-12} {1,5}" -f $_.Name, $_.Count)
}

$hojas = @()
if ($noActivas.Count)     { $hojas += @{ Nombre='NO ACTIVAS';     Filas=$noActivas;     Columnas=@($noActivas[0].PSObject.Properties.Name) } }
if ($noEncontradas.Count) { $hojas += @{ Nombre='NO ENCONTRADAS'; Filas=$noEncontradas; Columnas=@($noEncontradas[0].PSObject.Properties.Name) } }
if ($enTramite.Count)     { $hojas += @{ Nombre='EN TRAMITE';     Filas=$enTramite;     Columnas=@($enTramite[0].PSObject.Properties.Name) } }
if ($cuota3.Count)        { $hojas += @{ Nombre='CUOTA 3';        Filas=$cuota3;        Columnas=@($cuota3[0].PSObject.Properties.Name) } }
if (-not $hojas.Count) { throw "No hay nada que agregar." }

# El entregable sale siempre con la misma estructura: la hoja de datos
# (CONSOLIDADO) mas las pestanas que agregamos. Se quitan las hojas de trabajo
# que arrastra cada consolidado de origen (RECHAZOS, ESTADO, GRABADO, EXITOSO,
# PLANES, Hoja1, etc.), que cambian de un mes a otro y ensucian el archivo.
$limpio = Join-Path $env:TEMP ("consolidado_limpio_" + [Guid]::NewGuid().ToString('N') + '.xlsx')
$sel = Select-HojasXlsx -Archivo $Verificado -Conservar @($Hoja) -Destino $limpio
if ($sel.Quitadas.Count) { Write-Host ("Se quitan del entregable las hojas de trabajo: {0}" -f ($sel.Quitadas -join ', ')) -ForegroundColor Gray }

$r2 = Add-HojasAXlsx -Origen $limpio -Destino $Salida -Hojas $hojas
Remove-Item -LiteralPath $limpio -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host ("Entregable: {0}" -f $r2.Archivo) -ForegroundColor Cyan
Write-Host ("  hojas agregadas: {0}" -f ($r2.Agregadas -join ', '))
if ($r2.Reemplazadas) { Write-Host ("  hojas reemplazadas: {0}" -f ($r2.Reemplazadas -join ', ')) }
