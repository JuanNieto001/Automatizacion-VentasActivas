<#
.SINOPSIS
    Arma el listado de lineas suspendidas por no pago, pensado para decidir si
    a la empresa le conviene pagar la deuda del cliente.

.DESCRIPCION
    Claro paga cada venta de portabilidad en tres cuotas, y solo si la linea
    sigue activa a los tres meses. Cuando una linea se suspende por falta de
    pago, la venta se pierde... salvo que alguien pague. De ahi la pregunta de
    negocio: cuesta menos pagar la deuda que perder la venta?

    Para responderla, cada fila trae lo que hace falta:

      CUANTO DEBE           lo que costaria salvarla, como NUMERO sumable
      DIAS PARA PERDERLA    cuanto falta para que cumpla los tres meses
      EN DEMANDA/CASTIGADA  si la deuda ya paso a cobro juridico o se dio por
                            perdida; en esos casos pagar no sirve
      CARGO FIJO            lo que factura el plan al mes, como referencia

    Se ordena por urgencia y, dentro de cada grupo, de menor a mayor deuda:
    arriba quedan las que se salvan mas barato y con menos tiempo.

.EJEMPLO
    .\Armar-SuspendidasParaPago.ps1
#>

[CmdletBinding()]
param(
    [string]   $Saldos,
    [string[]] $Consolidados,
    [string]   $Salida,
    [datetime] $Hoy = (Get-Date).Date
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$raiz\lib\Excel.ps1"
. "$raiz\lib\ExcelNuevo.ps1"
. "$raiz\lib\ExcelHojas.ps1"
function Digitos { param($s) [regex]::Replace("$s", '[^0-9]', [string]::Empty) }

$dirSalida = Join-Path $raiz 'salida'
if (-not $Saldos) {
    $Saldos = (Get-ChildItem -LiteralPath $dirSalida -Filter 'saldos_*.csv' -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}
if (-not $Saldos) { throw "No se encontro ningun saldos_*.csv en $dirSalida" }
if (-not $Consolidados) {
    $Consolidados = @(Get-ChildItem -LiteralPath $dirSalida -Filter 'Consolidado *.xlsx' -ErrorAction SilentlyContinue |
                      ForEach-Object { $_.FullName })
}
if (-not $Salida) { $Salida = Join-Path $dirSalida ("SUSPENDIDAS - decidir si pagar - " + $Hoy.ToString('yyyy-MM-dd') + ".xlsx") }

Write-Host ("Saldos: {0}" -f (Split-Path -Leaf $Saldos))

$saldo = @{}
foreach ($r in (Import-Csv -LiteralPath $Saldos -Encoding UTF8)) { $saldo[$r.NUMERO] = $r }
Write-Host ("  lineas con saldo consultado: {0}" -f $saldo.Count)

# datos de la venta, de los consolidados
$base = [datetime]'1899-12-30'
$venta = @{}
foreach ($x in $Consolidados) {
    $c = Get-XlsxCeldas -Path $x -Hoja 'CONSOLIDADO'
    $ult = 0
    foreach ($k in $c.Keys) { if ($k -match '^[A-Z]{1,2}(\d+)$') { $f=[int]$Matches[1]; if ($f -gt $ult) { $ult=$f } } }
    for ($f = 2; $f -le $ult; $f++) {
        $n = Digitos $c["E$f"]; if ($n.Length -ne 10) { continue }
        if (-not $saldo.ContainsKey($n)) { continue }
        $s = "$($c["B$f"])".Trim(); $d = 0
        if (-not ([double]::TryParse($s,[ref]$d) -and $d -gt 1000)) { continue }
        $vta = $base.AddDays([Math]::Floor($d))
        # si el numero esta en dos consolidados se queda con la venta mas vieja,
        # que es la que vence primero
        if ($venta.ContainsKey($n) -and $venta[$n].Venta -le $vta) { continue }
        $venta[$n] = [pscustomobject]@{
            Venta = $vta; Mes = [System.IO.Path]::GetFileNameWithoutExtension($x) -replace '^Consolidado ',''
            Cliente = "$($c["M$f"])".Trim(); Cedula = "$($c["K$f"])".Trim()
            Asesor = "$($c["D$f"])".Trim(); Ciudad = "$($c["I$f"])".Trim()
            Plan = "$($c["G$f"])".Trim(); CargoFijo = "$($c["T$f"])".Trim()
        }
    }
}
Write-Host ("  lineas cruzadas con su venta: {0}" -f $venta.Count)

$filas = @()
foreach ($n in $saldo.Keys) {
    $s = $saldo[$n]
    $v = $venta[$n]
    $vence = if ($v) { $v.Venta.AddMonths(3) } else { $null }
    $dias  = if ($vence) { [int](($vence - $Hoy).TotalDays) } else { $null }

    $debe = $null
    if ($s.DEBE) { $t = 0.0; if ([double]::TryParse($s.DEBE, [ref]$t)) { $debe = $t } }

    # Recomendacion: lo que decide es si todavia se puede salvar y cuanto cuesta.
    # El primer caso es el mejor y hay que distinguirlo: la linea ya volvio a
    # estar activa, o sea que el cliente pago solo y la venta se salvo sin que
    # la empresa gaste nada. Si no se separa, esas filas parecen pendientes.
    $activa = ($s.ESTADO -and $s.ESTADO.Trim().ToUpper() -eq 'ACTIVO')
    $rec = if ($activa -and $null -ne $debe -and $debe -le 0) {
               'RESUELTA: pago y la linea esta activa, no hay que hacer nada'
           }
           elseif ($activa) {
               # sigue debiendo pero todavia no la suspenden: no hay que pagarla
               # hoy, pero puede caerse en cualquier momento
               'ACTIVA CON DEUDA: vigilar, puede volver a suspenderse'
           }
           elseif ($s.EN_DEMANDA -eq 'SI') { 'NO PAGAR: la deuda esta en cobro juridico' }
           elseif ($s.CASTIGADA -eq 'SI') { 'NO PAGAR: deuda castigada, la venta no se recupera' }
           elseif ($null -eq $debe) { 'REVISAR A MANO: no se pudo leer el saldo en AC' }
           elseif ($debe -le 0) { 'NO DEBE NADA: revisar por que figura suspendida' }
           elseif ($null -eq $dias) { 'REVISAR A MANO: no se pudo calcular el vencimiento' }
           elseif ($dias -lt 0) { 'YA VENCIO: pagar ya no salva la cuota' }
           elseif ($dias -le 7) { 'DECIDIR HOY: vence esta semana' }
           elseif ($dias -le 30) { 'EVALUAR: hay tiempo pero se acerca' }
           else { 'HAY TIEMPO: vence en mas de un mes' }

    $orden = if ($rec -like 'DECIDIR HOY*') { 1 }
             elseif ($rec -like 'EVALUAR*') { 2 }
             elseif ($rec -like 'HAY TIEMPO*') { 3 }
             elseif ($rec -like 'NO DEBE*') { 4 }
             elseif ($rec -like 'REVISAR*') { 5 }
             elseif ($rec -like 'RESUELTA*') { 7 }
             else { 6 }

    $filas += [pscustomobject]@{
        'QUE HACER'            = $rec
        'CUANTO DEBE'          = $debe
        'DIAS PARA PERDERLA'   = $dias
        'CUMPLE 3 MESES EL'    = if ($vence) { $vence.ToString('yyyy-MM-dd') } else { '' }
        'NUMERO'               = $n
        'CLIENTE'              = if ($v) { $v.Cliente } else { $s.NOMBRE }
        'CEDULA'               = if ($v) { $v.Cedula } else { '' }
        'ESTADO EN AC'         = $s.ESTADO
        'EN DEMANDA'           = $s.EN_DEMANDA
        'CASTIGADA'            = $s.CASTIGADA
        'CARGO FIJO DEL PLAN'  = if ($v) { $v.CargoFijo } else { '' }
        'PLAN'                 = if ($v) { $v.Plan } else { '' }
        'FECHA VENTA'          = if ($v) { $v.Venta.ToString('yyyy-MM-dd') } else { '' }
        'MES'                  = if ($v) { $v.Mes } else { '' }
        'ASESOR'               = if ($v) { $v.Asesor } else { '' }
        'CIUDAD'               = if ($v) { $v.Ciudad } else { '' }
        'SALDO SERVICIOS'      = $s.SALDO_SERVICIOS
        'SALDO EQUIPO'         = $s.SALDO_EQUIPO
        'CONSULTADO EN AC'     = $s.CONSULTADO
        'VERIFICADO A MANO'    = ''
        'DECISION (PAGAR/NO)'  = ''
        'OBSERVACION'          = ''
        '_orden'               = $orden
    }
}

# urgencia primero y, dentro de cada grupo, de mas barata a mas cara
$filas = @($filas | Sort-Object _orden, 'CUANTO DEBE')

Write-Host ""
Write-Host "=== REPARTO ===" -ForegroundColor Cyan
$filas | Group-Object 'QUE HACER' | Sort-Object { ($_.Group | Select-Object -First 1)._orden } | ForEach-Object {
    $t = ($_.Group | Where-Object { $null -ne $_.'CUANTO DEBE' } | Measure-Object 'CUANTO DEBE' -Sum).Sum
    Write-Host ("  {0,-48} {1,4} lineas   `${2,12:N0}" -f $_.Name, $_.Count, $t)
}
$tot = ($filas | Where-Object { $null -ne $_.'CUANTO DEBE' } | Measure-Object 'CUANTO DEBE' -Sum).Sum
$resueltas = @($filas | Where-Object { $_.'QUE HACER' -like 'RESUELTA*' })
if ($resueltas.Count) {
    Write-Host ""
    Write-Host ("  {0} se resolvieron solas: el cliente pago y la linea volvio a estar activa." -f $resueltas.Count) -ForegroundColor Green
}
$salvables = @($filas | Where-Object { $_.'QUE HACER' -like 'DECIDIR HOY*' -or $_.'QUE HACER' -like 'EVALUAR*' -or $_.'QUE HACER' -like 'HAY TIEMPO*' })
$totSalv = ($salvables | Where-Object { $null -ne $_.'CUANTO DEBE' } | Measure-Object 'CUANTO DEBE' -Sum).Sum
Write-Host ""
Write-Host ("  DEUDA TOTAL de todas          : `${0:N0}" -f $tot)
Write-Host ("  Ventas que AUN se pueden salvar: {0}   cuestan `${1:N0}" -f $salvables.Count, $totSalv) -ForegroundColor Yellow
if ($salvables.Count) { Write-Host ("  Promedio por venta salvable    : `${0:N0}" -f ($totSalv/$salvables.Count)) }

$cols = @('QUE HACER','CUANTO DEBE','DIAS PARA PERDERLA','CUMPLE 3 MESES EL','NUMERO','CLIENTE','CEDULA',
          'ESTADO EN AC','EN DEMANDA','CASTIGADA','CARGO FIJO DEL PLAN','PLAN','FECHA VENTA','MES',
          'ASESOR','CIUDAD','SALDO SERVICIOS','SALDO EQUIPO','CONSULTADO EN AC',
          'VERIFICADO A MANO','DECISION (PAGAR/NO)','OBSERVACION')
$numericas = @('CUANTO DEBE','DIAS PARA PERDERLA','CARGO FIJO DEL PLAN')

$r = New-XlsxSimple -Filas $filas -Columnas $cols -Destino $Salida -Hoja 'SUSPENDIDAS' -Numericas $numericas
Write-Host ""
Write-Host ("Archivo: {0}" -f $r.Archivo) -ForegroundColor Cyan
Write-Host ("  {0} filas, {1} columnas" -f $r.Filas, $r.Columnas)

