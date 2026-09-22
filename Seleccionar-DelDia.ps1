<#
.SINOPSIS
    Decide que numeros hay que consultar hoy, en vez de consultarlos todos.

.DESCRIPCION
    La forma mas rapida de acortar la corrida no es consultar mas rapido, es
    consultar menos. Una venta que vence dentro de dos meses no cambia de un
    dia para otro, y una que ya paso sus tres meses no va a cambiar nunca: su
    resultado ya es definitivo.

    Por eso la frecuencia se ata a lo que falta para el vencimiento de la
    tercera cuota:

      ya vencio          no se consulta mas, el resultado es definitivo
      faltan 0-7 dias    todos los dias, es la ultima oportunidad de recuperar
      faltan 8-30 dias   cada 3 dias
      faltan 31+ dias    una vez por semana

    Medido sobre los tres meses cargados (5.026 ventas): 995 consultas al dia
    en vez de 5.026, un 80% menos.

    La ultima fecha de consulta sale de la columna CONSULTADO de los CSV de
    resultados, asi que el reparto se acomoda solo: lo que se consulto ayer no
    se vuelve a mirar hasta que le toque.

.EJEMPLO
    .\Seleccionar-DelDia.ps1 -Salida .\salida\hoy.txt
    .\Automatizar-Ventas.ps1 -Numeros (Get-Content .\salida\hoy.txt) -Instancias 5 -SoloEstado
#>

[CmdletBinding()]
param(
    [string[]] $Consolidados,
    [string]   $Salida,
    [int]      $DiasDiario  = 7,
    [int]      $DiasCada3   = 30,
    [datetime] $Hoy = (Get-Date).Date,
    [switch]   $SoloResumen
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$raiz\lib\Excel.ps1"

function Digitos { param($s) [regex]::Replace("$s", '[^0-9]', [string]::Empty) }

$dirSalida = Join-Path $raiz 'salida'
if (-not $Consolidados) {
    $Consolidados = @(Get-ChildItem -LiteralPath $dirSalida -Filter 'Consolidado *.xlsx' -ErrorAction SilentlyContinue |
                      ForEach-Object { $_.FullName })
}
if (-not $Consolidados) { throw "No se encontro ningun 'Consolidado *.xlsx' en $dirSalida" }
if (-not $Salida) { $Salida = Join-Path $dirSalida 'consultar_hoy.txt' }

# Ultima vez que se consulto cada numero, de todos los CSV de resultados
$ultima = @{}
foreach ($csv in (Get-ChildItem -LiteralPath $dirSalida -Filter 'ventas_*.csv' -ErrorAction SilentlyContinue)) {
    foreach ($r in (Import-Csv -LiteralPath $csv.FullName -Encoding UTF8)) {
        if (-not $r.NUMERO) { continue }
        $c = $null
        try { $c = [datetime]$r.CONSULTADO } catch { }
        if (-not $c) { continue }
        if (-not $ultima.ContainsKey($r.NUMERO) -or $c -gt $ultima[$r.NUMERO]) { $ultima[$r.NUMERO] = $c }
    }
}
Write-Host ("Numeros con consulta previa: {0}" -f $ultima.Count)

$base = [datetime]'1899-12-30'
$vistos = @{}
$seleccion = @()
$stats = [ordered]@{ 'vencida (no se consulta)'=0; 'diario'=0; 'cada 3 dias'=0; 'semanal'=0; 'sin fecha'=0 }
$alDia  = 0

foreach ($x in $Consolidados) {
    $c = Get-XlsxCeldas -Path $x -Hoja 'CONSOLIDADO'
    $ult = 0
    foreach ($k in $c.Keys) { if ($k -match '^[A-Z]{1,2}(\d+)$') { $f=[int]$Matches[1]; if ($f -gt $ult) { $ult=$f } } }
    Write-Host ("  {0,-34} {1,5} filas" -f (Split-Path -Leaf $x), ($ult-1))

    for ($f = 2; $f -le $ult; $f++) {
        $n = Digitos $c["E$f"]
        if ($n.Length -ne 10 -or $vistos.ContainsKey($n)) { continue }
        $vistos[$n] = $true

        $s = "$($c["B$f"])".Trim(); $v = 0
        if (-not ([double]::TryParse($s, [ref]$v) -and $v -gt 1000)) { $stats['sin fecha']++; $seleccion += $n; continue }
        $vence = $base.AddDays([Math]::Floor($v)).AddMonths(3)
        $faltan = [int](($vence - $Hoy).TotalDays)

        if ($faltan -lt 0) { $stats['vencida (no se consulta)']++; continue }

        $cada = if ($faltan -le $DiasDiario) { 1 } elseif ($faltan -le $DiasCada3) { 3 } else { 7 }
        $etiq = switch ($cada) { 1 { 'diario' } 3 { 'cada 3 dias' } default { 'semanal' } }
        $stats[$etiq]++

        $u = $ultima[$n]
        if ($u -and ($Hoy - $u.Date).TotalDays -lt $cada) { $alDia++; continue }
        $seleccion += $n
    }
}

Write-Host ""
Write-Host "=== REPARTO POR CADENCIA ===" -ForegroundColor Cyan
foreach ($k in $stats.Keys) { Write-Host ("  {0,-26} {1,5}" -f $k, $stats[$k]) }
Write-Host ""
Write-Host ("  Total en seguimiento      : {0}" -f $vistos.Count)
Write-Host ("  Ya consultadas al dia     : {0}" -f $alDia)
Write-Host ("  A CONSULTAR HOY           : {0}" -f $seleccion.Count) -ForegroundColor Yellow
if ($vistos.Count) {
    Write-Host ("  Ahorro frente a todo      : {0:N0}%" -f (100 - 100*$seleccion.Count/$vistos.Count))
}
# 3,54 s por numero medido con 5 instancias
Write-Host ("  Tiempo estimado           : {0:N0} min   (todo seria {1:N0} min)" -f `
            ($seleccion.Count*3.54/60), ($vistos.Count*3.54/60))

if (-not $SoloResumen) {
    Set-Content -LiteralPath $Salida -Value $seleccion -Encoding ASCII
    Write-Host ""
    Write-Host ("Lista guardada en: {0}" -f $Salida) -ForegroundColor Cyan
    Write-Host  "Para consultarlos:"
    Write-Host ("  .\Automatizar-Ventas.ps1 -Numeros (Get-Content '{0}') -Instancias 5 -SoloEstado" -f $Salida) -ForegroundColor Gray
}
