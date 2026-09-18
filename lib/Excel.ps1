# =====================================================================
#  Excel.ps1 - Lectura de la columna de numeros desde el archivo de ventas
#
#  En este equipo NO hay Microsoft Excel instalado, asi que el .xlsx se lee
#  directamente: un xlsx es un ZIP con XML dentro. Tambien acepta .csv y .txt.
# =====================================================================

function Get-XlsxCeldas {
    <#  Devuelve una tabla hash "E2" -> valor de la primera hoja del libro.  #>
    param([Parameter(Mandatory=$true)][string]$Path, [string]$Hoja)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tmp = Join-Path $env:TEMP ("xlsx_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $tmp)

        # --- cadenas compartidas ---
        $shared = @()
        $ssPath = Join-Path $tmp "xl\sharedStrings.xml"
        if (Test-Path -LiteralPath $ssPath) {
            [xml]$ss = Get-Content -LiteralPath $ssPath -Encoding UTF8
            foreach ($si in $ss.sst.si) {
                if ($si.t -is [string])      { $shared += $si.t }
                elseif ($si.t.'#text')       { $shared += $si.t.'#text' }
                elseif ($si.r)               { $shared += (($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '') }
                else                         { $shared += '' }
            }
        }

        # --- hoja a leer ---
        $hojas = Get-ChildItem -LiteralPath (Join-Path $tmp "xl\worksheets") -Filter *.xml -ErrorAction SilentlyContinue |
                 Sort-Object Name
        if (-not $hojas) { throw "El archivo no contiene hojas de calculo." }

        $archivoHoja = $hojas[0].FullName
        if ($Hoja) {
            $wbPath = Join-Path $tmp "xl\workbook.xml"
            if (Test-Path -LiteralPath $wbPath) {
                [xml]$wb = Get-Content -LiteralPath $wbPath -Encoding UTF8
                $i = 0
                foreach ($s in $wb.workbook.sheets.sheet) {
                    if ($s.name -eq $Hoja -and $i -lt $hojas.Count) { $archivoHoja = $hojas[$i].FullName }
                    $i++
                }
            }
        }

        [xml]$sh = Get-Content -LiteralPath $archivoHoja -Encoding UTF8
        $celdas = @{}
        foreach ($fila in $sh.worksheet.sheetData.row) {
            foreach ($c in $fila.c) {
                $ref = $c.r
                if (-not $ref) { continue }
                $val = $null
                if ($c.t -eq 's') {
                    $idx = [int]$c.v
                    if ($idx -ge 0 -and $idx -lt $shared.Count) { $val = $shared[$idx] }
                } elseif ($c.t -eq 'inlineStr') {
                    $val = $c.is.t
                    if ($val -isnot [string]) { $val = $c.is.t.'#text' }
                } elseif ($null -ne $c.v) {
                    $val = $c.v
                    if ($val -isnot [string]) { $val = $c.v.'#text' }
                }
                if ($null -ne $val) { $celdas[$ref] = [string]$val }
            }
        }
        return $celdas
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}

function ConvertTo-NumeroLimpio {
    <#  Normaliza a 10 digitos: quita espacios, guiones, indicativos y notacion
        cientifica que Excel a veces aplica a numeros largos.  #>
    param([string]$Valor)
    if (-not $Valor) { return $null }
    $v = $Valor.Trim()
    if ($v -match '^[0-9.]+E\+?[0-9]+$') {
        try { $v = ([decimal]$v).ToString("F0") } catch { }
    }
    $v = ($v -replace '[^0-9]', '')
    if ($v.Length -eq 12 -and $v.StartsWith('57')) { $v = $v.Substring(2) }   # indicativo pais
    if ($v.Length -eq 11 -and $v.StartsWith('0'))  { $v = $v.Substring(1) }
    if ($v.Length -ne 10) { return $null }
    return $v
}

function Get-NumerosDesdeArchivo {
    <#  Extrae los numeros de la columna indicada (por defecto E, "NUMERO ACTIVAR").  #>
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [string]$Columna = 'E',
        [int]$FilaInicio = 2,
        [string]$Hoja
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "No existe el archivo: $Path" }
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    $crudos = @()

    switch ($ext) {
        '.xlsx' {
            $celdas = Get-XlsxCeldas -Path $Path -Hoja $Hoja
            $filas = $celdas.Keys |
                     Where-Object { $_ -match "^$Columna(\d+)$" } |
                     ForEach-Object { [int]($_ -replace "^$Columna", '') } |
                     Sort-Object
            foreach ($f in $filas) {
                if ($f -lt $FilaInicio) { continue }
                $crudos += $celdas["$Columna$f"]
            }
        }
        '.xlsm' {
            $celdas = Get-XlsxCeldas -Path $Path -Hoja $Hoja
            $filas = $celdas.Keys |
                     Where-Object { $_ -match "^$Columna(\d+)$" } |
                     ForEach-Object { [int]($_ -replace "^$Columna", '') } |
                     Sort-Object
            foreach ($f in $filas) {
                if ($f -lt $FilaInicio) { continue }
                $crudos += $celdas["$Columna$f"]
            }
        }
        '.csv' {
            $idx = [int][char]$Columna.ToUpper()[0] - 65
            $lineas = Get-Content -LiteralPath $Path -Encoding UTF8
            for ($i = $FilaInicio - 1; $i -lt $lineas.Count; $i++) {
                $sep = if ($lineas[$i] -match ';') { ';' } else { ',' }
                $partes = $lineas[$i].Split($sep)
                if ($idx -lt $partes.Count) { $crudos += $partes[$idx] }
            }
        }
        default {
            # .txt u otro: un numero por linea
            $crudos = Get-Content -LiteralPath $Path -Encoding UTF8
        }
    }

    $resultado = New-Object System.Collections.ArrayList
    $vistos = @{}
    $descartados = 0
    foreach ($c in $crudos) {
        $n = ConvertTo-NumeroLimpio -Valor $c
        if (-not $n) { if ($c -and $c.Trim()) { $descartados++ }; continue }
        if ($vistos.ContainsKey($n)) { continue }
        $vistos[$n] = $true
        [void]$resultado.Add($n)
    }

    return [pscustomobject]@{
        Numeros     = @($resultado)
        Descartados = $descartados
        Origen      = $Path
    }
}
