# =====================================================================
#  ExcelHojas.ps1 - Agrega hojas nuevas a un .xlsx que ya existe
#
#  ExcelNuevo.ps1 crea un libro desde cero y ExcelEscribir.ps1 agrega columnas
#  a una hoja existente. Esto es lo que faltaba: meter hojas ADICIONALES en el
#  libro del consolidado, para poder entregar un unico archivo por mes con
#  todo dentro (el consolidado verificado mas los cortes que alimentan el
#  tablero) en vez de varios archivos sueltos.
#
#  Un .xlsx es un ZIP con XML dentro, y agregar una hoja obliga a tocar cuatro
#  partes a la vez; si una queda sin actualizar, Excel da el archivo por danado:
#
#    xl/worksheets/hojaN.xml     el contenido
#    xl/workbook.xml             el nombre y el sheetId
#    xl/_rels/workbook.xml.rels  el vinculo entre ambos
#    [Content_Types].xml         el tipo de la parte nueva
# =====================================================================

. "$PSScriptRoot\ExcelNuevo.ps1"

function Get-NombreHojaValido {
    <#  Excel no admite : \ / ? * [ ] en el nombre, ni mas de 31 caracteres.  #>
    param([Parameter(Mandatory=$true)][string]$Nombre)
    $n = [regex]::Replace($Nombre, '[:\\/?*\[\]]', ' ')
    $n = ($n -replace '\s+', ' ').Trim()
    if ($n.Length -gt 31) { $n = $n.Substring(0, 31).Trim() }
    if (-not $n) { $n = 'Hoja' }
    return $n
}

function New-HojaXml {
    <#  Construye el XML de una hoja a partir de objetos, como cadenas en linea
        (inlineStr) para no depender del sharedStrings del libro original.  #>
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Filas,
        [Parameter(Mandatory=$true)][string[]]$Columnas
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

    [void]$sb.Append('<row r="1">')
    for ($i = 0; $i -lt $Columnas.Count; $i++) {
        $ref = (Convert-IndiceAColumna ($i + 1)) + '1'
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + (ConvertTo-XmlTexto $Columnas[$i]) + '</t></is></c>')
    }
    [void]$sb.Append('</row>')

    $n = 1
    foreach ($f in $Filas) {
        $n++
        [void]$sb.Append('<row r="' + $n + '">')
        for ($i = 0; $i -lt $Columnas.Count; $i++) {
            $ref = (Convert-IndiceAColumna ($i + 1)) + $n
            [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' +
                             (ConvertTo-XmlTexto $f.($Columnas[$i])) + '</t></is></c>')
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData></worksheet>')
    return $sb.ToString()
}

function Add-HojasAXlsx {
    <#  Agrega (o reemplaza) hojas en un libro existente.

        -Hojas  lista de @{ Nombre=''; Filas=@(); Columnas=@() }

        Si ya existe una hoja con ese nombre se reemplaza su contenido, para que
        volver a correr el proceso sobre el mismo archivo no duplique pestanas.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Origen,
        [Parameter(Mandatory=$true)][string]$Destino,
        [Parameter(Mandatory=$true)][hashtable[]]$Hojas
    )

    if (-not (Test-Path -LiteralPath $Origen)) { throw "No existe el libro: $Origen" }
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tmp = Join-Path $env:TEMP ("xlsxh_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null

    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Origen, $tmp)

        $rutaWb   = Join-Path $tmp 'xl\workbook.xml'
        $rutaRels = Join-Path $tmp 'xl\_rels\workbook.xml.rels'
        $rutaCt   = Join-Path $tmp '[Content_Types].xml'
        foreach ($r in @($rutaWb, $rutaRels, $rutaCt)) {
            if (-not (Test-Path -LiteralPath $r)) { throw "El libro no tiene $r; no parece un .xlsx valido" }
        }

        $wb   = [xml](Get-Content -LiteralPath $rutaWb -Raw)
        $rels = [xml](Get-Content -LiteralPath $rutaRels -Raw)
        $ct   = [xml](Get-Content -LiteralPath $rutaCt -Raw)

        $nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $nsR    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $nsRel  = 'http://schemas.openxmlformats.org/package/2006/relationships'
        $nsCt   = 'http://schemas.openxmlformats.org/package/2006/content-types'

        # Numeros ya usados, para no pisar nada
        $maxSheetId = 0
        foreach ($s in $wb.workbook.sheets.sheet) {
            $id = 0; if ([int]::TryParse($s.sheetId, [ref]$id) -and $id -gt $maxSheetId) { $maxSheetId = $id }
        }
        $maxRid = 0
        foreach ($r in $rels.Relationships.Relationship) {
            if ($r.Id -match '^rId(\d+)$' -and [int]$Matches[1] -gt $maxRid) { $maxRid = [int]$Matches[1] }
        }
        $maxHoja = 0
        foreach ($f in (Get-ChildItem -LiteralPath (Join-Path $tmp 'xl\worksheets') -Filter '*.xml' -ErrorAction SilentlyContinue)) {
            if ($f.Name -match '^sheet(\d+)\.xml$' -and [int]$Matches[1] -gt $maxHoja) { $maxHoja = [int]$Matches[1] }
        }

        $agregadas = @(); $reemplazadas = @()

        foreach ($h in $Hojas) {
            $nombre = Get-NombreHojaValido $h.Nombre
            $xml = New-HojaXml -Filas @($h.Filas) -Columnas $h.Columnas

            $existente = $wb.workbook.sheets.sheet | Where-Object { $_.name -eq $nombre } | Select-Object -First 1
            if ($existente) {
                # Reutilizar la hoja: se localiza su archivo por el r:id
                $rid = $existente.GetAttribute('id', $nsR)
                $rel = $rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid } | Select-Object -First 1
                if (-not $rel) { throw "La hoja '$nombre' existe pero no tiene relacion $rid" }
                $destinoXml = Join-Path $tmp ('xl\' + ($rel.Target -replace '/', '\'))
                [System.IO.File]::WriteAllText($destinoXml, $xml, [System.Text.Encoding]::UTF8)
                $reemplazadas += $nombre
                continue
            }

            $maxHoja++; $maxSheetId++; $maxRid++
            $archivo = "sheet$maxHoja.xml"
            [System.IO.File]::WriteAllText((Join-Path $tmp "xl\worksheets\$archivo"), $xml, [System.Text.Encoding]::UTF8)

            $nodo = $wb.CreateElement('sheet', $nsMain)
            $nodo.SetAttribute('name', $nombre)
            $nodo.SetAttribute('sheetId', [string]$maxSheetId)
            $nodo.SetAttribute('id', $nsR, "rId$maxRid")
            [void]$wb.workbook.sheets.AppendChild($nodo)

            $nr = $rels.CreateElement('Relationship', $nsRel)
            $nr.SetAttribute('Id', "rId$maxRid")
            $nr.SetAttribute('Type', "$nsR/worksheet")
            $nr.SetAttribute('Target', "worksheets/$archivo")
            [void]$rels.Relationships.AppendChild($nr)

            $nc = $ct.CreateElement('Override', $nsCt)
            $nc.SetAttribute('PartName', "/xl/worksheets/$archivo")
            $nc.SetAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            [void]$ct.Types.AppendChild($nc)

            $agregadas += $nombre
        }

        $wb.Save($rutaWb); $rels.Save($rutaRels); $ct.Save($rutaCt)

        # calcChain se vuelve invalido al cambiar las hojas; Excel lo rehace solo
        $cc = Join-Path $tmp 'xl\calcChain.xml'
        if (Test-Path -LiteralPath $cc) {
            Remove-Item -LiteralPath $cc -Force
            $nodo = $ct.Types.Override | Where-Object { $_.PartName -eq '/xl/calcChain.xml' }
            if ($nodo) { [void]$ct.Types.RemoveChild($nodo); $ct.Save($rutaCt) }
        }

        if (Test-Path -LiteralPath $Destino) { Remove-Item -LiteralPath $Destino -Force }
        Write-ZipOpc -Carpeta $tmp -Destino $Destino

        return [pscustomobject]@{
            Archivo      = $Destino
            Agregadas    = $agregadas
            Reemplazadas = $reemplazadas
        }
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}

function Rename-HojaXlsx {
    <#  Cambia el nombre de una hoja, dejando el resto del libro igual.

        Sirve para normalizar los consolidados: llegan unas veces con la hoja
        principal llamada CONSOLIDADO y otras como Hoja1, y el entregable debe
        verse igual todos los meses para que el tablero no tenga que adivinar.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Archivo,
        [Parameter(Mandatory=$true)][string]$De,
        [Parameter(Mandatory=$true)][string]$A,
        [string]$Destino
    )

    if (-not (Test-Path -LiteralPath $Archivo)) { throw "No existe el libro: $Archivo" }
    if (-not $Destino) { $Destino = $Archivo }
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tmp = Join-Path $env:TEMP ("xlsxr_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Archivo, $tmp)
        $rutaWb = Join-Path $tmp 'xl\workbook.xml'
        $wb = [xml](Get-Content -LiteralPath $rutaWb -Raw)

        $nuevo = Get-NombreHojaValido $A
        $hoja = $wb.workbook.sheets.sheet | Where-Object { $_.name -eq $De } | Select-Object -First 1
        if (-not $hoja) { return [pscustomobject]@{ Archivo = $Destino; Cambio = $false; Motivo = "no existe la hoja '$De'" } }
        if ($wb.workbook.sheets.sheet | Where-Object { $_.name -eq $nuevo }) {
            return [pscustomobject]@{ Archivo = $Destino; Cambio = $false; Motivo = "ya existe una hoja '$nuevo'" }
        }

        $hoja.SetAttribute('name', $nuevo)
        $wb.Save($rutaWb)

        # definedNames y calcChain pueden referirse al nombre viejo
        $cc = Join-Path $tmp 'xl\calcChain.xml'
        if (Test-Path -LiteralPath $cc) { Remove-Item -LiteralPath $cc -Force }

        if (Test-Path -LiteralPath $Destino) { Remove-Item -LiteralPath $Destino -Force }
        Write-ZipOpc -Carpeta $tmp -Destino $Destino
        return [pscustomobject]@{ Archivo = $Destino; Cambio = $true; Motivo = "'$De' -> '$nuevo'" }
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}
