# =====================================================================
#  ExcelNuevo.ps1 - Crea un .xlsx desde cero (sin Excel instalado)
#
#  Escribe los valores como cadenas en linea (inlineStr), asi no hace falta
#  sharedStrings.xml. Reutiliza Write-ZipOpc de ExcelEscribir.ps1, que empaqueta
#  con rutas validas (ZipFile.CreateFromDirectory las escribe con barra
#  invertida y Excel da el archivo por danado).
# =====================================================================

. "$PSScriptRoot\ExcelEscribir.ps1"

function ConvertTo-XmlTexto {
    param([string]$Texto)
    if ($null -eq $Texto) { return '' }
    $t = [string]$Texto
    # se quitan los caracteres de control que XML 1.0 no admite
    $t = [regex]::Replace($t, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return ($t -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;')
}

function New-XlsxSimple {
    <#  Crea un libro de una hoja a partir de objetos.

        -Filas      objetos a volcar
        -Columnas   nombres de propiedad, en orden; tambien son los encabezados
        -Hoja       nombre de la hoja  #>
    param(
        [Parameter(Mandatory=$true)][object[]]$Filas,
        [Parameter(Mandatory=$true)][string[]]$Columnas,
        [Parameter(Mandatory=$true)][string]$Destino,
        [string]$Hoja = 'Hoja1'
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tmp = Join-Path $env:TEMP ("xlsxn_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path (Join-Path $tmp "_rels") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tmp "xl\_rels") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $tmp "xl\worksheets") | Out-Null

    try {
        $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
              '<Default Extension="xml" ContentType="application/xml"/>' +
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
              '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
              '</Types>'
        [System.IO.File]::WriteAllText((Join-Path $tmp '[Content_Types].xml'), $ct, [System.Text.Encoding]::UTF8)

        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>'
        [System.IO.File]::WriteAllText((Join-Path $tmp "_rels\.rels"), $rels, [System.Text.Encoding]::UTF8)

        $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
              '<sheets><sheet name="' + (ConvertTo-XmlTexto $Hoja) + '" sheetId="1" r:id="rId1"/></sheets></workbook>'
        [System.IO.File]::WriteAllText((Join-Path $tmp "xl\workbook.xml"), $wb, [System.Text.Encoding]::UTF8)

        $wbr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
               '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
               '</Relationships>'
        [System.IO.File]::WriteAllText((Join-Path $tmp "xl\_rels\workbook.xml.rels"), $wbr, [System.Text.Encoding]::UTF8)

        # estilo 1 = negrita, para la fila de encabezados
        $st = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
              '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
              '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
              '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
              '<borders count="1"><border/></borders>' +
              '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
              '<cellXfs count="2"><xf xfId="0"/><xf fontId="1" applyFont="1" xfId="0"/></cellXfs>' +
              '</styleSheet>'
        [System.IO.File]::WriteAllText((Join-Path $tmp "xl\styles.xml"), $st, [System.Text.Encoding]::UTF8)

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

        [void]$sb.Append('<row r="1">')
        for ($i = 0; $i -lt $Columnas.Count; $i++) {
            $ref = (Convert-IndiceAColumna ($i + 1)) + '1'
            [void]$sb.Append('<c r="' + $ref + '" s="1" t="inlineStr"><is><t>' +
                             (ConvertTo-XmlTexto $Columnas[$i]) + '</t></is></c>')
        }
        [void]$sb.Append('</row>')

        $n = 1
        foreach ($f in $Filas) {
            $n++
            [void]$sb.Append('<row r="' + $n + '">')
            for ($i = 0; $i -lt $Columnas.Count; $i++) {
                $ref = (Convert-IndiceAColumna ($i + 1)) + $n
                $val = $f.($Columnas[$i])
                [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' +
                                 (ConvertTo-XmlTexto $val) + '</t></is></c>')
            }
            [void]$sb.Append('</row>')
        }
        [void]$sb.Append('</sheetData></worksheet>')
        [System.IO.File]::WriteAllText((Join-Path $tmp "xl\worksheets\sheet1.xml"), $sb.ToString(), [System.Text.Encoding]::UTF8)

        if (Test-Path -LiteralPath $Destino) { Remove-Item -LiteralPath $Destino -Force }
        Write-ZipOpc -Carpeta $tmp -Destino $Destino

        return [pscustomobject]@{ Archivo = $Destino; Hoja = $Hoja; Filas = $Filas.Count; Columnas = $Columnas.Count }
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}
