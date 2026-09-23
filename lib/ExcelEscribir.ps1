# =====================================================================
#  ExcelEscribir.ps1 - Devuelve el MISMO Excel con columnas de verificacion
#
#  No hay Excel instalado en el equipo, asi que el .xlsx se modifica como lo
#  que es: un ZIP con XML. Se trabaja sobre una copia del original y solo se
#  agregan celdas nuevas al final de cada fila de la hoja indicada, de modo
#  que el resto del libro (formatos, formulas, otras hojas) queda intacto.
#
#  Los valores nuevos se escriben como cadenas en linea (inlineStr) para no
#  tener que tocar sharedStrings.xml ni recalcular sus contadores.
# =====================================================================

function Convert-IndiceAColumna {
    <#  1 -> A, 27 -> AA  #>
    param([Parameter(Mandatory=$true)][int]$Indice)
    $s = ''
    while ($Indice -gt 0) {
        $r = ($Indice - 1) % 26
        $s = [char](65 + $r) + $s
        $Indice = [int](($Indice - $r - 1) / 26)
    }
    return $s
}

function Convert-ColumnaAIndice {
    <#  A -> 1, AA -> 27  #>
    param([Parameter(Mandatory=$true)][string]$Columna)
    $n = 0
    foreach ($ch in $Columna.ToUpper().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $n = $n * 26 + ([int][char]$ch - 64)
    }
    return $n
}

function Resolve-HojaXlsx {
    <#  Devuelve la ruta del XML de una hoja dentro de un xlsx ya extraido.
        El nombre de archivo no corresponde al nombre de la hoja: hay que
        resolverlo por el r:id contra workbook.xml.rels.  #>
    param([Parameter(Mandatory=$true)][string]$Raiz, [string]$Hoja)

    [xml]$wb = Get-Content -LiteralPath (Join-Path $Raiz "xl\workbook.xml") -Encoding UTF8
    [xml]$rl = Get-Content -LiteralPath (Join-Path $Raiz "xl\_rels\workbook.xml.rels") -Encoding UTF8
    $destinos = @{}
    foreach ($r in $rl.Relationships.Relationship) { $destinos[$r.Id] = $r.Target }

    $nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    $hojas = @()
    foreach ($s in $wb.workbook.sheets.sheet) {
        $rid = $s.id
        if (-not $rid) { $rid = $s.GetAttribute("id", $nsRel) }
        $hojas += [pscustomobject]@{ Nombre = $s.name; Target = $destinos[$rid] }
    }

    $elegida = $null
    if ($Hoja) {
        $elegida = $hojas | Where-Object { $_.Nombre -and $_.Nombre.Trim() -ieq $Hoja.Trim() } | Select-Object -First 1
        if (-not $elegida) { $elegida = $hojas | Where-Object { $_.Nombre -and $_.Nombre -ilike "*$($Hoja.Trim())*" } | Select-Object -First 1 }
        if (-not $elegida) { throw "No existe la hoja '$Hoja'. Disponibles: $(($hojas | ForEach-Object { $_.Nombre }) -join ', ')" }
    } else {
        $elegida = $hojas[0]
    }
    $rel = $elegida.Target -replace '^/xl/', '' -replace '^/', ''
    return [pscustomobject]@{
        Nombre = $elegida.Nombre
        Ruta   = (Join-Path $Raiz ("xl\" + ($rel -replace '/', '\')))
    }
}

function Write-ZipOpc {
    <#  Reempaqueta una carpeta como .xlsx valido.

        No se usa ZipFile.CreateFromDirectory: en esta version de .NET escribe
        los nombres de entrada con barra invertida (docProps\app.xml), lo que
        incumple el formato OPC y hace que Excel de el archivo por danado.
        Aqui se escriben con barra normal y respetando el orden de entradas
        del libro original.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Carpeta,
        [Parameter(Mandatory=$true)][string]$Destino,
        [string]$Original
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $archivos = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $Carpeta -Recurse -File -Force)) {
        $rel = $f.FullName.Substring($Carpeta.Length).TrimStart('\', '/') -replace '\\', '/'
        $archivos[$rel] = $f.FullName
    }

    # orden original; lo que no estuviera, al final
    $orden = New-Object System.Collections.ArrayList
    if ($Original -and (Test-Path -LiteralPath $Original)) {
        $zo = [System.IO.Compression.ZipFile]::OpenRead($Original)
        try {
            foreach ($e in $zo.Entries) {
                $n = $e.FullName -replace '\\', '/'
                if ($archivos.ContainsKey($n)) { [void]$orden.Add($n) }
            }
        } finally { $zo.Dispose() }
    }
    foreach ($k in $archivos.Keys) { if (-not $orden.Contains($k)) { [void]$orden.Add($k) } }

    # [Content_Types].xml debe ser la primera entrada del paquete OPC.
    foreach ($primero in @('[Content_Types].xml', '_rels/.rels')) {
        if ($orden.Contains($primero)) {
            $orden.Remove($primero)
            $orden.Insert(0, $primero)
        }
    }
    if ($orden.Contains('[Content_Types].xml')) {
        $orden.Remove('[Content_Types].xml')
        $orden.Insert(0, '[Content_Types].xml')
    }

    $zip = [System.IO.Compression.ZipFile]::Open($Destino, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($nombre in $orden) {
            $entrada = $zip.CreateEntry($nombre, [System.IO.Compression.CompressionLevel]::Optimal)
            $salida = $entrada.Open()
            try {
                $bytes = [System.IO.File]::ReadAllBytes($archivos[$nombre])
                $salida.Write($bytes, 0, $bytes.Length)
            } finally { $salida.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Add-VerificacionAExcel {
    <#  Copia el libro original y agrega a la hoja indicada las columnas de
        verificacion, emparejando por el numero de la columna de origen.

        -Resultados es una tabla hash: numero normalizado -> objeto con
        Verificacion, Estado, Motivo y Detalle.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Origen,
        [Parameter(Mandatory=$true)][string]$Destino,
        [string]$Hoja,
        [string]$Columna = 'E',
        [int]$FilaInicio = 2,
        [Parameter(Mandatory=$true)][hashtable]$Resultados,
        [string]$Fecha = (Get-Date -Format 'yyyy-MM-dd')
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (-not (Test-Path -LiteralPath $Origen)) { throw "No existe el archivo de origen: $Origen" }

    $dirDestino = Split-Path -Parent $Destino
    if ($dirDestino -and -not (Test-Path -LiteralPath $dirDestino)) {
        New-Item -ItemType Directory -Force -Path $dirDestino | Out-Null
    }

    $tmp = Join-Path $env:TEMP ("xlsxw_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Origen, $tmp)

        $info = Resolve-HojaXlsx -Raiz $tmp -Hoja $Hoja

        # --- cadenas compartidas, para poder leer la columna de numeros ---
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

        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $false
        $doc.Load($info.Ruta)
        $ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $mgr.AddNamespace("d", $ns)

        $filas = $doc.SelectNodes("//d:sheetData/d:row", $mgr)
        if (-not $filas -or $filas.Count -eq 0) { throw "La hoja '$($info.Nombre)' no tiene filas." }

        # --- primera columna libre ---
        # Se toma la ultima columna CON ENCABEZADO, no la ultima celda del
        # archivo. Algunos consolidados traen restos sueltos mucho mas a la
        # derecha -el de septiembre del 23-sep tenia celdas hasta la BC aunque
        # sus encabezados terminan en la AG- y midiendo por la ultima celda las
        # columnas de verificacion caian en BD, lejos del resto y en una letra
        # distinta a la de los demas meses.
        #
        # Si la fila de encabezados no se puede leer se vuelve al criterio
        # anterior, que al menos garantiza no pisar datos.
        $maxCab = 0
        $filaCab = $filas | Where-Object { $_.GetAttribute("r") -eq '1' } | Select-Object -First 1
        if ($filaCab) {
            foreach ($c in $filaCab.ChildNodes) {
                if ($c.LocalName -ne 'c') { continue }
                $ref = $c.GetAttribute("r")
                if ($ref -notmatch '^([A-Z]+)\d+$') { continue }
                # solo cuenta si la celda tiene contenido
                $tiene = $false
                foreach ($h in $c.ChildNodes) { if ($h.LocalName -in @('v','is') -and "$($h.InnerText)".Trim()) { $tiene = $true } }
                if (-not $tiene) { continue }
                $i = Convert-ColumnaAIndice $matches[1]
                if ($i -gt $maxCab) { $maxCab = $i }
            }
        }

        $maxCol = 0
        foreach ($f in $filas) {
            foreach ($c in $f.ChildNodes) {
                if ($c.LocalName -ne 'c') { continue }
                $ref = $c.GetAttribute("r")
                if ($ref -match '^([A-Z]+)\d+$') {
                    $i = Convert-ColumnaAIndice $matches[1]
                    if ($i -gt $maxCol) { $maxCol = $i }
                }
            }
        }
        if ($maxCab -gt 0) { $maxCol = $maxCab }

        $encabezados = @('VERIFICACION', 'ESTADO AC', 'MOTIVO', 'FECHA VERIFICACION')
        $colsNuevas = @()
        for ($i = 0; $i -lt $encabezados.Count; $i++) {
            $colsNuevas += (Convert-IndiceAColumna ($maxCol + 1 + $i))
        }

        function New-CeldaTexto {
            param([string]$Ref, [string]$Texto)
            $c = $doc.CreateElement("c", $ns)
            $c.SetAttribute("r", $Ref)
            $c.SetAttribute("t", "inlineStr")
            $is = $doc.CreateElement("is", $ns)
            $t  = $doc.CreateElement("t", $ns)
            # No se pone xml:space: SetAttribute con el namespace reservado de
            # XML genera un prefijo inventado (d6p1) que deja el XML invalido y
            # Excel da el archivo por danado. Los valores se recortan en su
            # lugar, que es lo que se necesita aqui.
            $t.InnerText = if ($null -eq $Texto) { "" } else { ([string]$Texto).Trim() }
            [void]$is.AppendChild($t)
            [void]$c.AppendChild($is)
            return $c
        }

        $colOrigenIdx = Convert-ColumnaAIndice $Columna
        $escritas = 0
        $sinResultado = 0

        foreach ($f in $filas) {
            $nFila = [int]$f.GetAttribute("r")

            # valor de la columna de origen en esta fila
            $valor = $null
            foreach ($c in $f.ChildNodes) {
                if ($c.LocalName -ne 'c') { continue }
                $ref = $c.GetAttribute("r")
                if ($ref -notmatch '^([A-Z]+)\d+$') { continue }
                if ((Convert-ColumnaAIndice $matches[1]) -ne $colOrigenIdx) { continue }
                $tipo = $c.GetAttribute("t")
                $nodoV = $c.SelectSingleNode("d:v", $mgr)
                if ($tipo -eq 's' -and $nodoV) {
                    $idx = [int]$nodoV.InnerText
                    if ($idx -ge 0 -and $idx -lt $shared.Count) { $valor = $shared[$idx] }
                } elseif ($tipo -eq 'inlineStr') {
                    $nodoT = $c.SelectSingleNode("d:is/d:t", $mgr)
                    if ($nodoT) { $valor = $nodoT.InnerText }
                } elseif ($nodoV) {
                    $valor = $nodoV.InnerText
                }
                break
            }

            if ($nFila -eq 1) {
                for ($i = 0; $i -lt $encabezados.Count; $i++) {
                    [void]$f.AppendChild((New-CeldaTexto -Ref ($colsNuevas[$i] + $nFila) -Texto $encabezados[$i]))
                }
                continue
            }
            if ($nFila -lt $FilaInicio) { continue }

            $num = ConvertTo-NumeroLimpio -Valor $valor
            if (-not $num) { continue }

            if ($Resultados.ContainsKey($num)) {
                $r = $Resultados[$num]
                $vals = @("$($r.Verificacion)", "$($r.Estado)", "$($r.Motivo)", $Fecha)
                $escritas++
            } else {
                $vals = @('NO VERIFICADO', '', '', $Fecha)
                $sinResultado++
            }
            for ($i = 0; $i -lt $vals.Count; $i++) {
                [void]$f.AppendChild((New-CeldaTexto -Ref ($colsNuevas[$i] + $nFila) -Texto $vals[$i]))
            }
        }

        # --- actualizar la dimension declarada, si existe ---
        $dim = $doc.SelectSingleNode("//d:dimension", $mgr)
        if ($dim) {
            $refDim = $dim.GetAttribute("ref")
            if ($refDim -match '^([A-Z]+\d+):([A-Z]+)(\d+)$') {
                $dim.SetAttribute("ref", ("{0}:{1}{2}" -f $matches[1], $colsNuevas[-1], $matches[3]))
            }
        }

        $doc.Save($info.Ruta)

        if (Test-Path -LiteralPath $Destino) { Remove-Item -LiteralPath $Destino -Force }
        Write-ZipOpc -Carpeta $tmp -Destino $Destino -Original $Origen

        return [pscustomobject]@{
            Archivo        = $Destino
            Hoja           = $info.Nombre
            Columnas       = $colsNuevas
            Encabezados    = $encabezados
            FilasEscritas  = $escritas
            SinResultado   = $sinResultado
        }
    } finally {
        Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}
