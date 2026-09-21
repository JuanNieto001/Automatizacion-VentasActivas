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
    [switch]   $MotivoRapido,
    [switch]   $SoloEstado,
    [string]   $Salida,
    [string]   $ExcelSalida,
    [int]      $Limite = 0,
    [int]      $MaxReintentos = 2,
    [ValidateRange(1, 12)]
    [int]      $Instancias = 1
)

$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path

. "$raiz\lib\Win32.ps1"
. "$raiz\lib\Remote.ps1"
. "$raiz\lib\AC.ps1"
. "$raiz\lib\ACParalelo.ps1"
. "$raiz\lib\Excel.ps1"
. "$raiz\lib\ExcelEscribir.ps1"

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
    # ASCII a proposito: el cifrado DPAPI es una cadena hexadecimal, y un BOM
    # de UTF-8 al principio del archivo rompe la lectura posterior.
    $sec | ConvertFrom-SecureString | Set-Content -LiteralPath $archivoClave -Encoding ASCII
    Write-Host "Clave guardada cifrada en $archivoClave" -ForegroundColor Green
    Write-Host "Solo se puede descifrar con tu usuario de Windows en este equipo." -ForegroundColor Gray
}

function Get-Clave {
    param([string]$Explicita)
    if ($Explicita) { return $Explicita }
    if (Test-Path -LiteralPath $archivoClave) {
        try {
            # se limpian BOM y espacios por si el archivo se guardo en UTF-8
            $txt = (Get-Content -LiteralPath $archivoClave -Raw).Trim([char]0xFEFF, [char]0x200B, ' ', "`r", "`n", "`t")
            $sec = $txt | ConvertTo-SecureString
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

$totalDisponibles = $listaNumeros.Count
if ($Limite -gt 0 -and $listaNumeros.Count -gt $Limite) {
    $listaNumeros = @($listaNumeros | Select-Object -First $Limite)
    Write-Host "Limitado a los primeros $Limite de $totalDisponibles numeros (-Limite)." -ForegroundColor Yellow
}

Write-Host "Origen : $origen"
if ($Hoja) { Write-Host "Hoja   : $Hoja" }
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
        HISTORIAL_MOVIMIENTOS = ''
        HISTORIAL_COMPLETO  = ''
        OBSERVACION         = ''
        CONSULTADO          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

# ---------------------------------------------------------------------
#  PASADA 1 - estado de cada numero (sin abrir fichas, sin Ticklers)
# ---------------------------------------------------------------------

Write-Titulo "PASADA 1 - ESTADO DE CADA NUMERO"
Write-Host "AC se maneja por mensajes, pero necesita el primer plano en cada clic:" -ForegroundColor Yellow
Write-Host "sus ventanas van a saltar al frente. No uses el equipo mientras corre." -ForegroundColor Yellow
Write-Host ""

# ojo: no llamar a esta variable $instancias, PowerShell no distingue
# mayusculas y pisaria el parametro -Instancias del script
$pool = Start-ACInstancias -Cantidad $Instancias -Password $clave -BaseDatos $BaseDatos
Write-Host ""

$hechos = 0
$total  = $listaNumeros.Count
$porNumero = @{}

$stats = Invoke-ACConsultaMasiva -Numeros $listaNumeros -Instancias $pool `
            -Password $clave -BaseDatos $BaseDatos -OnResultado {
    param($r)
    $script:hechos++
    $reg = New-Registro -Numero $r.Numero

    if (-not $r.Encontrado) {
        $reg.ENCONTRADO  = 'NO'
        $reg.ACTIVA      = 'NO'
        $reg.ESTADO      = 'NO ENCONTRADO'
        $reg.MOTIVO      = $r.Mensaje
        $reg.OBSERVACION = 'AC no devolvio ninguna linea para este numero.'
        $txt = "NO ENCONTRADO"; $col = "Yellow"
    } else {
        # si el cliente tiene varias lineas, se toma la del numero consultado
        $fila = $r.Filas | Where-Object { ($_.Campos.MIN -replace '[^0-9]','') -eq $r.Numero } | Select-Object -First 1
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
        $txt = "$($c.NOMBRE) - $($c.ESTADO)"
        $col = if ($reg.ACTIVA -eq 'SI') { "Green" } else { "Yellow" }
    }

    [void]$script:resultados.Add([pscustomobject]$reg)
    Write-Host ("  [{0}/{1}] {2}  {3}  ({4}s, inst {5})" -f `
        $script:hechos, $script:total, $r.Numero, $txt, $r.Segundos, $r.Instancia) -ForegroundColor $col

    # Guardado parcial: en una corrida larga no se puede perder todo por un
    # fallo al final. Se reescribe el CSV cada 50 numeros.
    if (($script:hechos % 50) -eq 0) {
        try { $script:resultados | Export-Csv -LiteralPath $script:Salida -NoTypeInformation -Encoding UTF8 } catch { }
        Write-Host ("       ... parcial guardado ({0} numeros)" -f $script:hechos) -ForegroundColor DarkGray
    }
}

Stop-ACSession -Todas

Write-Host ""
Write-Host ("Pasada 1: {0} numeros en {1} s  |  {2} s por numero  |  {3} instancia(s)" -f `
    $stats.Resultados.Count, $stats.SegundosTotales, $stats.SegundosPorNumero, $stats.InstanciasVivas) -ForegroundColor Cyan
Write-Host ("Latencia media de cada consulta: {0} s" -f $stats.LatenciaMedia) -ForegroundColor Gray
if ($Instancias -gt 1 -and $stats.SegundosPorNumero -gt 0) {
    Write-Host ("Rendimiento: {0} numeros por minuto" -f [Math]::Round(60 / $stats.SegundosPorNumero, 1)) -ForegroundColor Gray
}

# ---------------------------------------------------------------------
#  PASADA 2 - motivo de las lineas que no estan activas
# ---------------------------------------------------------------------

if ($SoloEstado) {
    $pendientes = @()
    Write-Host ""
    Write-Host "Modo -SoloEstado: no se consultan motivos." -ForegroundColor Yellow
} else {
    $pendientes = @($resultados | Where-Object {
        $_.ENCONTRADO -eq 'SI' -and ($HistorialSiempre -or $_.ACTIVA -ne 'SI')
    })
}

if ($pendientes.Count -gt 0) {
    Write-Titulo "PASADA 2 - MOTIVO (HISTORIAL, Ctrl+Shift+H)"
    Write-Host "Lineas por revisar: $($pendientes.Count)"
    Write-Host "Cada ficha consume una instancia de AC (no se puede cerrar sin guardar un Tickler)," -ForegroundColor Gray
    Write-Host "asi que los reinicios se encadenan entre instancias para solapar la espera." -ForegroundColor Gray
    Write-Host "Usa el raton: no toques el equipo mientras corre." -ForegroundColor Yellow
    Write-Host ""

    $swP2 = [Diagnostics.Stopwatch]::StartNew()
    $j = 0
    foreach ($reg in $pendientes) {
        $j++
        Write-Host ("[{0}/{1}] {2} ({3})" -f $j, $pendientes.Count, $reg.NUMERO, $reg.ESTADO) -ForegroundColor White
        try {
            Stop-ACSession -Todas
            $ctx = Start-ACSession -Password $clave -BaseDatos $BaseDatos

            $r = Invoke-ACBusqueda -Numero $reg.NUMERO
            if (-not $r.Encontrado) {
                $reg.OBSERVACION = (("$($reg.OBSERVACION) No se pudo reabrir para el motivo: $($r.Mensaje)").Trim())
                Write-Paso "No se pudo reabrir: $($r.Mensaje)" "WARN"
                continue
            }

            $ctxF = Open-ACFicha

            if ($MotivoRapido) {
                # La ficha muestra el motivo vigente en el recuadro resaltado
                # junto a "Estado Contrato". Verificado: coincide con la ultima
                # fila del HISTORIAL. Ahorra los ~40 s que cuesta abrirlo, pero
                # ese campo no trae la fecha del movimiento.
                #
                # Se localiza el control por su posicion relativa (barato) y se
                # sondea SOLO ese campo: recorrer todos los campos por UI
                # Automation en cada ciclo costaba tanto como el historial.
                $f = (Get-ACContext).Ficha
                $campo = Get-ChildHandles -Parent $f.Handle | Where-Object {
                    $_.Class -like '*TextBox*' -and
                    [Math]::Abs(($_.X - $f.X) - 279) -le 12 -and
                    [Math]::Abs(($_.Y - $f.Y) -  80) -le 12
                } | Select-Object -First 1

                $valor = $null
                if ($campo) {
                    $valor = Wait-Condition -TimeoutSeg 30 -IntervaloMs 150 -Condicion {
                        $v = Get-UiaName -Handle ([int]$campo.Handle)
                        if ($v) { $v }
                    }
                }
                if ($valor) {
                    $reg.MOTIVO = $valor
                    Write-Paso "Motivo: $valor (leido de la ficha)" "OK"
                } else {
                    $reg.OBSERVACION = (("$($reg.OBSERVACION) La ficha no mostro el motivo.").Trim())
                    Write-Paso "La ficha no mostro el motivo." "WARN"
                }
            }
            else {
                $h = Get-ACHistorial
                $filas = @($h.EstadoContrato.Rows)
                if ($filas.Count -gt 0) {
                    # El HISTORIAL viene en orden cronologico: la PRIMERA fila es
                    # el movimiento mas antiguo y la ULTIMA el estado vigente.
                    # Verificado contra la grilla en varios casos.
                    $ultima = $filas[-1].Campos
                    $reg.MOTIVO               = $ultima.MOTIVO
                    $reg.FECHA_ESTADO         = $ultima.'VALIDO DESDE'
                    $reg.USUARIO_ESTADO       = $ultima.USUARIO
                    $reg.HISTORIAL_MOVIMIENTOS = $filas.Count
                    $reg.HISTORIAL_COMPLETO   = (($filas | ForEach-Object {
                        "$($_.Campos.ESTADO) / $($_.Campos.MOTIVO) / $($_.Campos.'VALIDO DESDE') / $($_.Campos.USUARIO)"
                    }) -join ' || ')
                    Write-Paso "Motivo: $($ultima.MOTIVO) ($($ultima.ESTADO), $($ultima.'VALIDO DESDE')) - $($filas.Count) movimientos" "OK"
                    [void](Save-WindowShot -Handle $h.VentanaHandle -Path (Join-Path $dirCapturas "historial_$($reg.NUMERO).png"))
                } else {
                    $reg.OBSERVACION = (("$($reg.OBSERVACION) El historial no devolvio movimientos.").Trim())
                    Write-Paso "El historial no devolvio movimientos." "WARN"
                }
            }
        } catch {
            $reg.OBSERVACION = (("$($reg.OBSERVACION) Error al consultar el motivo: $($_.Exception.Message)").Trim())
            Write-Paso "Error: $($_.Exception.Message)" "ERROR"
        }
    }
    $swP2.Stop()
    Write-Host ""
    Write-Host ("Pasada 2: {0} lineas en {1:N0} s  |  {2:N1} s por linea" -f `
        $pendientes.Count, $swP2.Elapsed.TotalSeconds,
        ($swP2.Elapsed.TotalSeconds / [Math]::Max(1, $pendientes.Count))) -ForegroundColor Cyan
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
Write-Host "Reporte CSV: $Salida" -ForegroundColor Cyan

# ---------------------------------------------------------------------
#  Excel de vuelta: el mismo libro con las columnas de verificacion
# ---------------------------------------------------------------------

if ($Archivo -and ([System.IO.Path]::GetExtension($Archivo) -in '.xlsx', '.xlsm')) {
    if (-not $ExcelSalida) {
        $nom = [System.IO.Path]::GetFileNameWithoutExtension($Archivo)
        $ext = [System.IO.Path]::GetExtension($Archivo)
        $ExcelSalida = Join-Path $dirSalida "$nom - VERIFICADO $sello$ext"
    }
    try {
        $mapa = @{}
        foreach ($r in $resultados) {
            $verif = switch ($r.ENCONTRADO) {
                'SI'    { if ($r.ACTIVA -eq 'SI') { 'EXITOSA' } else { 'NO EXITOSA' } }
                'NO'    { 'NO ENCONTRADA' }
                default { 'ERROR' }
            }
            # el motivo: primero el del historial, si no el mensaje de AC
            $motivo = if ($r.MOTIVO) { $r.MOTIVO } else { $r.OBSERVACION }
            if ($r.FECHA_ESTADO) { $motivo = "$motivo ($($r.FECHA_ESTADO))".Trim() }
            $mapa[$r.NUMERO] = [pscustomobject]@{
                Verificacion = $verif
                Estado       = $r.ESTADO
                Motivo       = $motivo
            }
        }

        $info = Add-VerificacionAExcel -Origen $Archivo -Destino $ExcelSalida -Hoja $Hoja `
                    -Columna $Columna -FilaInicio $FilaInicio -Resultados $mapa `
                    -Fecha (Get-Date -Format 'yyyy-MM-dd')

        Write-Host ("Excel      : {0}" -f $info.Archivo) -ForegroundColor Cyan
        Write-Host ("             hoja '{0}', columnas {1} ({2})" -f `
            $info.Hoja, ($info.Columnas -join ', '), ($info.Encabezados -join ' / ')) -ForegroundColor Gray
        Write-Host ("             {0} filas con verificacion, {1} sin verificar" -f `
            $info.FilasEscritas, $info.SinResultado) -ForegroundColor Gray
    } catch {
        Write-Host "No se pudo generar el Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "El CSV con los resultados si quedo generado." -ForegroundColor Yellow
    }
}

if (Test-Path -LiteralPath $dirCapturas) { Write-Host "Capturas   : $dirCapturas" -ForegroundColor Cyan }

try { Stop-ACSession -Todas } catch { }



