# =====================================================================
#  AC.ps1 - Flujo de la aplicacion "AC Administracion de Clientes" V22.8.0
#
#  Implementa el paso a paso documentado en "PASO A PASO AC.docx":
#    1. Tomar el numero de la columna E del Excel
#    2. Ingresar a AC
#    3. Consultar el numero (criterio MIN / MSISDN)
#    4. Seleccionar la linea
#    5. Validar si la linea esta activa o suspendida
#    6. Ctrl+Shift+H sobre la ficha -> HISTORIAL -> motivo del estado
#
#  IMPORTANTE - solo lectura:
#  AC exige diligenciar y GUARDAR un "Solicitud Tickler" para cerrar una
#  ficha de cliente. Guardar ese tickler escribiria un registro en
#  produccion, por lo que la automatizacion NUNCA lo hace: cuando necesita
#  cerrar una ficha reinicia la aplicacion.
# =====================================================================

. "$PSScriptRoot\Win32.ps1"
. "$PSScriptRoot\Remote.ps1"
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

$script:VK_CONTROL = 0x11
$script:VK_SHIFT   = 0x10
$script:VK_H       = 0x48

function Write-Paso {
    param([string]$Mensaje, [string]$Nivel = "INFO")
    $color = switch ($Nivel) {
        "OK"    { "Green" }
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        default { "Gray" }
    }
    Write-Host ("  [{0:HH:mm:ss}] {1}" -f (Get-Date), $Mensaje) -ForegroundColor $color
}

# ---------------------------------------------------------------------
#  Localizacion del ejecutable y del proceso
# ---------------------------------------------------------------------

function Get-ACExePath {
    foreach ($base in @("C:\Program Files (x86)", "C:\Program Files")) {
        if (-not (Test-Path -LiteralPath $base)) { continue }
        $dir = Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like "AC Admin*" } | Select-Object -First 1
        if ($dir) {
            $exe = Get-ChildItem -LiteralPath $dir.FullName -File -ErrorAction SilentlyContinue |
                   Where-Object { $_.Extension -eq ".exe" } | Select-Object -First 1
            if ($exe) { return $exe.FullName }
        }
    }
    throw "No se encontro el ejecutable de AC Administracion de Clientes."
}

function Get-ACProcess {
    param([int]$ProcessId = 0)
    if ($ProcessId -gt 0) {
        return (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)
    }
    return (Get-Process | Where-Object { $_.ProcessName -like 'AC Admin*' } | Select-Object -First 1)
}

function Get-ACProcesses {
    return @(Get-Process | Where-Object { $_.ProcessName -like 'AC Admin*' })
}

function Wait-Condition {
    param([scriptblock]$Condicion, [int]$TimeoutSeg = 30, [int]$IntervaloMs = 500)
    $fin = (Get-Date).AddSeconds($TimeoutSeg)
    while ((Get-Date) -lt $fin) {
        $r = & $Condicion
        if ($r) { return $r }
        Start-Sleep -Milliseconds $IntervaloMs
    }
    return $null
}

# ---------------------------------------------------------------------
#  Dialogos emergentes (#32770)
# ---------------------------------------------------------------------

function Get-ACDialogs {
    param([int]$ProcessId)
    return @(Get-TopWindows -ProcessId $ProcessId | Where-Object { $_.Visible -and $_.Class -eq '#32770' })
}

function Read-DialogText {
    param([IntPtr]$Handle)
    $partes = Get-ChildHandles -Parent $Handle |
              Where-Object { $_.Class -eq 'Static' -and $_.Text } |
              ForEach-Object { $_.Text }
    return ($partes -join ' ')
}

function Close-ACDialogs {
    <#  Cierra los cuadros de dialogo y devuelve los mensajes encontrados.  #>
    param([Parameter(Mandatory=$true)][int]$ProcessId, [int]$MaxIteraciones = 6)
    $mensajes = @()
    for ($i = 0; $i -lt $MaxIteraciones; $i++) {
        $dlgs = Get-ACDialogs -ProcessId $ProcessId
        if (-not $dlgs -or $dlgs.Count -eq 0) { break }
        foreach ($d in $dlgs) {
            $txt = Read-DialogText -Handle $d.Handle
            if ($txt) { $mensajes += $txt }
            $btn = Get-ChildHandles -Parent $d.Handle |
                   Where-Object { $_.Class -eq 'Button' -and $_.Visible } |
                   Select-Object -First 1
            if ($btn) {
                # por mensaje: no roba el raton ni exige primer plano, asi que
                # funciona con varias instancias corriendo a la vez
                Invoke-PostClick -Handle $btn.Handle
            } else {
                [void][W32]::PostMessage($d.Handle, [W32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            Start-Sleep -Milliseconds 600
        }
    }
    return $mensajes
}

function Close-ACTickler {
    <#  Cierra el formulario "Solicitud Tickler" SIN guardar nada.  #>
    param([Parameter(Mandatory=$true)][int]$ProcessId)
    $tk = @(Get-TopWindows -ProcessId $ProcessId | Where-Object { $_.Visible -and $_.Text -like "*Tickler*" })
    foreach ($t in $tk) {
        [void][W32]::PostMessage($t.Handle, [W32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 800
    }
    return ($tk.Count -gt 0)
}

# ---------------------------------------------------------------------
#  Contexto: descubrimiento dinamico de ventanas y controles
#  (los handles cambian en cada ejecucion, nunca se codifican fijos)
# ---------------------------------------------------------------------

function Get-ACContext {
    <#  Descubre ventanas y controles de UNA instancia de AC.
        Con -ProcessId apunta a una instancia concreta, lo que permite
        manejar varias a la vez.  #>
    param([int]$ProcessId = 0)

    $p = Get-ACProcess -ProcessId $ProcessId
    if (-not $p) { return $null }

    $tops = Get-TopWindows -ProcessId $p.Id | Where-Object Visible
    $ctx = [ordered]@{
        Proceso     = $p
        ProcessId   = $p.Id
        Login       = ($tops | Where-Object { $_.Class -eq 'ThunderRT6FormDC' -and $_.Text -like '*Administraci*n de Clientes (AC)*' } | Select-Object -First 1)
        Principal   = ($tops | Where-Object { $_.Class -eq 'ThunderRT6MDIForm' } | Select-Object -First 1)
        Tickler     = ($tops | Where-Object { $_.Text -like '*Solicitud Tickler*' } | Select-Object -First 1)
        Dialogos    = @($tops | Where-Object { $_.Class -eq '#32770' })
    }

    if ($ctx.Principal) {
        $hijos = Get-ChildHandles -Parent $ctx.Principal.Handle
        $ctx.MDIClient = ($hijos | Where-Object { $_.Class -eq 'MDIClient' } | Select-Object -First 1)

        $panel = $hijos | Where-Object { $_.Class -eq 'ThunderRT6FormDC' -and $_.Text -eq 'Criterios' } | Select-Object -First 1
        $ctx.PanelCriterios = $panel
        if ($panel) {
            $pc = Get-ChildHandles -Parent $panel.Handle
            $ctx.ComboCriterio = ($pc | Where-Object { $_.Class -like '*ComboBox*' } | Select-Object -First 1)
            $ctx.EditCriterio  = ($pc | Where-Object { $_.Class -eq 'Edit' } | Select-Object -First 1)
            $ctx.RadioMin      = ($pc | Where-Object { $_.Class -like '*OptionButton*' -and $_.Text -like '*MIN*MSISDN*' } | Select-Object -First 1)
            $ctx.BotonBuscar   = ($pc | Where-Object { $_.Class -eq 'ThunderRT6UserControlDC' -and $_.Visible -and $_.W -gt 60 } | Select-Object -First 1)
        }

        if ($ctx.MDIClient) {
            $mdi = Get-ChildHandles -Parent $ctx.MDIClient.Handle |
                   Where-Object { $_.Visible -and $_.Class -eq 'ThunderRT6FormDC' }
            $ctx.VentanasMDI = @($mdi)
            $ctx.Resultados  = ($mdi | Where-Object { $_.Text -like 'Resultados de la busqueda*' } | Select-Object -First 1)
            $ctx.Ficha       = ($mdi | Where-Object { $_.Text -like 'Datos del Cliente*' } | Select-Object -First 1)
            $ctx.Historial   = ($mdi | Where-Object { $_.Text -eq 'HISTORIAL' } | Select-Object -First 1)
        }
    }
    return [pscustomobject]$ctx
}

# ---------------------------------------------------------------------
#  Inicio y cierre de sesion
# ---------------------------------------------------------------------

function Stop-ACSession {
    <#  Cierra una instancia concreta o todas.  #>
    param([int]$ProcessId = 0, [switch]$Todas)

    if ($Todas) {
        foreach ($p in (Get-ACProcesses)) { try { $p.Kill() } catch { } }
        Start-Sleep -Seconds 2
        $null = Wait-Condition -TimeoutSeg 15 -Condicion { if ((Get-ACProcesses).Count -eq 0) { $true } }
        return
    }

    $p = Get-ACProcess -ProcessId $ProcessId
    if ($p) { try { $p.Kill(); Start-Sleep -Seconds 3 } catch { } }
    $null = Wait-Condition -TimeoutSeg 15 -Condicion {
        if (-not (Get-ACProcess -ProcessId $ProcessId)) { $true }
    }
}

function Start-ACInstancia {
    <#  Lanza UNA instancia nueva de AC sin tocar las que ya esten corriendo,
        e inicia sesion usando solo mensajes (sin raton ni teclado globales).
        Devuelve el contexto de esa instancia.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Password,
        [string]$BaseDatos = "AC_PRODUCCION",
        [int]$TimeoutSeg = 300,
        [switch]$Silencioso
    )
    $exe = Get-ACExePath
    $p = Start-Process -FilePath $exe -WorkingDirectory (Split-Path -Parent $exe) -PassThru
    return (Connect-ACInstancia -ProcessId $p.Id -Password $Password -BaseDatos $BaseDatos `
                                -TimeoutSeg $TimeoutSeg -Silencioso:$Silencioso)
}

function Start-ACSession {
    <#  Deja UNA sola instancia de AC con sesion iniciada (cierra las previas).
        Es lo que usa la pasada del historial, que no se puede paralelizar.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Password,
        [string]$BaseDatos = "AC_PRODUCCION",
        [int]$TimeoutSeg = 300
    )

    if ((Get-ACProcesses).Count -gt 0) {
        Write-Paso "Cerrando instancias previas de AC..."
        Stop-ACSession -Todas
    }
    $exe = Get-ACExePath
    Write-Paso "Abriendo AC: $exe"
    $p = Start-Process -FilePath $exe -WorkingDirectory (Split-Path -Parent $exe) -PassThru
    return (Connect-ACInstancia -ProcessId $p.Id -Password $Password -BaseDatos $BaseDatos -TimeoutSeg $TimeoutSeg)
}

function Connect-ACInstancia {
    <#  Inicia sesion en una instancia de AC ya lanzada.  #>
    param(
        [Parameter(Mandatory=$true)][int]$ProcessId,
        [Parameter(Mandatory=$true)][string]$Password,
        [string]$BaseDatos = "AC_PRODUCCION",
        # AC carga decenas de parametros ("Switch") contra Oracle uno por uno al
        # arrancar. Con la base descargada tarda ~20 s, pero el 22-sep tardaba
        # cerca de 3 minutos: con el plazo en 120 s el script se rendia justo
        # antes de que terminara e informaba "AC no respondio al iniciar
        # sesion", que era falso. En el log de AC se ve la sesion autenticandose
        # y conectando bien unos segundos despues de que el script se rindio.
        # Esperar de mas no cuesta nada: en cuanto aparece la ventana principal
        # se sigue de largo.
        [int]$TimeoutSeg = 300,
        [switch]$Silencioso
    )

    function Log { param($m, $n = "INFO") if (-not $Silencioso) { Write-Paso $m $n } }

    $ctx = Wait-Condition -TimeoutSeg $TimeoutSeg -Condicion {
        $c = Get-ACContext -ProcessId $ProcessId
        if ($c -and $c.Login) { $c }
    }
    if (-not $ctx) { throw "No aparecio la ventana de conexion de AC (PID $ProcessId)." }

    Set-WindowFocus -Handle $ctx.Login.Handle
    $ctrls = Get-ChildHandles -Parent $ctx.Login.Handle

    # Base de datos: el combo es de solo lectura; CB_SETCURSEL basta porque el
    # formulario lee el valor al pulsar Aceptar.
    $combo = $ctrls | Where-Object { $_.Class -like '*ComboBox*' } | Select-Object -First 1
    if (-not $combo) { throw "No se encontro el combo de Base de Datos." }
    $items = Get-ComboItems -Handle $combo.Handle
    $idx = [array]::IndexOf($items, $BaseDatos)
    if ($idx -lt 0) { throw "Base de datos '$BaseDatos' no disponible. Opciones: $($items -join ', ')" }
    [void][W32]::SendMessage($combo.Handle, [W32]::CB_SETCURSEL, [IntPtr]$idx, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 300
    $sel = [int64][W32]::SendMessage($combo.Handle, [W32]::CB_GETCURSEL, [IntPtr]::Zero, [IntPtr]::Zero)
    if ($sel -ne $idx) { throw "No se pudo seleccionar la base de datos '$BaseDatos'." }
    Log "Base de datos: $BaseDatos"

    # Contrasena: unico TextBox habilitado (Dominio y Usuario vienen bloqueados)
    $txtPass = $ctrls | Where-Object { $_.Class -like '*TextBox*' -and $_.Enabled -and $_.Visible } | Select-Object -First 1
    if (-not $txtPass) { throw "No se encontro el campo de contrasena." }

    [void](Set-CtrlText -Handle $txtPass.Handle -Text $Password)
    $escrito = Get-UiaName -Handle ([int]$txtPass.Handle)

    if ($escrito -ne $Password) {
        # Respaldo: escribir tecla por tecla. El campo se verifica siempre antes
        # de pulsar Aceptar para no gastar intentos contra el bloqueo de cuenta.
        Log "WM_SETTEXT no cargo el campo; escribiendo por teclado..." "WARN"
        Set-WindowFocus -Handle $ctx.Login.Handle
        Invoke-ClickControl -Handle $txtPass.Handle -SettleMs 300
        [System.Windows.Forms.SendKeys]::SendWait("{END}")
        for ($k = 0; $k -lt 40; $k++) { [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}"); Start-Sleep -Milliseconds 15 }
        foreach ($ch in $Password.ToCharArray()) {
            # solo +^%~(){}[] requieren llaves en SendKeys
            $t = if ('+^%~(){}[]'.Contains($ch)) { '{' + $ch + '}' } else { [string]$ch }
            [System.Windows.Forms.SendKeys]::SendWait($t)
            Start-Sleep -Milliseconds 60
        }
        Start-Sleep -Milliseconds 300
        $escrito = Get-UiaName -Handle ([int]$txtPass.Handle)
    }

    if ($escrito -ne $Password) {
        throw ("La contrasena no quedo escrita correctamente en el formulario " +
               "(el campo tiene $("$escrito".Length) caracteres y la clave tiene $($Password.Length)). " +
               "No se pulso Aceptar para no gastar intentos de inicio de sesion.")
    }
    Log "Credenciales cargadas y verificadas en el formulario"

    # Aceptar: de los dos UserControl del pie, el de menor X
    $botones = $ctrls | Where-Object { $_.Class -eq 'ThunderRT6UserControlDC' -and $_.Visible } | Sort-Object X
    if ($botones.Count -lt 1) { throw "No se encontro el boton Aceptar." }

    # Aceptar solo responde al raton real (probado: por mensaje no reacciona, y
    # ademas deja el control en un estado que ignora el clic fisico posterior).
    # No es problema para el paralelismo: el login ocurre una sola vez por
    # instancia y se hace de a una; las consultas si van por mensajes.
    Invoke-ClickControl -Handle $botones[0].Handle

    $res = Wait-Condition -TimeoutSeg $TimeoutSeg -Condicion {
        $c = Get-ACContext -ProcessId $ProcessId
        if (-not $c) { return $null }
        if ($c.Principal) { return @{ Ok = $true; Ctx = $c } }
        if ($c.Dialogos -and $c.Dialogos.Count -gt 0) { return @{ Ok = $false; Ctx = $c } }
    }
    if (-not $res) { throw "AC no respondio al iniciar sesion (tiempo de espera agotado)." }

    if (-not $res.Ok) {
        $msg = ($res.Ctx.Dialogos | ForEach-Object { Read-DialogText -Handle $_.Handle }) -join ' | '
        [void](Close-ACDialogs -ProcessId $ProcessId)
        throw "AC rechazo el inicio de sesion: $msg"
    }

    # dar tiempo a que cargue el perfil y el panel de criterios
    $ctx = Wait-Condition -TimeoutSeg ([Math]::Max(90, $TimeoutSeg)) -Condicion {
        $c = Get-ACContext -ProcessId $ProcessId
        if ($c -and $c.PanelCriterios -and $c.BotonBuscar) { $c }
    }
    if (-not $ctx) { throw "AC inicio sesion pero no cargo el panel de criterios de busqueda." }

    Log "Sesion iniciada en AC (PID $ProcessId)" "OK"
    return $ctx
}

# ---------------------------------------------------------------------
#  Paso 3 y 4: consultar el numero y leer la grilla de resultados
# ---------------------------------------------------------------------

function Close-ACResultados {
    param($Contexto, [int]$ProcessId = 0)
    if ($ProcessId -eq 0 -and $Contexto) { $ProcessId = $Contexto.ProcessId }
    $c = Get-ACContext -ProcessId $ProcessId
    if ($c -and $c.Resultados) {
        [void][W32]::PostMessage($c.Resultados.Handle, [W32]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 600
    }
}

function Invoke-ACBusqueda {
    <#  Escribe el numero, pulsa BUSCAR y devuelve las filas de la grilla.
        Todo por mensajes: no usa raton ni teclado, funciona con AC en segundo
        plano y no interfiere con otras instancias.
        No abre la ficha del cliente: no genera ningun Tickler.  #>
    param(
        [Parameter(Mandatory=$true)][string]$Numero,
        [int]$ProcessId = 0,
        [int]$TimeoutSeg = 45
    )

    $ctx = Get-ACContext -ProcessId $ProcessId
    if (-not $ctx -or -not $ctx.Principal) { throw "La sesion de AC no esta disponible." }
    if ($ProcessId -eq 0) { $ProcessId = $ctx.ProcessId }
    if ($ctx.Ficha) { throw "Hay una ficha de cliente abierta; AC no permite buscar hasta cerrarla." }

    if ($ctx.Resultados) { Close-ACResultados -ProcessId $ProcessId; $ctx = Get-ACContext -ProcessId $ProcessId }

    # Criterio MIN / MSISDN
    if ($ctx.RadioMin) { Invoke-PostClick -Handle $ctx.RadioMin.Handle -SettleMs 150 }

    # El numero se escribe con WM_SETTEXT: no depende del foco del teclado,
    # que en AC se desvia a un PictureBox al hacer clic en el panel.
    if (-not $ctx.EditCriterio) { throw "No se encontro el campo Criterio." }

    # AC deja el numero de la busqueda anterior en el campo, y a veces lo repone
    # despues de escribir: si se da por buena la primera escritura se termina
    # consultando el numero equivocado, o falla con "quedo '<otro numero>'".
    # Se vacia primero y se reintenta verificando lo que quedo.
    $leido = ''
    for ($intento = 1; $intento -le 4; $intento++) {
        [void](Set-CtrlText -Handle $ctx.EditCriterio.Handle -Text '')
        Start-Sleep -Milliseconds 80
        $leido = Set-CtrlText -Handle $ctx.EditCriterio.Handle -Text $Numero
        if ($leido -eq $Numero) {
            # releer tras una pausa: si AC lo repone, aqui se nota
            Start-Sleep -Milliseconds 150
            $leido = Get-CtrlText -Handle $ctx.EditCriterio.Handle
            if ("$leido".Trim() -eq $Numero) { break }
        }
        Start-Sleep -Milliseconds 250
    }
    if ("$leido".Trim() -ne $Numero) { throw "No se pudo escribir el numero en el campo Criterio (quedo '$leido')." }

    Invoke-PostClick -Handle $ctx.BotonBuscar.Handle -SettleMs 150

    # Esperar la ventana de resultados o un dialogo
    $r = Wait-Condition -TimeoutSeg $TimeoutSeg -IntervaloMs 150 -Condicion {
        $c = Get-ACContext -ProcessId $ProcessId
        if ($c.Resultados) { return @{ Tipo = 'Resultados'; Ctx = $c } }
        if ($c.Dialogos -and $c.Dialogos.Count -gt 0) { return @{ Tipo = 'Dialogo'; Ctx = $c } }
        return $null
    }

    if (-not $r) {
        return [pscustomobject]@{ Encontrado = $false; Mensaje = "AC no devolvio resultados ni mensaje (tiempo agotado)."; Filas = @() }
    }

    if ($r.Tipo -eq 'Dialogo') {
        $msgs = Close-ACDialogs -ProcessId $r.Ctx.ProcessId
        return [pscustomobject]@{ Encontrado = $false; Mensaje = ($msgs -join ' | '); Filas = @() }
    }

    # ----- Leer la ventana de resultados -----
    $ctx = $r.Ctx

    # La ventana aparece antes de que sus controles esten poblados: hay que
    # esperar a que el arbol tenga nodos, no basta con que exista la ventana.
    $ctrls = Wait-Condition -TimeoutSeg 20 -IntervaloMs 150 -Condicion {
        $k = Get-ChildHandles -Parent $ctx.Resultados.Handle
        if (($k | Where-Object { $_.Class -like 'ListView*' }) -and
            ($k | Where-Object { $_.Class -like 'TreeView*' })) { $k }
    }
    if (-not $ctrls) { throw "La ventana de resultados no cargo sus grillas." }

    $tv = $ctrls | Where-Object { $_.Class -like 'TreeView*' } | Select-Object -First 1
    $lv = $ctrls | Where-Object { $_.Class -like 'ListView*' } | Select-Object -First 1

    $nodos = @(Wait-Condition -TimeoutSeg 20 -IntervaloMs 150 -Condicion {
        $n = @(Get-TreeViewItems -Hwnd $tv.Handle)
        if ($n.Count -gt 0) { ,$n }
    })

    # Paso 4: seleccionar la linea en el arbol para que se llene la grilla.
    # TVM_SELECTITEM no dispara el evento de AC; si lo hace un clic enviado
    # como mensaje al propio TreeView. Se reintenta por si el arbol aun pinta.
    $datos = $null
    if ($nodos.Count -gt 0) {
        for ($intento = 1; $intento -le 4 -and -not $datos; $intento++) {
            $tvAhora = Get-ChildHandles -Parent $ctx.Resultados.Handle |
                       Where-Object { $_.Class -like 'TreeView*' } | Select-Object -First 1
            if (-not $tvAhora) { break }
            Invoke-PostClick -Handle $tvAhora.Handle -X 60 -Y 10 -SettleMs 150
            $datos = Wait-Condition -TimeoutSeg 6 -IntervaloMs 150 -Condicion {
                $d = Get-ListViewData -Hwnd $lv.Handle
                if ($d.Rows.Count -gt 0) { $d }
            }
        }
    }

    if (-not $datos) {
        $msgs = Close-ACDialogs -ProcessId $ProcessId
        return [pscustomobject]@{
            Encontrado = $false
            Mensaje    = if ($msgs) { $msgs -join ' | ' } else { "La busqueda no devolvio ninguna linea." }
            Filas      = @()
            Nodos      = $nodos
        }
    }

    return [pscustomobject]@{
        Encontrado = $true
        Mensaje    = ""
        Columnas   = $datos.Columns
        Filas      = $datos.Rows
        Nodos      = $nodos
    }
}

# ---------------------------------------------------------------------
#  Paso 5: abrir la ficha del suscriptor
# ---------------------------------------------------------------------

function Open-ACFicha {
    param([int]$TimeoutSeg = 60)
    $ctx = Get-ACContext
    if (-not $ctx.Resultados) { throw "No hay ventana de resultados abierta." }

    $ctrls = Get-ChildHandles -Parent $ctx.Resultados.Handle
    # De los UserControl del pie, "Consultar" es el de menor X (el otro es "Salir");
    # se descarta la X de cierre por su tamano.
    $botones = @($ctrls | Where-Object { $_.Class -eq 'ThunderRT6UserControlDC' -and $_.Visible -and $_.W -gt 60 } | Sort-Object X)
    if ($botones.Count -lt 1) { throw "No se encontro el boton Consultar." }

    # Este es el UNICO paso que no funciona por mensaje: probado, Consultar no
    # reacciona a WM_LBUTTONDOWN/UP y exige el raton real. Por eso la pasada
    # del historial no se puede paralelizar.
    Set-WindowFocus -Handle $ctx.Principal.Handle
    Invoke-ClickControl -Handle $botones[0].Handle -SettleMs 800

    $r = Wait-Condition -TimeoutSeg $TimeoutSeg -Condicion {
        $c = Get-ACContext
        if ($c.Ficha) { return $c }
        if ($c.Dialogos -and $c.Dialogos.Count -gt 0) { return $c }
    }
    if (-not $r) { throw "No se abrio la ficha del cliente." }
    if (-not $r.Ficha) {
        $msgs = Close-ACDialogs -ProcessId $r.ProcessId
        throw "AC no abrio la ficha: $($msgs -join ' | ')"
    }
    return $r
}

function Get-ACDatosFicha {
    <#  Lee los campos de la ficha por UI Automation, ubicandolos por su
        posicion relativa al formulario (los handles cambian en cada consulta).  #>
    param([Parameter(Mandatory=$true)]$Ficha)

    $campos = [ordered]@{}
    try {
        $cond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty, [int]$Ficha.Handle)
        $form = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants, $cond)
        if (-not $form) { return [pscustomobject]$campos }

        $todos = $form.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                               [System.Windows.Automation.Condition]::TrueCondition)
        $items = @()
        foreach ($e in $todos) {
            try {
                $r = $e.Current.BoundingRectangle
                $items += [pscustomobject]@{
                    Nombre = $e.Current.Name
                    Clase  = $e.Current.ClassName
                    RelX   = [int]($r.X - $Ficha.X)
                    RelY   = [int]($r.Y - $Ficha.Y)
                }
            } catch { }
        }

        # posiciones relativas estables dentro del formulario
        $mapa = @{
            'EstadoContrato' = @(279, 80)   # etiqueta resaltada junto a "Estado Contrato"
            'Nombre'         = @(105, 107)
            'Apellido'       = @(105, 130)
            'Identificacion' = @(105, 153)
            'Plan'           = @(455, 106)
            'TipoCliente'    = @(119, 435)
            'FechaActivacion'= @(120, 383)
            'ComportamientoPago' = @(120, 462)
        }
        foreach ($k in $mapa.Keys) {
            $x = $mapa[$k][0]; $y = $mapa[$k][1]
            $cand = $items | Where-Object { [Math]::Abs($_.RelX - $x) -le 12 -and [Math]::Abs($_.RelY - $y) -le 12 -and $_.Nombre } |
                    Select-Object -First 1
            $campos[$k] = if ($cand) { $cand.Nombre } else { "" }
        }
    } catch { }
    return [pscustomobject]$campos
}

# ---------------------------------------------------------------------
#  Paso 6: Ctrl+Shift+H -> HISTORIAL
# ---------------------------------------------------------------------

function Get-ACHistorial {
    <#  Abre el historial con Ctrl+Shift+H sobre la ficha y lee las dos grillas:
        ESTADO CONTRATO y DISPOSITIVOS (Estado del ESN/ICCID).  #>
    param([int]$TimeoutSeg = 45)

    $ctx = Get-ACContext
    if (-not $ctx.Ficha) { throw "No hay ficha de cliente abierta." }

    # La ficha se muestra antes de terminar de cargar sus datos, y mientras
    # tanto ignora el atajo. Se espera a que los campos tengan contenido.
    [void](Wait-Condition -TimeoutSeg 30 -Condicion {
        $f = (Get-ACContext).Ficha
        if (-not $f) { return $null }
        $conTexto = @(Get-ChildHandles -Parent $f.Handle |
                      Where-Object { $_.Class -like '*TextBox*' } |
                      ForEach-Object { Get-UiaName -Handle ([int]$_.Handle) } |
                      Where-Object { $_ })
        if ($conTexto.Count -ge 5) { $true }
    })
    Start-Sleep -Milliseconds 800

    # El atajo puede perderse si AC sigue ocupado: se reintenta.
    $ctxH = $null
    for ($intento = 1; $intento -le 4 -and -not $ctxH; $intento++) {
        $ficha = (Get-ACContext).Ficha
        if (-not $ficha) { throw "La ficha del cliente se cerro inesperadamente." }
        Set-WindowFocus -Handle $ctx.Principal.Handle
        Invoke-ClickAt -X ($ficha.X + 30) -Y ($ficha.Y + 117) -SettleMs 400
        Send-Hotkey -Modifiers @($script:VK_CONTROL, $script:VK_SHIFT) -Key $script:VK_H
        $ctxH = Wait-Condition -TimeoutSeg ([Math]::Max(8, [int]($TimeoutSeg / 4))) -Condicion {
            $c = Get-ACContext
            if ($c.Historial) { $c }
        }
        if (-not $ctxH) { Write-Paso "El atajo Ctrl+Shift+H no respondio (intento $intento); reintentando..." "WARN" }
    }
    if (-not $ctxH) { throw "No se abrio la ventana HISTORIAL (Ctrl+Shift+H)." }
    $ctx = $ctxH

    $ctrls = Get-ChildHandles -Parent $ctx.Historial.Handle
    $grids = @($ctrls | Where-Object { $_.Class -like 'ListView*' } | Sort-Object Y)
    if ($grids.Count -lt 1) { throw "No se encontraron las grillas del HISTORIAL." }

    $estado = Get-ListViewData -Hwnd $grids[0].Handle
    $disp   = if ($grids.Count -gt 1) { Get-ListViewData -Hwnd $grids[1].Handle } else { $null }

    # cabecera del historial (custcode, nombre, min, numero de contrato)
    $textos = @($ctrls | Where-Object { $_.Class -like '*TextBox*' } | Sort-Object Y, X |
                ForEach-Object { Get-UiaName -Handle ([int]$_.Handle) })

    return [pscustomobject]@{
        VentanaHandle   = $ctx.Historial.Handle
        Cabecera        = $textos
        EstadoContrato  = $estado
        Dispositivos    = $disp
    }
}

function Get-ACIndicadoresCobranza {
    <#  Lee las casillas de cobranza de la ficha abierta.

        "En Demanda" y "Castigada" son las que importan para una linea
        suspendida: dicen si la deuda ya paso a cobro juridico o si el operador
        la dio por perdida. En esos casos la venta no se recupera contactando
        al cliente.  #>
    $ctx = Get-ACContext
    if (-not $ctx.Ficha) { throw "No hay ficha de cliente abierta." }

    $BM_GETCHECK = 0x00F0
    $k = Get-ChildHandles -Parent $ctx.Ficha.Handle
    $r = [ordered]@{}
    foreach ($n in @('Respon. Pago', 'Flag No Cobrar', 'En Demanda', 'Castigada')) {
        $cb = $k | Where-Object { $_.Class -like '*CheckBox*' -and "$($_.Text)".Trim() -eq $n } | Select-Object -First 1
        $r[$n] = if ($cb) {
            if ([int][W32]::SendMessage($cb.Handle, $BM_GETCHECK, [IntPtr]::Zero, [IntPtr]::Zero) -eq 1) { 'SI' } else { 'NO' }
        } else { '' }
    }
    return [pscustomobject]$r
}

function Get-ACSaldo {
    <#  Pulsa el boton "Saldo" de la ficha abierta y lee lo que debe la linea.

        La ventana trae tres importes, en este orden de arriba a abajo:
        Saldo Equipo, Saldo Servicios y Saldo Total. Las etiquetas son labels
        de VB6 -no son ventanas reales y no se pueden leer por mensaje- asi que
        los campos se identifican por su posicion vertical, que es fija.

        Solo lectura: abrir el saldo no modifica nada.  #>
    param([int]$TimeoutSeg = 30)

    $ctx = Get-ACContext
    if (-not $ctx.Ficha) { throw "No hay ficha de cliente abierta." }

    $k = Get-ChildHandles -Parent $ctx.Ficha.Handle
    $btn = $k | Where-Object { $_.Class -like '*CommandButton*' -and "$($_.Text)".Trim() -eq 'Saldo' } |
           Select-Object -First 1
    if (-not $btn) { throw "La ficha no tiene boton 'Saldo' (perfil sin acceso?)." }

    $antes = @(Get-ChildHandles -Parent $ctx.MDIClient.Handle | ForEach-Object { $_.Handle })
    # Como Consultar, este boton exige el raton real
    Invoke-ClickControl -Handle $btn.Handle -SettleMs 500

    $vent = Wait-Condition -TimeoutSeg $TimeoutSeg -IntervaloMs 300 -Condicion {
        $v = @(Get-ChildHandles -Parent $ctx.MDIClient.Handle |
               Where-Object { $antes -notcontains $_.Handle -and $_.Visible }) | Select-Object -First 1
        if (-not $v) {
            $v = @(Get-TopWindows -ProcessId $ctx.ProcessId |
                   Where-Object { $_.Handle -ne $ctx.Principal.Handle -and $_.Visible -and $_.Class -eq 'ThunderRT6FormDC' }) |
                 Select-Object -First 1
        }
        if (-not $v) { return $null }
        $tb = @(Get-ChildHandles -Parent $v.Handle | Where-Object { $_.Class -like '*TextBox*' -and $_.Visible })
        if ($tb.Count -ge 3) { return @{ Vent = $v; Campos = $tb } }
    }
    if (-not $vent) { throw "No aparecio la ventana de saldo." }

    $orden = @($vent.Campos | Sort-Object Y)
    function Num { param($s)
        $t = "$s" -replace '[^\d,.\-]', ''
        $t = $t -replace ',', ''
        $d = 0.0
        if ([double]::TryParse($t, [ref]$d)) { return $d }
        return $null
    }
    $eq = Get-CtrlText -Handle $orden[0].Handle
    $sv = Get-CtrlText -Handle $orden[1].Handle
    $to = Get-CtrlText -Handle $orden[2].Handle

    return [pscustomobject]@{
        SaldoEquipo    = "$eq".Trim()
        SaldoServicios = "$sv".Trim()
        SaldoTotal     = "$to".Trim()
        EquipoNum      = Num $eq
        ServiciosNum   = Num $sv
        TotalNum       = Num $to
    }
}
