# =====================================================================
#  Remote.ps1 - Lectura de ListView / TreeView de OTRO proceso
#
#  Los grids de AC son common controls (ListView20WndClass / TreeView20WndClass)
#  que NO exponen su contenido por UI Automation. La unica forma de leerlos
#  como datos -y no por OCR- es enviarles los mensajes LVM_* / TVM_* con un
#  buffer reservado dentro del propio proceso de AC.
#
#  AC es un proceso de 32 bits: las estructuras LVITEM / TVITEM / LVCOLUMN
#  se construyen a mano con punteros de 4 bytes, independientemente de que
#  PowerShell corra en 64 bits.
# =====================================================================

if (-not ("RP" -as [type])) {
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class RP {
    [DllImport("kernel32.dll")]
    public static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
    [DllImport("kernel32.dll")]
    public static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")]
    public static extern IntPtr VirtualAllocEx(IntPtr hProc, IntPtr addr, uint size, uint type, uint protect);
    [DllImport("kernel32.dll")]
    public static extern bool VirtualFreeEx(IntPtr hProc, IntPtr addr, uint size, uint type);
    [DllImport("kernel32.dll")]
    public static extern bool WriteProcessMemory(IntPtr hProc, IntPtr addr, byte[] buf, uint size, out UIntPtr written);
    [DllImport("kernel32.dll")]
    public static extern bool ReadProcessMemory(IntPtr hProc, IntPtr addr, byte[] buf, uint size, out UIntPtr read);
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam,
                                                   uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@
}

# SendMessage se queda bloqueado para siempre si la ventana destino dejo de
# atender su cola de mensajes. Con varias instancias eso es fatal: el bucle es
# de un solo hilo, asi que una instancia colgada congela a todas las demas.
#
# OJO con SMTO_ABORTIFHUNG (0x0002): NO se usa a proposito. Windows da por
# "colgada" a toda ventana que lleve unos 5 s sin atender su cola, y AC es VB6:
# congela su propia ventana durante cada consulta a Oracle. Con esa bandera el
# sondeo devolvia error mientras AC estaba trabajando normalmente, se contaban
# esas respuestas como instancia muda, se la mataba a media consulta y el
# numero salia como "NO ENCONTRADO" siendo falso. Paso el 22-sep: el 6% de una
# tanda salio mal por esto. SMTO_NORMAL espera el plazo completo, que es lo
# correcto: lo que se quiere es no bloquearse para siempre, no rendirse pronto.
$script:SMTO_NORMAL = 0x0000
# Holgado a proposito: AC tarda en contestar mientras consulta, y un plazo
# corto solo produce falsos negativos.
$script:SONDEO_MS   = 5000

$script:PROC_ACCESS  = 0x0008 -bor 0x0010 -bor 0x0020 -bor 0x0400   # VM_OPERATION|VM_READ|VM_WRITE|QUERY_INFO
$script:MEM_RESERVE  = 0x1000 -bor 0x2000
$script:PAGE_RW      = 0x04
$script:MEM_RELEASE  = 0x8000

# Mensajes
$script:LVM_GETITEMCOUNT = 0x1004
$script:LVM_GETITEMTEXTA = 0x102D
$script:LVM_GETITEMSTATE = 0x102C
$script:LVM_GETCOLUMNA   = 0x1019
$script:LVIS_SELECTED    = 0x0002
$script:TVM_GETNEXTITEM  = 0x110A
$script:TVM_GETITEMA     = 0x110C
$script:TVGN_ROOT        = 0x0000
$script:TVGN_NEXT        = 0x0001
$script:TVGN_CHILD       = 0x0004

function Open-TargetProcess {
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
    $procId = 0
    [void][RP]::GetWindowThreadProcessId($Hwnd, [ref]$procId)
    $h = [RP]::OpenProcess($script:PROC_ACCESS, $false, $procId)
    if ($h -eq [IntPtr]::Zero) { throw "No se pudo abrir el proceso $procId (codigo $([Runtime.InteropServices.Marshal]::GetLastWin32Error()))" }
    return $h
}

function New-RemoteBuffer {
    param([Parameter(Mandatory=$true)][IntPtr]$HProc, [uint32]$Size = 8192)
    $p = [RP]::VirtualAllocEx($HProc, [IntPtr]::Zero, $Size, $script:MEM_RESERVE, $script:PAGE_RW)
    if ($p -eq [IntPtr]::Zero) { throw "VirtualAllocEx fallo" }
    return $p
}

function Read-RemoteString {
    param([Parameter(Mandatory=$true)][IntPtr]$HProc, [Parameter(Mandatory=$true)][IntPtr]$Addr, [int]$Bytes = 256)
    $buf  = New-Object byte[] $Bytes
    $read = [UIntPtr]::Zero
    [void][RP]::ReadProcessMemory($HProc, $Addr, $buf, [uint32]$Bytes, [ref]$read)
    $s = [System.Text.Encoding]::Default.GetString($buf)
    $z = $s.IndexOf([char]0)
    if ($z -ge 0) { $s = $s.Substring(0, $z) }
    return $s.Trim()
}

function Send-MensajeConPlazo {
    <#  Como SendMessage pero sin poder colgarse: si la ventana no contesta en
        el plazo, devuelve -1 en vez de esperar indefinidamente.

        Es lo que hace la diferencia entre "una instancia de AC se atasco" y
        "se atasco la corrida entera": el bucle de Invoke-ACConsultaMasiva es de
        un solo hilo y atiende a las instancias por turnos, de modo que un
        SendMessage bloqueado en una de ellas deja a las otras sin atencion, sus
        numeros vencen por tiempo y la corrida se desploma. Visto el 22-sep en
        la corrida de julio: los tiempos pasaron de 10 s a mas de 300 s por
        numero con un plazo configurado de 75 s, justo porque el vencimiento
        tampoco alcanzaba a evaluarse.  #>
    param(
        [Parameter(Mandatory=$true)][IntPtr]$Hwnd,
        [Parameter(Mandatory=$true)][uint32]$Mensaje,
        [IntPtr]$WParam = [IntPtr]::Zero,
        [IntPtr]$LParam = [IntPtr]::Zero,
        [int]$PlazoMs = 0
    )
    if ($PlazoMs -le 0) { $PlazoMs = $script:SONDEO_MS }
    $res = [IntPtr]::Zero
    $ok = [RP]::SendMessageTimeout($Hwnd, $Mensaje, $WParam, $LParam,
                                   $script:SMTO_NORMAL, [uint32]$PlazoMs, [ref]$res)
    if ($ok -eq [IntPtr]::Zero) { return -1 }
    return [int]$res
}

function Get-ListViewRowCount {
    <#  Cuenta las filas con UN solo mensaje y sin reservar memoria remota.
        Devuelve -1 si la ventana no contesta dentro del plazo.

        Es lo que se usa para sondear: llamar a Get-ListViewData en cada ciclo
        satura el bucle de mensajes de AC (decenas de SendMessage sincronos por
        segundo mas VirtualAllocEx) y llega a impedir que procese los clics que
        se le envian.  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd, [int]$PlazoMs = 0)
    return Send-MensajeConPlazo -Hwnd $Hwnd -Mensaje $script:LVM_GETITEMCOUNT -PlazoMs $PlazoMs
}

function Get-TreeViewCount {
    <#  Numero de nodos del arbol con un solo mensaje (TVM_GETCOUNT).
        Devuelve -1 si la ventana no contesta dentro del plazo.  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd, [int]$PlazoMs = 0)
    return Send-MensajeConPlazo -Hwnd $Hwnd -Mensaje 0x1105 -PlazoMs $PlazoMs
}

function Test-VentanaResponde {
    <#  Sondeo barato para saber si una ventana sigue atendiendo mensajes.  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd, [int]$PlazoMs = 0)
    # WM_NULL: no hace nada, solo comprueba que la cola se atiende
    return (Send-MensajeConPlazo -Hwnd $Hwnd -Mensaje 0x0000 -PlazoMs $PlazoMs) -ne -1
}

function Get-ListViewColumns {
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd, [int]$Max = 20)
    $hProc = Open-TargetProcess -Hwnd $Hwnd
    try {
        $pRem  = New-RemoteBuffer -HProc $hProc -Size 4096
        $pText = [IntPtr]([int64]$pRem + 256)
        try {
            $cols = @()
            for ($c = 0; $c -lt $Max; $c++) {
                # LVCOLUMN (32 bits): mask(0) fmt(4) cx(8) pszText(12) cchTextMax(16) iSubItem(20)
                $lc = New-Object byte[] 44
                [Array]::Copy([BitConverter]::GetBytes([uint32]0x0004), 0, $lc, 0, 4)     # LVCF_TEXT
                [Array]::Copy([BitConverter]::GetBytes([uint32]$pText.ToInt32()), 0, $lc, 12, 4)
                [Array]::Copy([BitConverter]::GetBytes([int32]200), 0, $lc, 16, 4)
                $w = [UIntPtr]::Zero
                [void][RP]::WriteProcessMemory($hProc, $pRem, $lc, 44, [ref]$w)
                if ([RP]::SendMessage($Hwnd, $script:LVM_GETCOLUMNA, [IntPtr]$c, $pRem) -eq [IntPtr]::Zero) { break }
                $cols += (Read-RemoteString -HProc $hProc -Addr $pText -Bytes 200)
            }
            return $cols
        } finally { [void][RP]::VirtualFreeEx($hProc, $pRem, 0, $script:MEM_RELEASE) }
    } finally { [void][RP]::CloseHandle($hProc) }
}

function Get-ListViewData {
    <#  Devuelve un objeto con Columns (nombres reales) y Rows (filas con celdas).  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)

    $cols  = Get-ListViewColumns -Hwnd $Hwnd
    $nCols = if ($cols.Count -gt 0) { $cols.Count } else { 10 }

    $hProc = Open-TargetProcess -Hwnd $Hwnd
    try {
        $pRem  = New-RemoteBuffer -HProc $hProc -Size 8192
        $pText = [IntPtr]([int64]$pRem + 512)
        try {
            $nRows = [int][RP]::SendMessage($Hwnd, $script:LVM_GETITEMCOUNT, [IntPtr]::Zero, [IntPtr]::Zero)
            $rows = @()
            for ($r = 0; $r -lt $nRows; $r++) {
                $cells = @()
                for ($c = 0; $c -lt $nCols; $c++) {
                    # LVITEM (32 bits): mask(0) iItem(4) iSubItem(8) state(12) stateMask(16) pszText(20) cchTextMax(24)
                    $lv = New-Object byte[] 60
                    [Array]::Copy([BitConverter]::GetBytes([uint32]0x0001), 0, $lv, 0, 4)   # LVIF_TEXT
                    [Array]::Copy([BitConverter]::GetBytes([int32]$r), 0, $lv, 4, 4)
                    [Array]::Copy([BitConverter]::GetBytes([int32]$c), 0, $lv, 8, 4)
                    [Array]::Copy([BitConverter]::GetBytes([uint32]$pText.ToInt32()), 0, $lv, 20, 4)
                    [Array]::Copy([BitConverter]::GetBytes([int32]256), 0, $lv, 24, 4)
                    $w = [UIntPtr]::Zero
                    [void][RP]::WriteProcessMemory($hProc, $pRem, $lv, 60, [ref]$w)
                    [void][RP]::SendMessage($Hwnd, $script:LVM_GETITEMTEXTA, [IntPtr]$r, $pRem)
                    $cells += (Read-RemoteString -HProc $hProc -Addr $pText -Bytes 256)
                }
                $st = [int][RP]::SendMessage($Hwnd, $script:LVM_GETITEMSTATE, [IntPtr]$r, [IntPtr]$script:LVIS_SELECTED)
                # objeto con las celdas accesibles por nombre de columna
                $o = [ordered]@{}
                for ($c = 0; $c -lt $nCols; $c++) {
                    $name = if ($c -lt $cols.Count -and $cols[$c]) { $cols[$c] } else { "COL$c" }
                    if (-not $o.Contains($name)) { $o[$name] = $cells[$c] }
                }
                $rows += [pscustomobject]@{
                    Index    = $r
                    Cells    = $cells
                    Campos   = [pscustomobject]$o
                    Selected = (($st -band $script:LVIS_SELECTED) -ne 0)
                }
            }
            return [pscustomobject]@{ Columns = $cols; Rows = $rows }
        } finally { [void][RP]::VirtualFreeEx($hProc, $pRem, 0, $script:MEM_RELEASE) }
    } finally { [void][RP]::CloseHandle($hProc) }
}

function Get-TreeViewItems {
    param([Parameter(Mandatory=$true)][IntPtr]$Hwnd)
    $hProc = Open-TargetProcess -Hwnd $Hwnd
    try {
        $pRem  = New-RemoteBuffer -HProc $hProc -Size 4096
        $pText = [IntPtr]([int64]$pRem + 256)
        try {
            $readText = {
                param([IntPtr]$hItem)
                # TVITEM (32 bits): mask(0) hItem(4) state(8) stateMask(12) pszText(16) cchTextMax(20)
                $tv = New-Object byte[] 40
                [Array]::Copy([BitConverter]::GetBytes([uint32]0x0001), 0, $tv, 0, 4)   # TVIF_TEXT
                [Array]::Copy([BitConverter]::GetBytes([uint32]$hItem.ToInt32()), 0, $tv, 4, 4)
                [Array]::Copy([BitConverter]::GetBytes([uint32]$pText.ToInt32()), 0, $tv, 16, 4)
                [Array]::Copy([BitConverter]::GetBytes([int32]200), 0, $tv, 20, 4)
                $w = [UIntPtr]::Zero
                [void][RP]::WriteProcessMemory($hProc, $pRem, $tv, 40, [ref]$w)
                [void][RP]::SendMessage($Hwnd, $script:TVM_GETITEMA, [IntPtr]::Zero, $pRem)
                return (Read-RemoteString -HProc $hProc -Addr $pText -Bytes 200)
            }
            $items = @()
            $h = [RP]::SendMessage($Hwnd, $script:TVM_GETNEXTITEM, [IntPtr]$script:TVGN_ROOT, [IntPtr]::Zero)
            while ($h -ne [IntPtr]::Zero) {
                $items += [pscustomobject]@{ Handle = $h; Nivel = 0; Texto = (& $readText $h) }
                $c = [RP]::SendMessage($Hwnd, $script:TVM_GETNEXTITEM, [IntPtr]$script:TVGN_CHILD, $h)
                while ($c -ne [IntPtr]::Zero) {
                    $items += [pscustomobject]@{ Handle = $c; Nivel = 1; Texto = (& $readText $c) }
                    $c = [RP]::SendMessage($Hwnd, $script:TVM_GETNEXTITEM, [IntPtr]$script:TVGN_NEXT, $c)
                }
                $h = [RP]::SendMessage($Hwnd, $script:TVM_GETNEXTITEM, [IntPtr]$script:TVGN_NEXT, $h)
            }
            return $items
        } finally { [void][RP]::VirtualFreeEx($hProc, $pRem, 0, $script:MEM_RELEASE) }
    } finally { [void][RP]::CloseHandle($hProc) }
}
