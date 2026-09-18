# =====================================================================
#  Win32.ps1 - Interoperabilidad con la API de Windows
#  Usado por la automatizacion de AC (aplicacion VB6 / ThunderRT6)
# =====================================================================

if (-not ("W32" -as [type])) {
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class W32 {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet=CharSet.Auto, EntryPoint="SendMessage")]
    public static extern IntPtr SendMessageSb(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);
    [DllImport("user32.dll", CharSet=CharSet.Auto, EntryPoint="SendMessage")]
    public static extern IntPtr SendMessageStr(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindowEnabled(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc cb, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X, Y; }

    public const uint WM_SETTEXT        = 0x000C;
    public const uint WM_GETTEXT        = 0x000D;
    public const uint WM_GETTEXTLENGTH  = 0x000E;
    public const uint WM_CLOSE          = 0x0010;
    public const uint WM_COMMAND        = 0x0111;
    public const uint WM_LBUTTONDOWN    = 0x0201;
    public const uint WM_LBUTTONUP      = 0x0202;
    public const uint BM_CLICK          = 0x00F5;
    public const uint CB_GETCOUNT       = 0x0146;
    public const uint CB_GETLBTEXT      = 0x0148;
    public const uint CB_GETLBTEXTLEN   = 0x0149;
    public const uint CB_SETCURSEL      = 0x014E;
    public const uint CB_GETCURSEL      = 0x0147;
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP   = 0x0004;
    public const uint KEYEVENTF_KEYUP      = 0x0002;
}
"@
}

# ---------------------------------------------------------------------
#  Enumeracion de ventanas
# ---------------------------------------------------------------------

function Get-ChildHandles {
    <#  Devuelve TODOS los descendientes de una ventana (no solo hijos directos).  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Parent)
    $list = New-Object System.Collections.ArrayList
    $cb = [W32+EnumWindowsProc]{
        param($h, $l)
        $cls = New-Object System.Text.StringBuilder 256
        [void][W32]::GetClassName($h, $cls, 256)
        $txt = New-Object System.Text.StringBuilder 512
        [void][W32]::GetWindowText($h, $txt, 512)
        $r = New-Object W32+RECT
        [void][W32]::GetWindowRect($h, [ref]$r)
        [void]$list.Add([pscustomobject]@{
            Handle  = $h
            Class   = $cls.ToString()
            Text    = $txt.ToString()
            X       = $r.Left
            Y       = $r.Top
            W       = ($r.Right  - $r.Left)
            H       = ($r.Bottom - $r.Top)
            Visible = [W32]::IsWindowVisible($h)
            Enabled = [W32]::IsWindowEnabled($h)
        })
        return $true
    }
    [void][W32]::EnumChildWindows($Parent, $cb, [IntPtr]::Zero)
    return $list
}

function Get-TopWindows {
    <#  Ventanas de nivel superior pertenecientes a un proceso.  #>
    param([Parameter(Mandatory=$true)][int]$ProcessId)
    $list = New-Object System.Collections.ArrayList
    $script:__pidFilter = $ProcessId
    $cb = [W32+EnumWindowsProc]{
        param($h, $l)
        $procId = 0
        [void][W32]::GetWindowThreadProcessId($h, [ref]$procId)
        if ($procId -eq $script:__pidFilter) {
            $cls = New-Object System.Text.StringBuilder 256
            [void][W32]::GetClassName($h, $cls, 256)
            $txt = New-Object System.Text.StringBuilder 512
            [void][W32]::GetWindowText($h, $txt, 512)
            $r = New-Object W32+RECT
            [void][W32]::GetWindowRect($h, [ref]$r)
            [void]$list.Add([pscustomobject]@{
                Handle  = $h
                Class   = $cls.ToString()
                Text    = $txt.ToString()
                X = $r.Left; Y = $r.Top
                W = ($r.Right - $r.Left); H = ($r.Bottom - $r.Top)
                Visible = [W32]::IsWindowVisible($h)
                Enabled = [W32]::IsWindowEnabled($h)
            })
        }
        return $true
    }
    [void][W32]::EnumWindows($cb, [IntPtr]::Zero)
    return $list
}

# ---------------------------------------------------------------------
#  Lectura / escritura de texto en controles
# ---------------------------------------------------------------------

function Get-CtrlText {
    param([Parameter(Mandatory=$true)][IntPtr]$Handle)
    $len = [int][W32]::SendMessage($Handle, [W32]::WM_GETTEXTLENGTH, [IntPtr]::Zero, [IntPtr]::Zero)
    if ($len -le 0) { return "" }
    $sb = New-Object System.Text.StringBuilder ($len + 2)
    [void][W32]::SendMessageSb($Handle, [W32]::WM_GETTEXT, [IntPtr]($len + 1), $sb)
    return $sb.ToString()
}

function Set-CtrlText {
    <#  Escribe texto sin depender del foco del teclado. Imprescindible en AC,
        donde el foco se va a un PictureBox al hacer clic en el panel.  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Handle, [string]$Text)
    [void][W32]::SendMessageStr($Handle, [W32]::WM_SETTEXT, [IntPtr]::Zero, $Text)
    Start-Sleep -Milliseconds 120
    return (Get-CtrlText -Handle $Handle)
}

function Get-UiaName {
    <#  Los TextBox de VB6 no responden a WM_GETTEXT, pero si exponen su
        contenido por UI Automation.  #>
    param([Parameter(Mandatory=$true)][int]$Handle)
    try {
        $cond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty, $Handle)
        $el = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants, $cond)
        if ($el) { return $el.Current.Name }
    } catch { }
    return $null
}

function Get-ComboItems {
    param([Parameter(Mandatory=$true)][IntPtr]$Handle)
    $n = [int][W32]::SendMessage($Handle, [W32]::CB_GETCOUNT, [IntPtr]::Zero, [IntPtr]::Zero)
    $items = @()
    for ($i = 0; $i -lt $n; $i++) {
        $len = [int][W32]::SendMessage($Handle, [W32]::CB_GETLBTEXTLEN, [IntPtr]$i, [IntPtr]::Zero)
        $sb = New-Object System.Text.StringBuilder ($len + 2)
        [void][W32]::SendMessageSb($Handle, [W32]::CB_GETLBTEXT, [IntPtr]$i, $sb)
        $items += $sb.ToString()
    }
    return $items
}

# ---------------------------------------------------------------------
#  Clics y teclado
# ---------------------------------------------------------------------

function Invoke-ClickAt {
    param([int]$X, [int]$Y, [int]$SettleMs = 250)
    [void][W32]::SetCursorPos($X, $Y)
    Start-Sleep -Milliseconds 120
    [W32]::mouse_event([W32]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [W32]::mouse_event([W32]::MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds $SettleMs
}

function Invoke-ClickControl {
    <#  Clic sobre un control. Los UserControl de VB6 no responden a BM_CLICK,
        por eso se usa el raton fisico sobre el centro del control.  #>
    param([Parameter(Mandatory=$true)][IntPtr]$Handle, [int]$SettleMs = 250)
    if (-not [W32]::IsWindow($Handle)) { throw "El control $Handle ya no existe" }
    $r = New-Object W32+RECT
    [void][W32]::GetWindowRect($Handle, [ref]$r)
    Invoke-ClickAt -X ([int](($r.Left + $r.Right) / 2)) -Y ([int](($r.Top + $r.Bottom) / 2)) -SettleMs $SettleMs
}

function Send-Hotkey {
    <#  Combinacion con modificadores mediante keybd_event.
        SendKeys no funciona con los formularios VB6 de AC.  #>
    param([byte[]]$Modifiers, [byte]$Key)
    foreach ($m in $Modifiers) { [W32]::keybd_event($m, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80 }
    [W32]::keybd_event($Key, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80
    [W32]::keybd_event($Key, 0, [W32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero); Start-Sleep -Milliseconds 80
    for ($i = $Modifiers.Count - 1; $i -ge 0; $i--) {
        [W32]::keybd_event($Modifiers[$i], 0, [W32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 60
    }
}

function Set-WindowFocus {
    param([Parameter(Mandatory=$true)][IntPtr]$Handle)
    [void][W32]::ShowWindow($Handle, 9)      # SW_RESTORE
    [void][W32]::SetForegroundWindow($Handle)
    Start-Sleep -Milliseconds 400
}

# ---------------------------------------------------------------------
#  Capturas de pantalla (evidencia y diagnostico)
# ---------------------------------------------------------------------

function Save-WindowShot {
    param([Parameter(Mandatory=$true)][IntPtr]$Handle, [Parameter(Mandatory=$true)][string]$Path)
    try {
        Add-Type -AssemblyName System.Drawing
        $r = New-Object W32+RECT
        [void][W32]::GetWindowRect($Handle, [ref]$r)
        $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
        if ($w -le 0 -or $h -le 0) { return $null }
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $bmp = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.CopyFromScreen($r.Left, $r.Top, 0, 0, (New-Object System.Drawing.Size $w, $h))
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose()
        return $Path
    } catch { return $null }
}
