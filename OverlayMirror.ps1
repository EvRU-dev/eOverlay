param(
    [switch]$StartClickThrough,
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$source = @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

public static class OverlayNative
{
    public const uint WDA_NONE = 0x00000000;
    public const uint WDA_MONITOR = 0x00000001;
    public const uint WDA_EXCLUDEFROMCAPTURE = 0x00000011;
    public const uint WM_MOUSEMOVE = 0x0200;
    public const uint WM_LBUTTONDOWN = 0x0201;
    public const uint WM_LBUTTONUP = 0x0202;
    public const uint WM_LBUTTONDBLCLK = 0x0203;
    public const uint WM_RBUTTONDOWN = 0x0204;
    public const uint WM_RBUTTONUP = 0x0205;
    public const uint WM_RBUTTONDBLCLK = 0x0206;
    public const uint WM_MBUTTONDOWN = 0x0207;
    public const uint WM_MBUTTONUP = 0x0208;
    public const uint WM_MBUTTONDBLCLK = 0x0209;
    public const uint WM_MOUSEWHEEL = 0x020A;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const uint WM_CHAR = 0x0102;
    public const uint WM_SYSKEYDOWN = 0x0104;
    public const uint WM_SYSKEYUP = 0x0105;
    public const uint MK_LBUTTON = 0x0001;
    public const uint MK_RBUTTON = 0x0002;
    public const uint MK_SHIFT = 0x0004;
    public const uint MK_CONTROL = 0x0008;
    public const uint MK_MBUTTON = 0x0010;

    private const int DWMWA_CLOAKED = 14;
    private const int GWL_EXSTYLE = -20;
    private const long WS_EX_TRANSPARENT = 0x20L;
    private const long WS_EX_LAYERED = 0x80000L;
    private const long WS_EX_TOOLWINDOW = 0x80L;
    private const uint CWP_SKIPINVISIBLE = 0x0001;
    private const uint CWP_SKIPDISABLED = 0x0002;
    private const uint CWP_SKIPTRANSPARENT = 0x0004;

    private const int DWM_TNP_RECTDESTINATION = 0x00000001;
    private const int DWM_TNP_OPACITY = 0x00000004;
    private const int DWM_TNP_VISIBLE = 0x00000008;
    private const int DWM_TNP_SOURCECLIENTAREAONLY = 0x00000010;

    public class WindowInfo
    {
        public IntPtr Handle { get; set; }
        public string Title { get; set; }
        public string ProcessName { get; set; }

        public override string ToString()
        {
            if (string.IsNullOrWhiteSpace(ProcessName))
            {
                return Title ?? "<Untitled>";
            }

            return string.Format("{0} [{1}]", Title ?? "<Untitled>", ProcessName);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SIZE
    {
        public int cx;
        public int cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DWM_THUMBNAIL_PROPERTIES
    {
        public int dwFlags;
        public RECT rcDestination;
        public RECT rcSource;
        public byte opacity;
        [MarshalAs(UnmanagedType.Bool)]
        public bool fVisible;
        [MarshalAs(UnmanagedType.Bool)]
        public bool fSourceClientAreaOnly;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetShellWindow();

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

    [DllImport("user32.dll")]
    private static extern IntPtr ChildWindowFromPointEx(IntPtr hwndParent, POINT pt, uint flags);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetWindowDisplayAffinity(IntPtr hWnd, uint dwAffinity);

    [DllImport("dwmapi.dll")]
    private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

    [DllImport("dwmapi.dll")]
    private static extern int DwmRegisterThumbnail(IntPtr dest, IntPtr src, out IntPtr thumb);

    [DllImport("dwmapi.dll")]
    private static extern int DwmUnregisterThumbnail(IntPtr thumb);

    [DllImport("dwmapi.dll")]
    private static extern int DwmQueryThumbnailSourceSize(IntPtr thumb, out SIZE size);

    [DllImport("dwmapi.dll")]
    private static extern int DwmUpdateThumbnailProperties(IntPtr hThumbnail, ref DWM_THUMBNAIL_PROPERTIES props);

    [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
    private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
    private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
    private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex)
    {
        return IntPtr.Size == 8
            ? GetWindowLongPtr64(hWnd, nIndex)
            : new IntPtr(GetWindowLong32(hWnd, nIndex));
    }

    private static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr newValue)
    {
        return IntPtr.Size == 8
            ? SetWindowLongPtr64(hWnd, nIndex, newValue)
            : new IntPtr(SetWindowLong32(hWnd, nIndex, newValue.ToInt32()));
    }

    public static WindowInfo[] EnumerateWindows(IntPtr excludeA, IntPtr excludeB)
    {
        IntPtr shell = GetShellWindow();
        var windows = new List<WindowInfo>();

        EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
        {
            if (hWnd == IntPtr.Zero || hWnd == shell || hWnd == excludeA || hWnd == excludeB)
            {
                return true;
            }

            if (!IsWindowVisible(hWnd) || IsCloaked(hWnd))
            {
                return true;
            }

            string title = GetWindowTitle(hWnd);
            if (string.IsNullOrWhiteSpace(title))
            {
                return true;
            }

            windows.Add(new WindowInfo
            {
                Handle = hWnd,
                Title = title,
                ProcessName = GetProcessName(hWnd)
            });

            return true;
        }, IntPtr.Zero);

        return windows
            .OrderBy(w => w.ProcessName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ThenBy(w => w.Title ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public static WindowInfo GetForegroundWindowInfo(IntPtr excludeA, IntPtr excludeB)
    {
        IntPtr hWnd = GetForegroundWindow();
        if (hWnd == IntPtr.Zero || hWnd == excludeA || hWnd == excludeB)
        {
            return null;
        }

        if (!IsWindowVisible(hWnd) || IsCloaked(hWnd))
        {
            return null;
        }

        string title = GetWindowTitle(hWnd);
        if (string.IsNullOrWhiteSpace(title))
        {
            return null;
        }

        return new WindowInfo
        {
            Handle = hWnd,
            Title = title,
            ProcessName = GetProcessName(hWnd)
        };
    }

    public static bool IsValidWindow(IntPtr hWnd)
    {
        return hWnd != IntPtr.Zero && IsWindow(hWnd);
    }

    public static bool TryGetWindowBounds(IntPtr hWnd, out int left, out int top, out int width, out int height)
    {
        left = 0;
        top = 0;
        width = 0;
        height = 0;

        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        RECT rect;
        if (!GetWindowRect(hWnd, out rect))
        {
            return false;
        }

        left = rect.left;
        top = rect.top;
        width = Math.Max(1, rect.right - rect.left);
        height = Math.Max(1, rect.bottom - rect.top);
        return true;
    }

    public static int RegisterThumbnail(IntPtr destinationWindow, IntPtr sourceWindow, out IntPtr thumbnail)
    {
        return DwmRegisterThumbnail(destinationWindow, sourceWindow, out thumbnail);
    }

    public static void UnregisterThumbnail(IntPtr thumbnail)
    {
        if (thumbnail != IntPtr.Zero)
        {
            DwmUnregisterThumbnail(thumbnail);
        }
    }

    public static int FitThumbnailToClient(IntPtr thumbnail, int clientWidth, int clientHeight)
    {
        if (thumbnail == IntPtr.Zero || clientWidth <= 0 || clientHeight <= 0)
        {
            return 0;
        }

        SIZE sourceSize;
        int hr = DwmQueryThumbnailSourceSize(thumbnail, out sourceSize);
        if (hr != 0)
        {
            return hr;
        }

        RECT destination = BuildLetterboxRect(sourceSize.cx, sourceSize.cy, clientWidth, clientHeight);

        var props = new DWM_THUMBNAIL_PROPERTIES
        {
            dwFlags = DWM_TNP_RECTDESTINATION | DWM_TNP_OPACITY | DWM_TNP_VISIBLE | DWM_TNP_SOURCECLIENTAREAONLY,
            rcDestination = destination,
            opacity = 255,
            fVisible = true,
            fSourceClientAreaOnly = false
        };

        return DwmUpdateThumbnailProperties(thumbnail, ref props);
    }

    public static bool ApplyCaptureExclusion(IntPtr hWnd, bool exclude)
    {
        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        if (!exclude)
        {
            return SetWindowDisplayAffinity(hWnd, WDA_NONE);
        }

        if (SetWindowDisplayAffinity(hWnd, WDA_EXCLUDEFROMCAPTURE))
        {
            return true;
        }

        return SetWindowDisplayAffinity(hWnd, WDA_MONITOR);
    }

    public static void SetClickThrough(IntPtr hWnd, bool enabled)
    {
        if (hWnd == IntPtr.Zero)
        {
            return;
        }

        long styles = GetWindowLongPtr(hWnd, GWL_EXSTYLE).ToInt64();

        if (enabled)
        {
            styles |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW;
        }
        else
        {
            styles &= ~WS_EX_TRANSPARENT;
        }

        SetWindowLongPtr(hWnd, GWL_EXSTYLE, new IntPtr(styles));
    }

    public static IntPtr ResolveMessageTarget(IntPtr rootWindow, int screenX, int screenY)
    {
        if (!IsValidWindow(rootWindow))
        {
            return IntPtr.Zero;
        }

        IntPtr current = rootWindow;
        POINT screenPoint = new POINT { X = screenX, Y = screenY };

        for (int depth = 0; depth < 12; depth++)
        {
            POINT localPoint = screenPoint;
            if (!ScreenToClient(current, ref localPoint))
            {
                break;
            }

            IntPtr child = ChildWindowFromPointEx(
                current,
                localPoint,
                CWP_SKIPINVISIBLE | CWP_SKIPDISABLED | CWP_SKIPTRANSPARENT
            );

            if (child == IntPtr.Zero || child == current)
            {
                break;
            }

            current = child;
        }

        return current;
    }

    public static bool TryScreenToClient(IntPtr hWnd, int screenX, int screenY, out int clientX, out int clientY)
    {
        clientX = 0;
        clientY = 0;

        if (hWnd == IntPtr.Zero)
        {
            return false;
        }

        POINT point = new POINT { X = screenX, Y = screenY };
        if (!ScreenToClient(hWnd, ref point))
        {
            return false;
        }

        clientX = point.X;
        clientY = point.Y;
        return true;
    }

    public static void PostMouseMessage(IntPtr hWnd, uint message, int clientX, int clientY, uint keyState)
    {
        if (hWnd == IntPtr.Zero)
        {
            return;
        }

        PostMessage(hWnd, message, new IntPtr(unchecked((int)keyState)), MakeLParam(clientX, clientY));
    }

    public static void PostMouseWheelMessage(IntPtr hWnd, int screenX, int screenY, int wheelDelta, uint keyState)
    {
        if (hWnd == IntPtr.Zero)
        {
            return;
        }

        uint packed = ((uint)(ushort)wheelDelta << 16) | (keyState & 0xFFFFu);
        PostMessage(hWnd, WM_MOUSEWHEEL, new IntPtr(unchecked((int)packed)), MakeLParam(screenX, screenY));
    }

    public static void PostKeyMessage(IntPtr hWnd, uint message, int virtualKey, int lParam)
    {
        if (hWnd == IntPtr.Zero)
        {
            return;
        }

        PostMessage(hWnd, message, new IntPtr(virtualKey), new IntPtr(lParam));
    }

    public static void PostCharMessage(IntPtr hWnd, int characterCode)
    {
        if (hWnd == IntPtr.Zero)
        {
            return;
        }

        PostMessage(hWnd, WM_CHAR, new IntPtr(characterCode), new IntPtr(1));
    }

    public static bool ActivateWindow(IntPtr hWnd)
    {
        return hWnd != IntPtr.Zero && SetForegroundWindow(hWnd);
    }

    public static int BuildKeyLParam(int virtualKey, bool isKeyUp, bool altDown)
    {
        uint scanCode = MapVirtualKey((uint)virtualKey, 0);
        int lParam = 1 | ((int)scanCode << 16);

        if (IsExtendedKey(virtualKey))
        {
            lParam |= 1 << 24;
        }

        if (altDown)
        {
            lParam |= 1 << 29;
        }

        if (isKeyUp)
        {
            lParam |= unchecked((int)0xC0000000);
        }

        return lParam;
    }

    private static RECT BuildLetterboxRect(int sourceWidth, int sourceHeight, int targetWidth, int targetHeight)
    {
        if (sourceWidth <= 0 || sourceHeight <= 0 || targetWidth <= 0 || targetHeight <= 0)
        {
            return new RECT
            {
                left = 0,
                top = 0,
                right = targetWidth,
                bottom = targetHeight
            };
        }

        double scale = Math.Min((double)targetWidth / sourceWidth, (double)targetHeight / sourceHeight);
        int width = Math.Max(1, (int)Math.Round(sourceWidth * scale));
        int height = Math.Max(1, (int)Math.Round(sourceHeight * scale));
        int left = (targetWidth - width) / 2;
        int top = (targetHeight - height) / 2;

        return new RECT
        {
            left = left,
            top = top,
            right = left + width,
            bottom = top + height
        };
    }

    private static bool IsCloaked(IntPtr hWnd)
    {
        int cloaked;
        return DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out cloaked, Marshal.SizeOf(typeof(int))) == 0 && cloaked != 0;
    }

    private static string GetWindowTitle(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        if (length <= 0)
        {
            return null;
        }

        StringBuilder builder = new StringBuilder(length + 1);
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString().Trim();
    }

    private static string GetProcessName(IntPtr hWnd)
    {
        try
        {
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);
            if (processId == 0)
            {
                return null;
            }

            using (Process process = Process.GetProcessById((int)processId))
            {
                return process.ProcessName;
            }
        }
        catch
        {
            return null;
        }
    }

    private static IntPtr MakeLParam(int lowWord, int highWord)
    {
        return new IntPtr((highWord << 16) | (lowWord & 0xFFFF));
    }

    private static bool IsExtendedKey(int virtualKey)
    {
        switch (virtualKey)
        {
            case 0x21: // PAGE UP
            case 0x22: // PAGE DOWN
            case 0x23: // END
            case 0x24: // HOME
            case 0x25: // LEFT
            case 0x26: // UP
            case 0x27: // RIGHT
            case 0x28: // DOWN
            case 0x2D: // INSERT
            case 0x2E: // DELETE
            case 0x5B: // LWIN
            case 0x5C: // RWIN
            case 0x5D: // APPS
            case 0x6F: // DIVIDE
            case 0x90: // NUM LOCK
            case 0xA3: // RCONTROL
            case 0xA5: // RMENU
                return true;
            default:
                return false;
        }
    }
}
"@

Add-Type -TypeDefinition $source -ReferencedAssemblies @(
    "System.dll",
    "System.Core.dll",
    "System.Windows.Forms.dll",
    "System.Drawing.dll"
)

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

$script:overlayForm = $null
$script:controlForm = $null
$script:windowPicker = $null
$script:statusLabel = $null
$script:overlayHint = $null
$script:overlayTextLabel = $null
$script:overlayModePicker = $null
$script:windowModePanel = $null
$script:textModePanel = $null
$script:webModePanel = $null
$script:textInputBox = $null
$script:textAlignmentPicker = $null
$script:textFontSizePicker = $null
$script:webUrlBox = $null
$script:webZoomPicker = $null
$script:webGoButton = $null
$script:webBackButton = $null
$script:webForwardButton = $null
$script:webReloadButton = $null
$script:webOpenExternalButton = $null
$script:webViewControl = $null
$script:webViewReady = $false
$script:webViewAvailable = $false
$script:webViewAssemblyLoaded = $false
$script:webViewAssemblyError = $null
$script:webPendingUrl = "https://www.google.com"
$script:webCurrentUrl = "https://www.google.com"
$script:webZoomPercent = 100
$script:opacitySlider = $null
$script:opacityValueLabel = $null
$script:clickThroughCheck = $null
$script:interactiveInputCheck = $null
$script:controlTopMostCheck = $null
$script:controlCaptureCheck = $null
$script:trayShowControlItem = $null
$script:trayToggleOverlayItem = $null
$script:notifyIcon = $null
$script:appContext = $null
$script:thumbnailHandle = [IntPtr]::Zero
$script:currentSourceHandle = [IntPtr]::Zero
$script:captureExcluded = $true
$script:clickThroughEnabled = [bool]$StartClickThrough
$script:interactiveInputEnabled = $false
$script:overlayMode = "Window"
$script:overlayOpacityPercent = 100
$script:controlTopMostEnabled = $false
$script:controlCaptureExcluded = $true
$script:textAlignment = "Center"
$script:textFontSize = 28
$script:isExiting = $false
$script:selectionTimer = $null
$script:watchdogTimer = $null
$script:shuttingDown = $false
$script:capturedMouseTarget = [IntPtr]::Zero
$script:capturedMouseButton = [System.Windows.Forms.MouseButtons]::None

function Set-Status {
    param([string]$Message)

    $script:statusLabel.Text = $Message
}

function Clear-Thumbnail {
    [OverlayNative]::UnregisterThumbnail($script:thumbnailHandle)
    $script:thumbnailHandle = [IntPtr]::Zero
    $script:currentSourceHandle = [IntPtr]::Zero
    $script:capturedMouseTarget = [IntPtr]::Zero
    $script:capturedMouseButton = [System.Windows.Forms.MouseButtons]::None
    if ($null -ne $script:overlayTextLabel) {
        $script:overlayTextLabel.Visible = $false
    }
    if ($null -ne $script:webViewControl) {
        $script:webViewControl.Visible = $false
    }
    if ($null -ne $script:overlayHint) {
        $script:overlayHint.Visible = $true
    }
}

function Update-ThumbnailLayout {
    if ($script:thumbnailHandle -eq [IntPtr]::Zero) {
        return
    }

    $clientSize = $script:overlayForm.ClientSize
    [void][OverlayNative]::FitThumbnailToClient(
        $script:thumbnailHandle,
        $clientSize.Width,
        $clientSize.Height
    )
}

function Apply-OverlayFlags {
    if (-not $script:overlayForm.IsHandleCreated) {
        return
    }

    [void][OverlayNative]::ApplyCaptureExclusion($script:overlayForm.Handle, $script:captureExcluded)
    [OverlayNative]::SetClickThrough($script:overlayForm.Handle, $script:clickThroughEnabled)
    Update-OverlayInteractionUi
}

function Apply-ControlFlags {
    if ($null -eq $script:controlForm -or -not $script:controlForm.IsHandleCreated) {
        return
    }

    $script:controlForm.TopMost = $script:controlTopMostEnabled
    [void][OverlayNative]::ApplyCaptureExclusion($script:controlForm.Handle, $script:controlCaptureExcluded)
}

function Initialize-WebView2Support {
    if ($script:webViewAssemblyLoaded) {
        return $script:webViewAvailable
    }

    $script:webViewAssemblyLoaded = $true

    $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {
        $_.GetName().Name -eq "Microsoft.Web.WebView2.WinForms"
    } | Select-Object -First 1

    if ($null -ne $alreadyLoaded) {
        $script:webViewAvailable = $true
        return $true
    }

    $baseDir = [System.AppDomain]::CurrentDomain.BaseDirectory
    $scriptDir = $PSScriptRoot
    $candidateDirs = @(
        $baseDir,
        (Join-Path $baseDir "vendor\webview2"),
        $scriptDir,
        (Join-Path $scriptDir "vendor\webview2"),
        "C:\Program Files\Logi\LogiPluginService",
        "C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin",
        "C:\Program Files\sinsam\app",
        "C:\Program Files\sinsam\app\browser"
    )

    foreach ($dir in $candidateDirs) {
        if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path $dir)) {
            continue
        }

        $coreDll = Join-Path $dir "Microsoft.Web.WebView2.Core.dll"
        $formsDll = Join-Path $dir "Microsoft.Web.WebView2.WinForms.dll"

        if ((Test-Path $coreDll) -and (Test-Path $formsDll)) {
            try {
                Add-Type -Path $coreDll
                Add-Type -Path $formsDll
                $script:webViewAvailable = $true
                return $true
            }
            catch {
                $script:webViewAssemblyError = $_.Exception.Message
            }
        }
    }

    $script:webViewAvailable = $false
    return $false
}

function Normalize-WebUrl {
    param([string]$Url)

    $value = if ([string]::IsNullOrWhiteSpace($Url)) { $script:webCurrentUrl } else { $Url.Trim() }
    if ([string]::IsNullOrWhiteSpace($value)) {
        return "https://www.google.com"
    }

    if ($value -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*://') {
        return "https://$value"
    }

    return $value
}

function Update-WebNavigationButtons {
    if ($null -eq $script:webBackButton -or $null -eq $script:webForwardButton) {
        return
    }

    $canUseHistory = ($script:webViewReady -and $null -ne $script:webViewControl -and $null -ne $script:webViewControl.CoreWebView2)
    $script:webBackButton.Enabled = ($canUseHistory -and $script:webViewControl.CoreWebView2.CanGoBack)
    $script:webForwardButton.Enabled = ($canUseHistory -and $script:webViewControl.CoreWebView2.CanGoForward)
}

function Set-WebZoom {
    param([int]$Percent)

    $Percent = [Math]::Max(25, [Math]::Min(300, $Percent))
    $script:webZoomPercent = $Percent

    if ($null -ne $script:webZoomPicker -and [int]$script:webZoomPicker.Value -ne $Percent) {
        $script:webZoomPicker.Value = $Percent
    }

    if ($script:webViewReady -and $null -ne $script:webViewControl -and $null -ne $script:webViewControl.CoreWebView2) {
        $script:webViewControl.ZoomFactor = $Percent / 100.0
    }
}

function Ensure-WebViewControl {
    if ($null -ne $script:webViewControl) {
        return $true
    }

    if (-not (Initialize-WebView2Support)) {
        return $false
    }

    $script:webViewControl = New-Object Microsoft.Web.WebView2.WinForms.WebView2
    $script:webViewControl.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:webViewControl.Visible = $false
    $script:webViewControl.DefaultBackgroundColor = [System.Drawing.Color]::Black

    $script:webViewControl.Add_CoreWebView2InitializationCompleted({
        param($sender, $eventArgs)

        if ($eventArgs.IsSuccess) {
            $script:webViewReady = $true
            $sender.CoreWebView2.Settings.IsStatusBarEnabled = $true
            $sender.CoreWebView2.Settings.AreDefaultContextMenusEnabled = $true
            $sender.CoreWebView2.Settings.AreBrowserAcceleratorKeysEnabled = $true
            Set-WebZoom -Percent $script:webZoomPercent
            Update-WebNavigationButtons

            if (-not [string]::IsNullOrWhiteSpace($script:webPendingUrl)) {
                $sender.CoreWebView2.Navigate($script:webPendingUrl)
            }
        }
        else {
            $script:webViewAssemblyError = $eventArgs.InitializationException.Message
            Set-Status ("WebView2 init failed: {0}" -f $script:webViewAssemblyError)
        }
    })

    $script:webViewControl.Add_NavigationCompleted({
        param($sender, $eventArgs)
        try {
            if ($null -ne $sender.Source) {
                $script:webCurrentUrl = $sender.Source.AbsoluteUri
                if ($null -ne $script:webUrlBox) {
                    $script:webUrlBox.Text = $script:webCurrentUrl
                }
            }
        }
        catch {
        }

        Update-WebNavigationButtons
        if ($eventArgs.IsSuccess) {
            Set-Status ("Web page loaded: {0}" -f $script:webCurrentUrl)
        }
        else {
            Set-Status "Web page failed to load."
        }
    })

    $script:webViewControl.Add_SourceChanged({
        param($sender, $eventArgs)
        try {
            if ($null -ne $sender.Source) {
                $script:webCurrentUrl = $sender.Source.AbsoluteUri
                if ($null -ne $script:webUrlBox) {
                    $script:webUrlBox.Text = $script:webCurrentUrl
                }
            }
        }
        catch {
        }

        Update-WebNavigationButtons
    })

    $script:overlayForm.Controls.Add($script:webViewControl)
    $script:webViewControl.BringToFront()

    if ($script:overlayForm.IsHandleCreated) {
        [void]$script:webViewControl.EnsureCoreWebView2Async($null)
    }

    return $true
}

function Show-WebPage {
    param([string]$Url)

    if (-not (Ensure-WebViewControl)) {
        $message = if ([string]::IsNullOrWhiteSpace($script:webViewAssemblyError)) {
            "WebView2 is not available on this PC."
        }
        else {
            "WebView2 is unavailable: $($script:webViewAssemblyError)"
        }

        Set-Status $message
        return
    }

    $normalizedUrl = Normalize-WebUrl -Url $Url
    $script:webPendingUrl = $normalizedUrl
    $script:webCurrentUrl = $normalizedUrl

    if ($script:clickThroughEnabled -and $null -ne $script:clickThroughCheck) {
        $script:clickThroughCheck.Checked = $false
    }

    Clear-Thumbnail
    $script:overlayHint.Visible = $false
    $script:overlayTextLabel.Visible = $false
    $script:webViewControl.Visible = $true
    Show-OverlayWindow

    if ($script:webViewReady -and $null -ne $script:webViewControl.CoreWebView2) {
        $script:webViewControl.CoreWebView2.Navigate($normalizedUrl)
        Set-WebZoom -Percent $script:webZoomPercent
    }
    else {
        [void]$script:webViewControl.EnsureCoreWebView2Async($null)
        Set-Status "Initializing WebView2..."
    }
}

function Get-LetterboxRect {
    param(
        [int]$SourceWidth,
        [int]$SourceHeight,
        [int]$TargetWidth,
        [int]$TargetHeight
    )

    if ($SourceWidth -le 0 -or $SourceHeight -le 0 -or $TargetWidth -le 0 -or $TargetHeight -le 0) {
        return [pscustomobject]@{
            X = 0
            Y = 0
            Width = [Math]::Max(1, $TargetWidth)
            Height = [Math]::Max(1, $TargetHeight)
        }
    }

    $scale = [Math]::Min($TargetWidth / [double]$SourceWidth, $TargetHeight / [double]$SourceHeight)
    $width = [Math]::Max(1, [int][Math]::Round($SourceWidth * $scale))
    $height = [Math]::Max(1, [int][Math]::Round($SourceHeight * $scale))
    $left = [int](($TargetWidth - $width) / 2)
    $top = [int](($TargetHeight - $height) / 2)

    return [pscustomobject]@{
        X = $left
        Y = $top
        Width = $width
        Height = $height
    }
}

function Get-CurrentSourceMap {
    if ($script:currentSourceHandle -eq [IntPtr]::Zero) {
        return $null
    }

    if (-not $script:overlayForm.IsHandleCreated) {
        return $null
    }

    $left = 0
    $top = 0
    $width = 0
    $height = 0
    if (-not [OverlayNative]::TryGetWindowBounds($script:currentSourceHandle, [ref]$left, [ref]$top, [ref]$width, [ref]$height)) {
        return $null
    }

    $client = $script:overlayForm.ClientSize
    if ($client.Width -le 0 -or $client.Height -le 0) {
        return $null
    }

    $thumbRect = Get-LetterboxRect -SourceWidth $width -SourceHeight $height -TargetWidth $client.Width -TargetHeight $client.Height
    return [pscustomobject]@{
        ScreenLeft = $left
        ScreenTop = $top
        SourceWidth = $width
        SourceHeight = $height
        ThumbX = $thumbRect.X
        ThumbY = $thumbRect.Y
        ThumbWidth = $thumbRect.Width
        ThumbHeight = $thumbRect.Height
    }
}

function Convert-OverlayPointToSource {
    param([System.Drawing.Point]$OverlayPoint)

    $map = Get-CurrentSourceMap
    if ($null -eq $map) {
        return $null
    }

    if (
        $OverlayPoint.X -lt $map.ThumbX -or
        $OverlayPoint.Y -lt $map.ThumbY -or
        $OverlayPoint.X -ge ($map.ThumbX + $map.ThumbWidth) -or
        $OverlayPoint.Y -ge ($map.ThumbY + $map.ThumbHeight)
    ) {
        return $null
    }

    $relativeX = ($OverlayPoint.X - $map.ThumbX) / [double]$map.ThumbWidth
    $relativeY = ($OverlayPoint.Y - $map.ThumbY) / [double]$map.ThumbHeight
    $sourceX = [Math]::Min($map.SourceWidth - 1, [Math]::Max(0, [int][Math]::Floor($relativeX * $map.SourceWidth)))
    $sourceY = [Math]::Min($map.SourceHeight - 1, [Math]::Max(0, [int][Math]::Floor($relativeY * $map.SourceHeight)))

    return [pscustomobject]@{
        ScreenX = $map.ScreenLeft + $sourceX
        ScreenY = $map.ScreenTop + $sourceY
    }
}

function Resolve-InputTarget {
    param(
        [int]$ScreenX,
        [int]$ScreenY,
        [IntPtr]$PreferredHandle = [IntPtr]::Zero
    )

    $targetHandle = $PreferredHandle
    if ($targetHandle -eq [IntPtr]::Zero) {
        $targetHandle = [OverlayNative]::ResolveMessageTarget($script:currentSourceHandle, $ScreenX, $ScreenY)
    }

    if ($targetHandle -eq [IntPtr]::Zero) {
        $targetHandle = $script:currentSourceHandle
    }

    $clientX = 0
    $clientY = 0
    if (-not [OverlayNative]::TryScreenToClient($targetHandle, $ScreenX, $ScreenY, [ref]$clientX, [ref]$clientY)) {
        return $null
    }

    return [pscustomobject]@{
        Handle = $targetHandle
        ClientX = $clientX
        ClientY = $clientY
        ScreenX = $ScreenX
        ScreenY = $ScreenY
    }
}

function Get-ModifierMouseMask {
    $mask = [uint32]0
    $modifiers = [System.Windows.Forms.Control]::ModifierKeys

    if (($modifiers -band [System.Windows.Forms.Keys]::Shift) -ne 0) {
        $mask = $mask -bor [OverlayNative]::MK_SHIFT
    }

    if (($modifiers -band [System.Windows.Forms.Keys]::Control) -ne 0) {
        $mask = $mask -bor [OverlayNative]::MK_CONTROL
    }

    return $mask
}

function Convert-MouseButtonsToMask {
    param([int]$ButtonsInt)

    $mask = [uint32]0

    if (($ButtonsInt -band [int][System.Windows.Forms.MouseButtons]::Left) -ne 0) {
        $mask = $mask -bor [OverlayNative]::MK_LBUTTON
    }

    if (($ButtonsInt -band [int][System.Windows.Forms.MouseButtons]::Right) -ne 0) {
        $mask = $mask -bor [OverlayNative]::MK_RBUTTON
    }

    if (($ButtonsInt -band [int][System.Windows.Forms.MouseButtons]::Middle) -ne 0) {
        $mask = $mask -bor [OverlayNative]::MK_MBUTTON
    }

    return $mask
}

function Get-MouseKeyState {
    param(
        [System.Windows.Forms.MouseButtons]$AddButton = [System.Windows.Forms.MouseButtons]::None,
        [System.Windows.Forms.MouseButtons]$RemoveButton = [System.Windows.Forms.MouseButtons]::None
    )

    $buttonsInt = [int][System.Windows.Forms.Control]::MouseButtons
    if ($AddButton -ne [System.Windows.Forms.MouseButtons]::None) {
        $buttonsInt = $buttonsInt -bor [int]$AddButton
    }

    if ($RemoveButton -ne [System.Windows.Forms.MouseButtons]::None) {
        $buttonsInt = $buttonsInt -band (-bnot [int]$RemoveButton)
    }

    return [uint32]((Convert-MouseButtonsToMask -ButtonsInt $buttonsInt) -bor (Get-ModifierMouseMask))
}

function Get-MouseMessage {
    param(
        [System.Windows.Forms.MouseButtons]$Button,
        [ValidateSet("Down", "Up", "DoubleClick")]
        [string]$Phase
    )

    switch ("{0}:{1}" -f $Button, $Phase) {
        "Left:Down" { return [OverlayNative]::WM_LBUTTONDOWN }
        "Left:Up" { return [OverlayNative]::WM_LBUTTONUP }
        "Left:DoubleClick" { return [OverlayNative]::WM_LBUTTONDBLCLK }
        "Right:Down" { return [OverlayNative]::WM_RBUTTONDOWN }
        "Right:Up" { return [OverlayNative]::WM_RBUTTONUP }
        "Right:DoubleClick" { return [OverlayNative]::WM_RBUTTONDBLCLK }
        "Middle:Down" { return [OverlayNative]::WM_MBUTTONDOWN }
        "Middle:Up" { return [OverlayNative]::WM_MBUTTONUP }
        "Middle:DoubleClick" { return [OverlayNative]::WM_MBUTTONDBLCLK }
        default { return $null }
    }
}

function Send-MouseInputToSource {
    param(
        [uint32]$Message,
        [System.Drawing.Point]$OverlayPoint,
        [uint32]$KeyState = 0,
        [int]$WheelDelta = 0,
        [switch]$UseCapturedTarget,
        [switch]$ActivateSource
    )

    if (-not $script:interactiveInputEnabled -or $script:currentSourceHandle -eq [IntPtr]::Zero) {
        return $false
    }

    $mappedPoint = Convert-OverlayPointToSource -OverlayPoint $OverlayPoint
    if ($null -eq $mappedPoint) {
        return $false
    }

    if ($ActivateSource) {
        [void][OverlayNative]::ActivateWindow($script:currentSourceHandle)
    }

    $preferredHandle = if ($UseCapturedTarget -and $script:capturedMouseTarget -ne [IntPtr]::Zero) { $script:capturedMouseTarget } else { [IntPtr]::Zero }
    $target = Resolve-InputTarget -ScreenX $mappedPoint.ScreenX -ScreenY $mappedPoint.ScreenY -PreferredHandle $preferredHandle
    if ($null -eq $target) {
        return $false
    }

    if ($Message -eq [OverlayNative]::WM_MOUSEWHEEL) {
        [OverlayNative]::PostMouseWheelMessage($target.Handle, $target.ScreenX, $target.ScreenY, $WheelDelta, $KeyState)
    }
    else {
        [OverlayNative]::PostMouseMessage($target.Handle, $Message, $target.ClientX, $target.ClientY, $KeyState)
    }

    return $true
}

function Send-KeyInputToSource {
    param(
        [System.Windows.Forms.KeyEventArgs]$EventArgs,
        [bool]$IsKeyUp
    )

    if (-not $script:interactiveInputEnabled -or $script:currentSourceHandle -eq [IntPtr]::Zero) {
        return
    }

    $altDown = (($EventArgs.Modifiers -band [System.Windows.Forms.Keys]::Alt) -ne 0)
    $message = if ($IsKeyUp) {
        if ($altDown) { [OverlayNative]::WM_SYSKEYUP } else { [OverlayNative]::WM_KEYUP }
    }
    else {
        if ($altDown) { [OverlayNative]::WM_SYSKEYDOWN } else { [OverlayNative]::WM_KEYDOWN }
    }

    $lParam = [OverlayNative]::BuildKeyLParam([int]$EventArgs.KeyCode, $IsKeyUp, $altDown)
    [void][OverlayNative]::ActivateWindow($script:currentSourceHandle)
    [OverlayNative]::PostKeyMessage($script:currentSourceHandle, $message, [int]$EventArgs.KeyCode, $lParam)
    $EventArgs.Handled = $true
    $EventArgs.SuppressKeyPress = $true
}

function Send-CharInputToSource {
    param([System.Windows.Forms.KeyPressEventArgs]$EventArgs)

    if (-not $script:interactiveInputEnabled -or $script:currentSourceHandle -eq [IntPtr]::Zero) {
        return
    }

    [void][OverlayNative]::ActivateWindow($script:currentSourceHandle)
    [OverlayNative]::PostCharMessage($script:currentSourceHandle, [int][char]$EventArgs.KeyChar)
    $EventArgs.Handled = $true
}

function Update-OverlayInteractionUi {
    if ($null -eq $script:overlayForm) {
        return
    }

    if ($script:interactiveInputEnabled) {
        $script:overlayForm.Cursor = [System.Windows.Forms.Cursors]::Cross
    }
    else {
        $script:overlayForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Set-OverlayOpacity {
    param([int]$Percent)

    $Percent = [Math]::Max(10, [Math]::Min(100, $Percent))
    $script:overlayOpacityPercent = $Percent

    if ($null -ne $script:opacitySlider -and $script:opacitySlider.Value -ne $Percent) {
        $script:opacitySlider.Value = $Percent
    }

    if ($null -ne $script:opacityValueLabel) {
        $script:opacityValueLabel.Text = ("{0}%" -f $Percent)
    }

    if ($null -ne $script:overlayForm) {
        $script:overlayForm.Opacity = $Percent / 100.0
    }
}

function Show-OverlayWindow {
    if ($null -eq $script:overlayForm -or $script:isExiting) {
        return
    }

    if (-not $script:overlayForm.Visible) {
        $script:overlayForm.Show()
    }

    $script:overlayForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    Apply-OverlayFlags
    Update-ThumbnailLayout
    Update-TrayMenu
}

function Hide-OverlayWindow {
    if ($null -eq $script:overlayForm) {
        return
    }

    if ($script:overlayForm.Visible) {
        $script:overlayForm.Hide()
    }

    Update-TrayMenu
}

function Show-ControlPanel {
    if ($null -eq $script:controlForm -or $script:isExiting) {
        return
    }

    if (-not $script:controlForm.Visible) {
        $script:controlForm.Show()
    }

    $script:controlForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    Apply-ControlFlags
    $script:controlForm.Activate()
    Update-TrayMenu
}

function Hide-ControlPanel {
    if ($null -eq $script:controlForm) {
        return
    }

    if ($script:controlForm.Visible) {
        $script:controlForm.Hide()
    }

    Update-TrayMenu
}

function Get-TextMarkupSegments {
    param([string]$Text)

    $segments = New-Object System.Collections.Generic.List[object]
    $value = if ($null -eq $Text) { "" } else { [string]$Text }
    $index = 0
    $isBold = $false

    while ($index -lt $value.Length) {
        $markerIndex = $value.IndexOf("**", $index, [System.StringComparison]::Ordinal)
        if ($markerIndex -lt 0) {
            $segments.Add([pscustomobject]@{
                Text = $value.Substring($index)
                Bold = $isBold
            })
            break
        }

        if ($markerIndex -gt $index) {
            $segments.Add([pscustomobject]@{
                Text = $value.Substring($index, $markerIndex - $index)
                Bold = $isBold
            })
        }

        $isBold = -not $isBold
        $index = $markerIndex + 2
    }

    if ($segments.Count -eq 0) {
        $segments.Add([pscustomobject]@{
            Text = ""
            Bold = $false
        })
    }

    return $segments
}

function Get-TextAlignmentValue {
    switch ($script:textAlignment) {
        "Left" { return [System.Windows.Forms.HorizontalAlignment]::Left }
        "Right" { return [System.Windows.Forms.HorizontalAlignment]::Right }
        default { return [System.Windows.Forms.HorizontalAlignment]::Center }
    }
}

function Render-OverlayText {
    if ($null -eq $script:overlayTextLabel) {
        return
    }

    $content = if ($null -ne $script:textInputBox -and -not [string]::IsNullOrWhiteSpace($script:textInputBox.Text)) {
        $script:textInputBox.Text
    }
    elseif (-not [string]::IsNullOrWhiteSpace($script:overlayTextLabel.Text)) {
        $script:overlayTextLabel.Text
    }
    else {
        "Sample overlay text"
    }

    $script:overlayTextLabel.SuspendLayout()
    $script:overlayTextLabel.Clear()
    $script:overlayTextLabel.SelectionStart = 0
    $script:overlayTextLabel.SelectionLength = 0

    foreach ($segment in Get-TextMarkupSegments -Text $content) {
        $style = if ($segment.Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
        $script:overlayTextLabel.SelectionFont = New-Object System.Drawing.Font("Segoe UI", [float]$script:textFontSize, $style)
        $script:overlayTextLabel.SelectionColor = [System.Drawing.Color]::White
        $script:overlayTextLabel.AppendText($segment.Text)
    }

    $script:overlayTextLabel.SelectAll()
    $script:overlayTextLabel.SelectionAlignment = Get-TextAlignmentValue
    $script:overlayTextLabel.DeselectAll()
    $script:overlayTextLabel.SelectionStart = 0
    $script:overlayTextLabel.ScrollToCaret()
    $script:overlayTextLabel.ResumeLayout()
}

function Set-OverlayText {
    param([string]$Text)

    $cleanText = if ([string]::IsNullOrWhiteSpace($Text)) { "Sample overlay text" } else { $Text }

    Clear-Thumbnail
    if ($null -ne $script:textInputBox) {
        $script:textInputBox.Text = $cleanText
    }

    Render-OverlayText
    $script:overlayTextLabel.Visible = $true
    $script:overlayHint.Visible = $false
    Show-OverlayWindow
    Set-Status "Overlay is showing text mode."
}

function Update-ModeUi {
    $windowMode = ($script:overlayMode -eq "Window")
    $textMode = ($script:overlayMode -eq "Text")
    $webMode = ($script:overlayMode -eq "Web")

    if ($null -ne $script:overlayModePicker) {
        $script:overlayModePicker.SelectedItem = if ($windowMode) {
            "Window preview"
        }
        elseif ($textMode) {
            "Text overlay"
        }
        else {
            "Web page"
        }
    }

    if ($null -ne $script:windowModePanel) {
        $script:windowModePanel.Visible = $windowMode
    }

    if ($null -ne $script:textModePanel) {
        $script:textModePanel.Visible = $textMode
    }

    if ($null -ne $script:webModePanel) {
        $script:webModePanel.Visible = $webMode
    }

    if ($null -ne $script:textAlignmentPicker -and $script:textAlignmentPicker.SelectedItem -ne $script:textAlignment) {
        $script:textAlignmentPicker.SelectedItem = $script:textAlignment
    }

    if ($null -ne $script:textFontSizePicker -and [int]$script:textFontSizePicker.Value -ne $script:textFontSize) {
        $script:textFontSizePicker.Value = $script:textFontSize
    }

    if ($null -ne $script:interactiveInputCheck) {
        $script:interactiveInputCheck.Enabled = $windowMode

        if (-not $windowMode -and $script:interactiveInputCheck.Checked) {
            $script:interactiveInputCheck.Checked = $false
        }
    }
}

function Set-OverlayMode {
    param([string]$Mode)

    if ([string]::IsNullOrWhiteSpace($Mode)) {
        return
    }

    $script:overlayMode = switch ($Mode) {
        "Text" { "Text" }
        "Web" { "Web" }
        default { "Window" }
    }
    Update-ModeUi

    if ($script:overlayMode -eq "Text") {
        $script:overlayHint.Text = "Text overlay mode is ready."
        if ($script:overlayTextLabel.Visible) {
            Render-OverlayText
        }
    }
    elseif ($script:overlayMode -eq "Web") {
        $script:overlayHint.Text = "Web page mode is ready."
    }
    else {
        $script:overlayHint.Text = "Pick a window in the control panel.`r`nThis overlay will be excluded from capture."
        if ($script:thumbnailHandle -eq [IntPtr]::Zero -and $null -ne $script:overlayTextLabel) {
            $script:overlayTextLabel.Visible = $false
            $script:overlayHint.Visible = $true
        }
    }
}

function Update-TrayMenu {
    if ($null -ne $script:trayShowControlItem) {
        $script:trayShowControlItem.Text = if ($script:controlForm.Visible) { "Hide control panel" } else { "Open control panel" }
    }

    if ($null -ne $script:trayToggleOverlayItem) {
        $script:trayToggleOverlayItem.Text = if ($script:overlayForm.Visible) { "Hide overlay" } else { "Show overlay" }
    }
}

function Exit-OverlayApplication {
    if ($script:isExiting) {
        return
    }

    $script:isExiting = $true
    $script:watchdogTimer.Stop()
    $script:selectionTimer.Stop()
    Clear-Thumbnail

    if ($null -ne $script:notifyIcon) {
        $script:notifyIcon.Visible = $false
        $script:notifyIcon.Dispose()
    }

    if ($null -ne $script:overlayForm -and -not $script:overlayForm.IsDisposed) {
        $script:overlayForm.Close()
        $script:overlayForm.Dispose()
    }

    if ($null -ne $script:controlForm -and -not $script:controlForm.IsDisposed) {
        $script:controlForm.Close()
        $script:controlForm.Dispose()
    }

    if ($null -ne $script:appContext) {
        $script:appContext.ExitThread()
    }
    else {
        [System.Windows.Forms.Application]::Exit()
    }
}

function Find-WindowItemByHandle {
    param([IntPtr]$Handle)

    foreach ($item in $script:windowPicker.Items) {
        if ($item.Handle -eq $Handle) {
            return $item
        }
    }

    return $null
}

function Refresh-WindowList {
    param([IntPtr]$PreferredHandle = [IntPtr]::Zero)

    $overlayHandle = if ($script:overlayForm -and $script:overlayForm.IsHandleCreated) { $script:overlayForm.Handle } else { [IntPtr]::Zero }
    $controlHandle = if ($script:controlForm -and $script:controlForm.IsHandleCreated) { $script:controlForm.Handle } else { [IntPtr]::Zero }
    $windows = [OverlayNative]::EnumerateWindows($overlayHandle, $controlHandle)

    $script:windowPicker.BeginUpdate()
    try {
        $script:windowPicker.Items.Clear()
        foreach ($window in $windows) {
            [void]$script:windowPicker.Items.Add($window)
        }
    }
    finally {
        $script:windowPicker.EndUpdate()
    }

    if ($script:windowPicker.Items.Count -eq 0) {
        Set-Status "No compatible windows found. Open a browser or app and click 'Refresh'."
        return
    }

    $preferred = $PreferredHandle
    if ($preferred -eq [IntPtr]::Zero -and $script:currentSourceHandle -ne [IntPtr]::Zero) {
        $preferred = $script:currentSourceHandle
    }

    if ($preferred -ne [IntPtr]::Zero) {
        $found = Find-WindowItemByHandle -Handle $preferred
        if ($null -ne $found) {
            $script:windowPicker.SelectedItem = $found
            return
        }
    }

    if ($script:windowPicker.SelectedIndex -lt 0) {
        $script:windowPicker.SelectedIndex = 0
    }
}

function Set-OverlaySource {
    param([OverlayNative+WindowInfo]$WindowInfo)

    Clear-Thumbnail

    if ($null -eq $WindowInfo) {
        Set-Status "No source window selected."
        return
    }

    if (-not $script:overlayForm.IsHandleCreated) {
        $null = $script:overlayForm.Handle
    }

    $thumbnail = [IntPtr]::Zero
    $hr = [OverlayNative]::RegisterThumbnail($script:overlayForm.Handle, $WindowInfo.Handle, [ref]$thumbnail)
    if ($hr -ne 0 -or $thumbnail -eq [IntPtr]::Zero) {
        Set-Status ("Could not attach window '{0}'. HRESULT: 0x{1}" -f $WindowInfo.Title, $hr.ToString("X8"))
        return
    }

    $script:thumbnailHandle = $thumbnail
    $script:currentSourceHandle = $WindowInfo.Handle
    $script:overlayTextLabel.Visible = $false
    $script:overlayHint.Visible = $false
    Show-OverlayWindow
    Update-ThumbnailLayout
    Set-Status ("Overlay is showing: {0}" -f $WindowInfo.ToString())
}

$script:overlayForm = New-Object System.Windows.Forms.Form
$script:overlayForm.Text = "Overlay Preview"
$script:overlayForm.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$script:overlayForm.Location = New-Object System.Drawing.Point(120, 120)
$script:overlayForm.Size = New-Object System.Drawing.Size(960, 540)
$script:overlayForm.MinimumSize = New-Object System.Drawing.Size(320, 180)
$script:overlayForm.BackColor = [System.Drawing.Color]::Black
$script:overlayForm.ForeColor = [System.Drawing.Color]::White
$script:overlayForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::SizableToolWindow
$script:overlayForm.TopMost = $true
$script:overlayForm.ShowInTaskbar = $false
$script:overlayForm.KeyPreview = $true

$script:overlayHint = New-Object System.Windows.Forms.Label
$script:overlayHint.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:overlayHint.Text = "Pick a window in the control panel.`r`nThis overlay will be excluded from capture."
$script:overlayHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$script:overlayHint.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Regular)
$script:overlayHint.ForeColor = [System.Drawing.Color]::FromArgb(180, 255, 255, 255)
$script:overlayForm.Controls.Add($script:overlayHint)

$script:overlayTextLabel = New-Object System.Windows.Forms.RichTextBox
$script:overlayTextLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$script:overlayTextLabel.ReadOnly = $true
$script:overlayTextLabel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$script:overlayTextLabel.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$script:overlayTextLabel.WordWrap = $true
$script:overlayTextLabel.DetectUrls = $false
$script:overlayTextLabel.ShortcutsEnabled = $true
$script:overlayTextLabel.Font = New-Object System.Drawing.Font("Segoe UI", 28, [System.Drawing.FontStyle]::Regular)
$script:overlayTextLabel.ForeColor = [System.Drawing.Color]::White
$script:overlayTextLabel.BackColor = [System.Drawing.Color]::Black
$script:overlayTextLabel.Visible = $false
$script:overlayForm.Controls.Add($script:overlayTextLabel)
$script:overlayTextLabel.BringToFront()

$script:controlForm = New-Object System.Windows.Forms.Form
$script:controlForm.Text = "Overlay Mirror Control"
$script:controlForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$script:controlForm.Size = New-Object System.Drawing.Size(620, 690)
$script:controlForm.MinimumSize = New-Object System.Drawing.Size(620, 690)
$script:controlForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$script:controlForm.MaximizeBox = $false
$script:controlForm.ShowInTaskbar = $false

$font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:controlForm.Font = $font

$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Location = New-Object System.Drawing.Point(12, 15)
$modeLabel.Size = New-Object System.Drawing.Size(80, 22)
$modeLabel.Text = "Overlay mode"
$script:controlForm.Controls.Add($modeLabel)

$script:overlayModePicker = New-Object System.Windows.Forms.ComboBox
$script:overlayModePicker.Location = New-Object System.Drawing.Point(102, 12)
$script:overlayModePicker.Size = New-Object System.Drawing.Size(180, 28)
$script:overlayModePicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$script:overlayModePicker.Items.Add("Window preview")
[void]$script:overlayModePicker.Items.Add("Text overlay")
[void]$script:overlayModePicker.Items.Add("Web page")
$script:controlForm.Controls.Add($script:overlayModePicker)

$script:windowModePanel = New-Object System.Windows.Forms.GroupBox
$script:windowModePanel.Location = New-Object System.Drawing.Point(12, 50)
$script:windowModePanel.Size = New-Object System.Drawing.Size(590, 118)
$script:windowModePanel.Text = "Window source"
$script:controlForm.Controls.Add($script:windowModePanel)

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Location = New-Object System.Drawing.Point(12, 27)
$sourceLabel.Size = New-Object System.Drawing.Size(120, 20)
$sourceLabel.Text = "Source window"
$script:windowModePanel.Controls.Add($sourceLabel)

$script:windowPicker = New-Object System.Windows.Forms.ComboBox
$script:windowPicker.Location = New-Object System.Drawing.Point(12, 50)
$script:windowPicker.Size = New-Object System.Drawing.Size(470, 28)
$script:windowPicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:windowModePanel.Controls.Add($script:windowPicker)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(488, 49)
$refreshButton.Size = New-Object System.Drawing.Size(90, 30)
$refreshButton.Text = "Refresh"
$script:windowModePanel.Controls.Add($refreshButton)

$applyWindowButton = New-Object System.Windows.Forms.Button
$applyWindowButton.Location = New-Object System.Drawing.Point(12, 84)
$applyWindowButton.Size = New-Object System.Drawing.Size(278, 28)
$applyWindowButton.Text = "Show selected window"
$script:windowModePanel.Controls.Add($applyWindowButton)

$captureActiveButton = New-Object System.Windows.Forms.Button
$captureActiveButton.Location = New-Object System.Drawing.Point(300, 84)
$captureActiveButton.Size = New-Object System.Drawing.Size(278, 28)
$captureActiveButton.Text = "Use active window in 3 sec"
$script:windowModePanel.Controls.Add($captureActiveButton)

$script:textModePanel = New-Object System.Windows.Forms.GroupBox
$script:textModePanel.Location = New-Object System.Drawing.Point(12, 176)
$script:textModePanel.Size = New-Object System.Drawing.Size(590, 190)
$script:textModePanel.Text = "Text overlay"
$script:controlForm.Controls.Add($script:textModePanel)

$textInputLabel = New-Object System.Windows.Forms.Label
$textInputLabel.Location = New-Object System.Drawing.Point(12, 25)
$textInputLabel.Size = New-Object System.Drawing.Size(200, 20)
$textInputLabel.Text = "Text to show in overlay"
$script:textModePanel.Controls.Add($textInputLabel)

$script:textInputBox = New-Object System.Windows.Forms.RichTextBox
$script:textInputBox.Location = New-Object System.Drawing.Point(12, 48)
$script:textInputBox.Size = New-Object System.Drawing.Size(566, 82)
$script:textInputBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$script:textInputBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$script:textInputBox.WordWrap = $true
$script:textInputBox.DetectUrls = $false
$script:textInputBox.Text = "Sample overlay text"
$script:textModePanel.Controls.Add($script:textInputBox)

$markupHintLabel = New-Object System.Windows.Forms.Label
$markupHintLabel.Location = New-Object System.Drawing.Point(12, 134)
$markupHintLabel.Size = New-Object System.Drawing.Size(566, 20)
$markupHintLabel.Text = "Use **double asterisks** for bold text. Default text stays regular."
$markupHintLabel.ForeColor = [System.Drawing.Color]::DimGray
$script:textModePanel.Controls.Add($markupHintLabel)

$alignmentLabel = New-Object System.Windows.Forms.Label
$alignmentLabel.Location = New-Object System.Drawing.Point(12, 158)
$alignmentLabel.Size = New-Object System.Drawing.Size(60, 20)
$alignmentLabel.Text = "Align"
$script:textModePanel.Controls.Add($alignmentLabel)

$script:textAlignmentPicker = New-Object System.Windows.Forms.ComboBox
$script:textAlignmentPicker.Location = New-Object System.Drawing.Point(74, 155)
$script:textAlignmentPicker.Size = New-Object System.Drawing.Size(110, 28)
$script:textAlignmentPicker.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$script:textAlignmentPicker.Items.Add("Left")
[void]$script:textAlignmentPicker.Items.Add("Center")
[void]$script:textAlignmentPicker.Items.Add("Right")
$script:textAlignmentPicker.SelectedItem = "Center"
$script:textModePanel.Controls.Add($script:textAlignmentPicker)

$fontSizeLabel = New-Object System.Windows.Forms.Label
$fontSizeLabel.Location = New-Object System.Drawing.Point(204, 158)
$fontSizeLabel.Size = New-Object System.Drawing.Size(62, 20)
$fontSizeLabel.Text = "Font size"
$script:textModePanel.Controls.Add($fontSizeLabel)

$script:textFontSizePicker = New-Object System.Windows.Forms.NumericUpDown
$script:textFontSizePicker.Location = New-Object System.Drawing.Point(272, 156)
$script:textFontSizePicker.Size = New-Object System.Drawing.Size(68, 24)
$script:textFontSizePicker.Minimum = 10
$script:textFontSizePicker.Maximum = 96
$script:textFontSizePicker.Value = 28
$script:textModePanel.Controls.Add($script:textFontSizePicker)

$applyTextButton = New-Object System.Windows.Forms.Button
$applyTextButton.Location = New-Object System.Drawing.Point(356, 154)
$applyTextButton.Size = New-Object System.Drawing.Size(222, 28)
$applyTextButton.Text = "Show text overlay"
$script:textModePanel.Controls.Add($applyTextButton)

$script:webModePanel = New-Object System.Windows.Forms.GroupBox
$script:webModePanel.Location = New-Object System.Drawing.Point(12, 176)
$script:webModePanel.Size = New-Object System.Drawing.Size(590, 190)
$script:webModePanel.Text = "Web page"
$script:webModePanel.Visible = $false
$script:controlForm.Controls.Add($script:webModePanel)

$webUrlLabel = New-Object System.Windows.Forms.Label
$webUrlLabel.Location = New-Object System.Drawing.Point(12, 25)
$webUrlLabel.Size = New-Object System.Drawing.Size(40, 20)
$webUrlLabel.Text = "URL"
$script:webModePanel.Controls.Add($webUrlLabel)

$script:webUrlBox = New-Object System.Windows.Forms.TextBox
$script:webUrlBox.Location = New-Object System.Drawing.Point(58, 22)
$script:webUrlBox.Size = New-Object System.Drawing.Size(430, 24)
$script:webUrlBox.Text = "https://www.google.com"
$script:webModePanel.Controls.Add($script:webUrlBox)

$script:webGoButton = New-Object System.Windows.Forms.Button
$script:webGoButton.Location = New-Object System.Drawing.Point(494, 20)
$script:webGoButton.Size = New-Object System.Drawing.Size(84, 28)
$script:webGoButton.Text = "Open"
$script:webModePanel.Controls.Add($script:webGoButton)

$script:webBackButton = New-Object System.Windows.Forms.Button
$script:webBackButton.Location = New-Object System.Drawing.Point(12, 58)
$script:webBackButton.Size = New-Object System.Drawing.Size(84, 28)
$script:webBackButton.Text = "Back"
$script:webBackButton.Enabled = $false
$script:webModePanel.Controls.Add($script:webBackButton)

$script:webForwardButton = New-Object System.Windows.Forms.Button
$script:webForwardButton.Location = New-Object System.Drawing.Point(102, 58)
$script:webForwardButton.Size = New-Object System.Drawing.Size(84, 28)
$script:webForwardButton.Text = "Forward"
$script:webForwardButton.Enabled = $false
$script:webModePanel.Controls.Add($script:webForwardButton)

$script:webReloadButton = New-Object System.Windows.Forms.Button
$script:webReloadButton.Location = New-Object System.Drawing.Point(192, 58)
$script:webReloadButton.Size = New-Object System.Drawing.Size(84, 28)
$script:webReloadButton.Text = "Reload"
$script:webModePanel.Controls.Add($script:webReloadButton)

$script:webOpenExternalButton = New-Object System.Windows.Forms.Button
$script:webOpenExternalButton.Location = New-Object System.Drawing.Point(282, 58)
$script:webOpenExternalButton.Size = New-Object System.Drawing.Size(128, 28)
$script:webOpenExternalButton.Text = "Open in browser"
$script:webModePanel.Controls.Add($script:webOpenExternalButton)

$webZoomLabel = New-Object System.Windows.Forms.Label
$webZoomLabel.Location = New-Object System.Drawing.Point(430, 63)
$webZoomLabel.Size = New-Object System.Drawing.Size(46, 20)
$webZoomLabel.Text = "Zoom"
$script:webModePanel.Controls.Add($webZoomLabel)

$script:webZoomPicker = New-Object System.Windows.Forms.NumericUpDown
$script:webZoomPicker.Location = New-Object System.Drawing.Point(482, 60)
$script:webZoomPicker.Size = New-Object System.Drawing.Size(60, 24)
$script:webZoomPicker.Minimum = 25
$script:webZoomPicker.Maximum = 300
$script:webZoomPicker.Increment = 25
$script:webZoomPicker.Value = 100
$script:webModePanel.Controls.Add($script:webZoomPicker)

$webZoomPercentLabel = New-Object System.Windows.Forms.Label
$webZoomPercentLabel.Location = New-Object System.Drawing.Point(546, 63)
$webZoomPercentLabel.Size = New-Object System.Drawing.Size(28, 20)
$webZoomPercentLabel.Text = "%"
$script:webModePanel.Controls.Add($webZoomPercentLabel)

$webHintLabel = New-Object System.Windows.Forms.Label
$webHintLabel.Location = New-Object System.Drawing.Point(12, 98)
$webHintLabel.Size = New-Object System.Drawing.Size(566, 44)
$webHintLabel.Text = "Open any page inside the overlay. Leave click-through off if you want to click, scroll, type, or use the page directly."
$webHintLabel.ForeColor = [System.Drawing.Color]::DimGray
$script:webModePanel.Controls.Add($webHintLabel)

$webHintLabel2 = New-Object System.Windows.Forms.Label
$webHintLabel2.Location = New-Object System.Drawing.Point(12, 146)
$webHintLabel2.Size = New-Object System.Drawing.Size(566, 34)
$webHintLabel2.Text = "You can still resize the overlay window itself, and zoom from 25% to 300% using the control panel."
$webHintLabel2.ForeColor = [System.Drawing.Color]::DimGray
$script:webModePanel.Controls.Add($webHintLabel2)

$opacityPanel = New-Object System.Windows.Forms.GroupBox
$opacityPanel.Location = New-Object System.Drawing.Point(12, 374)
$opacityPanel.Size = New-Object System.Drawing.Size(590, 82)
$opacityPanel.Text = "Overlay opacity"
$script:controlForm.Controls.Add($opacityPanel)

$opacityLabel = New-Object System.Windows.Forms.Label
$opacityLabel.Location = New-Object System.Drawing.Point(12, 32)
$opacityLabel.Size = New-Object System.Drawing.Size(120, 20)
$opacityLabel.Text = "Transparency"
$opacityPanel.Controls.Add($opacityLabel)

$script:opacitySlider = New-Object System.Windows.Forms.TrackBar
$script:opacitySlider.Location = New-Object System.Drawing.Point(128, 22)
$script:opacitySlider.Size = New-Object System.Drawing.Size(380, 45)
$script:opacitySlider.Minimum = 10
$script:opacitySlider.Maximum = 100
$script:opacitySlider.TickFrequency = 10
$script:opacitySlider.SmallChange = 5
$script:opacitySlider.LargeChange = 10
$script:opacitySlider.Value = 100
$opacityPanel.Controls.Add($script:opacitySlider)

$script:opacityValueLabel = New-Object System.Windows.Forms.Label
$script:opacityValueLabel.Location = New-Object System.Drawing.Point(518, 32)
$script:opacityValueLabel.Size = New-Object System.Drawing.Size(60, 20)
$script:opacityValueLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$script:opacityValueLabel.Text = "100%"
$opacityPanel.Controls.Add($script:opacityValueLabel)

$behaviorPanel = New-Object System.Windows.Forms.GroupBox
$behaviorPanel.Location = New-Object System.Drawing.Point(12, 464)
$behaviorPanel.Size = New-Object System.Drawing.Size(590, 116)
$behaviorPanel.Text = "Behavior"
$script:controlForm.Controls.Add($behaviorPanel)

$excludeCheck = New-Object System.Windows.Forms.CheckBox
$excludeCheck.Location = New-Object System.Drawing.Point(12, 22)
$excludeCheck.Size = New-Object System.Drawing.Size(290, 24)
$excludeCheck.Checked = $true
$excludeCheck.Text = "Hide overlay from OBS / Discord / Zoom"
$behaviorPanel.Controls.Add($excludeCheck)

$script:clickThroughCheck = New-Object System.Windows.Forms.CheckBox
$script:clickThroughCheck.Location = New-Object System.Drawing.Point(310, 22)
$script:clickThroughCheck.Size = New-Object System.Drawing.Size(200, 24)
$script:clickThroughCheck.Checked = $script:clickThroughEnabled
$script:clickThroughCheck.Text = "Click-through mode"
$behaviorPanel.Controls.Add($script:clickThroughCheck)

$script:interactiveInputCheck = New-Object System.Windows.Forms.CheckBox
$script:interactiveInputCheck.Location = New-Object System.Drawing.Point(12, 50)
$script:interactiveInputCheck.Size = New-Object System.Drawing.Size(200, 24)
$script:interactiveInputCheck.Checked = $false
$script:interactiveInputCheck.Text = "Interactive input mode"
$behaviorPanel.Controls.Add($script:interactiveInputCheck)

$topMostCheck = New-Object System.Windows.Forms.CheckBox
$topMostCheck.Location = New-Object System.Drawing.Point(310, 50)
$topMostCheck.Size = New-Object System.Drawing.Size(220, 24)
$topMostCheck.Checked = $true
$topMostCheck.Text = "Keep overlay above all windows"
$behaviorPanel.Controls.Add($topMostCheck)

$script:controlTopMostCheck = New-Object System.Windows.Forms.CheckBox
$script:controlTopMostCheck.Location = New-Object System.Drawing.Point(12, 78)
$script:controlTopMostCheck.Size = New-Object System.Drawing.Size(220, 24)
$script:controlTopMostCheck.Checked = $false
$script:controlTopMostCheck.Text = "Keep control panel above all windows"
$behaviorPanel.Controls.Add($script:controlTopMostCheck)

$script:controlCaptureCheck = New-Object System.Windows.Forms.CheckBox
$script:controlCaptureCheck.Location = New-Object System.Drawing.Point(310, 78)
$script:controlCaptureCheck.Size = New-Object System.Drawing.Size(230, 24)
$script:controlCaptureCheck.Checked = $true
$script:controlCaptureCheck.Text = "Hide control panel from capture"
$behaviorPanel.Controls.Add($script:controlCaptureCheck)

$helpLabel = New-Object System.Windows.Forms.Label
$helpLabel.Location = New-Object System.Drawing.Point(12, 587)
$helpLabel.Size = New-Object System.Drawing.Size(590, 20)
$helpLabel.Text = "Close buttons hide windows to tray. Control panel can be topmost and excluded from supported capture apps too."
$helpLabel.ForeColor = [System.Drawing.Color]::DimGray
$script:controlForm.Controls.Add($helpLabel)

$script:statusLabel = New-Object System.Windows.Forms.Label
$script:statusLabel.Location = New-Object System.Drawing.Point(12, 612)
$script:statusLabel.Size = New-Object System.Drawing.Size(590, 20)
$script:statusLabel.Text = "Ready."
$script:controlForm.Controls.Add($script:statusLabel)

$script:overlayModePicker.Add_SelectedIndexChanged({
    if ($script:overlayModePicker.SelectedItem -eq "Text overlay") {
        Set-OverlayMode -Mode "Text"
        Set-Status "Text overlay mode selected."
    }
    elseif ($script:overlayModePicker.SelectedItem -eq "Web page") {
        Set-OverlayMode -Mode "Web"
        Set-Status "Web page mode selected."
    }
    else {
        Set-OverlayMode -Mode "Window"
        Set-Status "Window preview mode selected."
    }
})

$refreshButton.Add_Click({
    Refresh-WindowList
})

$applyWindowButton.Add_Click({
    if ($script:windowPicker.SelectedItem -is [OverlayNative+WindowInfo]) {
        Set-OverlayMode -Mode "Window"
        Set-OverlaySource -WindowInfo $script:windowPicker.SelectedItem
    }
    else {
        Set-Status "Pick a source window first."
    }
})

$captureActiveButton.Add_Click({
    Set-OverlayMode -Mode "Window"
    Set-Status "Switch to the target window now. Active window will be captured in 3 seconds."
    $script:selectionTimer.Stop()
    $script:selectionTimer.Start()
})

$applyTextButton.Add_Click({
    $script:textAlignment = [string]$script:textAlignmentPicker.SelectedItem
    $script:textFontSize = [int]$script:textFontSizePicker.Value
    Set-OverlayMode -Mode "Text"
    Set-OverlayText -Text $script:textInputBox.Text
})

$script:textAlignmentPicker.Add_SelectedIndexChanged({
    if ($null -ne $script:textAlignmentPicker.SelectedItem) {
        $script:textAlignment = [string]$script:textAlignmentPicker.SelectedItem
    }

    if ($script:overlayMode -eq "Text" -and $script:overlayTextLabel.Visible) {
        Render-OverlayText
        Set-Status ("Text alignment set to {0}." -f $script:textAlignment.ToLowerInvariant())
    }
})

$script:textFontSizePicker.Add_ValueChanged({
    $script:textFontSize = [int]$script:textFontSizePicker.Value

    if ($script:overlayMode -eq "Text" -and $script:overlayTextLabel.Visible) {
        Render-OverlayText
        Set-Status ("Text font size set to {0}." -f $script:textFontSize)
    }
})

$script:webGoButton.Add_Click({
    Set-OverlayMode -Mode "Web"
    Show-WebPage -Url $script:webUrlBox.Text
})

$script:webUrlBox.Add_KeyDown({
    param($sender, $eventArgs)

    if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $eventArgs.Handled = $true
        $eventArgs.SuppressKeyPress = $true
        Set-OverlayMode -Mode "Web"
        Show-WebPage -Url $script:webUrlBox.Text
    }
})

$script:webBackButton.Add_Click({
    if ($script:webViewReady -and $null -ne $script:webViewControl.CoreWebView2 -and $script:webViewControl.CoreWebView2.CanGoBack) {
        $script:webViewControl.CoreWebView2.GoBack()
    }
})

$script:webForwardButton.Add_Click({
    if ($script:webViewReady -and $null -ne $script:webViewControl.CoreWebView2 -and $script:webViewControl.CoreWebView2.CanGoForward) {
        $script:webViewControl.CoreWebView2.GoForward()
    }
})

$script:webReloadButton.Add_Click({
    if ($script:webViewReady -and $null -ne $script:webViewControl.CoreWebView2) {
        $script:webViewControl.CoreWebView2.Reload()
        Set-Status "Reloading web page."
    }
    else {
        Set-OverlayMode -Mode "Web"
        Show-WebPage -Url $script:webUrlBox.Text
    }
})

$script:webOpenExternalButton.Add_Click({
    $url = Normalize-WebUrl -Url $script:webUrlBox.Text
    Start-Process $url
})

$script:webZoomPicker.Add_ValueChanged({
    Set-WebZoom -Percent ([int]$script:webZoomPicker.Value)

    if ($script:overlayMode -eq "Web" -and $script:webViewControl.Visible) {
        Set-Status ("Web zoom set to {0}%." -f $script:webZoomPercent)
    }
})

$script:opacitySlider.Add_Scroll({
    Set-OverlayOpacity -Percent $script:opacitySlider.Value
    Set-Status ("Overlay opacity set to {0}%." -f $script:opacitySlider.Value)
})

$excludeCheck.Add_CheckedChanged({
    $script:captureExcluded = $excludeCheck.Checked
    Apply-OverlayFlags

    if ($script:captureExcluded) {
        Set-Status "Overlay capture exclusion is enabled for apps that honor Windows Display Affinity."
    }
    else {
        Set-Status "Capture exclusion is disabled."
    }
})

$script:clickThroughCheck.Add_CheckedChanged({
    $script:clickThroughEnabled = $script:clickThroughCheck.Checked
    Apply-OverlayFlags

    if ($script:clickThroughEnabled) {
        Set-Status "Click-through is enabled. Disable it if you need to move or resize the overlay."
    }
    else {
        Set-Status "Click-through is disabled."
    }
})

$script:interactiveInputCheck.Add_CheckedChanged({
    $script:interactiveInputEnabled = $script:interactiveInputCheck.Checked

    if ($script:interactiveInputEnabled) {
        if ($script:clickThroughCheck.Checked) {
            $script:clickThroughCheck.Checked = $false
        }

        $script:clickThroughCheck.Enabled = $false
        Show-OverlayWindow
        Set-Status "Interactive mode is enabled. Click the overlay to control the source window."
    }
    else {
        $script:clickThroughCheck.Enabled = $true
        $script:capturedMouseTarget = [IntPtr]::Zero
        $script:capturedMouseButton = [System.Windows.Forms.MouseButtons]::None
        $script:overlayForm.Capture = $false
        Set-Status "Interactive mode is disabled."
    }

    Update-OverlayInteractionUi
    Apply-OverlayFlags
})

$topMostCheck.Add_CheckedChanged({
    $script:overlayForm.TopMost = $topMostCheck.Checked
})

$script:controlTopMostCheck.Add_CheckedChanged({
    $script:controlTopMostEnabled = $script:controlTopMostCheck.Checked
    Apply-ControlFlags

    if ($script:controlTopMostEnabled) {
        Set-Status "Control panel topmost is enabled."
    }
    else {
        Set-Status "Control panel topmost is disabled."
    }
})

$script:controlCaptureCheck.Add_CheckedChanged({
    $script:controlCaptureExcluded = $script:controlCaptureCheck.Checked
    Apply-ControlFlags

    if ($script:controlCaptureExcluded) {
        Set-Status "Control panel capture exclusion is enabled."
    }
    else {
        Set-Status "Control panel capture exclusion is disabled."
    }
})

$script:overlayForm.Add_MouseDown({
    param($sender, $eventArgs)

    if (-not $script:interactiveInputEnabled) {
        return
    }

    $message = Get-MouseMessage -Button $eventArgs.Button -Phase Down
    if ($null -eq $message) {
        return
    }

    $keyState = Get-MouseKeyState -AddButton $eventArgs.Button
    if (Send-MouseInputToSource -Message $message -OverlayPoint $eventArgs.Location -KeyState $keyState -ActivateSource) {
        $mappedPoint = Convert-OverlayPointToSource -OverlayPoint $eventArgs.Location
        if ($null -ne $mappedPoint) {
            $resolvedTarget = Resolve-InputTarget -ScreenX $mappedPoint.ScreenX -ScreenY $mappedPoint.ScreenY
            if ($null -ne $resolvedTarget) {
                $script:capturedMouseTarget = $resolvedTarget.Handle
                $script:capturedMouseButton = $eventArgs.Button
                $script:overlayForm.Capture = $true
            }
        }

        $script:overlayForm.Focus()
    }
})

$script:overlayForm.Add_MouseUp({
    param($sender, $eventArgs)

    if (-not $script:interactiveInputEnabled) {
        return
    }

    $message = Get-MouseMessage -Button $eventArgs.Button -Phase Up
    if ($null -eq $message) {
        return
    }

    $keyState = Get-MouseKeyState -RemoveButton $eventArgs.Button
    [void](Send-MouseInputToSource -Message $message -OverlayPoint $eventArgs.Location -KeyState $keyState -UseCapturedTarget)

    if ($script:capturedMouseButton -eq $eventArgs.Button) {
        $script:capturedMouseTarget = [IntPtr]::Zero
        $script:capturedMouseButton = [System.Windows.Forms.MouseButtons]::None
        $script:overlayForm.Capture = $false
    }
})

$script:overlayForm.Add_MouseMove({
    param($sender, $eventArgs)

    if (-not $script:interactiveInputEnabled) {
        return
    }

    $useCapture = ($script:capturedMouseTarget -ne [IntPtr]::Zero)
    $keyState = Get-MouseKeyState
    [void](Send-MouseInputToSource -Message ([OverlayNative]::WM_MOUSEMOVE) -OverlayPoint $eventArgs.Location -KeyState $keyState -UseCapturedTarget:$useCapture)
})

$script:overlayForm.Add_MouseWheel({
    param($sender, $eventArgs)

    if (-not $script:interactiveInputEnabled) {
        return
    }

    $keyState = Get-MouseKeyState
    [void](Send-MouseInputToSource -Message ([OverlayNative]::WM_MOUSEWHEEL) -OverlayPoint $eventArgs.Location -KeyState $keyState -WheelDelta $eventArgs.Delta -ActivateSource)
})

$script:overlayForm.Add_MouseDoubleClick({
    param($sender, $eventArgs)

    if (-not $script:interactiveInputEnabled) {
        return
    }

    $message = Get-MouseMessage -Button $eventArgs.Button -Phase DoubleClick
    if ($null -eq $message) {
        return
    }

    $keyState = Get-MouseKeyState -AddButton $eventArgs.Button
    [void](Send-MouseInputToSource -Message $message -OverlayPoint $eventArgs.Location -KeyState $keyState -ActivateSource)
})

$script:overlayForm.Add_KeyDown({
    param($sender, $eventArgs)
    Send-KeyInputToSource -EventArgs $eventArgs -IsKeyUp:$false
})

$script:overlayForm.Add_KeyUp({
    param($sender, $eventArgs)
    Send-KeyInputToSource -EventArgs $eventArgs -IsKeyUp:$true
})

$script:overlayForm.Add_KeyPress({
    param($sender, $eventArgs)
    Send-CharInputToSource -EventArgs $eventArgs
})

$script:selectionTimer = New-Object System.Windows.Forms.Timer
$script:selectionTimer.Interval = 3000
$script:selectionTimer.Add_Tick({
    $script:selectionTimer.Stop()
    $overlayHandle = if ($script:overlayForm.IsHandleCreated) { $script:overlayForm.Handle } else { [IntPtr]::Zero }
    $controlHandle = if ($script:controlForm.IsHandleCreated) { $script:controlForm.Handle } else { [IntPtr]::Zero }
    $foreground = [OverlayNative]::GetForegroundWindowInfo($overlayHandle, $controlHandle)

    if ($null -eq $foreground) {
        Set-Status "Could not capture the active window. Try again."
        return
    }

    Refresh-WindowList -PreferredHandle $foreground.Handle
    Set-OverlaySource -WindowInfo $foreground
})

$script:watchdogTimer = New-Object System.Windows.Forms.Timer
$script:watchdogTimer.Interval = 1500
$script:watchdogTimer.Add_Tick({
    if ($script:overlayMode -eq "Window" -and $script:currentSourceHandle -ne [IntPtr]::Zero -and -not [OverlayNative]::IsValidWindow($script:currentSourceHandle)) {
        Clear-Thumbnail
        Set-Status "The source window was closed. Pick another one."
        Refresh-WindowList
    }
    elseif ($script:overlayMode -eq "Window" -and $script:currentSourceHandle -ne [IntPtr]::Zero) {
        Update-ThumbnailLayout
    }
})

$script:overlayForm.Add_HandleCreated({
    Apply-OverlayFlags
    Set-OverlayOpacity -Percent $script:overlayOpacityPercent

    if ($null -ne $script:webViewControl -and -not $script:webViewReady) {
        [void]$script:webViewControl.EnsureCoreWebView2Async($null)
    }
})

$script:overlayForm.Add_Shown({
    Apply-OverlayFlags
    Set-OverlayOpacity -Percent $script:overlayOpacityPercent
    Update-ThumbnailLayout
    Update-TrayMenu
})

$script:overlayForm.Add_SizeChanged({
    Update-ThumbnailLayout
})

$script:overlayForm.Add_VisibleChanged({
    Update-TrayMenu
})

$script:overlayForm.Add_FormClosing({
    param($sender, $eventArgs)

    if ($script:isExiting) {
        return
    }

    $eventArgs.Cancel = $true
    Hide-OverlayWindow
    Set-Status "Overlay hidden to tray."
})

$script:controlForm.Add_VisibleChanged({
    if ($script:controlForm.Visible) {
        Apply-ControlFlags
        Refresh-WindowList
    }

    Update-TrayMenu
})

$script:controlForm.Add_HandleCreated({
    Apply-ControlFlags
})

$script:controlForm.Add_FormClosing({
    param($sender, $eventArgs)

    if ($script:isExiting) {
        return
    }

    $eventArgs.Cancel = $true
    Hide-ControlPanel
    Set-Status "Control panel hidden to tray."
})

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:trayShowControlItem = New-Object System.Windows.Forms.ToolStripMenuItem("Open control panel")
$script:trayToggleOverlayItem = New-Object System.Windows.Forms.ToolStripMenuItem("Show overlay")
$exitTrayItem = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
[void]$trayMenu.Items.Add($script:trayShowControlItem)
[void]$trayMenu.Items.Add($script:trayToggleOverlayItem)
[void]$trayMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$trayMenu.Items.Add($exitTrayItem)

$script:trayShowControlItem.Add_Click({
    if ($script:controlForm.Visible) {
        Hide-ControlPanel
    }
    else {
        Show-ControlPanel
    }
})

$script:trayToggleOverlayItem.Add_Click({
    if ($script:overlayForm.Visible) {
        Hide-OverlayWindow
    }
    else {
        Show-OverlayWindow
    }
})

$exitTrayItem.Add_Click({
    Exit-OverlayApplication
})

$script:notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$script:notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
$script:notifyIcon.Text = "Overlay Mirror"
$script:notifyIcon.ContextMenuStrip = $trayMenu
$script:notifyIcon.Visible = $true
$script:notifyIcon.Add_DoubleClick({
    Show-ControlPanel
})

Set-OverlayMode -Mode "Window"
Set-OverlayOpacity -Percent $script:overlayOpacityPercent
Refresh-WindowList
$script:watchdogTimer.Start()
Update-TrayMenu
Set-Status "Overlay Mirror is running in the tray."

if ($SmokeTest) {
    Apply-OverlayFlags
    Apply-ControlFlags
    Set-OverlayOpacity -Percent $script:overlayOpacityPercent
    $script:textAlignment = "Right"
    $script:textFontSize = 24
    Set-OverlayMode -Mode "Text"
    Set-OverlayText -Text "Smoke **test** text"
    Set-Status "Smoke test passed."
    if ($null -ne $script:notifyIcon) {
        $script:notifyIcon.Visible = $false
        $script:notifyIcon.Dispose()
    }
    if ($null -ne $script:webViewControl) {
        $script:webViewControl.Dispose()
    }
    if ($null -ne $script:overlayForm) {
        $script:overlayForm.Dispose()
    }
    if ($null -ne $script:controlForm) {
        $script:controlForm.Dispose()
    }
    return
}

$script:appContext = New-Object System.Windows.Forms.ApplicationContext
$script:notifyIcon.ShowBalloonTip(
    3000,
    "Overlay Mirror",
    "Overlay Mirror is running in the tray. Double-click the tray icon to open controls.",
    [System.Windows.Forms.ToolTipIcon]::Info
)
[void][System.Windows.Forms.Application]::Run($script:appContext)
