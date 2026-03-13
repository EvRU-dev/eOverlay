# eOverlay

Tray-first Windows overlay app that can:

- mirror a live desktop window or browser into a separate overlay;
- show a formatted text overlay with scroll, alignment, font size, and `**bold**` markup;
- open a normal embedded web page with navigation controls and zoom;
- autosave UI state, overlay settings, text, web page state, window positions, and hotkeys;
- register customizable global hotkeys that work across Windows;
- hide the overlay and control panel from supported capture apps;
- run without a visible console window.

## Main Modes

### 1. Window Preview

Use this when you want to mirror any visible app window into the overlay.

Controls:

- `Refresh`: reload the list of windows.
- `Show selected window`: attach the chosen window to the overlay.
- `Use active window in 3 sec`: switch to another app and capture it after a short delay.
- `Interactive input mode`: forward mouse, wheel, and basic keyboard input from the overlay into the source window.
- `Click-through mode`: let the mouse pass through the overlay instead of interacting with it.

### 2. Text Overlay

Use this when you want the overlay to show custom text instead of a captured window.

Controls:

- text editor with vertical scrolling;
- `**double asterisks**` for bold fragments;
- `Align`: `Left`, `Center`, or `Right`;
- `Font size`: from `10` to `96`;
- `Show text overlay`: apply the current text and formatting.

### 3. Web Page

Use this when you want the overlay itself to show a web page.

Controls:

- URL field;
- `Open`;
- `Back`;
- `Forward`;
- `Reload`;
- `Open in browser`;
- zoom from `25%` to `300%`.

This mode uses WebView2. The repository includes the required WinForms wrapper DLLs, and the target machine still needs the Microsoft Edge WebView2 Runtime.

## Overlay And Control Panel

The app has two separate windows:

- `eOverlay Preview`: the actual overlay window;
- `eOverlay Control`: the control panel.

Behavior options:

- `Overlay opacity`;
- `Keep overlay above all windows`;
- `Hide overlay from OBS / Discord / Zoom`;
- `Keep control panel above all windows`;
- `Hide control panel from capture`.

The close button does not exit the app. It hides the window to the tray.

## Autosave And Hotkeys

The app automatically saves its state to:

- `%APPDATA%\eOverlay\settings.json`

Saved items include:

- control-panel position;
- overlay position and size;
- overlay mode and opacity;
- text content, alignment, and font size;
- web URL and zoom;
- capture exclusion, topmost, click-through, and interactive input flags;
- custom global hotkeys.

Default hotkeys:

- `Ctrl + Alt + O`: toggle overlay;
- `Ctrl + Alt + P`: toggle control panel;
- `Ctrl + Alt + C`: toggle click-through.

You can change or clear these hotkeys in the `Global hotkeys` section of the control panel.

## Tray Behavior

After launch, the app lives in the system tray.

Tray actions:

- double-click tray icon: open the control panel;
- `Open control panel` / `Hide control panel`;
- `Show overlay` / `Hide overlay`;
- `Exit`.

## Local Run

Use one of these:

- `Run-eOverlay.vbs`
- `Run-eOverlay.cmd`
- `eOverlay.ps1`

`Run-eOverlay.cmd` forwards to the VBS launcher, so it does not leave a PowerShell console window open.

## Limits

- built for Windows 10/11 with DWM enabled;
- some protected-content apps may refuse DWM thumbnails;
- minimized source windows may stop updating;
- capture exclusion depends on whether the target app honors supported Windows APIs;
- interactive input is synthetic message forwarding, so some apps and games may ignore part of it;
- embedded web mode depends on WebView2 Runtime availability.
