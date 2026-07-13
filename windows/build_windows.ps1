$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Resolve-Path (Join-Path $ScriptRoot "..")
$IconPath = Join-Path $ScriptRoot "dist-assets\Tabledown.ico"
$SourceIcon = Join-Path $ProjectRoot "assets\generated\tablemark_app_1024.png"
$EntryPoint = Join-Path $ScriptRoot "run_windows.py"
$AssetData = "$SourceIcon;assets\generated"

# Build with the project venv's interpreter, never a bare `python` on PATH. A
# system Python without PyInstaller/Pillow does not fail the build — it skips
# the (failed) PyInstaller step and lets makeappx pack a STALE onedir, so the
# .msix silently ships old code (the 0.2.6 build hit exactly this). Fail loud
# instead.
$Python = Join-Path $ScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
  throw "venv python not found at $Python — create it: python -m venv .venv; .venv\Scripts\pip install -r requirements.txt"
}

& $Python (Join-Path $ScriptRoot "tools\make_icon.py") --source $SourceIcon --output $IconPath

Push-Location $ScriptRoot
try {
  # --collect-all winsdk: the StartupTask API (Open-at-Login) is reached
  # through winsdk, whose namespace modules load lazily and whose code lives in
  # a native _winrt.pyd — PyInstaller's static analysis misses both, so collect
  # the whole package (submodules + binaries) or the toggle silently vanishes
  # from the packaged build. The Excel formula action also imports COM lazily
  # on its worker thread; list pythoncom/win32com.client explicitly so that
  # deferred path is present in the frozen app.
  & $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "Tabledown-Windows" `
    --paths $ProjectRoot `
    --paths $ScriptRoot `
    --add-data $AssetData `
    --collect-all winsdk `
    --hidden-import pythoncom `
    --hidden-import win32com.client `
    --icon $IconPath `
    $EntryPoint
}
finally {
  Pop-Location
}
