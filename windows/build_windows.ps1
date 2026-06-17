$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Resolve-Path (Join-Path $ScriptRoot "..")
$IconPath = Join-Path $ScriptRoot "dist-assets\Tabledown.ico"
$SourceIcon = Join-Path $ProjectRoot "assets\generated\tablemark_app_1024.png"
$EntryPoint = Join-Path $ScriptRoot "run_windows.py"
$AssetData = "$SourceIcon;assets\generated"

python (Join-Path $ScriptRoot "tools\make_icon.py") --source $SourceIcon --output $IconPath

Push-Location $ScriptRoot
try {
  # --collect-all winsdk: the StartupTask API (Open-at-Login) is reached
  # through winsdk, whose namespace modules load lazily and whose code lives in
  # a native _winrt.pyd — PyInstaller's static analysis misses both, so collect
  # the whole package (submodules + binaries) or the toggle silently vanishes
  # from the packaged build.
  python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "Tabledown-Windows" `
    --paths $ProjectRoot `
    --paths $ScriptRoot `
    --add-data $AssetData `
    --collect-all winsdk `
    --icon $IconPath `
    $EntryPoint
}
finally {
  Pop-Location
}
