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
  python -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --name "Tabledown-Windows" `
    --paths $ProjectRoot `
    --paths $ScriptRoot `
    --add-data $AssetData `
    --icon $IconPath `
    $EntryPoint
}
finally {
  Pop-Location
}
