<#
.SYNOPSIS
  Tabledown 을 full-trust(mediumIL) MSIX 패키지로 빌드한다.

.DESCRIPTION
  흐름:
    1. PyInstaller onedir 빌드(build_windows.ps1) — 생략하려면 -SkipPyInstaller
    2. MSIX 타일/로고 PNG 생성(tools\make_msix_assets.py)
    3. staging 폴더 구성: app\ + Assets\ + AppxManifest.xml(토큰 치환)
    4. makeappx pack 으로 .msix 생성
    5. (-SelfSign) 로컬 sideload 테스트용 자체 서명

  Store 제출용 .msix 는 서명하지 않아도 된다(Microsoft 가 인증 후 재서명).
  로컬에서 설치해 동작을 확인하려면 -SelfSign 으로 자체 서명한다.

.PARAMETER StoreName
  Identity/@Name. Partner Center 가 부여한 Package identity Name 으로 교체.
  (로컬 테스트 기본값: Tabledown.Dev)

.PARAMETER StorePublisher
  Identity/@Publisher. Partner Center 값(예: CN=ABCD...) 으로 교체.
  -SelfSign 시에는 자체 서명 인증서 Subject 와 정확히 일치해야 한다.

.PARAMETER StorePublisherDisplay
  Properties/PublisherDisplayName. Partner Center 의 표시 이름.

.EXAMPLE
  # 로컬 테스트(자체 서명 + 설치)
  .\build_msix.ps1 -SelfSign

.EXAMPLE
  # Store 제출용(Partner Center 값으로)
  .\build_msix.ps1 -StoreName "12345LIMOD.Tabledown" `
                   -StorePublisher "CN=AB12CD34-..." `
                   -StorePublisherDisplay "LIMOD"
#>
param(
  [string]$StoreName = "Tabledown.Dev",
  [string]$StorePublisher = "CN=Tabledown.Dev",
  [string]$StorePublisherDisplay = "Tabledown",
  [switch]$SelfSign,
  [switch]$SkipPyInstaller
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ScriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Resolve-Path (Join-Path $ScriptRoot "..")
$DistDir     = Join-Path $ScriptRoot "dist"
$Staging     = Join-Path $DistDir "msix-staging"
$AppDir      = Join-Path $Staging "app"
$AssetsDir   = Join-Path $Staging "Assets"
$PyInstallerOut = Join-Path $DistDir "Tabledown-Windows"
$ManifestSrc = Join-Path $ScriptRoot "packaging\AppxManifest.xml"
$MasterIcon  = Join-Path $ProjectRoot "assets\generated\tablemark_app_1024.png"

# --- 0. 버전 (tablemark.__version__ -> 4-part, revision=0) ---
Push-Location $ProjectRoot
try {
  $rawVersion = (python -c "import tablemark; print(tablemark.__version__)").Trim()
} finally {
  Pop-Location
}
$parts = $rawVersion.Split(".")
while ($parts.Count -lt 3) { $parts += "0" }
$Version = "{0}.{1}.{2}.0" -f $parts[0], $parts[1], $parts[2]
Write-Host "Version: $Version (from $rawVersion)" -ForegroundColor Cyan

# --- 1. PyInstaller onedir ---
if (-not $SkipPyInstaller) {
  Write-Host "==> PyInstaller build" -ForegroundColor Cyan
  & (Join-Path $ScriptRoot "build_windows.ps1")
}
if (-not (Test-Path $PyInstallerOut)) {
  throw "PyInstaller output not found: $PyInstallerOut (run without -SkipPyInstaller first)"
}

# --- 2. Assets ---
Write-Host "==> MSIX assets" -ForegroundColor Cyan
python (Join-Path $ScriptRoot "tools\make_msix_assets.py") --source $MasterIcon --output-dir $AssetsDir

# --- 3. Staging ---
Write-Host "==> Staging" -ForegroundColor Cyan
if (Test-Path $AppDir) { Remove-Item $AppDir -Recurse -Force }
New-Item -ItemType Directory -Path $AppDir -Force | Out-Null
Copy-Item (Join-Path $PyInstallerOut "*") $AppDir -Recurse -Force

# AppxManifest 토큰 치환
$manifest = Get-Content $ManifestSrc -Raw
$manifest = $manifest.Replace("__IDENTITY_NAME__", $StoreName)
$manifest = $manifest.Replace("__IDENTITY_PUBLISHER__", $StorePublisher)
$manifest = $manifest.Replace("__VERSION__", $Version)
$manifest = $manifest.Replace("__PUBLISHER_DISPLAY__", $StorePublisherDisplay)
$manifestOut = Join-Path $Staging "AppxManifest.xml"
Set-Content -Path $manifestOut -Value $manifest -Encoding UTF8
Write-Host "    Identity Name      : $StoreName"
Write-Host "    Identity Publisher : $StorePublisher"

# --- 4. makeappx pack ---
function Find-Sdk-Tool([string]$name) {
  $cmd = Get-Command $name -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  $base = "${env:ProgramFiles(x86)}\Windows Kits\10\bin"
  if (Test-Path $base) {
    $found = Get-ChildItem $base -Recurse -Filter $name -ErrorAction SilentlyContinue |
             Where-Object { $_.FullName -match "\\x64\\" } |
             Sort-Object FullName -Descending | Select-Object -First 1
    if ($found) { return $found.FullName }
  }
  throw "$name not found. Install the Windows 10/11 SDK (makeappx/signtool)."
}

$makeappx = Find-Sdk-Tool "makeappx.exe"
$msixPath = Join-Path $DistDir "Tabledown-$Version.msix"
Write-Host "==> makeappx pack" -ForegroundColor Cyan
& $makeappx pack /d $Staging /p $msixPath /o
if ($LASTEXITCODE -ne 0) { throw "makeappx failed ($LASTEXITCODE)" }
Write-Host "    -> $msixPath" -ForegroundColor Green

# --- 5. (옵션) 로컬 테스트용 자체 서명 ---
if ($SelfSign) {
  Write-Host "==> Self-sign (local sideload test only)" -ForegroundColor Cyan
  if ($StorePublisher -notlike "CN=*") {
    throw "Self-sign 시 -StorePublisher 는 'CN=...' 형식이어야 한다 (현재: $StorePublisher)"
  }
  $pfx = Join-Path $DistDir "Tabledown-dev.pfx"
  $pwd = ConvertTo-SecureString "tabledown" -AsPlainText -Force
  $cert = New-SelfSignedCertificate -Type Custom -Subject $StorePublisher `
            -KeyUsage DigitalSignature -FriendlyName "Tabledown Dev" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
  Export-PfxCertificate -Cert $cert -FilePath $pfx -Password $pwd | Out-Null

  $signtool = Find-Sdk-Tool "signtool.exe"
  & $signtool sign /fd SHA256 /a /f $pfx /p "tabledown" $msixPath
  if ($LASTEXITCODE -ne 0) { throw "signtool failed ($LASTEXITCODE)" }

  Write-Host ""
  Write-Host "로컬 설치 방법:" -ForegroundColor Yellow
  Write-Host "  1) 인증서를 신뢰: 관리자 PowerShell 에서" -ForegroundColor Yellow
  Write-Host "     Import-Certificate -FilePath (Get-Item '$pfx') -CertStoreLocation Cert:\LocalMachine\Root" -ForegroundColor Yellow
  Write-Host "     (또는 .pfx 의 공개 인증서를 신뢰된 루트에 추가)" -ForegroundColor Yellow
  Write-Host "  2) Add-AppxPackage '$msixPath'" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "DONE: $msixPath" -ForegroundColor Green
