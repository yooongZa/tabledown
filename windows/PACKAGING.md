# Tabledown Windows — MSIX 패키징 & Microsoft Store 출시 가이드

이 문서는 Windows 포트를 **full-trust(mediumIL) MSIX** 로 패키징해 Microsoft Store 에
출시하는 전체 절차다. 비용은 **0원**(등록비 면제 + MSIX 는 Microsoft 가 자동 서명).

## 왜 full-trust(mediumIL) 인가

Tabledown 은 백그라운드에서 클립보드를 **상시 감시**하다가 표를 감지하면 클립보드를
다시 써넣는다. UWP/AppContainer 의 WinRT `Clipboard` API 는 "앱이 포커스를 가졌을
때만 접근 가능" 이라는 제약이 있어 **백그라운드 감시가 불가능**하다. 따라서
`runFullTrust` capability 를 선언한 full-trust 데스크톱 패키지로만 동작한다.
(근거: learn.microsoft.com `Clipboard.GetContent`, `desktop-to-uwp-behind-the-scenes`)

---

## 사전 준비 (Windows PC)

- **PowerShell 7 (`pwsh`)** — `powershell` 5.1 이 아니라 반드시 `pwsh`. 빌드 스크립트
  (`build_msix.ps1` / `build_windows.ps1`)는 BOM 없는 UTF-8 로 저장돼 있고 한글 주석과
  em-dash(—)를 포함한다. 5.1 은 이를 레거시 ANSI 코드페이지로 디코드해 파서가 깨지고,
  7 은 UTF-8 로 읽어 정상 동작한다. (CI 도 `pwsh -File` 로 실행 — `windows-build.yml` 참고.)
  https://github.com/PowerShell/PowerShell 또는 `winget install Microsoft.PowerShell`
- **빌드용 venv** — 두 빌드 스크립트는 시스템 `python` 이 아니라 `windows\.venv` 인터프리터를
  강제한다(없으면 `throw` 로 즉시 중단). 시스템 Python 에 Pillow/PyInstaller 가 없어도
  빌드가 "성공"한 척 STALE onedir 를 패킹하는 사고를 막기 위함이다. **Python 3.10+** 로
  먼저 생성·설치할 것(공용 변환 코드가 런타임에 PEP 604 `str | None` 애노테이션을 쓰므로
  3.10 미만은 안 됨 — `setup.py` 의 `python_requires=">=3.10"` 과 동일 요건):

  ```powershell
  cd windows
  py -3 -m venv .venv          # Python 3.10+ (py -3 는 설치된 최신 Python 3 선택)
  .\.venv\Scripts\python -m pip install -r requirements.txt
  ```

- **Windows 10/11 SDK** (makeappx.exe / signtool.exe 포함) — Visual Studio Installer 또는
  https://developer.microsoft.com/windows/downloads/windows-sdk/

---

## 1. 로컬 빌드 & 동작 테스트 (계정 없이)

```powershell
cd windows
pwsh -File .\build_msix.ps1 -SelfSign
```

이 한 줄이: PyInstaller onedir 빌드 → 타일 PNG 생성 → staging 구성 →
`dist\Tabledown-<버전>.msix` 생성 → 로컬 테스트용 자체 서명까지 수행한다.
(파일명의 `<버전>` 은 `tabledown_windows.__version__` 기반 4-part — 빌드 출력에 찍힌
실제 이름을 그대로 쓸 것.)

설치(자체 서명이라 인증서를 먼저 신뢰해야 함). `-SelfSign` 은 개인키가 든
**`dist\Tabledown-dev.pfx`** 만 만든다(공개 `.cer` 는 안 만듦). `.pfx` 는
`Import-Certificate`(공개 `.cer` 전용)로는 못 넣으므로 **`Import-PfxCertificate`** 로
신뢰된 루트에 등록한다:

```powershell
# 관리자 PowerShell — .pfx(개인키 포함)를 신뢰된 루트에 등록
$pw = ConvertTo-SecureString "tabledown" -AsPlainText -Force
Import-PfxCertificate -FilePath dist\Tabledown-dev.pfx -Password $pw -CertStoreLocation Cert:\LocalMachine\Root

# 빌드 출력에 찍힌 실제 파일명 사용 (예: Tabledown-<버전>.msix)
Add-AppxPackage dist\Tabledown-<버전>.msix
```

> 개인키를 루트에 두기 싫으면, `.pfx` 에서 공개 `.cer` 만 추출해
> `Import-Certificate` 로 신뢰해도 된다(CI `windows-build.yml` 이 쓰는 방식):
> ```powershell
> $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
> $c.Import("$PWD\dist\Tabledown-dev.pfx", "tabledown", 'DefaultKeySet')
> [IO.File]::WriteAllBytes("$PWD\dist\Tabledown-dev.cer", $c[0].Export('Cert'))
> Import-Certificate -FilePath dist\Tabledown-dev.cer -CertStoreLocation Cert:\LocalMachine\Root
> ```

설치 후 확인할 것:
- 알림 영역(트레이)에 Tabledown 아이콘이 뜨는가
- Excel/Sheets 표 복사 → 마크다운 에디터 붙여넣기 → `| ... |` 변환되는가
- 마크다운 표 복사 → Excel 붙여넣기 → 셀 분리되는가
- 로그: `%LOCALAPPDATA%` 아래 Tabledown 로그 (win 포트 logger 경로)

제거: `Get-AppxPackage *Tabledown* | Remove-AppxPackage`

> **macOS 개발자라면**: PyInstaller 는 크로스컴파일이 안 되므로 macOS 에서 이
> `.msix`(또는 포터블 exe)를 만들 수 없다. 대신 GitHub Actions 워크플로
> `.github/workflows/windows-build.yml` 을 **수동 트리거**(Actions 탭 → *Windows build
> (test installer)* → *Run workflow*, `workflow_dispatch`)해 `windows-latest` 러너에서
> 빌드하고, 산출물(자체 서명 `.msix` + 공개 `.cer` + 포터블 zip + `INSTALL.txt`)을
> artifact 로 내려받아 Windows PC 에서 테스트한다.

---

## 2. Microsoft Store 계정 등록 (0원)

1. **반드시** https://storedeveloper.microsoft.com 진입점으로 시작 (신규 무료 플로우).
   - Partner Center 등 다른 경로로 들어가면 옛 유료 화면이 보일 수 있음.
2. 개인(Individual) 계정 선택 — 1인/개인 프로젝트 권장 대상.
3. 신원 확인: 정부 발급 신분증 + 셀피. (등록비 $0)

> 무료 앱이면 payout/tax 프로필·원천징수(W-8BEN) 이슈 없음. 유료 전환 시에만 필요.

---

## 3. 앱 예약 → Identity 값 받기

Partner Center 에서 앱 이름(Tabledown) 예약 후 **Product identity** 에서 확인:

| 매니페스트 위치 | Partner Center 항목 | 예시 |
|---|---|---|
| `Identity/@Name` | Package/Identity/Name | `12345LIMOD.Tabledown` |
| `Identity/@Publisher` | Package/Identity/Publisher | `CN=AB12CD34-...` |
| `Properties/PublisherDisplayName` | Publisher display name | `LIMOD` |

---

## 4. Store 제출용 MSIX 빌드

```powershell
cd windows
.\build_msix.ps1 `
  -StoreName "12345LIMOD.Tabledown" `
  -StorePublisher "CN=AB12CD34-..." `
  -StorePublisherDisplay "LIMOD"
```

→ `dist\Tabledown-<버전>.msix` (서명 안 함 — Microsoft 가 인증 후 재서명).

---

## 5. 제출

Partner Center 에서 새 submission:

1. **Packages**: 위 `.msix` 업로드.
2. **Properties → restricted capabilities**: `runFullTrust` 사용 사유 기재. 예시 문구:
   > Tabledown is a clipboard utility that watches the system clipboard and rewrites
   > copied spreadsheet tables into Markdown (and Markdown tables back into a
   > spreadsheet-compatible format). Reading and writing the global clipboard from a
   > background tray process requires full-trust (Win32 clipboard APIs); the UWP
   > clipboard API only permits access while the app window has focus, which is
   > incompatible with passive background conversion. No data leaves the device.
3. **Privacy policy URL**: 필수(클립보드를 읽는 앱). README 의 개인정보 처리방침을
   GitHub Pages/README 링크로 제공. 핵심: "모든 변환은 로컬, 외부 전송 없음".
4. **Store listing**: 설명·스크린샷·검색 키워드(축 5 의 SEO 키워드 재활용:
   excel to markdown, csv to markdown, clipboard, menu bar/tray, obsidian).
5. **Age ratings(연령등급)**: IARC 설문 작성 → 등급 자동 생성(아래 5.1).
6. 제출 → 인증 **최대 3 영업일**. 통과 후 약 15분 내 노출.

### 5.1 연령등급(IARC age rating) — 발급 완료

제출 과정의 **IARC(International Age Rating Coalition) 설문**을 작성하면 연령등급이
자동 생성된다. Tabledown 은 콘텐츠 없는 생산성 유틸리티라 전 연령으로 발급됐고,
**2026-06-16 Microsoft storefront 에 live(발효)** 됐다.

| 항목 | 값 |
|---|---|
| Global Rating ID | `3370e471-8485-81a1-8c70-3f4acbb3cce6` |
| 제품 / 회사 | Tabledown / LIMOD |
| 등급일 / 스토어 | 2026-06-16 (화) / Microsoft (live) |

- **Global Rating ID 는 보관할 것** ⭐ — IARC 라이선스를 받은 **다른 스토어**(Steam·
  Google Play 등)에 올릴 때 온보딩에서 "Global Rating ID" 또는 "IARC Certificate ID"
  를 물으면 이 값을 그대로 입력 → **설문 재작성 없이 등급 이식**.
- **재설문 조건**: 앱을 크게 바꿔 IARC questionnaire(설문) 답이 달라질 정도면 설문을
  다시 작성하고 새 등급을 받아야 한다. 단순 업데이트·기능 추가로 답이 안 바뀌면 기존
  등급이 그대로 유효.
- 등급이 틀렸다고 보이면 출시 후 제공되는 **"request a rating check"** 링크로 재검토
  요청(시작까지 1~3 영업일, 추가 자료 요청 가능).

---

## 6. winget (보너스, 0원)

Store 출시 후 `msstore` 소스를 통해 거의 자동으로 노출된다:

```powershell
winget install --source msstore Tabledown
```

`winget` 커뮤니티 매니페스트(`microsoft/winget-pkgs`)에 별도 PR 도 가능하나,
Store 출시만으로 `msstore` 소스 검색이 되므로 우선순위 낮음.

---

## 흔한 함정 체크리스트

- [ ] **AppContainer 금지** — 반드시 `runFullTrust` + `EntryPoint="Windows.FullTrustApplication"`.
- [ ] 설정/로그는 패키지 내부가 아니라 `%LOCALAPPDATA%` 에 기록(패키지 폴더는 읽기 전용).
- [ ] `Identity/@Version` 4-part, revision(4번째)=0 (`build_msix.ps1` 이 자동 처리).
- [ ] PyInstaller exe 가 Defender 에 오탐될 수 있음 → Store 재서명으로 대부분 해소.
- [ ] 인증 노트(Notes for certification)에 클립보드 권한 사유·테스트 방법 기재.
- [ ] GitHub Releases 직접 다운로드(.msix/.exe)를 병행하면 **Store 와 동일 최신 버전** 유지.
- [x] 연령등급(IARC) 발급 — Global Rating ID `3370e471-8485-81a1-8c70-3f4acbb3cce6` 보관(§5.1).
