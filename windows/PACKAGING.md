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

- Python 3.10+ , `pip install -r windows/requirements.txt`
- **Windows 10/11 SDK** (makeappx.exe / signtool.exe 포함) — Visual Studio Installer 또는
  https://developer.microsoft.com/windows/downloads/windows-sdk/

---

## 1. 로컬 빌드 & 동작 테스트 (계정 없이)

```powershell
cd windows
.\build_msix.ps1 -SelfSign
```

이 한 줄이: PyInstaller onedir 빌드 → 타일 PNG 생성 → staging 구성 →
`dist\Tabledown-<버전>.msix` 생성 → 로컬 테스트용 자체 서명까지 수행한다.

설치(자체 서명이라 인증서를 먼저 신뢰해야 함):

```powershell
# 관리자 PowerShell
Import-Certificate -FilePath (Get-Item dist\Tabledown-dev.pfx) -CertStoreLocation Cert:\LocalMachine\Root
Add-AppxPackage dist\Tabledown-0.2.4.0.msix
```

설치 후 확인할 것:
- 알림 영역(트레이)에 Tabledown 아이콘이 뜨는가
- Excel/Sheets 표 복사 → 마크다운 에디터 붙여넣기 → `| ... |` 변환되는가
- 마크다운 표 복사 → Excel 붙여넣기 → 셀 분리되는가
- 로그: `%LOCALAPPDATA%` 아래 Tabledown 로그 (win 포트 logger 경로)

제거: `Get-AppxPackage *Tabledown* | Remove-AppxPackage`

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
5. 제출 → 인증 **최대 3 영업일**. 통과 후 약 15분 내 노출.

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
