# Tabledown Windows

Windows용 Tabledown 이식 버전입니다. 기존 macOS 앱 파일은 그대로 두고, 이 폴더 안에 Windows tray(시스템 트레이) 앱과 build(빌드) 설정만 분리했습니다.

## 개발 실행

PowerShell에서 실행합니다.

```powershell
cd windows
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run_windows.py
```

실행 후 Windows notification area(알림 영역)에 Tabledown 아이콘이 표시됩니다.

## 사용법

1. Excel 또는 Google Sheets에서 표를 복사합니다. (`Ctrl+C`)
2. Markdown editor(마크다운 에디터)에 붙여넣습니다. (`Ctrl+V`)
3. Excel HTML clipboard format(클립보드 형식)이 Markdown table(마크다운 표) text로 보강됩니다.

역방향도 지원합니다.

1. Markdown table(마크다운 표)을 복사합니다.
2. Excel에서 붙여넣습니다.
3. HTML table clipboard format(클립보드 형식)을 Excel이 읽어 셀 단위로 붙여넣습니다.

### 표의 수식을 포함해 XML로 복사

1. Excel desktop app에서 수식이 포함된 단일 사각형 영역을 선택합니다. (`Ctrl+C` 불필요)
2. 트레이 메뉴의 **‘표의 수식을 포함해 XML로 복사’** 를 누르거나 전역 단축키 **`Ctrl+Alt+E`** 를 누릅니다.
3. LLM/문서에 붙여넣으면 모든 셀의 주소·값·빈칸과 수식 셀의 현재 결과·A1·R1C1 수식이 표 구조 그대로 XML text로 들어갑니다. 같은 통합문서의 same-sheet/cross-sheet static A1 reference(같은 시트/다른 시트 정적 A1 참조)는 `<참조범위>` 아래 현재 값도 함께 연결됩니다.

Tabledown은 수식을 실행하지 않고 Excel의 현재 값을 읽습니다. 빈 셀도 XML에 남아 표 모양을 보존합니다. `INDIRECT`·`OFFSET`·defined name(정의된 이름)·structured reference(구조화 참조)·3-D·외부 통합문서·한도 초과 참조는 추측하지 않고 `참조상태="일부"`로 표시합니다. multi-area selection(다중 영역 선택)은 지원하지 않고 한 번에 최대 10,000셀, 직접 참조 최대 256범위·총 10,000셀·범위당 2,048셀, 셀 값 합계 5,000,000자, A1·R1C1 수식 합계 1,000,000자, 최종 XML 10MB까지 처리합니다. 잘못된 Excel instance(인스턴스)를 읽지 않도록 화면에 보이는 Excel process(프로세스)가 하나일 때만 동작합니다. Excel desktop app 전용이며 Google Sheets·LibreOffice는 지원하지 않습니다.

## 트레이 메뉴 / 기능

알림 영역(트레이) 아이콘을 클릭하면 메뉴가 열립니다.

- **Tabledown 사용** (자동 변환 토글) — 클립보드 감시를 켜고 끕니다. 전역 핫키 `Ctrl+Alt+T`
  (Accessibility 권한 불필요, user32 `RegisterHotKey`)로도 토글되며, 꺼지면
  트레이 아이콘에 빨간 사선이 표시됩니다. 토글은 "일시정지"용이라 매 실행 시
  켜진 상태로 시작합니다(영속 안 함).
- **표의 수식을 포함해 XML로 복사** — 현재 Excel 선택 영역의 값·빈칸·주소·A1/R1C1 수식과 같은 통합문서의 직접 A1 참조값을 표 구조 그대로 XML text로 복사합니다. 전역 단축키는 `Ctrl+Alt+E`입니다.
- **빈칸을 자동 채우기** (0.2.7) — 병합 셀이 남긴 헤더 빈칸을 마크다운 변환 시
  forward-fill 합니다(헤더 프레임만, 값 영역은 보존). 기본 꺼짐이며 켜면 설정에
  영속됩니다. macOS 0.5.0 의 `fill_blanks` 포팅.
- **언어** — 표시 언어를 고릅니다(설정에 영속).
- **로그인 시 자동 실행** (0.2.5) — Windows 로그인 시 앱을 자동 시작합니다.
  **MSIX 로 설치했을 때만** 메뉴에 표시됩니다(WinRT StartupTask 는 패키지 identity
  가 필요 — 소스/포터블 실행에는 항목이 숨겨짐).
- **문제 신고용 로그 열기** — 개인정보를 지운(scrubbed) 로컬 로그를 만들어 탐색기로
  폴더를 엽니다. **외부로 전송되지 않습니다**(네트워크·텔레메트리 없음).
- **도움말 / 종료**

또한 트레이 앱은 **단일 인스턴스 가드**(0.2.6, named mutex)로 두 번째 실행을
거부합니다 — 클립보드 워처가 둘이면 서로 변환을 덮어쓰기 때문입니다(macOS 는
LaunchServices 가 기본 제공).

## EXE 빌드 (포터블 onedir)

```powershell
cd windows
.\build_windows.ps1
```

PyInstaller **onedir**(`--onefile` 아님) 빌드라 산출물은 단일 exe 가 아니라
폴더입니다: `windows\dist\Tabledown-Windows\Tabledown-Windows.exe` (폴더째로
배포·실행). 이 스크립트는 위 "개발 실행"에서 만든 `windows\.venv` 인터프리터를
사용합니다(없으면 오류로 멈춤).

## MSIX 패키징 / Microsoft Store

설치형 MSIX 패키지(로그인 자동 실행 토글 포함, Store 제출용)는 별도 스크립트로
만듭니다.

```powershell
cd windows
.\build_msix.ps1 -SelfSign     # 로컬 sideload 테스트용 자체 서명
```

전체 절차(자체 서명 설치, 인증서 신뢰, Store 제출)는
[`PACKAGING.md`](PACKAGING.md) 참고. PyInstaller 는 크로스컴파일이 안 되므로
macOS 개발자는 GitHub Actions 워크플로(`.github/workflows/windows-build.yml`,
수동 트리거)로 테스트 빌드를 받습니다.

## 파일 구조

```text
windows/
├── run_windows.py
├── requirements.txt
├── build_windows.ps1          # 포터블 onedir EXE 빌드(PyInstaller)
├── build_msix.ps1             # MSIX 패키징(+ -SelfSign 로컬 테스트 서명)
├── PACKAGING.md               # MSIX / Microsoft Store 출시 가이드
├── PRIVACY.md                 # 개인정보 처리방침(영/한)
├── STORE_LISTING.md           # Microsoft Store 등록 문구(설명·카테고리·검색 키워드)
├── packaging/                 # MSIX 매니페스트 템플릿(build_msix.ps1 이 토큰 치환)
│   ├── AppxManifest.xml       # full-trust MSIX 매니페스트(StartupTask·runFullTrust)
│   └── Assets/                # Store 타일/로고 PNG
├── tabledown_windows/
│   ├── app.py                 # Windows tray(시스템 트레이) 앱 · 메뉴 · 클립보드 워처
│   ├── conversion.py          # 플랫폼 독립 변환 흐름
│   ├── diagnostics.py         # 로컬 진단 로그(scrubbed) 내보내기 + 크래시 훅(외부 전송 없음)
│   ├── hotkey.py              # 전역 핫키 Ctrl+Alt+T/E (user32 RegisterHotKey)
│   ├── html_clipboard.py      # Windows CF_HTML format(HTML 클립보드 형식)
│   ├── i18n.py                # Windows locale(로캘) + 언어 저장
│   ├── logger.py              # Windows 로그 경로
│   ├── settings.py            # 사용자 설정 저장(%APPDATA% JSON — 언어·첫 실행·빈칸 채우기)
│   ├── single_instance.py     # 단일 인스턴스 가드(named mutex — 트레이 중복 실행 방지)
│   ├── startup_task.py        # 로그인 시 자동 실행(WinRT StartupTask, MSIX)
│   └── win_clipboard.py       # pywin32 clipboard(클립보드) 래퍼
├── tests/                     # 단위 테스트(test_windows_port.py)
└── tools/                     # 아이콘/타일 PNG 생성 스크립트
```

## 테스트

Windows 포트 단위 테스트(변환 · i18n · CF_HTML · 핫키 · 단일 인스턴스 등):

```powershell
cd windows\tests
..\.venv\Scripts\python -m unittest test_windows_port -v
```

트레이·StartupTask 테스트는 Windows 전용이라 다른 OS 에서는 자동 skip 됩니다.
위 커맨드는 CI(`.github/workflows/windows-build.yml`)가 쓰는 것과 동일하며,
**macOS 에서도 리포 루트의 `.venv` 로 실행됩니다**(변환·CF_HTML 로직 검증):

```bash
# macOS — 리포 루트에서
cd windows/tests && ../../.venv/bin/python -m unittest test_windows_port -v
```

## 제한사항

- 이 폴더의 clipboard(클립보드) 기능은 Windows에서만 실행됩니다.
- macOS에서 검증 가능한 부분은 conversion(변환), i18n(다국어), CF_HTML format(HTML 클립보드 형식) 단위 테스트로 확인합니다.
- Windows clipboard(클립보드)의 일부 handle(핸들) 기반 format(형식)은 pywin32로 안전하게 복사할 수 없어서 보존 대상에서 제외됩니다.
