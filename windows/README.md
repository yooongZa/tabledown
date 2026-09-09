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

### 셀 값·수식·참조를 XML로 복사

1. Excel desktop app에서 수식이 포함된 단일 사각형 영역을 선택합니다. (`Ctrl+C` 불필요)
2. 트레이 메뉴의 **‘셀 값·수식·참조를 XML로 복사’** 를 누르거나 전역 단축키 **`Ctrl+Alt+E`** 를 누릅니다.
3. LLM/문서에 붙여넣으면 모든 셀의 주소·값·빈칸·`값종류`와 수식 셀의 현재 결과·A1·R1C1 수식이 표 구조 그대로 XML text로 들어갑니다. 같은 통합문서의 same-sheet/cross-sheet static A1 reference(같은 시트/다른 시트 정적 A1 참조)는 `<참조범위>` 아래 현재 값과 타입도 함께 연결됩니다. 지원하는 참조를 모두 안정적으로 읽은 경우에만 `값대입수식`에 참조를 현재 값으로 바꾼 설명용 표현도 추가됩니다. 숫자는 `3.15E+5` 대신 `315000` 같은 일반 십진수로 기록하며 숫자처럼 보이는 text는 그대로 둡니다.

Tabledown은 수식을 실행하거나 재계산하지 않고 Excel의 현재 raw value(원시값)를 읽습니다. `계산모드`·`계산상태`(`done`·`calculating`·`pending`·`unknown`)와 `계산결과상태`가 함께 기록됩니다. `done`은 `snapshot_stable`, `calculating`·`pending`은 `calculation_incomplete`, 알 수 없거나 읽지 못한 상태는 `freshness_unverified`로 표시합니다. 두 번 같은 snapshot(스냅샷)이더라도 manual calculation(수동 계산) 통합문서가 최신 재계산됐다고 보장하지는 않습니다. 화면 표시값·숫자 서식·병합 계층은 이 명령에 포함하지 않고 XML에 `표시정보상태="미포함"`·`병합정보상태="미포함"`으로 명시합니다.

`값대입수식`도 optional(선택적) display-only(표시 전용) 표현이며 `값대입수식동등성="보장안함"`으로 표시됩니다. 참조 셀 자체가 수식이어도 현재 결과만 한 단계 대입하고 recursive reference(재귀 참조)는 펼치지 않습니다. 실제 빈 셀은 그 표현에서 `BLANK()`로 구분하고, 빈 셀 자체도 XML에 남아 표 모양을 보존합니다. `INDIRECT`·`OFFSET`·defined name(정의된 이름)·structured reference(구조화 참조)·3-D·외부 통합문서·읽기 실패·한도 초과 참조는 추측하지 않고 `참조상태="일부"`·`참조포함범위수`·`참조누락이유`를 표시하며 `값대입수식`을 만들지 않습니다. 누락 code(코드)는 `dynamic_reference`·`calculated_range_reference`·`external_or_structured_reference`·`three_dimensional_reference`·`whole_row_or_column_reference`·`defined_name_or_unsupported_syntax`·`invalid_a1_reference`·`range_count_limit`·`cell_count_limit`·`range_size_limit`·`read_failed`·`value_size_limit`·`unspecified`입니다. 다만 함수 호출 모양의 workbook-defined LAMBDA/UDF(통합문서 정의 LAMBDA/사용자 정의 함수)는 새 Excel 내장 함수와 text만으로 안전하게 구별할 수 없어, 식에 직접 적힌 A1 참조만 값으로 바꾸며 함수 내부의 숨은 참조는 펼치지 않습니다. multi-area selection(다중 영역 선택)은 지원하지 않고 한 번에 최대 10,000셀, 직접 참조 최대 256개 **고유** 범위·총 10,000셀·범위당 2,048셀, 셀 값 합계 5,000,000자, 원본 A1+R1C1 수식 합계 1,000,000자, 파생 값대입수식 별도 합계 1,000,000자, 최종 XML 10MB까지 처리합니다. 같은 참조를 여러 수식이 공유하면 Excel에서는 한 번만 읽고 각 수식 셀에는 수식에 적힌 순서로 연결합니다. 파생식 글자 한도를 넘긴 셀은 `값대입수식상태="omitted_character_limit"`을 기록하고, 최종 XML 한도 때문에 모든 파생식을 뺀 경우 루트에 `값대입수식상태="omitted_xml_size_limit"`을 기록합니다. 두 경우 모두 원본 export는 유지합니다. 그래도 크면 값 타입·계산 상태·누락 사유를 조용히 제거하지 않고 내보내기를 중단합니다. 잘못된 Excel instance(인스턴스)를 읽지 않도록 화면에 보이는 Excel process(프로세스)가 하나일 때만 동작합니다. Excel desktop app 전용이며 Google Sheets·LibreOffice는 지원하지 않습니다.

내보내기는 background worker(백그라운드 작업 스레드)에서 실행됩니다. 이미 실행 중일 때 메뉴나 단축키를 다시 누르면 진행 중임을 안내하고, 준비 중 다른 내용을 복사하면 새 clipboard를 덮어쓰지 않고 XML 쓰기를 취소합니다. 실제 기록된 text와 marker(마커)를 다시 확인한 뒤에만 성공으로 안내합니다. 일부 참조값을 가져오지 못했거나 Excel이 계산 중인 경우에는 같은 완료창에서 짧게 알려줍니다. 원본 수식과 현재 결과는 유지하며, `A1:INDEX(...)`처럼 계산으로 정해지는 범위는 추측하지 않고 일부 참조로 표시합니다.

### AI용 간결 복사

수식이 포함된 영역을 선택한 뒤 **‘AI용 간결 복사’(Copy compact XML for AI)** 를 누릅니다. 제목과 항목명도 함께 선택하면 값의 의미를 전달하기 좋습니다. 기존 수식 복사와 같은 stable snapshot(안정된 스냅샷) 읽기 경로를 사용하며, 값·수식·참조값·계산 상태·누락 사유와 기존 크기 한도를 보존합니다.

- 선택 안의 단순한 열 제목·행 항목을 원본 주소에 연결하고 **‘추정’ 또는 ‘미확인’**으로 표시합니다. 복잡한 다단 헤더, 중복·빈 제목, 일반 텍스트 데이터는 미확인으로 남을 수 있으며, 맥락을 찾기 위해 주변 셀을 추가로 읽지 않습니다.
- 시트와 주소가 정확히 같은 참조범위만 `<참조목록>`에 한 번 기록하고 각 수식에서 `<참조 ref="…">`로 연결합니다. 일부만 겹치는 범위와 수식별 참조 순서는 유지합니다. 반복 참조가 많을수록 길이가 줄고 작은 표는 맥락 정보 때문에 더 길어질 수 있습니다.
- 기존 수식 XML 메뉴·`Ctrl+Alt+E`·자동 변환과 기본 설정은 유지됩니다. 새 설정이나 단축키는 없습니다. 결과는 로컬 클립보드에 복사하며 Tabledown이 AI로 전송하지 않습니다.

## 트레이 메뉴 / 기능

알림 영역(트레이) 아이콘을 클릭하면 메뉴가 열립니다.

- **Tabledown 사용** (자동 변환 토글) — 클립보드 감시를 켜고 끕니다. 전역 핫키 `Ctrl+Alt+T`
  (Accessibility 권한 불필요, user32 `RegisterHotKey`)로도 토글되며, 꺼지면
  트레이 아이콘에 빨간 사선이 표시됩니다. 토글은 "일시정지"용이라 매 실행 시
  켜진 상태로 시작합니다(영속 안 함).
- **셀 값·수식·참조를 XML로 복사** — 현재 Excel 선택 영역의 값 타입·원시값·계산 상태·주소·A1/R1C1 수식, 같은 통합문서의 직접 A1 참조값과 설명용 `값대입수식`을 cell-grid XML(셀 그리드 XML)로 복사합니다. 병합 계층·표시 서식은 포함하지 않으며 그 사실도 XML에 표시합니다. 전역 단축키는 `Ctrl+Alt+E`입니다.
- **AI용 간결 복사** — 기존 수식 정보에 선택 안의 추정 맥락을 보강하고 공통 참조값의 반복을 줄입니다. 메뉴에서만 실행합니다.
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
