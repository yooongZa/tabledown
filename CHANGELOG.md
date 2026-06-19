# Changelog

이 프로젝트의 모든 주요 변경 사항은 이 파일에 기록됩니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르며,
버전 번호는 [Semantic Versioning](https://semver.org/lang/ko/) 을 사용합니다.

## [Unreleased]

### Removed
- **유료화 전면 제거 — 앱을 완전 무료로 전환.** StoreKit IAP 래퍼(`store.py`)를 삭제하고 XML 변환의 Pro 구독 게이팅을 해제 — 이제 누구나 ‘표를 XML로 복사’와 전역 단축키 ⌘⌃C 를 제한 없이 쓴다. 기부(Consumable 3종) 메뉴·‘구매 복원’ 항목·구독 시트·구매/복원 알럿을 모두 제거하고, `settings.py` 의 `pro_active`/`pro_expires` 캐시, `i18n` 의 결제·후원·Pro 문자열, `setup.py` 의 `StoreKit` include, `app.py` 의 관련 콜백·잠금(🔒) 표시도 함께 정리. 전역 단축키(`hotkey.py`)와 XML 변환 로직 자체는 그대로 — 게이팅만 사라졌다. 클립보드 변환 불변식·테스트 영향 없음(변환 테스트 38/38).

### Changed
- **XML 표 변환 형식을 중첩 계층(v2)으로 재설계.** 0.3.0 의 평면형(`<dataset><row><cell name="…">`, 가로 그룹만 `<group>` 중첩)을 대체. 이제 **가로·세로 양방향 다단 헤더를 모두 중첩**으로 보존한다: 루트 `<표>`, 세로 상위 그룹 `<직급그룹 이름="부장">`, 행 `<행 직책="대족장">`, 가로 상위 그룹 `<열그룹 이름="1분기">`, 리프 `<열 n="1">동</열>`. 세로 키는 v1 처럼 매 행에 반복(forward-fill)하지 않고 **부모 노드로 묶는다**(트레이드오프: 행이 자기완결 레코드가 아님 — 그룹 값이 부모에만 있음). 가로 차원 이름이 표에 없으면 지어내지 않고 일반 태그(`열그룹`/`열`)로 보존. 헤더가 XML 이름 규칙에 안 맞는 열은 세로 키로 승격하지 않아 정규화 없이 무손실 roundtrip 유지. `model_to_xml`·`table_xml_to_model`·`is_table_xml` 재작성, 변환 테스트 38/38.
- **메뉴 구조 정리 (UX P2)**: 메뉴 순서를 동작(토글·XML) → ‘설정 ▸’ → 도움말·종료로 정리(첫인상은 유틸리티). 변환 토글(`enabled`)은 **의도적으로 비영속**(매 실행 켜짐 — "꺼둔 걸 잊고 고장 오인" 방지) — 정책을 CLAUDE.md "설정 영속성 정책" 에 명문화.

### Added
- **로컬 진단 — 로컬 전용 크래시 캡처 + ‘문제 신고용 로그’ (외부 전송 0줄).** 직배포 환경에서 안 보이던 실패를 **로컬 로그**로 끌어온다. 외부로 보내는 코드는 한 줄도 없어 `windows/PRIVACY.md` 의 "no network connections, no telemetry … nothing is sent to any external server" 약속을 그대로 지킨다 (DiskOUT 식 *원격* 수집은 이 약속과 충돌해 의도적으로 채택하지 않음). 신규 `tablemark/diagnostics.py`·`windows/tabledown_windows/diagnostics.py`(동형): ① `install_crash_hooks()` — `sys.excepthook` + **`threading.excepthook`**(지금까지 무음으로 사라지던 **클립보드 워처 데몬 스레드의 미포착 예외**를 포착 — 가장 핫한 실패면) + `faulthandler`(네이티브 폴트 — `enable()` 가 덤프 후 기본 핸들러로 넘겨 OS 크래시리포터(.ips)도 그대로 동작), 전부 `Tabledown.log`/`Tabledown.crash` 에만 기록. **크래시 기록은 예외 타입+프레임 위치만 남기고 예외 메시지·페이로드는 기록하지 않는다**(공유 로그 안전성). ② 메뉴 ‘문제 신고용 로그 열기’(mac·Windows) → **스크럽된**(홈경로·유저명·볼륨/드라이브명·secret 제거) `Tabledown-diagnostics.txt` 생성 후 Finder/Explorer 로 열기 — 사용자가 직접 버그리포트에 첨부, 클립보드 미접촉(워처 레이스 없음). ③ **로그 누출 하드닝**: 클립보드 변환 catch 사이트(`copy as xml failed`/`clipboard update failed`)의 `{exc}` → `{type(exc).__name__}`, `converter/table_xml.py` 의 에러 메시지에서 클립보드 파생 텍스트(파서 detail·헤더 태그명) 제거 — 공유 로그에 표 데이터가 새지 않게. 회귀 테스트 추가: converter 누출 메시지 2 + diagnostics(scrub 3·threading.excepthook 1) + i18n 키 1.
- **자동 변환 피드백 — 무반응 제거 (P0 후속).** 그동안 클립보드 자동 변환(Excel↔Markdown)은 **아무 표시 없이** 백그라운드에서 일어나, 사용자가 "변환됐다"는 걸 알 수 없고 붙여넣기 결과가 예상과 다를 때 원인이 Tabledown 인지조차 몰랐다(UX 리뷰 최상위 이슈). 이제 워처가 실제로 표를 변환하면 **메뉴바 아이콘이 0.5초간 체크 표시로 깜빡**인다(수동 XML 의 1초 플래시보다 짧게 — 매 복사마다 떠도 거슬리지 않게). 워처는 백그라운드 스레드라 `AppHelper.callAfter` 로 메인 스레드에 넘겨 아이콘을 건드린다. 변환 결과엔 `mark_generated` 가 찍혀 다음 워처 틱이 no-op 이 되므로 플래시가 무한 반복되지 않는다. 팝업·소리·시스템 알림 없음("권한 0개" 유지).
- **전역 단축키 ⌘⌃C (`hotkey.py`)**: Carbon `RegisterEventHotKey` 로 클립보드의 표를 XML 로 바로 변환 — Accessibility·Input Monitoring 등 권한 불필요("권한 0개" 유지). 등록 실패해도 메뉴 항목은 그대로 동작(graceful). (유료 게이팅 없이 누구나 사용.)
- **UX 정리 1차 — 무반응 제거 (P0, `outputs/tabledown-uxui-plan.md`):**
  - **XML 변환 성공 플래시**: ‘표를 XML로 복사’(메뉴·⌘⌃C) 성공 시 메뉴바 아이콘이 1초간 체크 표시로 바뀜 — 전역 단축키의 "성공했는지 알 수 없음" 해소. 시스템 알림 대신 아이콘 플래시인 이유: 권한 프롬프트 없이("권한 0개" 유지). (에셋 `tablemark_menu_40_check.png`, 생성기 `scripts/make_menu_icons.py`)
- **UX 정리 2차 — 첫인상 (P1):**
  - **첫 실행 환영 안내 (macOS·Windows)**: 메뉴바/트레이 전용 앱이라 설치 후 "아무 일도 안 일어난 것처럼" 보이던 문제 — 1회만 아이콘 위치 + 사용법 알럿 표시(표시 전 플래그 마킹으로 반복 방지).
  - **도움말 개선**: 타이틀에 실행 중 버전 표기(버그 제보용 — 앱에서 버전 볼 곳이 없었음), ⌘⌃C 단축키·체크 플래시 설명 추가, "GitHub 열기" 버튼.
  - **Windows 토글 명확화**: "활성화 ✓/비활성화" 라벨 교체 방식(현재 상태인지 누를 동작인지 모호) → macOS 와 같은 고정 라벨 "Tabledown 사용" + 체크마크.
  - **Windows 끔 상태 아이콘**: 변환 꺼짐을 트레이 아이콘 자체에 표시(빨간 사선 + 펀치 — macOS off 관례 이식).
  - **Windows 트레이 아이콘 선명도**: 고정 64px 소스를 Windows 가 재축소해 흐릿하던 것 → `SM_CXSMICON`(DPI 반영) 정확한 크기로 직접 렌더.
  - Windows 설정 저장을 read-modify-write JSON(`tabledown_windows/settings.py`)으로 분리 — 기존 언어 저장기는 파일 전체를 덮어써 다른 키를 소실시킬 구조였음.
- **Windows Store(MSIX) 출시 문서**: `windows/PRIVACY.md`(개인정보 처리방침 — 호스팅용 EN/KO)와 `windows/STORE_LISTING.md`(Partner Center 제출 문구 — 짧은/긴 설명·검색 키워드·`runFullTrust` 사유·심사 노트·스크린샷 가이드) 추가. 패키징 절차 자체는 기존 `windows/PACKAGING.md`. (실제 Windows 에서 full-trust MSIX 빌드 — PyInstaller → makeappx → 자체서명 — 와 frozen 앱 트레이·클립보드 변환을 검증.)
- **Windows ‘로그인 시 자동 실행’ 토글 (0.2.5 — macOS 패리티)**: 트레이 메뉴에 ‘로그인 시 자동 실행’ 체크 항목 추가(`windows/tabledown_windows/startup_task.py`). MSIX 매니페스트에 이미 선언된 `windows.startupTask`(`TaskId="TabledownStartup"`, 기본 꺼짐)를 WinRT `StartupTask` API(`winsdk`)로 켜고 끈다. 상태는 OS 의 **설정 ▸ 앱 ▸ 시작 프로그램** 에 저장(별도 JSON 그림자 없음 — macOS `login_item` 의 SMAppService 와 동형). 모든 WinRT 호출은 모듈 내부의 짧은 전용 스레드에서 돌려(`asyncio.run`) 호출자(특히 pystray 펌프 스레드)의 COM apartment·메시지 루프와 분리. 패키지 ID 가 없는 소스/개발 실행과 비-MSIX exe 에서는 `is_supported()` 가 False → 메뉴에서 항목을 숨김(graceful fallback). 사용자가 작업 관리자에서 꺼둔 경우(`disabled_by_user`)·정책 제어(`disabled_by_policy`)는 read-back 상태로 감지해 체크를 켜지 않고 안내 알림을 띄움. `winsdk` 의존성 추가 + `build_windows.ps1` 에 `--collect-all winsdk`(lazy import·native `_winrt.pyd` 라 정적 분석으로 안 잡힘 — 빠지면 토글이 조용히 사라짐).

### Fixed
- **Windows: 트레이 앱이 여러 개 실행되던 문제 — 단일 인스턴스 가드 추가 (0.2.6)**: macOS 는 LaunchServices 가 `.app` 두 번째 실행을 막아 단일 인스턴스가 공짜지만, Windows 는 아무 것도 막지 않아 **수동 실행 + 로그인 자동 실행(StartupTask), 더블클릭 중복, 크래시가 남긴 유령 프로세스**가 각각 트레이 아이콘 + 클립보드 워처를 하나씩 더 띄웠다. 워처가 둘이면 클립보드를 서로 덮어써 변환 불변식(0~4)이 깨질 수 있다. 수정: `main()` 이 앱을 만들기 전 named mutex(`CreateMutexW("Local\\TabledownSingleInstance")`)로 가드 — 첫 프로세스가 만들고, 이후 프로세스는 `ERROR_ALREADY_EXISTS` 를 보고 조용히 종료(`windows/tabledown_windows/single_instance.py`). 세션 한정(`Local\`) 이라 빠른 사용자 전환 시 사용자별 1개 허용. kernel32 접근은 함수 안으로 미뤄 비-Windows 테스트 러너에서 import 가 안 깨지게 했고(거기선 가드 없이 실행 허용), 회귀 테스트로 first/second/no-kernel 3경로 + `main()` 단축 동작 검증.
- **Windows: Excel/Sheets 표가 마크다운으로 변환 안 되던 핵심 버그 수정 (0.2.5)**: Excel·Google Sheets 는 CF_HTML 의 `StartFragment`/`EndFragment` 마커를 **`<table>` 엘리먼트 *안쪽*** 에 둔다 — 그래서 `extract_cf_html` 이 뽑은 fragment 는 표 내부(`<col>`/`<tr>`/`<td>`)뿐이고 `<table>` 여는/닫는 태그가 빠진다. 표 감지(`has_html_table`)는 `<table>` 문자열을 보므로, **실제 Excel 표를 복사하면 "표 아님"으로 판정 → 변환 없이 원본 TSV 가 그대로** 남았다(`converted_clipboard` 가 None). 웹·채팅 표(fragment 에 `<table>` 포함)와 마크다운→Excel 방향은 영향이 없어 일부만 동작해 보였고, 기존 테스트는 `<table>` 을 포함한 *이상적* HTML 만 써서 잡지 못했다. 수정: `extract_cf_html` 이 행(`<tr>`)은 있는데 `<table>` 래퍼가 없으면 `<table>…</table>` 로 감싼다(비-표 HTML 은 건드리지 않음). 실제 Windows 클립보드를 떠서 원인 확인·수정·검증, 회귀 테스트는 Excel 형식 CF_HTML(마커가 `<table>` 뒤를 가리킴)을 재현.
- **Windows 트레이 앱 시작 크래시 수정**: 언어 서브메뉴를 `lambda _icon, _item, language=code` 로 만들던 코드가 기본 인자(`language=code`)까지 포함해 인자 3개가 되어, 0/1/2 인자만 허용하는 pystray `_assert_action` 에 걸려 **매 실행 시작 즉시 `ValueError` 로 크래시**하던 문제(`windows/tabledown_windows/app.py`). 인자 2개 클로저를 돌려주는 `_language_action(code)` 팩토리로 교체. 실제 Windows 에서 트레이 앱을 처음 구동하며 발견 — 기존 테스트는 크로스플랫폼 변환 계층만 검증해 잡지 못했음.
- **Windows 도움말 창이 안 닫히던 문제 수정 (0.2.5)**: 도움말/환영 알림이 쓰는 모달 `MessageBox` 가 ① pystray 펌프 스레드에서 동기 실행돼 트레이를 막고, ② owner=NULL·foreground 권한 없음이라 활성 창 뒤로 떠서 안 보이고, 그래서 사용자가 다시 클릭하면 박스가 쌓여 “안 꺼지는” 것처럼 보이던 문제(`windows/tabledown_windows/app.py`). 전용 데몬 스레드(`_show_message_box_async`) + 단일 인스턴스 락(중복 클릭이 한 창으로 합쳐짐) + `MB_SETFOREGROUND | MB_TOPMOST`(항상 앞으로) 로 해결.

## [0.3.0] - 2026-06-08

### Added
- XML 표 변환 추가 — LLM 프롬프트에 넣기 좋은 **레코드형 + 구조 태그 XML**. 루트 `<dataset>`, 데이터 행마다 `<row>`, 값마다 `<cell name="컬럼명">값</cell>`. 다단(그룹) 헤더는 `<group name="…">` 로 **중첩**해 표의 계층을 그대로 보존. 각 값이 컬럼명을 달고 있어 모델이 잘 인식하면서, 헤더의 공백·기호·숫자에 영향받지 않고 어떤 표준 XML 파서로도 안전.
  - **표 → XML (메뉴 ‘표를 XML로 복사’)**: 현재 clipboard 의 표(HTML 표·마크다운 표·XML)를 XML 로 변환. 사용자의 **명시적 클릭** 동작이므로 text 슬롯에 XML 을 넣고 HTML 은 제거 — 어디에 붙여도 의도한 XML 이 나옴. **자동 XML→표 역변환은 두지 않음**(워처가 일반·설정 XML 을 표로 오인해 클립보드를 건드릴 위험 회피).
  - **이름은 태그가 아니라 `name=` 속성에**: 헤더는 데이터라 공백·앞자리 숫자·기호(`( ) % /`)·중복·예약어 `xml` 등 XML 이름 규칙에 안 맞는 게 흔하다. 현실의 표 표준(OOXML SpreadsheetML·HTML `<td>`·ODF) 처럼 **고정 구조 태그 + 헤더는 속성**으로 두어 태그 변형·이름충돌을 원천 제거. 한국어/CJK 도 그대로.
  - **루트는 `<dataset>`(HTML 비충돌)**: `<table>`/`<tr>` 은 진짜 HTML 태그라, 이 XML 이 HTML 로 렌더링되는 곳(브라우저·Obsidian 미리보기·리치텍스트)에 가면 표 파싱이 걸려 셀이 표 밖으로 밀려나며 **비어 보이는** 문제가 있다. `dataset`/`row`/`cell`/`group` 은 HTML 에 없는 이름이라 트리가 보존됨.
  - **병합 셀 인식 (Excel→XML 전용 `html_table_to_model`)**: Excel 의 병합 표를 XML 로 바꿀 때 마크다운을 거치지 않고 병합을 직접 처리. ① 세로 병합(rowspan) 값을 아래 행에 채워(forward-fill) 모든 행이 완전한 레코드가 되게 하고, ② 전체 열을 차지하는 상단 제목 행은 건너뛰며, ③ 2단 그룹 헤더(colspan, `<th>` 없는 실제 Excel 포함)를 평면화하지 않고 **`<group>` 으로 중첩 보존**(가로 계층). 세로 병합은 계층이 아니므로 중첩하지 않고 값만 채움. (마크다운 변환 경로는 그대로 — 병합/계층은 XML 경로에서만.)
  - **‘XML: 빈칸을 자동 채우기’ 옵션 (설정 ▸, 기본 꺼짐, `NSUserDefaults` 영속)**: 병합을 안 하고 빈칸으로 그룹을 표현한 표(직급을 그룹 첫 행에만 쓰고 아래는 비움 — 현실에서 흔함)를 표 → XML 변환할 때, **왼쪽 키 열의 빈칸만** 바로 위(세로) → 좌측(가로) 값으로 채움(병합의 rowspan→위·colspan→좌측 origin 과 같은 원리). 데이터(값) 열의 빈칸은 “진짜 없음”일 수 있어 보존(pandas·Power Query 관례). 메뉴 항목에 설명 tooltip.
  - **설정 서브메뉴**: ‘빈칸 채우기’·언어·로그인 항목을 ‘설정 ▸’ 하위로 묶어 메인 메뉴 정리.

## [0.2.4] - 2026-05-29

### Changed
- 순수 Excel/Sheets 표를 마크다운으로 변환할 때도 HTML `<table>` 슬롯을 유지하도록 통일 — 0.2.3 까지는 이 경우(Excel 표 → 마크다운 에디터)만 HTML 을 제거했으나, 그 표를 다시 Excel·Word 에 붙이면 표 형식이 깨지는 손실이 있었음. 이제 모든 표 케이스(웹표·문서·Excel 표)에서 **HTML 유지 + text 슬롯 마크다운 보강**으로 동작해, 한 번 복사로 Excel·Word 는 표 형식을, 마크다운 에디터는 마크다운을 받음. PNG/PDF/RTF 등 rendered format(렌더링 형식) 만 제거. (트레이드오프: HTML 자동변환이 켜진 마크다운 에디터는 HTML 을 우선해 리치 표로 붙일 수 있음 — 도착지 앱 정책)

## [0.2.3] - 2026-05-29

### Changed
- 표가 일부 포함된 문서(웹·채팅·Word 등) 를 붙여넣을 때, 0.2.2 의 "원본 그대로 보존" 대신 **text(일반 텍스트) 슬롯의 표 부분만 마크다운 표로 보강**하고 HTML `<table>` 슬롯은 그대로 유지하도록 변경 (`convert_document_tables`). 마크다운 에디터(Obsidian 등) 는 text 를 받아 표가 마크다운 표로 들어가고, Word·Excel 등 리치 에디터는 HTML 슬롯의 원본 `<table>` 을 받아 표 형식으로 그대로 붙음 — 양쪽 도착지를 동시에 만족. 표 외 헤딩·문단·리스트는 plain text 로 유지(마크다운 `#`·`-` 문법은 추가하지 않음). PNG/PDF/RTF 등 rendered format(렌더링 형식) 만 제거하고 HTML 은 절대 제거하지 않음

## [0.2.2] - 2026-05-29

### Fixed
- 표가 일부만 포함된 문서(웹·채팅·Word 등) 를 복사·붙여넣을 때 표만 남고 나머지 텍스트(헤딩·문단·리스트) 가 사라지던 문제 수정 — clipboard HTML 에 `<table>` 외 의미있는 콘텐츠가 있으면 '문서' 로 보고 변환을 건너뛰어 원본을 그대로 보존. Excel/Sheets 의 순수 표(표 외 콘텐츠 없음) 만 마크다운으로 변환 (`html_has_content_outside_table`)

## [0.2.1] - 2026-05-29

### Fixed
- 웹·채팅(Claude 등) 에서 복사한 표를 Excel 에 붙여넣을 때 마크다운 원문이 한 셀에 박히던 회귀 수정 — 이런 표는 clipboard 에 마크다운 text 와 HTML `<table>` 이 함께 실려오는데, 0.2.0 에서 추가된 `is_markdown_table` 의 cell count(셀 개수) 일치 검사 때문에 칸수가 어긋나면 markdown 표로 인정되지 않아 HTML 이 제거되고 마크다운으로 변환되던 문제. HTML `<table>` 이 동반된 경우 cell count 검사를 건너뛰고(`strict` 파라미터) 원본 clipboard 를 보존하도록 수정. cell count 검사는 HTML `<table>` 이 없는 순수 텍스트의 false positive 차단 용도로 유지

## [0.2.0] - 2026-05-28

### Added
- 메뉴바 아이콘에 disabled state(상태) 시각화 — 변환 OFF 시 사선(slash) 오버레이된 아이콘으로 전환되어 메뉴를 열지 않아도 ON/OFF 를 식별 가능

### Changed
- 토글 메뉴 라벨을 macOS HIG(휴먼 인터페이스 가이드라인) 컨벤션에 맞게 단일화 — "활성화 ✓" / "비활성화" 두 라벨 대신 항상 "Tabledown 사용" (en: "Use Tabledown") 한 라벨을 쓰고 NSMenuItem `state` 의 체크마크로 ON/OFF 표현. 로그인 자동 실행 항목과 언어 선택 항목도 동일 방식으로 통일
- Excel/Sheets 셀 내부 줄바꿈(Alt+Enter) 을 공백 대신 `<br>` 로 보존 — Obsidian / GitHub Flavored Markdown 에서 셀 안 줄바꿈이 그대로 렌더링됨

### Fixed
- `is_markdown_table` heuristic(휴리스틱) 의 false positive(거짓양성) 감소 — 헤더와 separator(구분선) 의 cell count(셀 개수) 가 일치할 때만 markdown 표로 판정. 두 번째 줄이 우연히 `-` 로 시작하는 일반 텍스트(예: shell 출력) 가 표로 오인되어 변환되던 케이스 차단
- Mac App Store 빌드의 nested executable(중첩 실행파일) 서명 수정 — py2app 이 번들에 넣는 `Contents/MacOS/python` 에 `application-identifier` 대신 sandbox `inherit` entitlements 를 적용. 메인 실행파일만 full entitlements 를 유지해 App Store Connect error 90885(중첩 실행파일이 application-identifier 를 가졌으나 provisioning profile 이 없음) 해결
- 동일 marketing version 재업로드 시 build number 충돌(App Store Connect error -19232) 회피 — `CFBundleVersion` 을 `TABLEDOWN_BUILD` 환경변수로 분리해, marketing version(`CFBundleShortVersionString`, `0.2.0`) 은 유지하면서 build number 만 증가 가능
- `CFBundleVersion` 형식 오류(App Store Connect error 236550) 수정 — build number 는 period(점)로 구분된 정수 최대 3개만 허용되므로 `0.2.0.1`(4개) 대신 `0.2.1`(3개) 형식 사용

## [0.1.1] - 2026-05-19

### Added
- macOS 로그인 시 자동 실행 toggle(토글)
- 메뉴 UI 한국어 / 영어 localization(로컬라이제이션)
- Mac App Store 배포용 packaging(패키징) script(스크립트) 와 privacy policy(개인정보 처리방침)
- English README
- 앱 시작 시 `NSUserDefaults` 의 stale `NSStatusItem Visible*` 키 정리 — 이전 빌드에서 아이콘 숨김 상태가 `NSUserDefaults` 에 저장되어 메뉴바에 돌아오지 않는 사용자 자동 복구
- App Store screenshot resize 유틸리티 `scripts/resize_screenshots.py` — raw screenshot 을 Mac App Store 표준 2880×1800 canvas 에 padding 처리로 맞춤

### Changed
- Mac App Store screenshot 을 실제 앱 UI capture(캡처) 로 교체 — 1차 거절(Guideline 2.3.3) 에 대응해 marketing/promotional 머티리얼을 제거하고 메뉴 드롭다운·언어 서브메뉴·변환 결과(Numbers + TextEdit) capture 로 교체

### Removed
- 메뉴바 아이콘 숨기기 메뉴 항목 — `LSUIElement` 앱에서 숨기면 종료 UI 자체가 사라지고 NSStatusItem autosave 가 visibility 를 영구 저장해 재실행으로도 복구 불가능했음

### Fixed
- Mac App Store package(패키지) validation(검증) 오류

## [0.1.0] - 2026-05-11

최초 공개 release(릴리스). Developer ID signing(개발자 ID 서명) 과 Apple notarization(애플 공증) 을 통과한 DMG 배포 시작.

### Added
- macOS menu bar 앱으로 동작하는 Tabledown 초기 버전
- Clipboard(클립보드) 감시를 통한 Excel / Google Sheets ↔ Markdown 표 자동 변환
  - Excel 표 복사 → Markdown source(마크다운 원문) 로 paste(붙여넣기)
  - Markdown 표 복사 → Excel 셀 단위로 paste
- Release hardening(릴리스 안정화): 서명·공증 빌드 파이프라인 구성
- Clipboard history(클립보드 히스토리) 앱 호환성 문서화
- Obsidian paste troubleshooting(붙여넣기 문제 해결) 가이드
- Markdown source paste 동작 문서

### Fixed
- Obsidian 에서 표 paste 시 발생하던 공백 처리 문제
- Excel 로 Markdown paste 시 HTML clipboard format(클립보드 형식) 충돌 — HTML format 을 제거하여 plain text 만 사용

[Unreleased]: https://github.com/yooongZa/tabledown/compare/v0.2.4...HEAD
[0.2.4]: https://github.com/yooongZa/tabledown/compare/v0.2.3...v0.2.4
[0.2.3]: https://github.com/yooongZa/tabledown/compare/v0.2.2...v0.2.3
[0.2.2]: https://github.com/yooongZa/tabledown/compare/v0.2.1...v0.2.2
[0.2.1]: https://github.com/yooongZa/tabledown/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/yooongZa/tabledown/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/yooongZa/tabledown/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yooongZa/tabledown/releases/tag/v0.1.0
