# Changelog

이 프로젝트의 모든 주요 변경 사항은 이 파일에 기록됩니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르며,
버전 번호는 [Semantic Versioning](https://semver.org/lang/ko/) 을 사용합니다.

## [Unreleased]

### Added
- **표의 값과 수식을 함께 LLM 친화 XML로 복사하는 수동 메뉴·전역 단축키 추가 (macOS·Windows).** Excel desktop app에서 단일 사각형 영역을 선택한 뒤 ‘표의 수식을 포함해 XML로 복사’를 실행하거나 macOS **⌘⌃E** / Windows **Ctrl+Alt+E** 를 누르면, 모든 셀의 값·빈칸·주소와 수식 셀의 현재 결과·A1/Formula2·R1C1 표현을 행·열 구조 그대로 XML text에 기록한다. Tabledown은 수식을 실행하지 않고 Excel의 현재 값을 읽으며, 자동 clipboard watcher의 기존 Excel↔Markdown 동작은 변경하지 않는다. macOS는 Excel Apple Events Automation, Windows는 실행 중인 Excel COM을 사용하고, 결과 clipboard는 명시적 text-only export로 교체한다.
- **수식 XML에 실제 계산 입력값을 direct reference(직접 참조)로 연결 (macOS·Windows).** `=C6*D6` 같은 같은 시트 참조뿐 아니라 `='단가 표'!B2*C6` 같은 같은 통합문서의 다른 시트 static A1 reference(정적 A1 참조)를 찾아, 각 수식 셀 아래 `<참조범위 시트="…" 주소="…"><참조셀 주소="…" 값="…" /></참조범위>`로 현재 값을 함께 기록한다. 참조 범위 값도 선택값·수식과 같은 immutable snapshot(불변 스냅샷)에 포함되어 두 번 일치할 때만 성공하며, 연결값만 바뀌어도 기존 최대 3회 bounded retry(제한 재시도)가 작동한다. `INDIRECT`·`OFFSET`·defined name(정의된 이름)·structured reference(구조화 참조)·3-D·외부 통합문서·한도 초과 참조는 추측하지 않고 기존 수식과 결과를 보존하면서 해당 셀에 `참조상태="일부"`를 표시한다. Excel의 다른 시트를 못 따라가는 `DirectPrecedents`와 사용자 선택을 바꾸는 `NavigateArrow`는 사용하지 않고, Formula2 A1 text를 보수적으로 해석해 UI 상태를 유지한다. 참조는 수식당 명시된 범위를 유지하면서 최대 256범위·총 10,000셀·범위당 2,048셀로 제한하고, 기존 선택 10,000셀·값 5,000,000자·XML 10MB·text-only clipboard·자동 변환 불변식은 유지한다.

### Fixed
- **일반 표→XML을 clipboard 사전 복사 방식에서 Excel 직접 선택 방식으로 수정 (macOS).** 이제 Excel에서 단일 사각형 표 범위를 선택하고 ‘선택한 표를 XML로 복사’ 또는 **⌘⌃X** 를 실행하면 `Cmd+C` 없이 Excel range의 `string value`로 데이터 셀 text의 유의미한 공백과 숫자·날짜·퍼센트·통화·사용자 지정 표시값, 빈칸·오류값·병합 구조를 직접 읽는다. clipboard 표 fallback을 제거해 오래된 복사본을 변환하던 UX를 차단했다. 선택·값·병합 영역이 같은 immutable snapshot(불변 스냅샷) 두 개가 연속될 때만 내보내며(최대 3회), Excel 미실행·Automation 거부·다중 영역·**10,000셀 초과**·선택 경계를 걸친 병합·읽는 중 변경은 fail-closed(실패 시 차단)로 writer 호출 전 기존 clipboard를 보존한다. 혼합 범위에서 거짓을 돌려주는 Excel의 merge scalar(병합 스칼라) 대신 모든 셀의 merge area 주소를 batch(일괄)로 읽어, 정확한 병합 구조와 수식 XML 명령과 같은 10,000셀 한도를 유지한다. HTML 구조 추론에는 opaque token(불투명 토큰)을 써 데이터 셀의 앞뒤·연속 공백과 literal `<br>`를 보존하고, XML 1.0 금지문자는 clipboard 쓰기 전에 차단한다. 기존 자동 Excel/Sheets↔Markdown watcher는 변경하지 않았다.
- **수동 일반 표→XML이 Excel에 원본 표로 다시 붙던 문제 수정 (macOS).** HTML slot(슬롯)만 제거해도 Excel이 만든 `com.microsoft.Embed-Source`·`DataObject`·OLE native format(네이티브 형식)이 clipboard에 남아, Excel이 XML text보다 원본 표를 우선 선택했다. 이제 명시적 `copy_as_xml` 성공 경로가 기존 형식을 모두 비우고 plain text(일반 텍스트) XML+생성 marker만 기록한다. Excel 선택 읽기·검증 실패는 writer 호출 전에 clipboard를 보존하며, writer 실패도 성공 flash(플래시) 없이 내용 비노출 오류로 처리한다. 자동 Excel↔Markdown 변환의 HTML 유지 불변식은 변경하지 않았다.
- **Excel의 `##` 표시를 실제 값으로 내보내던 데이터 손실 차단 (macOS).** Excel `string value`는 날짜·통화·퍼센트·사용자 지정 숫자가 열 너비를 넘거나 날짜·시간을 표시할 수 없으면 원문 대신 `##`만 반환한다. 직접 선택 reader가 hash-only 표시를 발견하면 raw scalar(원시값)와 대조해 사용자가 입력한 literal `##`는 보존하고, 숫자/날짜 display overflow(표시 넘침)는 writer 호출 전에 `display_overflow`로 중단해 전체 값이 보이도록 수정할 것을 안내한다. hash 후보는 range-level batch 1회로 판별하고, Excel 오류 때문에 batch shape가 깨질 때만 최대 32셀의 bounded fallback을 써 대량 literal `##`에서도 cell-wise Apple Event 성능 회귀를 막는다. 실패 시 기존 clipboard와 성공 flash는 그대로 보존된다.
- **일반 표→XML 변환의 조용한 데이터 손실 방지.** 1열 표의 모든 행을 전체폭 제목으로 오인해 비우던 문제, 본문의 `colspan` 전체폭 행을 삭제하던 문제, 명시적으로 존재하는 오른쪽 끝 빈 열을 잘라내던 문제를 수정했다. 현 직접 선택 경로가 쓰는 구조 parser와 내부 호환 helper에서 다중·중첩 HTML 표를 첫 표만 조용히 내보내지 않으며, Markdown 호환 parser는 둘째 줄의 header separator(헤더 구분선)만 제거해 뒤쪽의 `-`·`--` 값 행을 데이터로 보존한다. 이 helper들은 clipboard fallback으로 노출하지 않는다. XML v2 계층·다단 헤더·선두 전체폭 제목 제거·빈칸 채우기 기본 OFF·자동 변환 clipboard 불변식은 유지한다.
- **수식 XML export의 일관된 snapshot(스냅샷) 보장 (macOS·Windows).** Excel에는 값과 A1/R1C1 수식을 한 transaction(트랜잭션)으로 읽는 API가 없으므로, 선택 범위 전체의 immutable snapshot(불변 스냅샷)을 두 번 연속 비교하고 불일치나 일시적 selection 변경 때 한 번만 재시도한다(최대 3회 읽기). 같은 주소에서 수식·계산 결과만 바뀌는 경우도 감지하며, 안정된 두 snapshot을 얻지 못하면 내용 없는 오류만 표시하고 clipboard는 쓰지 않는다. 권한·Excel 미실행·COM/Automation 실패는 재시도하지 않는다.

### Verification
- **2026-07-21 자동 검증 — 수식 direct reference(직접 참조) 값 보강:** 공용 serializer/A1 parser + macOS Excel formula reader/AppKit/AppleScript compile 55/55, 일반 Excel→XML macOS 회귀 19/19(합계 74/74), 기본 test matrix 82/82, Windows 포트 88건 중 67 pass + 21 skip(macOS에서 Windows tray·WinRT 제외), Python `compileall`, `git diff --check`를 통과했다. 같은 시트 `=C6*D6`, 다른 시트 `=C6*'단가 표'!D6`, 범위·문자열 속 가짜 주소·structured/external/dynamic/3-D reference 오인 차단, 미해결/한도 초과 partial marker(일부 표시), 참조값-only race 재시도, 참조 오류값 이름, XML escape/빈값과 기존 수식 export 출력을 검증했다. 실제 macOS Excel의 다른 시트 참조와 실제 Windows Excel/COM 수동 검증은 이번 자동 테스트 범위에서 실행하지 않았다.
- **2026-07-17 출시 후보 실기기 재검증:** 권한 허용 뒤 OS-level key event로 macOS 전역 단축키 **⌘⌃E**(수식 XML)·**⌘⌃X**(일반 XML)·**⌘⌃T**(자동 변환 토글)가 하나의 Carbon handler에서 각각 올바른 callback으로 분기하는 것을 확인했다. 실제 Excel `A3:H10` 64셀은 수식 27개·R1C1 27개·다른 시트 참조 6개를 text-only `<표범위>` XML로 보존했고, 일반 XML은 plain text+생성 marker만 남겼다. 실제 `Cmd+C` 자동 변환은 Markdown text를 보강하면서 Excel HTML·OLE 형식을 유지했다. 수식 없음·헤더만·10,001셀 초과 경로는 기존 clipboard를 보존했다. macOS 0.6.0 / build 0.6.3 App Store 후보 패키지(16MB)는 deep codesign·installer signature·sandbox/Excel Automation entitlement·bundle metadata 검증을 통과했다. Apple Distribution 서명본은 App Store 처리 전 로컬 실행할 수 없으므로 동일 build의 최종 runtime smoke test는 TestFlight 처리 후 수행한다.
- **2026-07-17 자동 검증:** macOS 전체 단위/AppKit/AppleScript compile 70/70, converter 51/51을 포함한 기본 test matrix 82/82, Windows 포트 85건 중 64 pass + 21 skip(macOS에서 Windows 전용 tray·WinRT 제외), Python `compileall`, `git diff --check`를 통과했다. 테스트 workbook도 export→re-import 후 8개 sheet·17개 사용자 시나리오·수식 보존·formula error 0건과 preview layout을 확인했다.
- **2026-07-17 실제 macOS Excel 사용자 흐름:** `Cmd+C` 없이 현재 선택을 일반 XML로 복사해 stale clipboard 대신 선택값이 들어가고, 성공 clipboard에는 plain text XML+marker만 남는 것을 확인했다. 기본·8개 병합/다단 헤더·빈칸 채우기 OFF/ON(값 빈칸 보존)·날짜/퍼센트/통화/사용자 지정 서식·앞뒤/연속 공백·literal `<br>`/실제 줄바꿈·`#DIV/0!`/`#N/A`/`#CALC!`/`#분산!` 오류·4,008셀 stable export·선택 변경 race(혼합 snapshot 없음)·10,521셀 fail-fast/clipboard 보존을 검증했다. 정확히 **10,000셀**은 두 snapshot+302,392-byte XML을 31.94초에 생성했고, 좁은 열의 `##`는 `display_overflow`로 차단했다. 별도 10,000셀 literal `##` scratch 표는 batch 경로로 두 snapshot을 45.77초에 읽고 모든 값을 보존했다. 실제 Windows 기기 검증은 이번 macOS 작업 범위에서 실행하지 않았다.
- **2026-07-16 자동 검증:** converter 49/49, macOS XML action(동작)+formula 단위 테스트 49/49, 기본 test matrix 80/80, Windows 포트 85건 중 64 pass + 21 skip(macOS에서 Windows 전용 tray·WinRT 테스트 제외), Python `compileall` 통과.
- **2026-07-16 실제 macOS Excel 사용자 흐름:** 원본 `A1:F5` 표 복사 시 HTML·OLE·DataObject가 공존하는 실제 clipboard를 재현했다. 일반 XML 변환 후에는 plain text+marker만 남았고, 새 Excel 문서 붙여넣기는 원본 표가 아닌 XML text `A1:A34`로 들어갔다. 자동 변환 직후 HTML slot이 유지되는 것도 함께 확인했다. 실제 Windows Excel/COM 수동 검증은 이번 환경에서 실행하지 못했다.
- **2026-07-16 수정 commit `ad9742f` 재검증:** macOS XML action+formula 49/49, 기본 test matrix 80/80, Windows 포트 64 pass + 21 skip, Python `compileall`을 다시 통과했다. 실제 macOS Excel에서도 native clipboard format(네이티브 클립보드 형식)이 text+marker로 교체되고, 새 문서에 원본 표 대신 XML `A1:A34`가 붙는 것을 두 번째로 확인했다. 테스트 중 코드 변경은 없었다.

### Distribution
- **2026-07-17: commit `2b19bf5`의 macOS 0.6.0을 TestFlight build 0.6.3으로 업로드.** 15:10에 Apple ID `6768205551`로 전송됐고 Transporter의 “앱 처리가 완료되었습니다” 상태를 통과했다. 동일 내부 tester 계정의 TestFlight 앱에도 **버전 0.6.0 / 빌드 0.6.3**과 업데이트 버튼이 노출되는 것을 확인했다. 기본 test matrix 82/82, macOS XML 집중 테스트 70/70, Windows 포트 64 pass + 21 skip, Python `compileall`, `git diff --check`를 통과했다. 16MB 패키지(SHA-256 `19ee511b9d1dd2c93195b2923289636d31d5f42dae2f801f6be8d15c02fc5e6f`)는 deep codesign·installer signature·sandbox/Excel Automation entitlement·nested Python sandbox inheritance·테스트 package 제외를 검증했다. `TABLEDOWN_BUILD=0.6.3`만 재정의했으며 macOS marketing version은 0.6.0으로 유지했다.
- **2026-07-16: macOS 0.6.0을 TestFlight build 0.6.2로 업로드.** Apple Transporter가 `6768205551` delivery를 16:53에 수신했으며 현재 App Store Connect processing(처리 중) 상태다. macOS XML action+formula 49/49와 기본 test matrix 80/80을 통과했고, 16MB App Store package의 deep codesign(심층 코드 서명)·installer signature(설치 프로그램 서명)·sandbox/Excel Automation entitlement·nested Python sandbox inheritance·테스트 package 제외·export compliance 선언을 검증했다. build number만 `TABLEDOWN_BUILD=0.6.2`로 재정의했으며 macOS marketing version 0.6.0과 Windows version은 변경하지 않았다. **이 build는 2026-07-17 일반 XML 직접 선택 수정 전 산출물이라 현재 build 0.6.3으로 대체됐다.**
- **2026-07-13: macOS 0.6.0을 TestFlight build 0.6.1로 업로드.** App Store Connect 처리와 Mac App Store signing(서명)·sandbox entitlement(샌드박스 권한) 검증을 통과했으며, 내부 테스트 그룹 `22`에 연결되어 ‘제출 준비 완료’ 상태다. 기능 단위 테스트 39건과 전체 test matrix(테스트 매트릭스) 75건을 통과했다.

## [0.5.0] - 2026-07-01

### Added
- **마크다운 자동 변환에도 ‘빈칸을 자동 채우기’ 적용 — 병합 헤더가 남긴 빈칸을 채운다.** 지금까지 Excel/Sheets 표를 마크다운으로 자동 변환할 때 병합 셀은 무조건 빈칸으로 남았다(세로 병합 `부장` 아래는 공백, 가로 병합 그룹 헤더 `1분기`는 한 칸에만). 이제 설정 ‘빈칸을 자동 채우기’(기본 꺼짐)가 켜져 있으면 마크다운 경로도 **헤더 프레임만** forward-fill 한다 — ① 그룹 헤더 밴드(colspan 라벨)를 가로로 펴고(`1분기`→`1분기 1분기 1분기`), ② 왼쪽 키 열(rowspan/그룹)을 세로로 편다(`부장`→아래로). **값(데이터) 영역의 빈칸은 그대로 보존**(전체폭 단일 *제목* 행도 제외) — XML 경로의 `forward_fill_key_columns` 와 동일한 "그룹 열만 채우고 값 열은 보존" 가드. 다단 헤더에서 리프 헤더 행이 본문으로 내려가는 평면 구조는 마크다운 표 문법상 그대로 두고(의도된 동작), **밴드의 빈칸만** 채운다. 신규 `converter/html_to_md.py:_fill_header_frame`, 회귀 테스트 4개(`markdown_fill_vertical_key_column`·`markdown_fill_horizontal_band_and_key`·`markdown_fill_keeps_value_blank`·`markdown_fill_off_keeps_blanks`). 토글 OFF(기본)면 종전과 100% 동일.
- **Windows: macOS 0.5.0 의 ‘빈칸을 자동 채우기’ 토글 포팅 (0.2.7 — macOS 패리티)**: 병합 헤더 빈칸을 채우는 로직 자체는 Windows 가 import 하는 공용 `tablemark.converter.html_to_md` 에 이미 있었고, 빠진 건 토글·설정·배선뿐이었다(Windows 엔 `fill_blanks` 옵션이 아예 없었음 — macOS 는 XML 경로에만 있었고 XML 은 macOS 전용). 이제 `conversion.py` 의 `converted_clipboard(content, fill_blanks=False)` 가 플래그를 `html_table_to_markdown`/`convert_document_tables` 로 넘기고(기본 꺼짐 → 동작 무변화, HTML 슬롯 유지 — 불변식 3), `app.py` 에 `FILL_BLANKS_KEY` + 영속 설정 `self.fill_blanks` + 체크 가능한 ‘빈칸을 자동 채우기’ 메뉴 항목(토글 뒤·언어 앞 — macOS 설정 순서와 동일) + `toggle_fill_blanks`(플립+영속+메뉴 갱신)를 추가하고, 워처가 `self.fill_blanks` 를 전달한다. `i18n.py` 에 `menu.fill_blanks`(한 "빈칸을 자동 채우기" / 영 "Auto-fill blank cells") 추가 — Windows 엔 XML 경로가 없어 "XML:" 접두 없음, 트레이 메뉴라 툴팁도 없음. 회귀 테스트: 토글 OFF 면 빈칸 유지·ON 이면 병합 키 열 채움(+HTML 유지)·라벨 번역·메뉴 포함·토글 영속.

### Changed
- **‘빈칸을 자동 채우기’ 토글을 XML·마크다운 두 경로 공통으로 통합.** 옛 라벨 "XML: 빈칸을 자동 채우기"에서 **"XML:" 을 제거**(한·영 `i18n.py`) — 하나의 설정(`fill_blanks`, `settings.py`/NSUserDefaults 영속, 기본 꺼짐)이 이제 XML 변환과 마크다운 자동 변환 양쪽의 빈칸 채우기를 함께 제어한다. 툴팁도 두 경로를 모두 언급하도록 갱신.

## [0.4.2] - 2026-06-30

### Changed
- **XML 메뉴 라벨을 ‘표를 XML로 복사’ → ‘복사한 표를 XML로 변환’ (EN: Copy table as XML → Convert copied table to XML).** 동작은 그대로(클립보드의 표를 XML 로 바꿔 클립보드에 둠)지만, ① 동사를 **‘변환’**으로 바꿔 실제 동작(표→XML 변환)을 정확히 표현하고 앱 도움말의 기존 어휘(“XML로 변환됩니다”)와 맞췄으며, ② **‘복사한 표’**로 전제조건(먼저 표를 복사해야 함 — 첫 사용자 최다 혼란)을 명시. 능동형 ‘복사**한**’은 수동형 ‘복사된’(앱이 복사한 듯 읽힘)보다 사용자의 사전 동작을 직접 가리켜 더 자연스러움. 라이브 라벨(`i18n.py` 한·영) + 도움말·빈칸채우기 툴팁의 어휘(복사→변환) + README 사용법(한·영) 일관 갱신. 4관점 카피 리뷰(한국어 문구·발견성·동작정확성·일관성) 중 3관점이 ‘변환’ 지지, 2관점이 ‘복사한’ 표현에 독립 수렴.

## [0.4.1] - 2026-06-29

### Changed
- **표→XML 전역 단축키를 ⌘⌃C → ⌘⌃X 로 변경.** X = **X**ML mnemonic 으로 더 직관적. macOS 기본 단축키(`⌃⌘Space`·`⌃⌘F`·`⌃⌘Q`·`⌃⌘D` 등)·앱 단축키와 안 겹치고, 토글 `⌘⌃T` 와도 키보드상 멀리 떨어져 오발 위험 낮음. 변환 동작 자체는 그대로(클립보드의 표→XML 복사, HTML drop). `hotkey.py`(`KEY_X = 7`, kVK_ANSI_X)·`app.py`(핫키 등록 + 메뉴 key)·도움말·회귀 테스트(`hotkey_keycodes_are_ansi_x_and_t`) 갱신. 변환·핫키 테스트 영향 없음(converter 40/40, hotkey 4/4).

## [0.4.0] - 2026-06-26

### Removed
- **유료화 전면 제거 — 앱을 완전 무료로 전환.** StoreKit IAP 래퍼(`store.py`)를 삭제하고 XML 변환의 Pro 구독 게이팅을 해제 — 이제 누구나 ‘표를 XML로 복사’와 전역 단축키 ⌘⌃C 를 제한 없이 쓴다. 기부(Consumable 3종) 메뉴·‘구매 복원’ 항목·구독 시트·구매/복원 알럿을 모두 제거하고, `settings.py` 의 `pro_active`/`pro_expires` 캐시, `i18n` 의 결제·후원·Pro 문자열, `setup.py` 의 `StoreKit` include, `app.py` 의 관련 콜백·잠금(🔒) 표시도 함께 정리. 전역 단축키(`hotkey.py`)와 XML 변환 로직 자체는 그대로 — 게이팅만 사라졌다. 클립보드 변환 불변식·테스트 영향 없음(변환 테스트 38/38).

### Changed
- **XML 표 변환 형식을 중첩 계층(v2)으로 재설계.** 0.3.0 의 평면형(`<dataset><row><cell name="…">`, 가로 그룹만 `<group>` 중첩)을 대체. 이제 **가로·세로 양방향 다단 헤더를 모두 중첩**으로 보존한다: 루트 `<표>`, 세로 상위 그룹 `<직급그룹 이름="부장">`, 행 `<행 직책="대족장">`, 가로 상위 그룹 `<열그룹 이름="1분기">`, 리프 `<열 n="1">동</열>`. 세로 키는 v1 처럼 매 행에 반복(forward-fill)하지 않고 **부모 노드로 묶는다**(트레이드오프: 행이 자기완결 레코드가 아님 — 그룹 값이 부모에만 있음). 가로 차원 이름이 표에 없으면 지어내지 않고 일반 태그(`열그룹`/`열`)로 보존. 헤더가 XML 이름 규칙에 안 맞는 열은 세로 키로 승격하지 않아 정규화 없이 무손실 roundtrip 유지. `model_to_xml`·`table_xml_to_model`·`is_table_xml` 재작성, 변환 테스트 38/38.
- **메뉴 구조 정리 (UX P2)**: 메뉴 순서를 동작(토글·XML) → ‘설정 ▸’ → 도움말·종료로 정리(첫인상은 유틸리티). 변환 토글(`enabled`)은 **의도적으로 비영속**(매 실행 켜짐 — "꺼둔 걸 잊고 고장 오인" 방지) — 정책을 CLAUDE.md "설정 영속성 정책" 에 명문화.

### Added
- **로컬 진단 — 로컬 전용 크래시 캡처 + ‘문제 신고용 로그’ (외부 전송 0줄).** 직배포 환경에서 안 보이던 실패를 **로컬 로그**로 끌어온다. 외부로 보내는 코드는 한 줄도 없어 `windows/PRIVACY.md` 의 "no network connections, no telemetry … nothing is sent to any external server" 약속을 그대로 지킨다 (DiskOUT 식 *원격* 수집은 이 약속과 충돌해 의도적으로 채택하지 않음). 신규 `tablemark/diagnostics.py`·`windows/tabledown_windows/diagnostics.py`(동형): ① `install_crash_hooks()` — `sys.excepthook` + **`threading.excepthook`**(지금까지 무음으로 사라지던 **클립보드 워처 데몬 스레드의 미포착 예외**를 포착 — 가장 핫한 실패면) + `faulthandler`(네이티브 폴트 — `enable()` 가 덤프 후 기본 핸들러로 넘겨 OS 크래시리포터(.ips)도 그대로 동작), 전부 `Tabledown.log`/`Tabledown.crash` 에만 기록. **크래시 기록은 예외 타입+프레임 위치만 남기고 예외 메시지·페이로드는 기록하지 않는다**(공유 로그 안전성). ② 메뉴 ‘문제 신고용 로그 열기’(mac·Windows) → **스크럽된**(홈경로·유저명·볼륨/드라이브명·secret 제거) `Tabledown-diagnostics.txt` 생성 후 Finder/Explorer 로 열기 — 사용자가 직접 버그리포트에 첨부, 클립보드 미접촉(워처 레이스 없음). ③ **로그 누출 하드닝**: 클립보드 변환 catch 사이트(`copy as xml failed`/`clipboard update failed`)의 `{exc}` → `{type(exc).__name__}`, `converter/table_xml.py` 의 에러 메시지에서 클립보드 파생 텍스트(파서 detail·헤더 태그명) 제거 — 공유 로그에 표 데이터가 새지 않게. 회귀 테스트 추가: converter 누출 메시지 2 + diagnostics(scrub 3·threading.excepthook 1) + i18n 키 1.
- **자동 변환 피드백 — 무반응 제거 (P0 후속).** 그동안 클립보드 자동 변환(Excel↔Markdown)은 **아무 표시 없이** 백그라운드에서 일어나, 사용자가 "변환됐다"는 걸 알 수 없고 붙여넣기 결과가 예상과 다를 때 원인이 Tabledown 인지조차 몰랐다(UX 리뷰 최상위 이슈). 이제 워처가 실제로 표를 변환하면 **메뉴바 아이콘이 0.5초간 체크 표시로 깜빡**인다(수동 XML 의 1초 플래시보다 짧게 — 매 복사마다 떠도 거슬리지 않게). 워처는 백그라운드 스레드라 `AppHelper.callAfter` 로 메인 스레드에 넘겨 아이콘을 건드린다. 변환 결과엔 `mark_generated` 가 찍혀 다음 워처 틱이 no-op 이 되므로 플래시가 무한 반복되지 않는다. 팝업·소리·시스템 알림 없음("권한 0개" 유지).
- **전역 단축키 ⌘⌃C (`hotkey.py`)**: Carbon `RegisterEventHotKey` 로 클립보드의 표를 XML 로 바로 변환 — Accessibility·Input Monitoring 등 권한 불필요("권한 0개" 유지). 등록 실패해도 메뉴 항목은 그대로 동작(graceful). (유료 게이팅 없이 누구나 사용.)
- **전역 단축키 — 자동변환 켜기/끄기 토글 (macOS ⌘⌃T · Windows Ctrl+Alt+T)**: 클립보드 자동 변환을 어느 앱에서든 키 하나로 일시정지/재개. 권한 불필요·등록 실패 시 메뉴로 graceful fallback(핫키는 액셀러레이터일 뿐). macOS 는 기존 단일 핫키(`tablemark/hotkey.py`)를 다중 핫키 매니저(`GlobalHotkeys`)로 일반화 — Carbon 핸들러 하나가 발화된 핫키 ID(`GetEventParameter`)로 ⌘⌃C/⌘⌃T 를 분기(핫키별 핸들러는 서로 이벤트를 삼키므로 금지). Windows 는 신규 `tabledown_windows/hotkey.py` — user32 `RegisterHotKey`(NULL hwnd) + 전용 메시지 루프 스레드(`MOD_NOREPEAT`), 비-Windows 에선 우아하게 degrade. 키 선택: ⌘⌃ 조합·`Ctrl+Alt` 은 흔한 시스템/앱 단축키와 안 겹치고 토글 글자는 양 플랫폼 T 로 통일. 회귀 테스트: macOS `run_hotkey_tests`(ID 분기·단일 fallback·오발 방지) + Windows `HotkeyTests`(등록/디스패치/degrade). 도움말·CLAUDE.md 갱신.
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

[Unreleased]: https://github.com/yooongZa/tabledown/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/yooongZa/tabledown/compare/v0.4.2...v0.5.0
[0.4.2]: https://github.com/yooongZa/tabledown/compare/v0.4.1...v0.4.2
[0.4.1]: https://github.com/yooongZa/tabledown/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/yooongZa/tabledown/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/yooongZa/tabledown/compare/v0.2.4...v0.3.0
[0.2.4]: https://github.com/yooongZa/tabledown/compare/v0.2.3...v0.2.4
[0.2.3]: https://github.com/yooongZa/tabledown/compare/v0.2.2...v0.2.3
[0.2.2]: https://github.com/yooongZa/tabledown/compare/v0.2.1...v0.2.2
[0.2.1]: https://github.com/yooongZa/tabledown/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/yooongZa/tabledown/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/yooongZa/tabledown/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yooongZa/tabledown/releases/tag/v0.1.0
