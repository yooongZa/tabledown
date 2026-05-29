# Changelog

이 프로젝트의 모든 주요 변경 사항은 이 파일에 기록됩니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르며,
버전 번호는 [Semantic Versioning](https://semver.org/lang/ko/) 을 사용합니다.

## [Unreleased]

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

[Unreleased]: https://github.com/yooongZa/tabledown/compare/v0.2.3...HEAD
[0.2.3]: https://github.com/yooongZa/tabledown/compare/v0.2.2...v0.2.3
[0.2.2]: https://github.com/yooongZa/tabledown/compare/v0.2.1...v0.2.2
[0.2.1]: https://github.com/yooongZa/tabledown/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/yooongZa/tabledown/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/yooongZa/tabledown/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yooongZa/tabledown/releases/tag/v0.1.0
