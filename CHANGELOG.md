# Changelog

이 프로젝트의 모든 주요 변경 사항은 이 파일에 기록됩니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/) 를 따르며,
버전 번호는 [Semantic Versioning](https://semver.org/lang/ko/) 을 사용합니다.

## [Unreleased]

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

[Unreleased]: https://github.com/yooongZa/tabledown/compare/v0.1.1...HEAD
[0.1.1]: https://github.com/yooongZa/tabledown/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/yooongZa/tabledown/releases/tag/v0.1.0
