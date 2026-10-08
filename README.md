# Tabledown

<img src="assets/generated/tablemark_app_1024.png" width="80" alt="Tabledown 아이콘">

**Excel·Google Sheets ↔ Markdown(마크다운) 표를 복사·붙여넣기로 변환합니다.**

macOS 메뉴바와 Windows 트레이에서 실행됩니다. Excel의 표 구조나 수식은 XML로 복사해 AI에 전달할 수 있습니다.

한국어 · [English](README.en.md)

## 다운로드

| 플랫폼 | 스토어 | 직접 다운로드 |
| --- | --- | --- |
| macOS 12 이상 · Apple Silicon | [Mac App Store](https://apps.apple.com/app/id6768205551) | [DMG](https://github.com/yooongZa/tabledown/releases/latest/download/Tabledown.dmg) |
| Windows 10/11 · x64 | [Microsoft Store](https://apps.microsoft.com/detail/9NGS4C0N2Z6L) | [Windows ZIP](https://github.com/yooongZa/tabledown/releases/tag/windows-v0.3.2) |

Mac은 DMG를 열고 앱을 Applications(응용 프로그램)에 옮기세요. 배포 파일은 Developer ID 서명과 Apple 공증을 통과했습니다.
Windows는 ZIP을 풀고 `Tabledown-Windows.exe`를 실행하세요. 별도의 Python 설치가 필요 없습니다. 자체 서명 MSIX 설치는 [Windows 안내](windows/README.md)를 참고하세요.

## 사용법

1. Tabledown을 실행하고 메뉴바 또는 트레이의 표 아이콘을 확인합니다.
2. Excel·Google Sheets에서 표를 복사합니다.
3. Obsidian이나 Markdown 문서에 붙여넣습니다. Markdown 표를 복사해 Excel에 붙이면 셀별로 들어갑니다.

Mac은 `Cmd+C` / `Cmd+V`, Windows는 `Ctrl+C` / `Ctrl+V`를 사용합니다. 별도의 Markdown 복사 메뉴는 없습니다. 원본 HTML 표도 유지하므로 Excel·Word에서는 표로 붙습니다. Markdown 원문이 필요하면 붙여넣는 앱의 **일반 텍스트로 붙여넣기**를 사용하세요.

## Excel 표를 AI에 전달하기

Excel 데스크톱 앱에서 **제목 행과 항목명을 포함한 사각형 영역**을 선택한 뒤 XML 복사 메뉴를 누르세요. 먼저 복사할 필요가 없습니다. 생성된 XML 전체를 원하는 AI에 직접 붙여넣습니다.

| 복사할 정보 | 포함되는 내용 | Mac | Windows |
| --- | --- | --- | --- |
| 일반 XML | 화면에 표시된 값·빈칸·병합 구조·출처 | `⌘⌃X` | `Ctrl+Alt+X` |
| 수식 XML | 현재 값·수식·직접 참조값·계산 상태 | `⌘⌃E` | `Ctrl+Alt+E` |

Mac의 **수식 포함 XML 변환 복사**는 반복 참조를 공통 목록으로 정리합니다. Windows는 **셀 값·수식·참조를 XML로 복사**와 **AI용 간결 복사**를 따로 제공합니다.

XML 복사는 최대 **10,000셀**, 셀 값 합계 **5,000,000자**, 출력 **10MB**까지 지원합니다. 수식 XML은 Excel의 현재 값을 읽으며 재계산하지 않습니다. 표시 서식·병합 구조가 필요하면 일반 XML을 사용하세요. 지원하지 않거나 읽지 못한 참조는 결과에 ‘일부’로 표시합니다.

## 설정

- **변환 켜기·끄기:** Mac `⌘⌃T`, Windows `Ctrl+Alt+T`. 아이콘에 사선이 보이면 꺼진 상태입니다.
- **그룹·분류 빈칸 채우기:** 기본값은 꺼짐입니다. 자동 Markdown에서는 병합된 열 그룹 제목과 왼쪽 분류 열을, 일반 XML에서는 왼쪽 분류 열을 채웁니다. 원본 Excel·HTML과 수식 XML은 바꾸지 않습니다.
- **언어·자동 실행:** 한국어/영어와 로그인 시 실행을 설정할 수 있습니다. Windows 자동 실행은 설치형 MSIX에서 지원합니다.

<details>
<summary>빈칸 채우기의 적용 기준</summary>

분류 열은 왼쪽부터 처음으로 빈칸 없는 본문 열의 앞까지로 추정합니다. 자동 Markdown은 열 그룹 제목을 가로로, 분류 열을 세로로 채웁니다. 일반 XML은 분류 열을 위쪽 값으로 채운 뒤 남은 빈칸을 왼쪽 값으로 채웁니다. 본문에서 그 열과 오른쪽 열의 빈칸은 유지하며, 일반 XML에는 채운 셀의 원본 A1 주소를 기록합니다.

</details>

## 문제가 생기면

- **변환 전 표가 붙는 경우:** 앱이 켜져 있는지 확인하고, 복사 후 잠시 기다렸다가 붙여넣으세요.
- **XML을 복사할 수 없는 경우:** Excel의 단일 사각형 범위를 선택하세요. Mac에서 권한 요청이 나오면 Automation(자동화) 설정에서 Tabledown의 Excel 접근을 허용하세요. 기본 표 변환에는 손쉬운 사용·입력 모니터링 권한이 필요 없습니다.
- **큰 표가 오래 걸리거나 `###`가 보이는 경우:** 범위를 나누거나 Excel 열 너비를 늘리세요. Mac의 큰 선택 영역은 수십 초가 걸릴 수 있습니다.

Markdown은 색상·글꼴·병합 모양을 표현하지 못합니다. XML 복사는 Excel 데스크톱 전용이며 Google Sheets·LibreOffice에서는 사용할 수 없습니다.

메뉴의 **문제 신고용 로그 열기**에서 진단 정보를 확인할 수 있습니다. [문제 제보](https://github.com/yooongZa/tabledown/issues) · [변경 이력과 검증 기록](CHANGELOG.md)

## 개인정보 처리방침

모든 변환은 기기 안에서 처리합니다. Tabledown은 개인정보를 수집·판매·공유하지 않으며 계정, 광고, 사용량 추적, 외부 서버 전송 기능이 없습니다.

자동 변환은 현재 클립보드를 읽고 같은 클립보드에 결과를 기록합니다. XML 복사는 선택한 Excel 셀을 읽습니다. **수식 복사에는 선택 영역 밖이나 같은 통합문서의 다른 시트에 있는 직접 A1 참조 셀 값도 포함될 수 있습니다.** 외부 통합문서를 열거나 참조를 재귀적으로 추적하지 않습니다.

셀 값과 수식은 진단 로그에 기록하지 않습니다. Mac의 로컬 로그는 `~/Library/Logs/Tabledown.log`에 저장되며 직접 삭제할 수 있습니다. [Windows 개인정보 처리방침](windows/PRIVACY.md)

<details>
<summary>개발자 안내</summary>

Mac에서 Python 3.10 이상으로 저장소 루트에서 실행합니다.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python run.py
```

앱 코드는 `tablemark/`, 공유 변환 로직은 `tablemark/converter/`, Windows 앱은 `windows/`에 있습니다.

[XML 형식](docs/xml-format-spec.md) · [개발·테스트·배포 지침](AGENTS.md) · [Mac 빌드](scripts/build_release.sh) · [Windows 개발·빌드](windows/README.md#개발)

</details>

[MIT License](LICENSE)
