# Tabledown for Windows

Excel·Google Sheets ↔ Markdown(마크다운) 표를 복사·붙여넣기로 변환하는 Windows 트레이 앱입니다.

[프로젝트 소개](../README.md) · [English](../README.en.md) · [변경 이력](../CHANGELOG.md)

## 설치

**Windows 10/11 · x64**

[Microsoft Store](https://apps.microsoft.com/detail/9NGS4C0N2Z6L) 또는 [GitHub Windows 릴리스](https://github.com/yooongZa/tabledown/releases/tag/windows-v0.3.2)에서 받으세요.

- **ZIP:** 압축을 풀고 `Tabledown-Windows.exe`를 실행합니다. Python 설치가 필요 없습니다. EXE에는 상용 코드 서명이 없어 게시자 확인 안내가 나타날 수 있습니다.
- **자체 서명 MSIX:** 설치 시험용입니다. 릴리스의 `Tabledown-dev.cer`와 `INSTALL.txt` 안내에 따라 설치하세요. Microsoft Store 제출 패키지는 별도로 빌드합니다.

다운로드 파일은 릴리스의 `SHA256SUMS.txt`로 확인할 수 있습니다.

## 사용법

1. 앱을 실행하고 시스템 트레이의 표 아이콘을 확인합니다.
2. Excel·Google Sheets에서 표를 `Ctrl+C`로 복사합니다.
3. Markdown 문서에 `Ctrl+V`로 붙여넣습니다. Markdown 표를 복사해 Excel에 붙이는 역방향도 지원합니다.

**표를 AI에 전달할 때**는 Excel 데스크톱 앱에서 제목 행과 항목명을 포함한 사각형 범위를 선택하고 아래 메뉴를 사용하세요. 먼저 `Ctrl+C`를 누를 필요가 없습니다. XML 전체를 원하는 AI에 직접 붙여넣습니다.

| 메뉴 | 포함되는 내용 | 단축키 |
| --- | --- | --- |
| XML 변환 복사 | 표시값·빈칸·병합 구조·출처 | `Ctrl+Alt+X` |
| 셀 값·수식·참조를 XML로 복사 | 현재 값·수식·직접 참조값·계산 상태 | `Ctrl+Alt+E` |
| AI용 간결 복사 | 수식 XML 정보를 공통 참조 목록과 추정 맥락으로 정리 | 메뉴에서 실행 |

XML 복사는 최대 **10,000셀**, 셀 값 합계 **5,000,000자**, 출력 **10MB**까지 지원합니다. 원본 셀을 수정하거나 재계산하지 않습니다. 읽는 동안 다른 내용을 복사하면 진행 중인 XML 쓰기를 취소합니다.

## 설정과 문제 해결

- **변환 켜기·끄기:** `Ctrl+Alt+T`. 아이콘에 사선이 있으면 꺼진 상태입니다.
- **그룹·분류 빈칸 채우기:** 기본 꺼짐. 자동 Markdown의 열 그룹 제목·왼쪽 분류 열과 일반 XML의 왼쪽 분류 열에 적용합니다. [자세한 적용 기준](../README.md#설정)
- **언어·로그인 시 실행:** 한국어/영어를 지원합니다. 로그인 시 실행은 설치형 MSIX에서 설정할 수 있습니다.
- **XML 복사 실패:** Excel을 하나만 실행하고 단일 사각형 범위를 선택하세요. 숫자나 날짜가 `###`로 보이면 열 너비를 늘리세요.
- **진단 로그:** 트레이의 **문제 신고용 로그 열기**를 사용하세요. 로그는 `%LOCALAPPDATA%\Tabledown\Tabledown.log`에 있습니다.

수식 XML에는 표시 서식·병합 구조를 넣지 않습니다. 현재 계산 결과를 읽으므로 Excel이 계산 중이면 결과가 달라질 수 있습니다. 지원하지 않거나 읽지 못한 참조는 ‘일부’로 표시합니다. XML 복사는 Google Sheets·LibreOffice에서 지원하지 않습니다.

## 개인정보

모든 처리는 PC 안에서 이루어지며 Tabledown이 AI 서비스나 외부 서버로 내용을 보내지 않습니다. **수식 복사에는 선택 영역 밖이나 같은 통합문서의 다른 시트에 있는 직접 A1 참조 셀 값도 포함될 수 있습니다.** 셀 값과 수식은 로그에 기록하지 않습니다.

[개인정보 처리방침](PRIVACY.md) · [문제 제보](https://github.com/yooongZa/tabledown/issues)

## 개발

<details>
<summary>실행·테스트·빌드 명령</summary>

Windows에서 Python 3.10 이상을 사용합니다. 저장소 루트에서 실행하세요. 빌드 스크립트에는 **PowerShell 7**이 필요합니다.

```powershell
cd windows
py -3 -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python run_windows.py
```

같은 `windows` 폴더에서 테스트와 빌드를 실행합니다.

```powershell
cd tests
..\.venv\Scripts\python -m unittest test_windows_port test_windows_table -v
cd ..
pwsh -File build_windows.ps1
```

결과는 `windows\dist\Tabledown-Windows\`에 생성됩니다. 폴더 전체를 배포하세요. Mac에서는 [Windows 빌드 워크플로](../.github/workflows/windows-build.yml)를 실행해 산출물을 받을 수 있습니다.

[MSIX·Microsoft Store 패키징](PACKAGING.md) · [개발 지침](../AGENTS.md)

</details>
