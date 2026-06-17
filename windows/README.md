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

## EXE 빌드

```powershell
cd windows
.\build_windows.ps1
```

빌드 결과는 `windows\dist\Tabledown-Windows.exe`에 생성됩니다.

## 파일 구조

```text
windows/
├── run_windows.py
├── requirements.txt
├── build_windows.ps1
├── tabledown_windows/
│   ├── app.py                 # Windows tray(시스템 트레이) 앱
│   ├── conversion.py          # 플랫폼 독립 변환 흐름
│   ├── html_clipboard.py      # Windows CF_HTML format(HTML 클립보드 형식)
│   ├── i18n.py                # Windows locale(로캘) + 언어 저장
│   ├── logger.py              # Windows 로그 경로
│   ├── startup_task.py        # 로그인 시 자동 실행(WinRT StartupTask, MSIX)
│   └── win_clipboard.py       # pywin32 clipboard(클립보드) 래퍼
├── tests/
└── tools/
```

## 제한사항

- 이 폴더의 clipboard(클립보드) 기능은 Windows에서만 실행됩니다.
- macOS에서 검증 가능한 부분은 conversion(변환), i18n(다국어), CF_HTML format(HTML 클립보드 형식) 단위 테스트로 확인합니다.
- Windows clipboard(클립보드)의 일부 handle(핸들) 기반 format(형식)은 pywin32로 안전하게 복사할 수 없어서 보존 대상에서 제외됩니다.
