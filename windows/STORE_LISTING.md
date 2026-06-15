# Tabledown — Microsoft Store Listing (Partner Center 제출용 문구)

Partner Center 의 각 입력란에 그대로 붙여넣을 수 있는 문구 모음. 영어/한국어 두 언어
리스팅을 모두 등록하면 좋다(매니페스트가 en-US, ko-KR 리소스를 선언함).

---

## Product name
`Tabledown`

## Category
Productivity (생산성)

## Age rating
3+ / Everyone — 사용자 생성 콘텐츠 없음, 네트워크 없음, 광고 없음.

---

## Short description (≤ ~100 chars)

**EN:** Convert Excel & Google Sheets tables to Markdown — and back — with just Ctrl+C / Ctrl+V.

**KO:** Excel·구글 시트 표를 Ctrl+C / Ctrl+V 만으로 마크다운으로, 다시 표로 변환합니다.

---

## Description

**EN**
```
Tabledown lives in your system tray and quietly converts tables on your clipboard.

• Copy a table from Excel or Google Sheets, paste into a Markdown editor (Obsidian,
  GitHub, Notion, a chat box) — it arrives as a clean Markdown table.
• Copy a Markdown table, paste into Excel — it lands in separate cells.
• One copy serves every destination: spreadsheets and Word still paste a real
  table, Markdown editors get Markdown. Nothing to click — it just works on Ctrl+C.

Toggle conversion on/off from the tray icon (a red slash means it's off), switch
between English and Korean, and that's it.

100% local. No account, no network, no telemetry, no ads. Your clipboard never
leaves your PC.
```

**KO**
```
Tabledown은 시스템 트레이에 상주하며 클립보드의 표를 조용히 변환합니다.

• Excel·구글 시트에서 표를 복사해 마크다운 에디터(Obsidian, GitHub, Notion, 채팅창)에
  붙여넣으면 깔끔한 마크다운 표로 들어갑니다.
• 마크다운 표를 복사해 Excel에 붙여넣으면 셀 단위로 분리되어 들어갑니다.
• 한 번 복사로 모든 곳에 대응: 스프레드시트·Word는 진짜 표로, 마크다운 에디터는
  마크다운으로 붙습니다. 누를 것 없이 Ctrl+C 만으로 동작합니다.

트레이 아이콘에서 변환 켜기/끄기(빨간 사선 = 꺼짐), 한국어/영어 전환만 하면 끝입니다.

100% 로컬. 계정·네트워크·텔레메트리·광고 없음. 클립보드는 PC를 벗어나지 않습니다.
```

---

## Search terms / keywords
`excel to markdown`, `csv to markdown`, `markdown table`, `clipboard`, `tray`,
`obsidian`, `google sheets`, `paste table`, `convert table`

---

## What's new (first release)
**EN:** First Windows release. Excel/Sheets ↔ Markdown table conversion from the
system tray, English/Korean UI, on/off toggle.
**KO:** 첫 Windows 릴리스. 시스템 트레이에서 Excel/시트 ↔ 마크다운 표 변환, 한/영 UI,
켜기/끄기 토글.

---

## Notes for certification (심사 노트 — Submission > Properties / Notes)

> Tabledown is a clipboard utility that watches the system clipboard and rewrites
> copied spreadsheet tables into Markdown (and Markdown tables back into a
> spreadsheet-compatible format). Reading and writing the global clipboard from a
> background tray process requires full trust (Win32 clipboard APIs); the UWP
> clipboard API only permits access while the app window has focus, which is
> incompatible with passive background conversion. No data leaves the device — no
> network calls, no telemetry.
>
> How to test: launch the app (a table icon appears in the notification area).
> Copy a cell range in Excel, then paste into Notepad — the paste is a Markdown
> table. Copy a Markdown table (e.g. `| a | b |` / `| --- | --- |` / `| 1 | 2 |`)
> and paste into Excel — it splits into cells. Right-click the tray icon for the
> on/off toggle and language menu.

## restricted capability 사유 (runFullTrust)
위 Notes 의 첫 문단을 그대로 사용. `runFullTrust` 는 거절 사유가 아니지만 심사가
길어질 수 있으므로 사유를 명확히 적는다.

---

## Privacy policy URL (필수)
클립보드를 읽는 앱이라 **필수**. 다음 중 하나를 사용:
- `https://github.com/yooongZa/tabledown/blob/main/windows/PRIVACY.md` (저장소에 추가한 파일)
- 또는 GitHub Pages 로 호스팅한 URL

> ⚠️ 제출 전 `windows/PRIVACY.md` 를 `main` 에 push 해서 위 링크가 실제로 열려야 한다.

---

## Screenshots (1~9장, 권장 1366×768 이상)
- 트레이 아이콘 + 우클릭 메뉴(Tabledown 사용 / 언어 / 도움말 / 종료)
- Excel 표 복사 → 마크다운 에디터에 마크다운 표로 붙은 화면 (before/after)
- 마크다운 표 복사 → Excel 셀에 분리되어 붙은 화면
- (참고) macOS 판 비교 스크린샷 `assets/tabledown-paste-comparison.png` 의 구도를 재활용 가능
