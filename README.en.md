# Tabledown

<p align="center">
  <img src="assets/generated/tablemark_app_1024.png" width="96" alt="Tabledown icon">
</p>

<p align="center">
  <a href="https://apps.apple.com/app/id6768205551"><img alt="Mac App Store" src="https://img.shields.io/itunes/v/6768205551?label=Mac%20App%20Store"></a>
  <a href="https://github.com/yooongZa/tabledown/releases/latest"><img alt="GitHub release" src="https://img.shields.io/github/v/release/yooongZa/tabledown?label=GitHub"></a>
  <img alt="Platform" src="https://img.shields.io/badge/platform-macOS-111111">
  <img alt="License" src="https://img.shields.io/badge/license-MIT-blue">
</p>

<p align="center">
  English | <a href="README.md">Korean</a>
</p>

> A macOS menu bar app that turns spreadsheet tables into Markdown source tables.

Tabledown converts tables copied from Excel or Google Sheets into `| ... |` tables that paste cleanly into Obsidian, GitHub README files, and Markdown editors. No export flow, no extra window. Just copy with `Cmd+C` and paste with `Cmd+V`.

## Download

The Mac App Store is the recommended way to install (automatic updates).

<p>
  <a href="https://apps.apple.com/app/id6768205551">
    <img src="https://toolbox.marketingtools.apple.com/api/v2/badges/download-on-the-mac-app-store/black/en-us" alt="Download on the Mac App Store" height="48">
  </a>
</p>

Or get it directly from GitHub Releases.

**[Download Tabledown.dmg](https://github.com/yooongZa/tabledown/releases/latest/download/Tabledown.dmg)**

Install (DMG):

1. Open `Tabledown.dmg`
2. Drag `Tabledown.app` into Applications
3. Launch the app and check for the table icon in the menu bar

The release DMG is built with Developer ID signing and Apple notarization. The Mac App Store build runs in the App Sandbox.

## Why Use It

| Input | Paste Target | Output |
|------|--------------|--------|
| Excel/Google Sheets table (`Cmd+C`) | Obsidian/GitHub/Markdown editor (`Cmd+V`) | Markdown source table |
| Markdown table (`Cmd+C`) | Excel (`Cmd+V`) | Spreadsheet cells |
| Select an Excel range, then “Copy Markdown” | Notes or Markdown documents (`Cmd+V`) | Selected columns, displayed values, and cell line breaks |
| Select an Excel table range, then “Copy as XML” or `⌘⌃X` | LLM prompt (`Cmd+V`) | Hierarchy, exact merge ranges, source, and displayed values as XML |
| Select an Excel formula range, then “Copy as XML with Formulas” or `⌘⌃E` | LLM prompt (`Cmd+V`) | Cell types, raw values, calculation state, A1/R1C1 formulas, and direct reference values as XML |

When Tabledown is on, spreadsheet tables paste as Markdown source in Markdown editors such as Obsidian.

```markdown
| Name | Task |
| --- | --- |
| Tabledown | Paste tables as Markdown |
```

Tabledown augments the clipboard's plain-text slot with Markdown while keeping the HTML table slot intact (since 0.2.4). The destination app picks the format it prefers, so a single copy pastes as Markdown source in Markdown editors and as a rendered table in rich text editors such as TextEdit, Word, or Excel.

<p align="center">
  <img src="assets/tabledown-paste-comparison.png" width="760" alt="The same copied table pasted as Markdown source and as a rendered table">
</p>

The screenshot above shows the two paste results — Markdown source (top) and a rendered table (bottom). (It was captured on a version before 0.2.4, when the on/off toggle made this difference; today both formats coexist on the clipboard and the destination app chooses.) In other words, Tabledown is not a pretty table renderer. It is a table converter for Markdown documents.

## Copy Markdown (Excel only)

Select a rectangular range including its header row in Excel, choose **“Copy Markdown”**, and paste into your document. There is no need to press `Cmd+C` first. The copy retains selected column order, trailing blank columns, displayed values, and cell line breaks. Markdown source preserves whitespace and escapes literal special characters. The first selected row becomes the table header; dates, percentages, and currency use Excel’s displayed strings.

Markdown represents merged cells as separate grid cells. Colors, fonts, column widths, and merged shapes cannot be reproduced by Markdown syntax. The existing “Fill blanks automatically” option applies only to group/header blanks, keeping data blanks intact. A UTF-8 HTML table from the same selection also carries merges and line breaks for destinations such as Excel and Notes. The destination chooses its format and controls how whitespace is displayed.

This uses the same stable Excel selection reader as regular XML: up to 10,000 cells and 5,000,000 value characters, with a combined UTF-8 Markdown+HTML limit of 10 MB. Read, conversion, and size failures stop before writing. Copying something new while the export is being prepared cancels its write. Success replaces older Excel-specific formats with Markdown and HTML from the same selection.

All three manual commands require the Excel desktop app. **Markdown focuses on the table’s presentation in documents; both XML commands organize information for AI.** Existing automatic Excel/Google Sheets ↔ Markdown copy and paste remains available. These menu descriptions reflect TestFlight 0.6.1/build 0.6.5 as of September 10, 2026.

## Copy as XML (for AI)

Tabledown organizes tables as **nested multi-level-header XML**, keeping header and group relationships connected when you paste into AI. Multi-level (group) headers are represented as an XML nesting hierarchy: the root is `<표>` (table), vertical groups are `<{header}그룹 이름="value">`, a row is `<행 {header}="value">`, horizontal groups are `<열그룹 이름="value">`, and a cell is `<열 n="headerValue">cellValue</열>`. This nesting is a structural inference from Excel merges and value layout; the exact source merge ranges are preserved separately in the root’s `병합범위` attribute. (The structural tags are Korean words — `표` table, `행` row, `열` column, `열그룹` column-group — so the XML never collides with real HTML elements.)

Example: a cross-table of two vertical levels (Rank ▸ Manager/Deputy, Title) × two horizontal levels (Q1 ▸ 1,2,3 / Q2 ▸ 4,5,6). For readability, this abbreviated example omits only the root source and merge metadata.

```xml
<표>
  <직급그룹 이름="부장">
    <행 직책="대족장">
      <열그룹 이름="1분기"><열 n="1">동</열><열 n="2">해</열><열 n="3">물</열></열그룹>
      <열그룹 이름="2분기"><열 n="4">과</열><열 n="5">백</열><열 n="6">두</열></열그룹>
    </행>
    <행 직책="족장"> … </행>
    <행 직책="추장"> … </행>
  </직급그룹>
  <직급그룹 이름="차장"> … </직급그룹>
</표>
```

- **Excel table structure and displayed values → XML (menu or ⌘⌃X)**: Select one rectangular table range in the Excel desktop app, then click **“Copy as XML”** or press **⌘⌃X**. Without first pressing `Cmd+C`, Tabledown reads Excel's displayed number, date, percentage, currency, and custom-formatted values, error values, significant data-cell text whitespace, blanks, and merge structure. Formula text is not included. In data cells, literal `<br>` text remains distinct from an actual in-cell line break. It never falls back to an older clipboard table. On success the menu bar icon briefly shows a checkmark and the previous clipboard formats are replaced with plain-text XML so every destination receives XML. Google Sheets, LibreOffice, and clipboard Markdown/XML are not inputs to this manual command. There is no automatic XML→table direction.
- **Source facts are separated from inference**: Root attributes `통합문서`, `시트`, `주소`, `행수`, and `열수` identify the source, while `병합범위` records Excel’s actual merge areas. In selections with at least two columns, `제목N주소`/`제목N값` preserve leading full-width titles omitted from the hierarchy; a vertical merge in a one-column selection is not misclassified as a title. `헤더기준="추정"` means Excel provides no definitive header marker, so Tabledown inferred headers from the first row and merge layout. **Include the header row in your selection.** Blank, whitespace-only, and duplicate horizontal leaf headers retain their original `n` value and use the auxiliary column index `i` to stay distinguishable.
- **Only a stable selection is exported**: Tabledown requires two consecutive matching snapshots of the selection, values, and merge structure (at most three reads) before writing the clipboard. Disjoint ranges, selections over **10,000 cells**, partially selected merged cells, and tables that change while being read are rejected while preserving the existing clipboard. If Excel shows a value as `##` because a column is too narrow or a date/time cannot be displayed, Tabledown asks you to make the full value visible instead of exporting damaged text (literal `##` text is preserved). Both general and formula XML are limited to 5,000,000 total cell-value characters and 10 MB of UTF-8; an over-limit export leaves the clipboard untouched and asks for a smaller range.
- **Progress state and clipboard-overwrite protection**: XML commands first show `…` beside the menu-bar icon, switch the menu to “Copying…”, and allow only one export at a time. Because macOS requires NSAppleScript on the main thread, a large selection—especially near 10,000 cells—can leave the menu unresponsive for tens of seconds, and an in-progress read cannot currently be cancelled. Split the range if it takes too long. If you copy something else while XML is being prepared, the final clipboard-generation check cancels the XML write. A success checkmark appears only after read-back verification. macOS exposes no atomic clipboard compare-and-swap, so an extremely small race window remains between the final check and clearing the pasteboard.
- **Both horizontal and vertical groups nest**: horizontal multi-level headers nest as `<열그룹>`, and vertical groups (Rank: Manager/Deputy) nest as `<{header}그룹>`, keeping the table's hierarchy. Horizontal headers live in `n=`/`이름=` attributes (not tag names), so spaces, symbols, or leading digits never mangle a tag and stay safe in any standard XML parser. The Korean `<표>` root (not `<table>`) means the content survives even where the text is rendered as HTML (a browser, an Obsidian preview).
- **Vertical groups nest as parent nodes (changed from the previous version)**: the previous version repeated each vertical key column (Rank) on every row, so each row was a self-contained record. Now vertical groups nest as parent nodes (`<직급그룹>`) — the hierarchy is preserved, but the group value lives only on the parent, so **a single row on its own no longer carries its Rank** (a deliberate trade: hierarchy preservation over self-contained rows).
- **“Auto-fill blank cells” (Settings ▸, off by default)**: When converting a table, this one toggle drives **both the Markdown and XML paths**, filling blank cells in the left grouping (key) columns and the header frame from the value above (vertical), then to the left (horizontal). Blanks in the data (value) region are left as-is. General XML records whether the rule ran and the exact source cells it changed in `빈칸채움`, `빈칸채움기준`, `빈칸채움수`, and `빈칸채움셀`, so generated values are not mistaken for original Excel input.

## Copy as XML with Formulas (Excel only)

Select one rectangular range containing formulas in the Excel desktop app, then click **“Copy as XML with Formulas”** in the menu bar or press the global shortcut **⌘⌃E**. Tabledown copies every selected cell in the original row-and-column shape: constant values, blanks, addresses, and each formula cell’s current result plus A1/R1C1 formulas. Direct static A1 references in the same workbook link their current values through a shared reference list, on both the current sheet and other sheets. Merge hierarchy and display formatting are not included. You do not need to press `Cmd+C` first.

Include headers and item names in the selection. Repeated references are stored once and linked to each formula. This command now includes the former compact AI export. **Paste the complete XML into AI** so values, formulas, sources, and inferred context stay connected.

```xml
<표범위 통합문서="Book1.xlsx" 시트="Sheet1" 주소="$A$1:$C$2" 행수="2" 열수="3" 형식버전="1" 계산모드="automatic" 계산상태="unavailable" 계산결과상태="freshness_unverified" 값기준="Excel현재원시값" 표시정보상태="미포함" 병합정보상태="미포함" 형식="AI간결수식">
  <맥락 상태="미확인" 기준="선택범위내단순헤더" 사유="ambiguous_header_labels" />
  <참조목록>
    <참조범위 시트="Sheet1" 주소="$A$1" id="r1">
      <참조셀 주소="$A$1" 값="10" 값종류="number" />
    </참조범위>
    <참조범위 시트="Sheet1" 주소="$B$1" id="r2">
      <참조셀 주소="$B$1" 값="20" 값종류="number" />
    </참조범위>
    <참조범위 시트="Sheet1" 주소="$C$1" id="r3">
      <참조셀 주소="$C$1" 값="30" 값종류="number" />
    </참조범위>
  </참조목록>
  <행 인덱스="1">
    <셀 주소="$A$1" 값="10" 값종류="number" />
    <셀 주소="$B$1" 값="20" 값종류="number" />
    <셀 주소="$C$1" 값="30" 값종류="number" 수식="=A1+B1" 수식R1C1="=RC[-2]+RC[-1]" 값대입수식="=10+20" 값대입수식동등성="보장안함">
      <참조 ref="r1" />
      <참조 ref="r2" />
    </셀>
  </행>
  <행 인덱스="2">
    <셀 주소="$A$2" 값종류="blank" />
    <셀 주소="$B$2" 값="5" 값종류="number" />
    <셀 주소="$C$2" 값="60" 값종류="number" 수식="=C1*2" 수식R1C1="=R[-1]C*2" 값대입수식="=30*2" 값대입수식동등성="보장안함">
      <참조 ref="r3" />
    </셀>
  </행>
</표범위>
```

- `값` is Excel’s current constant or calculated formula result. `값종류` distinguishes `blank`, `number`, `text`, `boolean`, and `error`, so a true blank, a formula returning an empty string, numeric `123`, and text `"123"` are not conflated. `수식` is the readable A1 expression, and `수식R1C1` exposes relative and absolute references.
- `값대입수식` is an LLM-oriented explanation generated only when every supported static A1 reference was read from the same stable snapshot with a known native type. It substitutes current values one level deep; it is not an Excel-runnable or generally calculation-equivalent formula, as stated by `값대입수식동등성="보장안함"`. The original `수식` and `값` remain authoritative.
- References with exactly matching sheets and addresses appear once in `<참조목록>` and link to each formula through `<참조 ref="…">`. Overlapping but differently addressed ranges remain separate, preserving each formula’s reference order. The `시트` attribute identifies cross-sheet references; `<참조셀>` entries are row-major. Repeated references benefit most, while context can make small tables longer.
- Simple column headers and row labels within the selection are linked to their source cells, with context marked **inferred or unconfirmed**. Complex multi-level headers, duplicate or blank labels, and ordinary text data may remain unconfirmed. No surrounding cells are read to find context.
- `계산모드` and `계산상태` capture Excel’s state at export time; Tabledown never triggers recalculation. On Windows, `done` maps to `계산결과상태="snapshot_stable"`, `calculating`/`pending` map to `calculation_incomplete`, and `unknown` maps to `freshness_unverified`. macOS reports `계산상태="unavailable"` and `계산결과상태="freshness_unverified"` because Excel’s macOS API does not expose calculation state. Even `snapshot_stable` proves only that two current snapshots matched, not that a manual-calculation workbook was freshly recalculated.
- `값기준="Excel현재원시값"` means this schema carries raw current values, not displayed strings. Formula XML intentionally omits display formatting and merge hierarchy and says so with `표시정보상태="미포함"` and `병합정보상태="미포함"`. Use the separate **Copy as XML** command when you need formatted dates/currency/percentages or exact merges.
- Native numeric values are serialized as ordinary decimal text such as `315000`, not scientific notation such as `3.15E+5`; literal text that merely looks numeric is preserved exactly.
- `INDIRECT`, `OFFSET`, defined names, structured references, 3-D references, external workbooks, read failures, and over-limit references are not guessed. The formula and current result remain intact, with `참조상태="일부"`, `참조포함범위수`, and `참조누락이유` explaining partial reference data. Reason codes are `dynamic_reference`, `calculated_range_reference`, `external_or_structured_reference`, `three_dimensional_reference`, `whole_row_or_column_reference`, `defined_name_or_unsupported_syntax`, `invalid_a1_reference`, `range_count_limit`, `cell_count_limit`, `range_size_limit`, `read_failed`, `value_size_limit`, and `unspecified`.
- A brief notice appears after copying if some reference values are missing or Excel is still calculating. Original formulas and current results remain in the copy. An unavailable calculation status alone does not trigger a notice. Computed ranges such as `A1:INDEX(...)` are marked as partial instead of guessing their boundaries.
- Tabledown does not calculate or execute formulas; it reads Excel’s current values. Blank cells remain in the grid without a `값` attribute.
- This requires the Excel desktop app. Google Sheets and LibreOffice are not supported.
- Discontiguous multi-area selections are rejected, and a selection is limited to 10,000 cells.
- On macOS, formula cells split across more than 64 separate rectangular blocks must be copied in smaller selections.
- Combined A1 and R1C1 formula text is limited to 1,000,000 characters per export. Derived `값대입수식` text has a separate 1,000,000-character limit. A formula cell omitted by that character budget records `값대입수식상태="omitted_character_limit"` on `<셀>`; if the final XML byte limit requires removing all derived expressions, the root records `값대입수식상태="omitted_xml_size_limit"`. The core export remains intact in both cases.
- Direct reference values are limited to 256 **unique ranges**, 10,000 cells total, and 2,048 cells per range. Larger references keep the existing formula export available and are marked `참조상태="일부"`.
- Cell values are limited to 5,000,000 characters in total, and the final UTF-8 XML is limited to 10 MB. Rendering stops as soon as the bound is exceeded. If omitting the derived `값대입수식` is still insufficient, Tabledown fails the export and preserves the clipboard instead of silently dropping value kinds, calculation state, or partial-reference reasons.
- On macOS, first use of an XML command may request Automation permission to read the selected range from Excel. Existing automatic table conversion does not need that permission.

## How Automatic Conversion Works

Tabledown watches the clipboard and augments table content with the text/html formats needed for the paste direction. Excel tables are converted to Markdown plain text. Markdown tables are augmented with an HTML table that Excel can read.

Existing clipboard formats are preserved, and only the required text/html formats are added or updated. Plain non-table text is left unchanged.

Rendered formats such as PNG, PDF, and RTF that come with copied Excel tables are removed because Markdown and chat apps may otherwise choose them and paste the table as an image. Excel native formats are preserved.

If the copied Excel range has a trailing column where every row is empty, that column is removed during Markdown conversion. Middle empty cells needed for merged-cell alignment are kept.

Markdown plain text generated from Excel tables is padded with a leading blank line so Markdown parsers can recognize it as a standalone block. This helps Obsidian recognize a table even when it is pasted directly after a paragraph or list.

Clipboard items augmented by Tabledown include `org.nspasteboard.AutoGeneratedType` and `com.tabledown.generated` markers. Clipboard history managers such as Maccy can use these markers to ignore automatically generated clipboard entries.

Tabledown does not convert clipboard items that already contain its generated marker. This prevents repeated conversion and keeps the watcher idempotent.

## Development Install

### 1. Install Dependencies

Requires Python 3.10 or later (the code uses runtime PEP 604 union syntax such as `str | None`).

```bash
cd Tabledown
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Run

```bash
python run.py
```

A 3x2 table icon appears in the macOS menu bar.

### 3. macOS Permissions

Basic table conversion does not require Accessibility or Input Monitoring permissions. “Copy Markdown,” “Copy as XML,” and “Copy as XML with Formulas” read the current Excel selection directly, so they may request Microsoft Excel Automation permission on first use.

## Usage

1. Select cells in Excel or Google Sheets, then press `Cmd+C`
2. Paste into Obsidian, a GitHub README, or a Markdown editor with `Cmd+V`
3. Confirm that the result is a `| ... |` Markdown table

Use the paste result or the diagnostic log to confirm behavior.

Reverse direction:

1. Select a full Markdown table and press `Cmd+C`
2. Select the starting cell in Excel and press `Cmd+V`
3. Confirm that the table is split into cells

To pause or resume auto-conversion, press the global shortcut **⌘⌃T** (or use the menu bar icon ▸ “Use Tabledown”). A slash through the menu bar icon means it is off.

## Troubleshooting

If the paste result is not what you expected, check whether Tabledown is running and whether the diagnostic log shows a conversion event.

If you paste immediately after copying, the original clipboard item may be pasted before the app has augmented it. Processing usually takes around 0.1 seconds.

If an Excel table pastes into Obsidian as pipe text instead of a table, make sure the latest app is running. In the Excel to Markdown path, Tabledown keeps the HTML table clipboard format and augments the plain-text slot with Markdown (since 0.2.4 — only rendered formats such as PNG, PDF, and RTF are removed).

Conversely, if a Markdown editor pastes a rich table instead of Markdown source, an editor with HTML auto-conversion enabled (such as Obsidian) took the kept HTML slot first. That is destination-app policy outside Tabledown's control, and the trade-off for never losing the table format when the same copy is pasted back into Excel or Word. In that case, use the editor's "Paste as plain text" command to get the Markdown from the text slot.

Chat input fields such as Codex or Claude, and some plain text editors, may show Markdown tables as raw `| ... |` text instead of a rendered table. In that case, test with baseline apps such as Obsidian, GitHub Markdown preview, Excel, or Google Sheets.

If Maccy still records converted clipboard items, check whether `org.nspasteboard.AutoGeneratedType` or `com.tabledown.generated` is included in Maccy's Ignore Pasteboard Types setting.

Diagnostic logs are written here:

```bash
tail -f ~/Library/Logs/Tabledown.log
```

When the diagnostic log exceeds 1 MB, it is rotated to `Tabledown.log.1` and a new log file is started.

## Privacy Policy

Tabledown does not collect, store, sell, or share personal information.

When you explicitly choose “Copy Markdown” or “Copy as XML,” Tabledown locally reads values, blanks, and merge structure from the current Excel selection. “Copy as XML with Formulas” reads values, blanks, formulas, and addresses from the selection plus current values from direct static A1 references in the same workbook. These commands write Markdown+HTML or XML to the same clipboard; cell values and formulas are never written to the diagnostic log or sent to an external server.

The app reads the current macOS clipboard locally and writes the text/html formats needed for table conversion back to the same clipboard. Conversion happens only on the user's Mac and is not sent to an external server.

Tabledown does not use account creation, analytics, ad tracking, location data, contacts, photos, or file uploads.

Diagnostic logs are stored only on the user's Mac at `~/Library/Logs/Tabledown.log` for behavior checks. Logs are not sent externally and can be deleted by the user.

## Changelog

- 2026-09-10: **macOS 0.6.1 / TestFlight build 0.6.5 is available to internal group `22`.** The three copy commands passed 375 candidate checks; 27 Windows-only tests were skipped. App and installer signatures, Apple validation and upload, VALID processing, and IN_BETA_TESTING were verified. Korean and English build notes explain how to move tables into documents and ask AI about data; the existing app description was preserved. No App Store review submission or public GitHub release was made in this step.

- 2026-09-10 (development source, unreleased): Unified the macOS actions as “Copy Markdown,” “Copy as XML,” and “Copy as XML with Formulas.” Formula XML now includes inferred context and shared references. Added direct Markdown copying and improved special-character and line-break round trips. Defaults, existing shortcuts, and automatic copy/paste remain available.

- 2026-09-09: **macOS 0.6.0 release (build 0.6.4).** The Apple Silicon DMG and ZIP passed signature, Apple notarization, and Gatekeeper checks; the frozen candidate passed 223 automated checks. Published [v0.6.0 installers](https://github.com/yooongZa/tabledown/releases/tag/v0.6.0) as GitHub Latest and verified the downloaded DMG/ZIP sizes and SHA-256 hashes, plus the permanent DMG link. The final check on 2026-09-09 showed the App Store version waiting for review, configured to release automatically after approval. The approved Korean description is preserved verbatim. No Windows installer is released in this update.

- 2026-09-09: **Uploaded macOS 0.6.0 / TestFlight build 0.6.4.** Includes the new AI copy option to bring headers, item names, and calculation values together, reduce repeated details, and fix Korean tables pasted into Apple Notes. Korean and English TestFlight copy covers the whole app: moving tables into notes, continuing edits in Excel, preparing complex tables for explanations, and keeping familiar copy and paste. The final headline is “Get more out of your AI,” with automatic Markdown conversion distinguished from Excel menu-based copying. The isolated candidate passed 223 automated checks plus package and Apple validation. Apple processing completed, the build is available to the existing internal test group, and both languages were saved and verified.

- 2026-09-09: **Fixed corrupted Korean text when pasting Markdown tables into Apple Notes.** Generated HTML now explicitly declares UTF-8. Original Markdown and existing Excel/web table HTML are preserved. The reinstalled macOS app passed real Notes and Excel paste checks for eight cells containing Korean, emoji, and special characters, plus Excel → Markdown copy-back. Automated checks passed: 1 macOS encoding test, 51 converter tests, and 80 Windows-port tests with 27 skipped on macOS. This change has not been tested on a physical Windows device.
- 2026-07-17: **Changed general Table→XML to direct Excel selection.** Select a table range in Excel and invoke “Copy table structure and displayed values as XML” or `⌘⌃X`; Tabledown now reads formatted values, error values, significant whitespace, blanks, and merge structure without `Cmd+C`. Clipboard fallback was removed to prevent exporting a stale copy, and only two matching snapshots are exported.
- 2026-07-13: **Added full Excel table export with values, blanks, and formulas as XML (macOS 0.6.0).** Select one rectangular range containing formulas in the Excel desktop app, then use “Copy cell values, formulas, and references as XML,” `⌘⌃E` on macOS, or `Ctrl+Alt+E` on Windows. Tabledown copies every cell address and current value, preserves blanks, and adds A1/R1C1 formulas to formula cells in a `<표범위>` XML document. True blanks, `=""` results, and Excel error values remain distinct; selection changes, discontiguous ranges, and size-limit failures are rejected fail-closed. macOS 0.6.0 was uploaded as **TestFlight build 0.6.1** and reached “Ready to Submit” in internal group `22`. All 39 focused tests and 75/75 test-matrix cases passed.
- 2026-07-01: **“Auto-fill blank cells” now applies to Markdown conversion too (0.5.0).** When converting an Excel/Sheets table to Markdown, the blanks left by merged cells are filled if the “Auto-fill blank cells” setting (off by default) is on — a group-header band (`1분기`) spreads right, a left key column (`부장`) spreads down. Only the **header frame** is filled; blanks in the value (data) region are preserved (same guard as the XML fill). The flattened multi-level header (a leaf header row demoted into the body) is left as-is, since Markdown table syntax can't draw a merge. This toggle now drives **both** the XML and Markdown paths (dropped the “XML:” prefix from the old label). With the toggle off, behavior is unchanged.
- 2026-07-01: **Windows: ported the 0.5.0 “Auto-fill blank cells” toggle to the tray** (parity with macOS, 0.2.7). The merged-header fill logic already lived in the shared `tablemark.converter.html_to_md` that Windows imports; what was missing was the toggle, the setting, and the wiring (Windows never had the `fill_blanks` option — macOS shipped it only on the XML path, which is macOS-only). Now `conversion.py` passes the `fill_blanks` flag into the Markdown conversion (`html_table_to_markdown`/`convert_document_tables`) — default off, so behavior is unchanged and the HTML slot is still preserved — and the tray gains a checkable “Auto-fill blank cells” menu item (persisted, placed after the toggle and before Language, mirroring the macOS Settings order). No “XML:” prefix, since Windows has no XML path. Regression tests: off keeps blanks, on fills the merged key column (and keeps HTML), label translations, menu inclusion, toggle persistence
- 2026-06-30: **Renamed the XML menu item “Copy table as XML” → “Convert copied table to XML” (0.4.2).** Same behavior, but the verb “Convert” describes what actually happens, and “copied table” names the precondition (copy a table first), addressing the most common first-time confusion. Help text, tooltip, and usage docs updated to match.
- 2026-06-29: **Changed the Table→XML global shortcut from ⌘⌃C to ⌘⌃X (0.4.1).** X = **X**ML, a clearer mnemonic. It avoids macOS system and app shortcuts and sits far from the toggle ⌘⌃T to reduce misfires. The conversion behavior is unchanged (table→XML copy, HTML dropped). Help text, menu display, and the regression test were updated.
- 2026-06-24: **Added a global shortcut to toggle auto-conversion on/off — macOS ⌘⌃T (Windows Ctrl+Alt+T) (0.4.0).** Pause or resume clipboard auto-conversion from any app with one key. No permissions required; if registration fails it falls back gracefully to the menu item. macOS generalizes the former single hotkey into a multi-hotkey manager (one Carbon handler dispatches ⌘⌃C/⌘⌃T by the fired hotkey id), so XML (⌘⌃C) and the toggle (⌘⌃T) coexist. Windows uses user32 `RegisterHotKey` + a dedicated message-loop thread. The ⌘⌃ / `Ctrl+Alt` combos were chosen to avoid common system and app shortcuts.
- 2026-06-19: **Added local diagnostics — crash capture + “Open logs for bug report” (nothing sent off-device) (0.4.0).** Surfaces failures that were invisible in direct distribution into the local log — there's no network code at all, so the “no network connections, no telemetry … nothing is sent to any external server” promise still holds. `sys.excepthook` + **`threading.excepthook`** (catches the clipboard-watcher daemon-thread crashes that used to vanish silently) + `faulthandler` (native faults) write only to `Tabledown.log`/`.crash`. The “Open logs for bug report” menu writes a **scrubbed** diagnostics file (home paths, usernames, volume/drive names, and secrets removed) and reveals it in Finder/Explorer for you to attach to a bug report — no clipboard access (so it can't race the watcher). Conversion error messages and logs were also hardened so table data can't leak (macOS + Windows). The DiskOUT-style *remote* collection was deliberately not adopted, since it would conflict with the zero-telemetry promise.
- 2026-06-19: **Added feedback for auto-conversion (0.4.0).** The Excel↔Markdown auto-conversion used to happen silently in the background, so you couldn't tell it ran (and an unexpected paste gave no hint Tabledown was the cause). Now the menu bar icon **briefly flashes a checkmark (0.5s, shorter than the 1s manual-XML flash)** whenever a table is converted. No popup, sound, or system notification (keeps the zero-permissions design).
- 2026-06-19: **Removed monetization — the app is now fully free (0.4.0).** Dropped the Pro-subscription gate on XML conversion (menu and ⌘⌃C), and removed the donation IAPs, “Restore Purchases”, the subscription sheet, and the 🔒 lock marker (`store.py` deleted; related code/strings/dependency cleaned out of `settings.py`, `i18n.py`, `setup.py`, `requirements.txt`). The global hotkey and XML conversion logic are unchanged — only the gate is gone, so everyone uses XML for free. No impact on the clipboard invariants or tests (38/38).
- 2026-06-18: **Windows: fixed multiple tray-app instances — added a single-instance guard** (0.2.6). On macOS, LaunchServices blocks a second `.app` launch, so single-instance is free; Windows blocks nothing, so a manual launch plus the login StartupTask, double-click duplicates, and ghost processes left by a crash each spawned another tray icon + clipboard watcher — two watchers can overwrite each other's clipboard and break conversion. `main()` now guards with a named mutex (`CreateMutexW("Local\TabledownSingleInstance")`) before creating the app and quietly exits a second instance (session-scoped, so fast user switching still allows one per user). The kernel32 access is deferred into a function so imports don't break on non-Windows test runners (which run without the guard), with first/second/no-kernel regression tests
- 2026-06-17: **Windows: fixed the core bug where Excel/Sheets tables didn't convert to Markdown** (0.2.5). Excel places the CF_HTML fragment markers (`StartFragment`/`EndFragment`) *inside* the `<table>` element, so the extracted HTML is missing the `<table>` tag → table detection fails → copying a real Excel table left the raw TSV unconverted. `extract_cf_html` now wraps a bare table-row fragment in `<table>` (root-caused and verified by capturing the real Windows clipboard; added an Excel-format CF_HTML regression test).
- 2026-06-17: Added an **“Open at Login” toggle to the Windows tray** (parity with macOS, 0.2.5). It flips the MSIX manifest's `windows.startupTask` via the WinRT `StartupTask` API (`winsdk`), and the OS keeps the state in Settings ▸ Apps ▸ Startup (no JSON shadow). The item is hidden on source/dev runs and the bare non-MSIX exe, where there's no package identity (graceful fallback). If the user disabled it in Task Manager (`disabled_by_user`), the checkmark stays off and a hint explains why. Also **fixed the Help window not closing** — the modal `MessageBox` blocked pystray's pump thread and, lacking foreground rights, opened behind the active window, so re-clicking stacked boxes; now it runs on its own thread with a single-instance guard and foreground/topmost flags
- 2026-06-15: First real run/verification of the Windows port on actual Windows, plus Microsoft Store release prep. **Fixed a startup crash** where the tray died immediately — the language-menu callback was a 3-arg lambda pystray rejects (the default arg counts toward the arg count) → replaced with a 2-arg closure factory (`_language_action`). Built and verified the full-trust MSIX packaging end to end (PyInstaller → tile assets → makeappx → self-sign; the frozen app's tray and clipboard conversion confirmed working). Added a privacy policy (`windows/PRIVACY.md`) and store-listing copy (`windows/STORE_LISTING.md`) for submission
- 2026-06-10: UX pass 1–2 — feedback for previously-silent flows. A 1-second checkmark flash on the menu bar icon when “Copy table as XML” (menu or ⌘⌃C) succeeds (an icon flash instead of a system notification — keeps the zero-permissions design), purchase-flow alerts (subscribe success / purchase failure / restore outcome — user cancellation stays silent), localized store prices on the donation items and in the subscription sheet (the hardcoded price is now only a pre-fetch fallback), and a 🔒 marker on “Copy table as XML” while Pro is inactive. One-time first-run welcome (a menu-bar/tray-only app otherwise looks like nothing happened after install), and Help now shows the running version, the ⌘⌃C shortcut, and an “Open GitHub” button. Menu reordered (Settings above donations; Restore Purchases moved to the bottom of the donation submenu). Windows: fixed-label “Use Tabledown” + checkmark toggle, a red-slash tray icon when off, the tray icon rendered at the exact DPI size (no more softening), a read-modify-write JSON settings store, and the same first-run welcome. The conversion toggle is deliberately non-persistent (starts enabled every launch — policy documented)
- 2026-06-08: Redesigned XML conversion to a **nested hierarchy** preserving multi-level headers on both axes (`<표>` ▸ `<직급그룹 이름>` ▸ `<행 직책>` ▸ `<열그룹 이름>` ▸ `<열 n>`). Vertical groups now nest as parent nodes (previously repeated per row; tradeoff: rows are no longer self-contained). 38/38 converter tests. (Separately: partial-monetization scaffolding — donation IAPs, XML Pro yearly subscription, ⌘⌃C global hotkey; incomplete, pending App Store Connect setup)
- 2026-06-08: Added XML table conversion (0.3.0). An LLM-friendly record-style XML with structural tags — root `<dataset>`, each value a `<cell name="…">`, multi-level group headers nested as `<group>`. Headers live in attributes (not tag names), so spaces/symbols/digits never mangle a tag and any standard XML parser reads it; the `<dataset>` root (not `<table>`) keeps the content intact even where it is rendered as HTML. The “Copy table as XML” menu turns the clipboard table (Excel/Markdown/XML) into XML — **click-only**, with no automatic XML→table direction (so the watcher never touches ordinary config/document XML). Recognizes Excel merged cells (vertical rowspan filled per row; horizontal colspan group headers nested as `<group>`). A “XML: Auto-fill blank cells” setting (off by default) fills blanks in unmerged group columns from the value above/left (value columns preserved). Blank-fill / language / login-item are grouped under a “Settings” submenu
- 2026-05-29: Submitted Tabledown 0.2.4 to the Mac App Store (TestFlight) as build 0.2.4, and attached the notarized DMG/zip to the GitHub Release (v0.2.4, Latest)
- 2026-05-29: Unified HTML-slot handling — a bare Excel/Sheets table now keeps its HTML `<table>` slot too (0.2.4). Every table case (web table, document, Excel table) now keeps HTML and augments the text slot with Markdown, so one copy gives Excel/Word a real table and Markdown editors Markdown
- 2026-05-29: Submitted Tabledown 0.2.3 to the Mac App Store (TestFlight) as build 0.2.3
- 2026-05-29: When pasting a document that partly contains a table, the text slot now gains a Markdown table for the table portion while the original HTML `<table>` slot is preserved (0.2.3). Markdown editors read the text and get a Markdown table; Word/Excel read the original HTML and paste a real table. Headings, paragraphs, and lists outside the table stay as plain text
- 2026-05-29: Fixed a bug (0.2.2) where copying a document that only partly contains a table (web page, chat answer, Word, etc.) pasted just the table and dropped the surrounding text (headings, paragraphs, lists). When the clipboard HTML has meaningful content beyond the `<table>`, it is treated as a document and left untouched; only a bare Excel/Sheets table is converted to Markdown
- 2026-05-29: Submitted Tabledown 0.2.1 to the Mac App Store (TestFlight). App Store Connect requires the build number (`CFBundleVersion`) to be higher than the previous upload (0.2.0's build 0.2.1) regardless of marketing version, so the build was bumped to 0.2.2 (`TABLEDOWN_BUILD=0.2.2`)
- 2026-05-29: Fixed a regression (0.2.1) where tables copied from web or chat apps (e.g. Claude) pasted into Excel as raw Markdown in a single cell. Such tables carry both Markdown text and an HTML `<table>` on the clipboard; the 0.2.0 `is_markdown_table` cell-count check rejected them when the separator column count was off, so the HTML was stripped and the text converted to Markdown. When an HTML `<table>` is present, the cell-count check is now skipped (`strict` parameter) and the original clipboard is preserved. The check still guards html-less text to block false positives
- 2026-05-28: Submitted Tabledown 0.2.0 (build 0.2.1) to the Mac App Store for review
- 2026-05-28: Standardized the toggle menu to macOS HIG conventions. A single "Use Tabledown" item with a state checkmark replaces the old "Enabled ✓" / "Disabled" labels, and the menu bar icon now shows a slash overlay when conversion is off, so the on/off state is visible without opening the menu. The login-at-startup and language items use the same checkmark pattern
- 2026-05-28: Preserve cell-internal line breaks (Excel Alt+Enter, Sheets Ctrl+Enter) as `<br>` instead of collapsing them to spaces, so Obsidian and GitHub Flavored Markdown render them correctly
- 2026-05-28: Reduced `is_markdown_table` false positives by requiring the header and separator rows to have matching cell counts, so ordinary text whose second line happens to start with `-` is no longer mistaken for a table
- 2026-05-28: Fixed the Mac App Store build. The py2app nested executable (`Contents/MacOS/python`) now receives sandbox `inherit` entitlements instead of an `application-identifier` (resolves App Store Connect error 90885), and `CFBundleVersion` is split out via the `TABLEDOWN_BUILD` env var so the build number can be bumped while keeping the marketing version (`0.2.0`), avoiding error -19232 on re-upload. The build number must be at most three period-separated integers, so `0.2.1` (three components) is used instead of `0.2.0.1` (four components, rejected with error 236550)
- 2026-05-19: Replaced the Mac App Store screenshots with real app UI captures. The first submission was rejected under Guideline 2.3.3 for using marketing/promotional materials; the four mockup graphics were removed and replaced with raw captures of the menu bar drop-down, the language submenu, and the conversion result alongside Numbers and TextEdit
- 2026-05-19: Removed the "Hide menu bar icon" menu action. In an `LSUIElement: True` app, the menu bar icon was the only UI, so hiding it left no way to even quit the app. Worse, `NSStatusItem`'s autosaveName persisted the visibility flag to `NSUserDefaults`, so relaunching the app could not bring the icon back. On startup the app now strips any stale `NSStatusItem Visible*` keys from `NSUserDefaults`, restoring the icon automatically for anyone affected by the previous build
- 2026-05-15: Added an "Open at Login" toggle backed by macOS 13+ `SMAppService.mainAppService`, which works under App Sandbox and Mac App Store builds. The menu item is hidden on macOS 12 (graceful fallback)
- 2026-05-15: Localized the menu bar UI and help dialog into Korean and English. Auto-detects the macOS system language and lets users switch manually via the "Language / 언어" submenu; the choice is persisted in `NSUserDefaults`
- 2026-05-11: Added a README privacy policy for App Store submission
- 2026-05-11: Added a README comparison screenshot showing the paste difference between Tabledown enabled and disabled
- 2026-05-11: Added an English README and connected language links between the Korean and English docs
- 2026-05-11: Added public GitHub repository download instructions, introduction copy, release badges, and install instructions
- 2026-05-11: Documented the difference between Markdown source conversion when Tabledown is enabled and rich text HTML table paste behavior when disabled
- 2026-05-11: Added block spacing to Markdown plain text generated from Excel tables and removed HTML table clipboard formats for more reliable Obsidian table recognition after paragraphs or lists
- 2026-05-11: Clarified troubleshooting and known limitations for cases where Obsidian shows pipe text
- 2026-05-11: Applied `COPYFILE_DISABLE=1` in the release build script to reduce macOS xattr-related codesign verification failures
- 2026-05-10: Detect Tabledown-generated clipboard markers to avoid converting the app's own output
- 2026-05-10: Adjusted changeCount handling so the watcher does not miss external clipboard changes immediately after conversion
- 2026-05-10: Added 1 MB diagnostic log rotation
- 2026-05-10: Removed unused global hotkey code and the Quartz dependency
- 2026-05-10: Centralized package versioning through `tablemark.__version__`
- 2026-05-10: Split build dependencies into `requirements-build.txt` and `pyproject.toml`, and added a release build script
- 2026-05-08: Removed rendered formats when copying Excel tables to reduce cases where chat or Markdown apps paste a PNG
- 2026-05-08: Trimmed trailing empty columns when converting copied Excel ranges to Markdown
- 2026-05-07: Added pasteboard markers so clipboard history managers such as Maccy can identify generated entries
- 2026-05-07: Renamed the app to Tabledown and updated bundle metadata, app bundle names, and log paths
- 2026-05-07: Replaced the `Cmd+Ctrl+M` global shortcut workflow with clipboard watching
- 2026-05-07: Added Markdown plain text when copying Excel tables so Markdown editors can paste them directly
- 2026-05-07: Added HTML table output when copying Markdown tables so Excel can paste them as cells
- 2026-05-07: Preserved Excel native formats for normal Excel to Excel paste compatibility
- 2026-05-07: Expanded Excel merged cells with empty cells to preserve row and column alignment in Markdown
- 2026-05-07: Added a 40 px retina template menu bar icon based on a 3x2 table motif
- 2026-05-07: Removed Accessibility and Input Monitoring permission requirements for basic usage
- 2026-05-06: Removed the crash-prone `pynput` keyboard hook and cleaned up the Quartz-based implementation

## App Build

```bash
pip install -r requirements-build.txt
scripts/build_release.sh
```

This creates `dist/Tabledown.app` and `dist/Tabledown.zip`. Public distribution uses the notarized `Tabledown.dmg` attached to the GitHub Release.

## Project Structure

```
Tabledown/
├── run.py                      # Entry point (macOS)
├── setup.py                    # py2app build settings
├── requirements.txt
├── requirements-build.txt      # Release build dependencies
├── pyproject.toml              # Build-system settings
├── scripts/                    # Test and release scripts
├── assets/                     # Menu bar and app icons
├── tablemark/                  # macOS app (import path stays tablemark)
│   ├── app.py                  # Menu bar app core (rumps)
│   ├── clipboard.py            # NSPasteboard wrapper and changeCount handling
│   ├── settings.py             # Settings persistence (NSUserDefaults)
│   ├── i18n.py                 # Korean/English localization
│   ├── hotkey.py               # Global hotkeys (Carbon)
│   ├── login_item.py           # Open at Login (SMAppService)
│   ├── diagnostics.py          # Local crash capture and diagnostics file
│   ├── logger.py               # Diagnostic logging
│   └── converter/
│       ├── html_to_md.py       # Excel HTML to Markdown
│       ├── md_to_tsv.py        # Markdown to TSV/HTML table
│       └── table_xml.py        # Table ↔ LLM-friendly XML
└── windows/                    # Windows tray port (independent version track)
    ├── run_windows.py          # Windows entry point
    ├── tabledown_windows/      # Tray app + clipboard conversion (pystray)
    └── tests/                  # Windows port tests
```

The internal Python package path remains `tablemark/` for import compatibility.

## Known Limitations

- Merged cells: Markdown cannot represent the merged shape, but empty cells are added to preserve row and column alignment
- Line breaks inside cells: replaced with spaces
- Single-row tables: converted as header-only tables
- Markdown table to Excel conversion adds an HTML table format for Excel compatibility
- Paste as Picture after copying an Excel table is not supported
- Converted entries may still appear in clipboard history depending on whether the clipboard history manager supports and respects pasteboard markers

## License

[MIT](LICENSE)
