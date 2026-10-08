# Tabledown

<img src="assets/generated/tablemark_app_1024.png" width="80" alt="Tabledown icon">

**Convert Excel and Google Sheets tables ↔ Markdown with copy and paste.**

Tabledown runs in the macOS menu bar or Windows system tray. You can also copy Excel table structure or formulas as XML to share with AI.

[한국어](README.md) · English

## Download

| Platform | Store | Direct download |
| --- | --- | --- |
| macOS 12 or later · Apple Silicon | [Mac App Store](https://apps.apple.com/app/id6768205551) | [DMG](https://github.com/yooongZa/tabledown/releases/latest/download/Tabledown.dmg) |
| Windows 10/11 · x64 | [Microsoft Store](https://apps.microsoft.com/detail/9NGS4C0N2Z6L) | [Windows ZIP](https://github.com/yooongZa/tabledown/releases/tag/windows-v0.3.2) |

On Mac, open the DMG and drag the app into Applications. The download is signed with Developer ID and notarized by Apple.
On Windows, extract the ZIP and run `Tabledown-Windows.exe`. No Python installation is needed. For the optional self-signed MSIX, see the [Windows guide](windows/README.md).

## Usage

1. Launch Tabledown and find the table icon in the menu bar or system tray.
2. Copy a table from Excel or Google Sheets.
3. Paste into Obsidian or a Markdown document. To go back, copy a Markdown table and paste it into Excel cells.

Use `Cmd+C` / `Cmd+V` on Mac or `Ctrl+C` / `Ctrl+V` on Windows. There is no separate Copy Markdown menu. The original HTML table is retained for Excel and Word. To get Markdown source, use your destination app’s **Paste as Plain Text** command.

## Share Excel tables with AI

In the Excel desktop app, select **one rectangular range including headings and item names**, then choose an XML copy command. You do not need to copy first. Paste the complete XML into your preferred AI yourself.

| Export | Included information | Mac | Windows |
| --- | --- | --- | --- |
| General XML | Displayed values, blanks, merge structure, and source | `⌘⌃X` | `Ctrl+Alt+X` |
| Formula XML | Current values, formulas, direct reference values, and calculation state | `⌘⌃E` | `Ctrl+Alt+E` |

On Mac, **Copy as XML with Formulas** puts repeated references in a shared list. Windows offers separate **Copy cell values, formulas, and references as XML** and **Copy compact XML for AI** commands.

XML exports support up to **10,000 selected cells**, **5,000,000 cell-value characters**, and **10 MB of output**. Formula XML reads Excel’s current values without recalculating. Use general XML for displayed formatting and merge structure. Unsupported or unreadable references are marked as partial.

## Settings

- **Conversion on/off:** `⌘⌃T` on Mac or `Ctrl+Alt+T` on Windows. A slash across the icon means conversion is off.
- **Fill group and category blanks:** Off by default. Automatic Markdown fills merged column-group headings and left category columns; general XML fills left category columns. Original Excel cells, original HTML, and formula XML are unchanged.
- **Language and startup:** Choose Korean or English and Open at Login. Windows startup support requires an installed MSIX package.

<details>
<summary>How blank filling chooses cells</summary>

Category columns are inferred as those before the first body column with no blanks, starting from the left. Automatic Markdown fills column-group headings horizontally and category columns vertically. General XML fills category blanks from above, then fills remaining blanks from the left. Body blanks in that first complete column and columns to its right are preserved. General XML records the original A1 addresses of filled cells.

</details>

## Troubleshooting

- **The original table pastes before conversion:** Check that Tabledown is enabled, then wait briefly after copying before pasting.
- **XML copying fails:** Select one rectangular Excel range. On Mac, allow Tabledown to access Excel in Automation settings when prompted. Basic table conversion needs no Accessibility or Input Monitoring permission.
- **A large table takes too long or shows `###`:** Split the selection or widen the Excel columns. Large selections on Mac can take tens of seconds.

Markdown cannot represent colors, fonts, or merged-cell shapes. XML copying requires desktop Excel; Google Sheets and LibreOffice are not supported for these commands.

Use **Open logs for bug report** in the menu to inspect diagnostics. [Report an issue](https://github.com/yooongZa/tabledown/issues) · [Changelog and verification records](CHANGELOG.md)

## Privacy Policy

All conversion happens on your device. Tabledown does not collect, sell, or share personal data. It has no account, ads, usage tracking, or external-server transmission.

Automatic conversion reads the current clipboard and writes results to the same clipboard. XML copying reads the selected Excel cells. **Formula copying can also include direct A1 reference values outside the selection or on other sheets in the same workbook.** It does not open external workbooks or follow references recursively.

Cell values and formulas are not written to diagnostic logs. The Mac log is stored locally at `~/Library/Logs/Tabledown.log` and can be deleted. [Windows privacy policy](windows/PRIVACY.md)

<details>
<summary>Developer guide</summary>

On Mac, use Python 3.10 or later and run these commands from the repository root.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python run.py
```

The Mac app lives in `tablemark/`, shared conversion code in `tablemark/converter/`, and the Windows app in `windows/`.

[XML format](docs/xml-format-spec.md) · [Development, testing, and release guide](AGENTS.md) · [Mac build](scripts/build_release.sh) · [Windows development and builds](windows/README.md#개발)

</details>

[MIT License](LICENSE)
