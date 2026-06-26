#!/usr/bin/env python3
"""Create and run Tabledown consistency test fixtures.

The matrix is intentionally split into:
- pure converter tests that do not touch the system clipboard
- direct NSPasteboard read/write tests
- optional watcher tests against a running Tabledown app
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT_DIR = ROOT / "outputs" / "tabledown_test_envs"

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString

from tablemark.clipboard import (
    GENERATED_MARKER_TYPES,
    HTML_TYPES,
    LEGACY_HTML_TYPE,
    RENDERED_TABLE_TYPES,
    TABLEDOWN_GENERATED_TYPE,
    read_clipboard,
    write_clipboard,
)
from tablemark.converter.html_to_md import (
    forward_fill_key_columns,
    html_table_to_markdown,
    html_table_to_model,
    html_table_to_rows,
)
from tablemark.converter.md_to_tsv import markdown_table_to_html, markdown_table_to_rows
from tablemark.converter.table_xml import (
    is_table_xml,
    model_to_xml,
    table_xml_to_markdown,
    table_xml_to_model,
)
from tablemark import login_item
from tablemark.hotkey import CMD, CONTROL, KEY_C, KEY_T, GlobalHotkeys
from tablemark.i18n import SUPPORTED_LANGUAGES, detect_system_language, t
from tablemark.app import TabledownApp


MARKDOWN_BASIC = "| Name | Score | Status |\n| --- | --- | --- |\n| Alice | 95 | Pass |\n| Bob | 82 | Pass |"
MARKDOWN_KOREAN = "| 이름 | 값 | 메모 |\n| --- | --- | --- |\n| 사과 | 1200 | A\\|B |\n| 배 | 980 | 공백 포함 |"
PLAIN_TEXT = "This is plain text, not a table."

HTML_BASIC = "<table><tr><th>Name</th><th>Score</th></tr><tr><td>Alice</td><td>95</td></tr></table>"
HTML_COLSPAN = "<table><tr><th colspan='2'>Group</th><th>Total</th></tr><tr><td>A</td><td>B</td><td>2</td></tr></table>"
HTML_ROWSPAN = "<table><tr><th rowspan='2'>Name</th><th>Q1</th></tr><tr><td>10</td></tr></table>"
HTML_TRAILING_EMPTY = "<table><tr><th>Step</th><th>Source</th><th></th></tr><tr><td>1</td><td>Copy_Basic</td><td></td></tr></table>"
HTML_NOISE = """
<html><body>
<p>Before</p>
<table><tr><th>제품</th><th>수량</th></tr><tr><td>키보드</td><td>3</td></tr></table>
<p>After</p>
</body></html>
"""

XML_TABLE = (
    "<dataset>\n"
    '  <row>\n    <cell name="Name">Alice</cell>\n    <cell name="Score">95</cell>\n  </row>\n'
    '  <row>\n    <cell name="Name">Bob</cell>\n    <cell name="Score">82</cell>\n  </row>\n'
    "</dataset>"
)
# v2 nested format (current output): root <표>, rows <행>, horizontal leaves <열>.
XML_TABLE_V2 = (
    "<표>\n"
    '  <행>\n    <열 n="Name">Alice</열>\n    <열 n="Score">95</열>\n  </행>\n'
    '  <행>\n    <열 n="Name">Bob</열>\n    <열 n="Score">82</열>\n  </행>\n'
    "</표>"
)
# A v2 cross-table: vertical group <{header}그룹> + <행 {h}=> + horizontal <열그룹>/<열>.
XML_TABLE_V2_NESTED = (
    "<표>\n"
    '  <직급그룹 이름="부장">\n'
    '    <행 직책="대족장">\n'
    '      <열그룹 이름="1분기"><열 n="1">동</열><열 n="2">해</열></열그룹>\n'
    '    </행>\n'
    '  </직급그룹>\n'
    "</표>"
)
# A config-shaped XML that merely looks nested — must NOT be mistaken for a table.
XML_NOT_TABLE = (
    "<config>\n  <server>\n    <host>localhost</host>\n    <port>8080</port>\n  </server>\n</config>"
)

# A real Excel paste (all <td>): a full-width title row, a two-row group header
# (colspan), and a left column with vertical merges (rowspan).
HTML_MERGED_EXCEL = (
    "<table>"
    "<tr><td colspan='4'>2024 인사 평가</td></tr>"
    "<tr><td></td><td></td><td colspan='2'>매출</td></tr>"
    "<tr><td>직급</td><td>직책</td><td>1분기</td><td>2분기</td></tr>"
    "<tr><td rowspan='2'>부장</td><td>팀장</td><td>12</td><td>34</td></tr>"
    "<tr><td>파트장</td><td>56</td><td>78</td></tr>"
    "</table>"
)

# Grouping expressed by BLANK cells (no merge): 직급 is filled only on the first
# row of each group, the rest left empty meaning "same as above". The 점수 column
# carries a genuine blank that must NOT be filled.
HTML_BLANK_GROUP = (
    "<table>"
    "<tr><td>직급</td><td>직책</td><td>점수</td></tr>"
    "<tr><td>부장</td><td>대족장</td><td>10</td></tr>"
    "<tr><td></td><td>족장</td><td></td></tr>"
    "<tr><td></td><td>추장</td><td>30</td></tr>"
    "</table>"
)

# A left key cell whose column ABOVE is also blank must inherit from the LEFT
# (horizontal fill) — the colspan-origin analogue. Row 1's blank c2/c3 take "A".
HTML_HFILL = (
    "<table>"
    "<tr><td>c1</td><td>c2</td><td>c3</td><td>값</td></tr>"
    "<tr><td>A</td><td></td><td></td><td>1</td></tr>"
    "<tr><td></td><td>x</td><td>y</td><td>2</td></tr>"
    "</table>"
)


@dataclass
class TestResult:
    name: str
    ok: bool
    category: str
    detail: str = ""
    elapsed_ms: int = 0


def write_fixture_files(output_dir: Path) -> None:
    fixtures = {
        "markdown/basic.md": MARKDOWN_BASIC,
        "markdown/korean_escaped_pipe.md": MARKDOWN_KOREAN,
        "text/plain.txt": PLAIN_TEXT,
        "html/basic_table.html": HTML_BASIC,
        "html/colspan_table.html": HTML_COLSPAN,
        "html/rowspan_table.html": HTML_ROWSPAN,
        "html/trailing_empty_table.html": HTML_TRAILING_EMPTY,
        "html/noisy_table.html": HTML_NOISE,
    }
    for relative_path, content in fixtures.items():
        path = output_dir / relative_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content + "\n", encoding="utf-8")


def run_converter_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "converter", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001 - test harness should report exact failure
            tests.append(TestResult(name, False, "converter", str(exc), _elapsed(started)))

    check(
        "html_basic_to_markdown",
        lambda: _assert_equal(
            html_table_to_markdown(HTML_BASIC),
            "| Name | Score |\n| --- | --- |\n| Alice | 95 |",
        ),
    )
    check(
        "html_colspan_alignment",
        lambda: _assert_equal(
            html_table_to_markdown(HTML_COLSPAN),
            "| Group |   | Total |\n| --- | --- | --- |\n| A | B | 2 |",
        ),
    )
    check(
        "html_rowspan_alignment",
        lambda: _assert_equal(
            html_table_to_markdown(HTML_ROWSPAN),
            "| Name | Q1 |\n| --- | --- |\n|   | 10 |",
        ),
    )
    check(
        "html_trailing_empty_column_trimmed",
        lambda: _assert_equal(
            html_table_to_markdown(HTML_TRAILING_EMPTY),
            "| Step | Source |\n| --- | --- |\n| 1 | Copy_Basic |",
        ),
    )
    check(
        "markdown_escaped_pipe_rows",
        lambda: _assert_equal(
            markdown_table_to_rows(MARKDOWN_KOREAN)[1],
            ["사과", "1200", "A|B"],
        ),
    )
    check(
        "markdown_to_html_table",
        lambda: _assert_contains(markdown_table_to_html(MARKDOWN_BASIC), "<table><tr><th>Name</th>"),
    )
    check(
        "generated_clipboard_skipped",
        lambda: _assert_equal(
            TabledownApp._converted_clipboard(None, {"generated": True, "text": MARKDOWN_BASIC}),
            None,
        ),
    )
    check(
        "html_clipboard_markdown_is_block_spaced",
        lambda: _assert_equal(
            TabledownApp._converted_clipboard(
                None,
                {"html": HTML_BASIC, "text": "Name\tScore\nAlice\t95"},
            )["text"],
            "\n| Name | Score |\n| --- | --- |\n| Alice | 95 |\n",
        ),
    )
    check(
        "html_clipboard_keeps_html",
        lambda: _assert_equal(
            # a bare Excel/Sheets table now KEEPS the HTML slot so Excel/Word
            # still paste a real table; only RENDERED image formats are dropped.
            TabledownApp._converted_clipboard(
                None,
                {"html": HTML_BASIC, "text": "Name\tScore\nAlice\t95"},
            ).get("drop_types"),
            RENDERED_TABLE_TYPES,
        ),
    )
    check(
        "html_table_with_markdown_text_preserved",
        lambda: _assert_equal(
            TabledownApp._converted_clipboard(
                None,
                # markdown text + html <table> = real table (web/chat); a mismatched
                # separator must NOT trigger html->md, which would strip html and
                # break Excel paste.
                {"html": HTML_BASIC, "text": "| # | A | B |\n| --- | --- |\n| 1 | x | y |"},
            ),
            None,
        ),
    )
    check(
        "html_table_in_document_augments_text",
        lambda: _assert_contains(
            # a table embedded in a document: the text slot gains a Markdown
            # table while the surrounding paragraphs stay as text.
            TabledownApp._converted_clipboard(
                None,
                {"html": HTML_NOISE, "text": "Before\n제품 수량\n키보드 3\nAfter"},
            )["text"],
            "| 제품 | 수량 |",
        ),
    )
    check(
        "html_table_in_document_keeps_html",
        lambda: _assert_equal(
            # the original HTML slot must be preserved so rich editors still
            # paste a real table — only RENDERED image formats are dropped.
            bool(
                TabledownApp._converted_clipboard(
                    None,
                    {"html": HTML_NOISE, "text": "Before\n제품 수량\n키보드 3\nAfter"},
                ).get("drop_types", set())
                & HTML_TYPES
            ),
            False,
        ),
    )
    check(
        "xml_table_to_markdown",
        lambda: _assert_equal(
            table_xml_to_markdown(XML_TABLE),
            "| Name | Score |\n| --- | --- |\n| Alice | 95 |\n| Bob | 82 |",
        ),
    )
    check(
        "model_to_xml_roundtrip",
        lambda: _assert_equal(
            # model -> XML -> model is lossless for a single-level header.
            table_xml_to_model(
                model_to_xml(
                    [markdown_table_to_rows(MARKDOWN_BASIC)[0]],
                    markdown_table_to_rows(MARKDOWN_BASIC)[1:],
                )
            ),
            (
                [markdown_table_to_rows(MARKDOWN_BASIC)[0]],
                markdown_table_to_rows(MARKDOWN_BASIC)[1:],
            ),
        ),
    )
    check(
        "model_to_xml_nested_roundtrip",
        lambda: _assert_equal(
            # a 2D (vertical + horizontal) header round-trips through the v2
            # nested XML: vertical keys are forward-filled back into each row so
            # re-conversion is stable (spec §5). Header level fill: keys repeat
            # on every level on the way back.
            table_xml_to_model(
                model_to_xml(
                    [["직급", "직책", "분기", "분기"], ["직급", "직책", "Q1", "Q2"]],
                    [["부장", "팀장", "1", "2"], ["부장", "파트장", "3", "4"]],
                )
            ),
            (
                [["직급", "직책", "분기", "분기"], ["직급", "직책", "Q1", "Q2"]],
                [["부장", "팀장", "1", "2"], ["부장", "파트장", "3", "4"]],
            ),
        ),
    )
    check(
        "model_to_xml_emits_v2_vertical_group",
        lambda: _assert_contains(
            # the cross-table emits a vertical <{header}그룹> wrapper, the v2 shape.
            model_to_xml(
                [["직급", "직책", "분기", "분기"], ["직급", "직책", "Q1", "Q2"]],
                [["부장", "팀장", "1", "2"], ["부장", "파트장", "3", "4"]],
            ),
            '<직급그룹 이름="부장">',
        ),
    )
    check(
        "model_to_xml_simple_table_bare_row",
        lambda: _assert_equal(
            # no horizontal group → simple table: zero vertical keys, bare <행>,
            # leaves directly under it (spec §4-2).
            model_to_xml([["name", "val"]], [["A", "1"]]),
            '<표>\n  <행>\n    <열 n="name">A</열>\n    <열 n="val">1</열>\n  </행>\n</표>',
        ),
    )
    check(
        "model_to_xml_header_in_attribute",
        lambda: _assert_contains(
            # a header illegal as an XML tag name (space, leading digit) goes
            # verbatim into the 열 leaf's n= attribute — no mangling (v2).
            model_to_xml([["Q1 2024"]], [["12"]]),
            '<열 n="Q1 2024">12</열>',
        ),
    )
    check(
        "is_table_xml_accepts_table",
        lambda: _assert_equal(is_table_xml(XML_TABLE), True),
    )
    check(
        "is_table_xml_accepts_v2_table",
        lambda: _assert_equal(is_table_xml(XML_TABLE_V2), True),
    )
    check(
        "is_table_xml_accepts_v2_nested",
        lambda: _assert_equal(is_table_xml(XML_TABLE_V2_NESTED), True),
    )
    check(
        "v2_xml_to_markdown_flattens_groups",
        lambda: _assert_equal(
            # the Markdown path flattens v2 nesting to a combined-header table:
            # 직급/직책 keys + composite "1분기 1"/"1분기 2" columns (spec §7).
            table_xml_to_markdown(XML_TABLE_V2_NESTED),
            "| 직급 | 직책 | 1분기 1 | 1분기 2 |\n| --- | --- | --- | --- |\n| 부장 | 대족장 | 동 | 해 |",
        ),
    )
    check(
        "is_table_xml_rejects_config",
        lambda: _assert_equal(is_table_xml(XML_NOT_TABLE), False),
    )
    check(
        "is_table_xml_rejects_plain_text",
        lambda: _assert_equal(is_table_xml(PLAIN_TEXT), False),
    )
    check(
        "table_xml_not_auto_converted",
        lambda: _assert_equal(
            # XML is produced only by the explicit "Copy as XML" action; the
            # watcher never turns clipboard XML back into a table.
            TabledownApp._converted_clipboard(None, {"text": XML_TABLE}),
            None,
        ),
    )
    check(
        "config_xml_left_untouched",
        lambda: _assert_equal(
            # non-table XML is likewise never touched (false-positive guard).
            TabledownApp._converted_clipboard(None, {"text": XML_NOT_TABLE}),
            None,
        ),
    )
    check(
        "clipboard_model_from_html_table",
        lambda: _assert_equal(
            TabledownApp._clipboard_table_model({"html": HTML_BASIC}),
            ([["Name", "Score"]], [["Alice", "95"]]),
        ),
    )
    check(
        "clipboard_model_none_for_plain_text",
        lambda: _assert_equal(
            TabledownApp._clipboard_table_model({"text": PLAIN_TEXT}),
            None,
        ),
    )
    check(
        "merged_excel_skips_title_and_builds_composite_header",
        lambda: _assert_equal(
            # title row dropped; the two-row group header (colspan) is combined.
            html_table_to_rows(HTML_MERGED_EXCEL)[0],
            ["직급", "직책", "매출 1분기", "매출 2분기"],
        ),
    )
    check(
        "merged_excel_forward_fills_rowspan",
        lambda: _assert_equal(
            # the vertically merged "부장" is repeated into both of its rows.
            html_table_to_rows(HTML_MERGED_EXCEL)[1:],
            [["부장", "팀장", "12", "34"], ["부장", "파트장", "56", "78"]],
        ),
    )
    check(
        "merged_excel_to_xml_nests_group_header",
        lambda: _assert_contains(
            # the colspan group header becomes a <열그룹>, sub-columns nested 열
            # leaves — the 매출 / 1분기,2분기 hierarchy a flat header would collapse.
            model_to_xml(*html_table_to_model(HTML_MERGED_EXCEL)),
            '<열그룹 이름="매출">',
        ),
    )
    check(
        "merged_excel_to_xml_keeps_subheader_cell",
        lambda: _assert_contains(
            model_to_xml(*html_table_to_model(HTML_MERGED_EXCEL)),
            '<열 n="2분기">34</열>',
        ),
    )
    check(
        "merged_excel_to_xml_nests_vertical_key",
        lambda: _assert_contains(
            # the left key column (직급) nests vertically: 부장 spans both rows as
            # a <직급그룹> wrapper — the v2 vertical grouping (spec §3, §7).
            model_to_xml(*html_table_to_model(HTML_MERGED_EXCEL)),
            '<직급그룹 이름="부장">',
        ),
    )
    check(
        "forward_fill_fills_left_key_column",
        lambda: _assert_equal(
            # blank 직급 cells inherit "부장" from the row above.
            [r[0] for r in forward_fill_key_columns(html_table_to_rows(HTML_BLANK_GROUP))[1:]],
            ["부장", "부장", "부장"],
        ),
    )
    check(
        "forward_fill_stops_at_data_column",
        lambda: _assert_equal(
            # 점수(data) keeps its genuine blank — fill stops at 직책 (a full column).
            [r[2] for r in forward_fill_key_columns(html_table_to_rows(HTML_BLANK_GROUP))[1:]],
            ["10", "", "30"],
        ),
    )
    check(
        "forward_fill_horizontal_left",
        lambda: _assert_equal(
            # c2/c3 blank with a blank column above → inherit "A" from the left;
            # the 값 (data) column is never touched.
            forward_fill_key_columns(html_table_to_rows(HTML_HFILL))[1],
            ["A", "A", "A", "1"],
        ),
    )
    check(
        "clipboard_model_fill_blanks_on",
        lambda: _assert_equal(
            # with the option ON, the 족장 row's blank 직급 becomes 부장.
            # model is (header_levels, data_rows); data_rows[1] is the 족장 row.
            TabledownApp._clipboard_table_model({"html": HTML_BLANK_GROUP}, fill_blanks=True)[1][1][0],
            "부장",
        ),
    )
    check(
        "clipboard_model_no_fill_by_default",
        lambda: _assert_equal(
            # default (option OFF): the blank stays blank — no invented data.
            TabledownApp._clipboard_table_model({"html": HTML_BLANK_GROUP})[1][1][0],
            "",
        ),
    )
    # Privacy: a malformed-table error must never echo clipboard-derived text
    # (parser detail or the unexpected tag name, which is a header value here)
    # because these messages can reach the user-shareable diagnostics log.
    check("xml_parse_error_message_payload_free", _xml_parse_error_message_payload_free)
    check("xml_unexpected_node_message_payload_free", _xml_unexpected_node_message_payload_free)

    return tests


def run_i18n_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "i18n", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "i18n", str(exc), _elapsed(started)))

    check("translate_korean_menu_help", lambda: _assert_equal(t("menu.help", "ko"), "도움말"))
    check("translate_english_menu_help", lambda: _assert_equal(t("menu.help", "en"), "Help"))
    check("translate_korean_toggle", lambda: _assert_equal(t("menu.toggle", "ko"), "Tabledown 사용"))
    check("translate_english_toggle", lambda: _assert_equal(t("menu.toggle", "en"), "Use Tabledown"))
    check(
        "unknown_key_returns_key",
        lambda: _assert_equal(t("does.not.exist", "ko"), "does.not.exist"),
    )
    check(
        "unknown_language_falls_back_to_english",
        lambda: _assert_equal(t("menu.help", "fr"), "Help"),
    )
    check(
        "detect_system_language_returns_supported",
        lambda: _assert_in(detect_system_language(), SUPPORTED_LANGUAGES),
    )
    check(
        "help_message_localized_differs",
        lambda: _assert_not_equal(t("help.message", "ko"), t("help.message", "en")),
    )
    check(
        "login_item_keys_localized",
        lambda: _assert_not_equal(
            t("menu.login_item", "ko"),
            t("menu.login_item", "en"),
        ),
    )
    check(
        "copy_xml_keys_localized",
        lambda: _assert_not_equal(
            t("menu.copy_xml", "ko"),
            t("menu.copy_xml", "en"),
        ),
    )
    check(
        "diagnostics_keys_localized",
        lambda: _assert_not_equal(
            t("menu.diagnostics", "ko"),
            t("menu.diagnostics", "en"),
        ),
    )

    return tests


def run_diagnostics_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "diagnostics", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "diagnostics", str(exc), _elapsed(started)))

    check("scrub_removes_home_and_user", _scrub_removes_home_and_user)
    check("scrub_removes_volume_name", _scrub_removes_volume_name)
    check("scrub_masks_secret_token", _scrub_masks_secret_token)
    check("scrub_masks_space_bearer_token", _scrub_masks_space_bearer_token)
    check("scrub_masks_underscored_secret_key", _scrub_masks_underscored_secret_key)
    check("scrub_masks_url_credentials", _scrub_masks_url_credentials)
    check("scrub_masks_cloud_key_prefixes", _scrub_masks_cloud_key_prefixes)
    check("crash_record_omits_exception_message", _crash_record_omits_exception_message)
    check("install_hooks_sets_thread_excepthook", _install_hooks_sets_thread_excepthook)

    return tests


def run_login_item_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "login_item", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "login_item", str(exc), _elapsed(started)))

    check(
        "is_supported_returns_bool",
        lambda: _assert_isinstance(login_item.is_supported(), bool),
    )
    check(
        "is_enabled_returns_bool",
        lambda: _assert_isinstance(login_item.is_enabled(), bool),
    )

    return tests


def run_hotkey_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "hotkey", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "hotkey", str(exc), _elapsed(started)))

    check("hotkey_keycodes_are_ansi_c_and_t", _hotkey_keycodes)
    check("hotkey_dispatch_routes_by_id", _hotkey_dispatch_routes_by_id)
    check("hotkey_single_binding_fallback", _hotkey_single_binding_fallback)
    check("hotkey_unknown_id_does_not_misfire", _hotkey_unknown_id_no_misfire)
    return tests


def _hotkey_keycodes() -> str:
    # The constants app.py composes the ⌘⌃C / ⌘⌃T combos from must stay correct.
    return _assert_equal((KEY_C, KEY_T, CMD, CONTROL), (8, 17, 0x0100, 0x1000))


def _hotkey_dispatch_routes_by_id() -> str:
    # One Carbon handler serves both hotkeys; it must call the callback for the
    # id that actually fired, not just the first/last bound (the two-handler bug
    # this design replaces). No real Carbon: we drive _on_hotkey directly.
    hk = GlobalHotkeys()
    fired = []
    hk._bindings = {1: lambda _s: fired.append("c"), 2: lambda _s: fired.append("t")}
    hk._fired_hotkey_id = lambda _event: 2
    hk._on_hotkey(None, object(), None)
    return _assert_equal(fired, ["t"])


def _hotkey_single_binding_fallback() -> str:
    # If the id read fails (0) but exactly one hotkey is bound, the press must
    # still fire — the event can only be ours, so never drop a real keystroke.
    hk = GlobalHotkeys()
    fired = []
    hk._bindings = {5: lambda _s: fired.append("only")}
    hk._fired_hotkey_id = lambda _event: 0
    hk._on_hotkey(None, object(), None)
    return _assert_equal(fired, ["only"])


def _hotkey_unknown_id_no_misfire() -> str:
    # Two bound + unknown id: must NOT guess, or it would fire the wrong action.
    hk = GlobalHotkeys()
    fired = []
    hk._bindings = {1: lambda _s: fired.append("c"), 2: lambda _s: fired.append("t")}
    hk._fired_hotkey_id = lambda _event: 0
    hk._on_hotkey(None, object(), None)
    return _assert_equal(fired, [])


def run_clipboard_direct_tests() -> list[TestResult]:
    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "clipboard_direct", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "clipboard_direct", str(exc), _elapsed(started)))

    check("preserve_custom_type", _test_preserve_custom_type)
    check("write_marker_types", _test_write_marker_types)
    check("read_generated_marker", _test_read_generated_marker)
    check("drop_rendered_table_types", _test_drop_rendered_table_types)
    check("read_text_and_html", _test_read_text_and_html)
    return tests


def run_watcher_tests(require_watcher: bool) -> list[TestResult]:
    if not _is_tabledown_running():
        result = TestResult(
            "watcher_available",
            not require_watcher,
            "watcher",
            "Tabledown process is not running",
        )
        return [result]

    tests = []

    def check(name: str, fn) -> None:
        started = time.perf_counter()
        try:
            detail = fn()
            tests.append(TestResult(name, True, "watcher", detail or "ok", _elapsed(started)))
        except Exception as exc:  # noqa: BLE001
            tests.append(TestResult(name, False, "watcher", str(exc), _elapsed(started)))

    check("watcher_markdown_adds_html", _test_watcher_markdown_adds_html)
    check("watcher_html_adds_markdown", _test_watcher_html_adds_markdown)
    check("watcher_plain_text_unchanged", _test_watcher_plain_text_unchanged)
    return tests


def _test_preserve_custom_type() -> str:
    pb = NSPasteboard.generalPasteboard()
    custom_type = "com.tabledown.test.custom"
    custom_value = "custom payload"
    pb.clearContents()
    pb.declareTypes_owner_([str(NSPasteboardTypeString), custom_type], None)
    pb.setString_forType_("original", NSPasteboardTypeString)
    pb.setString_forType_(custom_value, custom_type)

    write_clipboard(text=MARKDOWN_BASIC, html=markdown_table_to_html(MARKDOWN_BASIC), mark_generated=True)
    preserved = pb.stringForType_(custom_type)
    if str(preserved) != custom_value:
        raise AssertionError(f"custom type lost: {preserved!r}")
    return "custom clipboard type preserved"


def _test_write_marker_types() -> str:
    pb = NSPasteboard.generalPasteboard()
    write_clipboard(text=MARKDOWN_BASIC, html=markdown_table_to_html(MARKDOWN_BASIC), mark_generated=True)
    missing = [pb_type for pb_type in GENERATED_MARKER_TYPES if not pb.stringForType_(pb_type)]
    if missing:
        raise AssertionError(f"missing marker types: {missing}")
    return "marker types written"


def _test_read_generated_marker() -> str:
    write_clipboard(text=MARKDOWN_BASIC, html=markdown_table_to_html(MARKDOWN_BASIC), mark_generated=True)
    content = read_clipboard()
    if not content.get("generated"):
        raise AssertionError(f"generated marker not detected: keys={list(content)}")
    return "generated marker detected"


def _test_drop_rendered_table_types() -> str:
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    pb.declareTypes_owner_([str(NSPasteboardTypeString), "public.png"], None)
    pb.setString_forType_("original", NSPasteboardTypeString)
    pb.setString_forType_("fake png payload", "public.png")

    write_clipboard(
        text=MARKDOWN_BASIC,
        html=markdown_table_to_html(MARKDOWN_BASIC),
        mark_generated=True,
        drop_types=RENDERED_TABLE_TYPES,
    )
    if "public.png" in [str(pb_type) for pb_type in pb.types() or []]:
        raise AssertionError("public.png was not dropped")
    return "rendered image type dropped"


def _test_read_text_and_html() -> str:
    content = read_clipboard()
    if not content.get("text") or "<table>" not in content.get("html", ""):
        raise AssertionError(f"unexpected clipboard content keys={list(content)}")
    return "read text/html"


def _test_watcher_markdown_adds_html() -> str:
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    pb.declareTypes_owner_([str(NSPasteboardTypeString)], None)
    pb.setString_forType_(MARKDOWN_BASIC, NSPasteboardTypeString)

    content = _wait_for(lambda: {
        "html": str(pb.stringForType_(NSPasteboardTypeHTML) or ""),
        "marker": str(pb.stringForType_(TABLEDOWN_GENERATED_TYPE) or ""),
    }, lambda value: "<table>" in value["html"] and value["marker"] == "Tabledown")
    return f"html_len={len(content['html'])}"


def _test_watcher_html_adds_markdown() -> str:
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    pb.declareTypes_owner_([str(NSPasteboardTypeString), str(NSPasteboardTypeHTML), LEGACY_HTML_TYPE, "public.png"], None)
    pb.setString_forType_("Name\tScore\nAlice\t95", NSPasteboardTypeString)
    pb.setString_forType_(HTML_BASIC, NSPasteboardTypeHTML)
    pb.setString_forType_(HTML_BASIC, LEGACY_HTML_TYPE)
    pb.setString_forType_("fake png payload", "public.png")

    content = _wait_for(
        lambda: {
            "text": str(pb.stringForType_(NSPasteboardTypeString) or ""),
            "html": str(pb.stringForType_(NSPasteboardTypeHTML) or ""),
            "types": [str(pb_type) for pb_type in pb.types() or []],
        },
        lambda value: (
            value["text"].startswith("\n| Name | Score |")
            and "| Alice | 95 |" in value["text"]
            and not value["html"]
            and all(pb_type not in value["types"] for pb_type in HTML_TYPES)
            and "public.png" not in value["types"]
        ),
    )
    return f"text_len={len(content['text'])}"


def _test_watcher_plain_text_unchanged() -> str:
    pb = NSPasteboard.generalPasteboard()
    pb.clearContents()
    pb.declareTypes_owner_([str(NSPasteboardTypeString)], None)
    pb.setString_forType_(PLAIN_TEXT, NSPasteboardTypeString)
    time.sleep(0.5)
    text = str(pb.stringForType_(NSPasteboardTypeString) or "")
    html = str(pb.stringForType_(NSPasteboardTypeHTML) or "")
    marker = str(pb.stringForType_(TABLEDOWN_GENERATED_TYPE) or "")
    if text != PLAIN_TEXT or html or marker:
        raise AssertionError(f"plain text changed text={text!r} html_len={len(html)} marker={marker!r}")
    return "plain text untouched"


def _xml_parse_error_message_payload_free() -> str:
    # Malformed XML carrying a recognizable token. The raised ValueError must NOT
    # echo the parser detail (line/column or source markup) — that can leak
    # clipboard content into a shared log.
    try:
        table_xml_to_model("<표><행>SECRETCELL12345</행")
    except ValueError as exc:
        msg = str(exc)
        if "line" in msg.lower() or "column" in msg.lower() or "SECRETCELL12345" in msg:
            raise AssertionError(f"parse detail leaked into message: {msg!r}")
        return f"clean: {msg!r}"
    raise AssertionError("expected ValueError for malformed xml")


def _xml_unexpected_node_message_payload_free() -> str:
    # The unexpected tag name is a header VALUE in this format, so it must not
    # appear in the error message (would leak table data into a shared log).
    secret = "SECRETHEADER12345"
    try:
        table_xml_to_model(f"<표><{secret}>x</{secret}></표>")
    except ValueError as exc:
        if secret in str(exc):
            raise AssertionError(f"tag name leaked into message: {str(exc)!r}")
        return f"clean: {str(exc)!r}"
    raise AssertionError("expected ValueError for unexpected node")


def _scrub_removes_home_and_user() -> str:
    import getpass
    from tablemark import diagnostics
    user = getpass.getuser() or "tester"
    out = diagnostics.scrub(f'File "/Users/{user}/Library/Logs/Tabledown.log", line 5')
    if user and user in out:
        raise AssertionError(f"username leaked: {out!r}")
    return out


def _scrub_removes_volume_name() -> str:
    from tablemark import diagnostics
    out = diagnostics.scrub("mounted at /Volumes/My Secret Drive/data")
    if "Secret" in out:
        raise AssertionError(f"volume label leaked: {out!r}")
    return out


def _scrub_masks_secret_token() -> str:
    from tablemark import diagnostics
    out = diagnostics.scrub("token=abcdef123456 and ghp_AAAAAAAAAAAAAAAAAA")
    if "abcdef123456" in out or "ghp_AAAAAAAAAAAAAAAAAA" in out:
        raise AssertionError(f"secret not masked: {out!r}")
    return out


def _scrub_masks_space_bearer_token() -> str:
    from tablemark import diagnostics
    out = diagnostics.scrub("Authorization: Bearer eyJhbGciOiJ.payloadpart9.signature99")
    if "eyJhbGciOiJ" in out:
        raise AssertionError(f"space-form bearer/JWT token leaked: {out!r}")
    return out


def _scrub_masks_underscored_secret_key() -> str:
    from tablemark import diagnostics
    out = diagnostics.scrub("aws_secret_access_key = wJalrXUtnFEMIabc123def")
    if "wJalrXUtnFEMIabc123def" in out:
        raise AssertionError(f"underscored secret key not masked: {out!r}")
    return out


def _scrub_masks_url_credentials() -> str:
    from tablemark import diagnostics
    out = diagnostics.scrub("postgres://admin:topsecretpw@db.example:5432/x")
    if "topsecretpw" in out:
        raise AssertionError(f"inline url credential leaked: {out!r}")
    return out


def _scrub_masks_cloud_key_prefixes() -> str:
    from tablemark import diagnostics
    samples = ["AKIAIOSFODNN7EXAMPLE", "sk-proj-AB12CD34EF56GH78", "ghp_ABCDEFGH12345678"]
    out = diagnostics.scrub(" ".join(samples))
    for s in samples:
        if s in out:
            raise AssertionError(f"cloud key not masked ({s}): {out!r}")
    return out


def _crash_record_omits_exception_message() -> str:
    # The crash record must keep the exception TYPE + frames but NEVER the
    # runtime exception value (which can carry clipboard/table payload). The
    # secret is built at runtime so it is not present in the raise's source line.
    from tablemark import diagnostics
    secret = "PAYLOADCELL" + "42PRIVATEROW"
    try:
        raise ValueError(secret)
    except ValueError as exc:
        rec = diagnostics._format_record("watcher", type(exc), exc.__traceback__)
    if secret in rec:
        raise AssertionError(f"runtime exception message leaked into crash record: {rec!r}")
    if "ValueError" not in rec:
        raise AssertionError(f"exception type missing from record: {rec!r}")
    return "type+frames recorded, runtime message omitted"


def _install_hooks_sets_thread_excepthook() -> str:
    # threading.excepthook is the only way to see uncaught exceptions on the
    # daemon clipboard-watcher thread — without it they vanish from the log.
    import sys as _sys
    import threading as _t
    from tablemark import diagnostics
    before_sys, before_thread = _sys.excepthook, _t.excepthook
    try:
        diagnostics._install_python_hooks()
        if _t.excepthook is before_thread:
            raise AssertionError("threading.excepthook not installed (watcher crashes would vanish)")
        if _sys.excepthook is before_sys:
            raise AssertionError("sys.excepthook not installed")
        return "sys + threading excepthook installed"
    finally:
        _sys.excepthook, _t.excepthook = before_sys, before_thread


def _wait_for(read_fn, done_fn, timeout: float = 2.0, interval: float = 0.05):
    deadline = time.monotonic() + timeout
    last_value = None
    while time.monotonic() < deadline:
        last_value = read_fn()
        if done_fn(last_value):
            return last_value
        time.sleep(interval)
    raise AssertionError(f"timed out waiting; last={last_value!r}")


def _is_tabledown_running() -> bool:
    result = subprocess.run(
        ["pgrep", "-f", "Tabledown.app/Contents/MacOS/Tabledown|python.*run.py"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    return bool(result.stdout.strip())


def _assert_equal(actual, expected) -> str:
    if actual != expected:
        raise AssertionError(f"expected {expected!r}, got {actual!r}")
    return "ok"


def _assert_contains(actual: str, expected: str) -> str:
    if expected not in actual:
        raise AssertionError(f"expected {expected!r} in {actual!r}")
    return "ok"


def _assert_in(value, expected) -> str:
    if value not in expected:
        raise AssertionError(f"expected one of {expected!r}, got {value!r}")
    return "ok"


def _assert_not_equal(actual, other) -> str:
    if actual == other:
        raise AssertionError(f"expected difference, both were {actual!r}")
    return "ok"


def _assert_isinstance(value, expected_type) -> str:
    if not isinstance(value, expected_type):
        raise AssertionError(f"expected {expected_type!r}, got {type(value).__name__}: {value!r}")
    return "ok"


def _elapsed(started: float) -> int:
    return round((time.perf_counter() - started) * 1000)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--watcher", action="store_true", help="Run tests against a running Tabledown app.")
    parser.add_argument("--require-watcher", action="store_true", help="Fail when watcher tests cannot run.")
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_fixture_files(args.output_dir)

    results = []
    results.extend(run_converter_tests())
    results.extend(run_i18n_tests())
    results.extend(run_diagnostics_tests())
    results.extend(run_login_item_tests())
    results.extend(run_hotkey_tests())
    results.extend(run_clipboard_direct_tests())
    if args.watcher or args.require_watcher:
        results.extend(run_watcher_tests(args.require_watcher))

    report_path = args.output_dir / "report.json"
    report_path.write_text(
        json.dumps([asdict(result) for result in results], ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    for result in results:
        status = "PASS" if result.ok else "FAIL"
        print(f"{status} {result.category}:{result.name} {result.detail} ({result.elapsed_ms}ms)")
    print(f"report={report_path}")

    return 0 if all(result.ok for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
