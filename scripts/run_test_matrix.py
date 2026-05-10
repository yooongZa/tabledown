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
    LEGACY_HTML_TYPE,
    RENDERED_TABLE_TYPES,
    TABLEDOWN_GENERATED_TYPE,
    read_clipboard,
    write_clipboard,
)
from tablemark.converter.html_to_md import html_table_to_markdown
from tablemark.converter.md_to_tsv import markdown_table_to_html, markdown_table_to_rows
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

    return tests


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
            "types": [str(pb_type) for pb_type in pb.types() or []],
        },
        lambda value: (
            value["text"].startswith("| Name | Score |")
            and "| Alice | 95 |" in value["text"]
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
