"""Verify a Windows release and exercise its EXE on a disposable CI desktop.

The native smoke checks replace the runner's clipboard. Require GitHub Actions
so this helper cannot accidentally run against someone's personal clipboard.
"""
from __future__ import annotations

import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import struct
import subprocess
import sys
import tempfile
import time
import types
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
WINDOWS = ROOT / "windows"
DIST = WINDOWS / "dist"
sys.path[:0] = [str(WINDOWS), str(ROOT)]


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def normalized(code):
    if not isinstance(code, types.CodeType):
        return code
    return code.replace(
        co_filename="<source>",
        co_consts=tuple(normalized(value) for value in code.co_consts),
    )


def wait_for(predicate, message, process=None, timeout=20):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if process is not None and process.poll() is not None:
            raise AssertionError(f"Packaged app exited early: {process.returncode}")
        if predicate():
            return
        time.sleep(0.1)
    raise AssertionError(message)


def verify():
    assert sys.platform == "win32" and os.environ.get("GITHUB_ACTIONS") == "true"
    from PyInstaller.archive.readers import CArchiveReader
    import win32api
    import win32clipboard
    import win32con
    from tabledown_windows import __version__
    from tabledown_windows.html_clipboard import build_cf_html
    from tabledown_windows.win_clipboard import HTML_FORMAT_ID, read_clipboard
    from tablemark.converter.md_to_tsv import markdown_table_to_rows

    report = {
        "status": "running", "version": __version__,
        "source_commit": os.environ["GITHUB_SHA"], "checks": {},
        "windows_excel_device_test": "not_performed",
        "msix_install_test": "not_performed",
        "dependencies": {d.metadata["Name"]: d.version for d in importlib.metadata.distributions()},
    }
    output = DIST / "release-verification.json"
    output.write_text(json.dumps(report, indent=2), encoding="utf-8")
    portable = DIST / f"Tabledown-Windows-portable-{__version__}.zip"
    msix = DIST / f"Tabledown-{__version__}.0.msix"
    try:
        with tempfile.TemporaryDirectory(prefix="tabledown-release-") as work:
            work = Path(work)
            with zipfile.ZipFile(portable) as archive, zipfile.ZipFile(msix) as package:
                files = {n: sha(archive.read(n)) for n in archive.namelist() if not n.endswith("/")}
                assert files and all(n.startswith("Tabledown-Windows/") for n in files)
                assert not any(n.lower().endswith((".pfx", ".key", ".pem")) for n in files)
                payload = {n: sha(package.read(n)) for n in package.namelist() if n.startswith("app/") and not n.endswith("/")}
                assert {n.replace("Tabledown-Windows/", "app/", 1): h for n, h in files.items()} == payload
                manifest = ET.fromstring(package.read("AppxManifest.xml"))
                ns = {"p": "http://schemas.microsoft.com/appx/manifest/foundation/windows10"}
                identity = manifest.find("p:Identity", ns).attrib
                assert identity == {"Name": "Tabledown.Dev", "Publisher": "CN=Tabledown.Dev", "Version": __version__ + ".0", "ProcessorArchitecture": "x64"}
                assert "AppxSignature.p7x" in package.namelist()
                report["checks"]["msix_zip_payload_identical"] = len(files)
                report["msix_identity"] = identity
                archive.extractall(work)

            exe = work / "Tabledown-Windows/Tabledown-Windows.exe"
            data = exe.read_bytes()
            pe = struct.unpack_from("<I", data, 0x3C)[0]
            assert data[pe:pe + 4] == b"PE\0\0" and struct.unpack_from("<H", data, pe + 4)[0] == 0x8664
            report["checks"]["exe_architecture"] = "x64"
            archive = CArchiveReader(str(exe))
            pyz = archive.open_embedded_archive("PYZ.pyz")
            modules = {}
            for base, prefix in [(WINDOWS / "tabledown_windows", "tabledown_windows"), (ROOT / "tablemark/converter", "tablemark.converter")]:
                for source in base.glob("*.py"):
                    name = prefix if source.stem == "__init__" else prefix + "." + source.stem
                    # General table XML is a macOS action. Windows imports the
                    # formula serializers, so PyInstaller omits this unused module.
                    if name == "tablemark.converter.table_xml":
                        continue
                    expected = compile(source.read_bytes(), str(source), "exec", dont_inherit=True, optimize=0)
                    assert normalized(pyz.extract(name)) == normalized(expected), name
                    modules[name] = sha(source.read_bytes())
            assert "win32com.client" in pyz.toc and "pythoncom" in pyz.toc
            assert any("winsdk" in n and n.endswith(".pyd") for n in files)
            report["checks"]["bundled_sources"] = modules
            report["checks"]["deferred_com_winrt_dependencies_present"] = True

            profile = work / "profile"
            settings = profile / "Tabledown/settings.json"
            settings.parent.mkdir(parents=True)
            settings.write_text(json.dumps({"welcome_shown": True, "language": "en", "fill_blanks": False}), encoding="utf-8")
            env = os.environ.copy()
            env.update(APPDATA=str(profile), LOCALAPPDATA=str(profile))
            log_path = profile / "Tabledown/Tabledown.log"

            def logs():
                return log_path.read_text(encoding="utf-8") if log_path.exists() else ""

            def seed(text, html=None):
                # Only a disposable runner uses this helper. Retry transient OS locks.
                for attempt in range(20):
                    try:
                        win32clipboard.OpenClipboard()
                        break
                    except Exception:
                        if attempt == 19:
                            raise
                        time.sleep(0.05)
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
                    if html:
                        win32clipboard.SetClipboardData(HTML_FORMAT_ID, build_cf_html(html))
                finally:
                    win32clipboard.CloseClipboard()

            def toggle():
                for key in (win32con.VK_CONTROL, win32con.VK_MENU, ord("T")):
                    win32api.keybd_event(key, 0, 0, 0)
                for key in (ord("T"), win32con.VK_MENU, win32con.VK_CONTROL):
                    win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)

            seed("Tabledown release check")
            process = subprocess.Popen([str(exe)], cwd=exe.parent, env=env)
            try:
                wait_for(lambda: "windows Excel formula hotkey registered" in logs(), "Packaged app did not start", process)
                assert "windows global hotkey registered" in logs()
                # This log is emitted after the frozen WinRT import succeeded;
                # unpackaged EXEs are expected to have no StartupTask identity.
                assert "StartupTask read failed:" in logs()
                report["checks"]["frozen_app_start_and_winrt_import"] = True
                second = subprocess.Popen([str(exe)], cwd=exe.parent, env=env)
                try:
                    assert second.wait(timeout=8) == 0
                finally:
                    if second.poll() is None:
                        second.terminate()
                        second.wait(timeout=5)
                assert process.poll() is None
                report["checks"]["second_instance_exits"] = True

                html = '<table><tr><td>항목</td><td>값</td></tr><tr><td>A</td><td>한글 😀 앞<span style="mso-spacerun:yes">\u00a0\r\n  </span>공백<br>\r\n<br>\r\n뒤</td></tr></table>'
                seed("source table", html)
                wait_for(lambda: read_clipboard().get("generated"), "HTML conversion timed out", process)
                copied = read_clipboard()
                assert copied["html"] == html
                assert markdown_table_to_rows(copied["text"]) == [["항목", "값"], ["A", "한글 😀 앞 공백\n\n뒤"]]
                report["checks"]["native_html_to_markdown_preserves_html_unicode_breaks"] = True

                markdown = '| 항목 | 값 |\n| --- | --- |\n| 한글 😀 | A\\|B |'
                seed(markdown)
                wait_for(lambda: read_clipboard().get("generated"), "Markdown conversion timed out", process)
                copied = read_clipboard()
                assert copied["text"] == markdown and "A|B" in copied["html"] and "한글 😀" in copied["html"]
                report["checks"]["native_markdown_to_html_unicode_literal_pipe"] = True

                for value in ["ordinary text 한글 😀", '<설정><값>keep</값></설정>']:
                    seed(value)
                    time.sleep(0.8)
                    assert read_clipboard() == {"text": value}
                report["checks"]["ordinary_text_and_xml_untouched"] = True

                toggle()
                wait_for(lambda: "conversion disabled" in logs(), "Pause hotkey timed out", process)
                seed("paused table", html)
                time.sleep(0.8)
                assert read_clipboard() == {"text": "paused table", "html": html}
                toggle()
                wait_for(lambda: "conversion enabled" in logs(), "Resume hotkey timed out", process)
                seed("resumed table", html)
                wait_for(lambda: read_clipboard().get("generated"), "Resume conversion timed out", process)
                report["checks"]["native_pause_resume_hotkey"] = True
                assert "Traceback" not in logs() and "clipboard update failed" not in logs()
                (DIST / "release-smoke.log").write_text(logs(), encoding="utf-8")
            finally:
                if process.poll() is None:
                    process.terminate()
                    process.wait(timeout=5)
                seed("")
        report["status"] = "passed"
        report["artifacts"] = {p.name: {"sha256": sha(p.read_bytes()), "bytes": p.stat().st_size} for p in (portable, msix, DIST / "Tabledown-dev.cer")}
    except Exception as error:
        report["status"] = "failed"
        report["error"] = repr(error)
        raise
    finally:
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": report["status"], "version": __version__, "checks": list(report["checks"])}))


if __name__ == "__main__":
    verify()
