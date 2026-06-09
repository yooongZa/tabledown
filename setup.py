"""py2app build script for Tabledown.

Build with:
    python setup.py py2app
"""
import ast
import os
from pathlib import Path

from setuptools import find_packages, setup

ROOT = Path(__file__).resolve().parent
VERSION_FILE = ROOT / "tablemark" / "__init__.py"


def read_version() -> str:
    module = ast.parse(VERSION_FILE.read_text(encoding="utf-8"))
    for node in module.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "__version__":
                return ast.literal_eval(node.value)
    raise RuntimeError("Unable to find __version__")


VERSION = read_version()
# CFBundleShortVersionString is the user-facing marketing version (VERSION).
# CFBundleVersion is the build number, which App Store Connect requires to be
# unique and monotonically increasing within a version train, and which must
# be at most three period-separated non-negative integers (e.g. 0.2.1 — NOT
# 0.2.0.1, which has four components and is rejected with error 236550). When
# re-submitting the same marketing version, override it per upload via
# TABLEDOWN_BUILD (e.g. TABLEDOWN_BUILD=0.2.1); defaults to VERSION otherwise.
BUILD = os.environ.get("TABLEDOWN_BUILD", VERSION)
APP = ["run.py"]
OPTIONS = {
    "argv_emulation": False,
    "excludes": ["setuptools", "pkg_resources", "pip", "wheel"],
    "plist": {
        "LSUIElement": True,  # Menu bar only — no Dock icon
        "CFBundleName": "Tabledown",
        "CFBundleDisplayName": "Tabledown",
        "CFBundleIdentifier": "com.tabledown.app",
        "CFBundleVersion": BUILD,
        "CFBundleShortVersionString": VERSION,
        "CFBundleSupportedPlatforms": ["MacOSX"],
        "LSApplicationCategoryType": "public.app-category.productivity",
        "LSMinimumSystemVersion": "12.0",
        "NSHumanReadableCopyright": "© 2026 Tabledown",
    },
    "packages": ["rumps", "bs4", "AppKit"],
    # StoreKit (IAP) is imported lazily/defensively in store.py, so name it
    # explicitly for py2app's dependency graph to bundle it. The global hotkey
    # (hotkey.py) binds the system Carbon.framework via ctypes at runtime and
    # needs no module include.
    "includes": ["StoreKit"],
    "resources": [
        "assets/generated/tablemark_menu_40.png",
        "assets/generated/tablemark_menu_40_off.png",
        "assets/generated/tablemark_menu_40_check.png",
    ],
    "iconfile": "assets/Tabledown.icns",
}

setup(
    app=APP,
    name="Tabledown",
    version=VERSION,
    packages=find_packages(),
    options={"py2app": OPTIONS},
)
