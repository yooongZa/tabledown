"""py2app build script for Tabledown.

Build with:
    python setup.py py2app
"""
import ast
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
APP = ["run.py"]
OPTIONS = {
    "argv_emulation": False,
    "excludes": ["setuptools", "pkg_resources", "pip", "wheel"],
    "plist": {
        "LSUIElement": True,  # Menu bar only — no Dock icon
        "CFBundleName": "Tabledown",
        "CFBundleDisplayName": "Tabledown",
        "CFBundleIdentifier": "com.tabledown.app",
        "CFBundleVersion": VERSION,
        "CFBundleShortVersionString": VERSION,
        "CFBundleSupportedPlatforms": ["MacOSX"],
        "LSApplicationCategoryType": "public.app-category.productivity",
        "LSMinimumSystemVersion": "12.0",
        "NSHumanReadableCopyright": "© 2026 Tabledown",
    },
    "packages": ["rumps", "bs4", "AppKit"],
    "resources": ["assets/generated/tablemark_menu_40.png"],
    "iconfile": "assets/Tabledown.icns",
}

setup(
    app=APP,
    name="Tabledown",
    version=VERSION,
    packages=find_packages(),
    options={"py2app": OPTIONS},
)
