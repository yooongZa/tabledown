"""py2app build script for Tabledown.

Build with:
    python setup.py py2app
"""
from setuptools import find_packages, setup

APP = ["run.py"]
OPTIONS = {
    "argv_emulation": False,
    "plist": {
        "LSUIElement": True,  # Menu bar only — no Dock icon
        "CFBundleName": "Tabledown",
        "CFBundleDisplayName": "Tabledown",
        "CFBundleIdentifier": "com.tabledown.app",
        "CFBundleVersion": "0.1.0",
        "CFBundleShortVersionString": "0.1.0",
        "NSHumanReadableCopyright": "© 2026 Tabledown",
    },
    "packages": ["rumps", "bs4", "AppKit", "Quartz"],
    "resources": ["assets/generated/tablemark_menu_40.png"],
    "iconfile": "assets/Tabledown.icns",
}

setup(
    app=APP,
    name="Tabledown",
    version="0.1.0",
    packages=find_packages(),
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)
