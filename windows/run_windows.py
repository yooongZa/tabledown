#!/usr/bin/env python3
"""Windows entry point for Tabledown."""
from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

WINDOWS_ROOT = Path(__file__).resolve().parent
if str(WINDOWS_ROOT) not in sys.path:
    sys.path.insert(0, str(WINDOWS_ROOT))

from tabledown_windows.app import main


if __name__ == "__main__":
    main()
