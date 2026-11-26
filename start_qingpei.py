"""ASCII launcher for the Chinese-named LAN management script."""

from __future__ import annotations

import runpy
from pathlib import Path


ROOT = Path(__file__).resolve().parent
scripts = list(ROOT.glob("*局域网管理系统.py"))

if len(scripts) != 1:
    raise SystemExit("Unable to locate the LAN management script.")

runpy.run_path(str(scripts[0]), run_name="__main__")
