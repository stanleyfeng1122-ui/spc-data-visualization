"""File paths for the SPC Data Viz app.

Logs, examples, and other path-based constants. Currently minimal —
the launchd service writes to /Users/zhefeng/Library/Logs/spc-streamlit.log
and that is configured outside this codebase (in the .plist file).
"""

from __future__ import annotations

from pathlib import Path

# Repository root — used by Quick Test page to discover xlsx files
REPO_ROOT: Path = Path(__file__).resolve().parents[3]

# Where sample xlsx fixtures will live after R8 (currently at REPO_ROOT)
EXAMPLES_DIR: Path = REPO_ROOT / "examples"
