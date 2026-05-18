"""File source helpers for uploaded and local Excel workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from spc_viz.config.paths import REPO_ROOT

LAST_FOLDER_PATH = REPO_ROOT / ".streamlit" / "last_data_folder.txt"


@dataclass(frozen=True)
class LocalXlsxFile:
    """Small adapter matching the uploaded-file methods used by the main app."""

    path: Path

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def size(self) -> int:
        return self.path.stat().st_size

    def getvalue(self) -> bytes:
        return self.path.read_bytes()


def default_local_data_dir() -> Path:
    """Return the best default folder for local workbook reloads."""
    candidate = REPO_ROOT / "New Data Repo"
    return candidate if candidate.exists() else REPO_ROOT


def read_last_data_folder() -> str:
    """Return the last used local data folder, falling back to the default."""
    try:
        value = LAST_FOLDER_PATH.read_text(encoding="utf-8").strip()
    except OSError:
        return str(default_local_data_dir())

    return value or str(default_local_data_dir())


def remember_last_data_folder(folder: str) -> None:
    """Persist the last used folder for refresh/code-update recovery."""
    try:
        LAST_FOLDER_PATH.parent.mkdir(parents=True, exist_ok=True)
        LAST_FOLDER_PATH.write_text(folder, encoding="utf-8")
    except OSError:
        pass


def discover_local_xlsx_files(folder: str) -> list[LocalXlsxFile]:
    """Discover non-temporary .xlsx files directly under a local folder."""
    data_dir = Path(folder).expanduser()
    if not data_dir.exists() or not data_dir.is_dir():
        return []

    files = [
        LocalXlsxFile(path)
        for path in sorted(data_dir.iterdir())
        if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
    ]
    return files

