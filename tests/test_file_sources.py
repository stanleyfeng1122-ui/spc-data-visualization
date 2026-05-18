from pathlib import Path

from spc_viz.ui.file_sources import LocalXlsxFile, discover_local_xlsx_files


def test_discover_local_xlsx_files_skips_temp_files(tmp_path: Path):
    workbook = tmp_path / "vendor.xlsx"
    temp = tmp_path / "~$vendor.xlsx"
    other = tmp_path / "notes.txt"
    workbook.write_bytes(b"xlsx")
    temp.write_bytes(b"temp")
    other.write_text("notes", encoding="utf-8")

    files = discover_local_xlsx_files(str(tmp_path))

    assert [f.name for f in files] == ["vendor.xlsx"]


def test_discover_local_xlsx_files_handles_missing_folder(tmp_path: Path):
    assert discover_local_xlsx_files(str(tmp_path / "missing")) == []


def test_local_xlsx_file_matches_uploaded_file_surface(tmp_path: Path):
    workbook = tmp_path / "vendor.xlsx"
    workbook.write_bytes(b"abc")

    local_file = LocalXlsxFile(workbook)

    assert local_file.name == "vendor.xlsx"
    assert local_file.size == 3
    assert local_file.getvalue() == b"abc"

