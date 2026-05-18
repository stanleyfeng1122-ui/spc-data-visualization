from spc_viz.ui.sheet_selection import choose_default_sheets


def test_choose_default_sheets_prefers_data_input():
    assert choose_default_sheets(["Raw data", "Data Input", "FAI"]) == ["Data Input"]


def test_choose_default_sheets_matches_data_input_case_insensitively():
    assert choose_default_sheets(["raw data", "data input", "FAI"]) == ["data input"]


def test_choose_default_sheets_falls_back_to_first_sheet():
    assert choose_default_sheets(["Raw Data-PP", "Cpk Summary"]) == ["Raw Data-PP"]


def test_choose_default_sheets_handles_empty_options():
    assert choose_default_sheets([]) == []

