"""Sheet selection helpers for upload-time parsing controls."""

from __future__ import annotations


def choose_default_sheets(
    display_sheets: list[str],
    preferred_names: tuple[str, ...] = ("Data Input",),
) -> list[str]:
    """Return the default sheet selection for the main upload workflow.

    The app may detect several parseable sheets across uploaded workbooks, but
    normal usage starts from one data sheet. Prefer an exact case-insensitive
    match such as "Data Input"; fall back to the first detected sheet so the app
    still starts with a valid selection for unusual vendor files.
    """
    if not display_sheets:
        return []

    normalized = {sheet.lower().strip(): sheet for sheet in display_sheets}
    for preferred in preferred_names:
        match = normalized.get(preferred.lower().strip())
        if match is not None:
            return [match]

    return [display_sheets[0]]

