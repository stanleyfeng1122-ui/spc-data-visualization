"""Centralized constants module.

Currently empty by design: after R3/R4 split, constants live with the code
that uses them — colors in `theme/css.py` and `charts/base.py`, parser
tuning in `parsers/header_detect.py`, sheet skip patterns in
`parsers/measurements.py`.

Move things here only when a constant becomes:
1. Used by 2+ unrelated layers (UI + parsers + charts), OR
2. Likely to be user-configurable in the future (env var / config file)

Until then, prefer co-location.
"""
