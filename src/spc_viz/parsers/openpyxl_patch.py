"""
openpyxl compatibility patches.

Two issues addressed here, both triggered at import time:

1. ExternalReference monkey-patch
   Strict-OOXML xlsx files have ``<externalReference r:id="..."/>`` which
   openpyxl 3.1.x deserializes as ``{'{ns}id': 'rId13'}`` but
   ``ExternalReference.__init__`` expects a positional ``id`` arg.  The fix
   makes ``id`` optional with a default of "".

2. Strict-OOXML namespace rewrite (``_open_strict_ooxml``)
   openpyxl 3.1.x cannot parse files using the strict OOXML namespace
   (``http://purl.oclc.org/ooxml/...``). The helper rewrites the XML
   namespaces in-memory to the transitional ones that openpyxl understands.
"""

from __future__ import annotations

import io

import openpyxl

# ---------------------------------------------------------------------------
# Monkey-patch openpyxl 3.1.x ExternalReference bug
# Strict-OOXML xlsx files have <externalReference r:id="..."/> which openpyxl
# deserializes as {'{ns}id': 'rId13'} but ExternalReference.__init__ expects
# a positional 'id' arg. The fix: make 'id' optional with a default.
# ---------------------------------------------------------------------------
try:
    import inspect as _inspect

    from openpyxl.packaging.workbook import ExternalReference as _ER

    _params = _inspect.signature(_ER.__init__).parameters
    if "id" in _params and _params["id"].default is _inspect.Parameter.empty:
        _ER.__init__.__defaults__ = ("",)
except Exception:
    pass


def _open_strict_ooxml(file_or_path):
    """Convert strict-OOXML xlsx to transitional namespace so openpyxl can read it.

    openpyxl 3.1.x cannot parse files using the strict OOXML namespace
    (http://purl.oclc.org/ooxml/...). This rewrites the XML namespaces
    in-memory to the transitional ones that openpyxl understands.
    """
    import zipfile

    _STRICT_TO_TRANSITIONAL = {
        "http://purl.oclc.org/ooxml/spreadsheetml/main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/officeDocument/relationships": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "http://purl.oclc.org/ooxml/drawingml/main": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "http://purl.oclc.org/ooxml/drawingml/chart": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/theme": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    }

    buf_in = io.BytesIO()
    if isinstance(file_or_path, str):
        with open(file_or_path, "rb") as f:
            buf_in.write(f.read())
    else:
        file_or_path.seek(0)
        buf_in.write(file_or_path.read())
    buf_in.seek(0)

    buf_out = io.BytesIO()
    with zipfile.ZipFile(buf_in, "r") as zin, zipfile.ZipFile(buf_out, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                text = data.decode("utf-8", errors="replace")
                # Also strip conformance="strict" attribute
                text = text.replace(' conformance="strict"', "")
                for strict_ns, trans_ns in _STRICT_TO_TRANSITIONAL.items():
                    text = text.replace(strict_ns, trans_ns)
                data = text.encode("utf-8")
            zout.writestr(item, data)

    buf_out.seek(0)
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return openpyxl.load_workbook(buf_out, data_only=True, read_only=True, keep_links=False)
