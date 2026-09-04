"""Convert in-memory .docx bytes to PDF bytes using Microsoft Word on Windows.

Word's COM API requires file paths, so a temporary directory is used for the
round-trip.  The temp files live for ~200 ms during the call and are destroyed
immediately afterwards by ``TemporaryDirectory``'s context manager.

Streamlit runs script re-runs in background worker threads whose COM apartment
is not initialised, so every call wraps the work between ``CoInitialize`` and
``CoUninitialize``.
"""

import sys
import tempfile
from pathlib import Path


def convert(docx_bytes: bytes) -> bytes:
    """Return *docx_bytes* rendered to PDF by Microsoft Word.

    Raises ``OSError`` on non-Windows platforms and ``RuntimeError`` when Word
    cannot be launched (not installed, COM error, etc.).
    """
    if sys.platform != "win32":
        raise OSError("Word COM automation is only available on Windows.")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            docx_file = tmp_path / "teamsheet.docx"
            pdf_file = tmp_path / "teamsheet.pdf"

            docx_file.write_bytes(docx_bytes)

            # DispatchEx starts a private Word process -- it will not attach to
            # a Word instance the user may already have open.
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone

            doc = word.Documents.Open(str(docx_file.resolve()), ReadOnly=True)
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_file.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,  # wdExportOptimizeForPrint
            )

            pdf_bytes = pdf_file.read_bytes()

            # Must close the document and quit Word BEFORE TemporaryDirectory exits,
            # otherwise Windows prevents deleting teamsheet.docx (WinError 32).
            doc.Close(SaveChanges=False)
            doc = None
            word.Quit()
            word = None

            return pdf_bytes
    except Exception as exc:
        raise RuntimeError(f"Word PDF conversion failed: {exc}") from exc
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
