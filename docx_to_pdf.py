"""Convert in-memory .docx bytes to PDF bytes.

Uses `dxpdf` for pure Python/Rust in-memory conversion across Linux, macOS,
and Windows. On Windows, falls back to Microsoft Word COM if dxpdf fails.
"""

import sys
import tempfile
from pathlib import Path


def convert(docx_bytes: bytes) -> bytes:
    """Return *docx_bytes* rendered to PDF.

    Works on Linux, macOS, and Windows.
    """
    # 1. Try pure-Python/Rust cross-platform converter (Linux, Windows, macOS)
    try:
        import dxpdf
        return dxpdf.convert(docx_bytes)
    except Exception:
        pass

    # 2. Fallback on Windows: Microsoft Word COM automation
    if sys.platform == "win32":
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

    raise RuntimeError("No PDF converter available on this platform.")
