"""Convert in-memory .docx bytes to PDF bytes.

Uses LibreOffice (soffice) when available (e.g. Streamlit Cloud via packages.txt
or Windows/Linux/macOS with LibreOffice installed).

On Windows, seamlessly falls back to Microsoft Word COM if LibreOffice is not installed.
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

STANDARD_SOFFICE_PATHS = [
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def convert(docx_bytes: bytes) -> bytes:
    """Return *docx_bytes* rendered to PDF.

    Tries LibreOffice first, then falls back to Microsoft Word COM on Windows.
    """
    # 1. Try LibreOffice (works on Cloud via packages.txt and locally if installed)
    soffice = (
        shutil.which("soffice")
        or shutil.which("libreoffice")
        or next((p for p in STANDARD_SOFFICE_PATHS if Path(p).exists()), None)
    )
    if soffice:
        try:
            return _convert_libreoffice(soffice, docx_bytes)
        except Exception:
            pass

    # 2. Fall back to Microsoft Word COM on Windows
    if sys.platform == "win32":
        try:
            return _convert_word_com(docx_bytes)
        except Exception as exc:
            raise RuntimeError(f"Word COM conversion failed: {exc}") from exc

    raise RuntimeError(
        "No PDF converter available. (LibreOffice not found and Word COM unavailable on non-Windows)."
    )


def _convert_libreoffice(soffice_bin: str, docx_bytes: bytes) -> bytes:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        in_file = tmp_path / "teamsheet.docx"
        in_file.write_bytes(docx_bytes)

        cmd = [
            str(soffice_bin),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(in_file),
        ]
        result = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=30
        )
        out_file = tmp_path / "teamsheet.pdf"
        if not out_file.exists():
            raise RuntimeError(
                f"LibreOffice output missing: {result.stderr.decode('utf-8', errors='ignore')}"
            )
        return out_file.read_bytes()


def _convert_word_com(docx_bytes: bytes) -> bytes:
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
            word.DisplayAlerts = 0

            doc = word.Documents.Open(str(docx_file.resolve()), ReadOnly=True)
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_file.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
            )

            pdf_bytes = pdf_file.read_bytes()

            doc.Close(SaveChanges=False)
            doc = None
            word.Quit()
            word = None

            return pdf_bytes
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
