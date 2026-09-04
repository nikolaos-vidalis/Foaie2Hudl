"""Convert in-memory .docx bytes to PDF bytes across platforms.

Tries conversion engines in order:
1. Microsoft Word COM automation (Windows only, exact template fidelity)
2. LibreOffice / soffice CLI (Linux / Streamlit Cloud with packages.txt)
3. dxpdf (pure Python/Rust cross-platform converter)
"""

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def convert(docx_bytes: bytes) -> bytes:
    """Return *docx_bytes* rendered to PDF.

    Tries Windows COM first, then LibreOffice headless, then dxpdf.
    """
    errors = []

    # 1. Windows: Microsoft Word COM automation
    if sys.platform == "win32":
        try:
            return _convert_word_com(docx_bytes)
        except Exception as exc:
            errors.append(f"Word COM: {exc}")

    # 2. Linux / macOS / Windows with LibreOffice installed
    try:
        return _convert_libreoffice(docx_bytes)
    except Exception as exc:
        errors.append(f"LibreOffice: {exc}")

    # 3. Pure Python / Rust fallback (dxpdf)
    try:
        import dxpdf
        return dxpdf.convert(docx_bytes)
    except Exception as exc:
        errors.append(f"dxpdf: {exc}")

    raise RuntimeError(" | ".join(errors) if errors else "No PDF converter available.")


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


def _convert_libreoffice(docx_bytes: bytes) -> bytes:
    soffice = (
        shutil.which("soffice")
        or shutil.which("libreoffice")
        or next((p for p in ["/usr/bin/soffice", "/usr/bin/libreoffice"] if Path(p).exists()), None)
    )
    if not soffice:
        raise RuntimeError("LibreOffice soffice binary not found.")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        in_file = tmp_path / "teamsheet.docx"
        in_file.write_bytes(docx_bytes)

        cmd = [
            str(soffice),
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
                f"Output PDF not generated: {result.stderr.decode('utf-8', errors='ignore')}"
            )
        return out_file.read_bytes()
