"""Turn an attachment into searchable text without keeping the file.

Nothing here writes a durable copy of the attachment: PDFs go through a temp
directory that is removed immediately, Office files are read straight out of the
zip container in memory. The caller caps size and decides what to keep.
"""

from __future__ import annotations

import io
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

PDF_TYPES = {"application/pdf", "application/x-pdf"}
DOCX_TYPES = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
}
XLSX_TYPES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
}
TEXT_SUFFIXES = {".txt", ".csv", ".md", ".json", ".html", ".htm", ".xml", ".ics"}

# OCR only the first pages: a 200-page scan is not worth minutes of CPU for a
# search hint, and the file itself is still on the server.
OCR_MAX_PAGES = 5
PDFTOTEXT_TIMEOUT = 60
OCR_TIMEOUT = 120
OCR_LANGS = "eng+chi_tra+chi_sim"


def _clean(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _run(command: list[str], *, timeout: int) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(command, capture_output=True, timeout=timeout, check=False)


def pdf_text(data: bytes) -> tuple[str, str]:
    """Text from a PDF: the text layer first, OCR for scans."""
    if shutil.which("pdftotext") is None:
        return "", "no-pdftotext"
    with tempfile.TemporaryDirectory(prefix="owa-att-") as work:
        source = Path(work) / "in.pdf"
        source.write_bytes(data)
        result = _run(
            ["pdftotext", "-layout", "-q", str(source), "-"],
            timeout=PDFTOTEXT_TIMEOUT,
        )
        text = _clean(result.stdout.decode("utf-8", errors="replace"))
        if len(text) >= 40:
            return text, "pdf-text-layer"
        if shutil.which("pdftoppm") is None or shutil.which("tesseract") is None:
            return text, "pdf-text-layer-short"
        # Scanned PDF: rasterise the first pages and OCR them.
        _run(
            [
                "pdftoppm",
                "-r",
                "200",
                "-l",
                str(OCR_MAX_PAGES),
                "-png",
                str(source),
                str(Path(work) / "page"),
            ],
            timeout=PDFTOTEXT_TIMEOUT,
        )
        chunks: list[str] = []
        for image in sorted(Path(work).glob("page*.png")):
            ocr = _run(
                ["tesseract", str(image), "stdout", "-l", OCR_LANGS],
                timeout=OCR_TIMEOUT,
            )
            chunks.append(ocr.stdout.decode("utf-8", errors="replace"))
        return _clean("\n".join(chunks)) or text, "pdf-ocr"


def _docx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        raw = archive.read("word/document.xml").decode("utf-8", errors="replace")
    raw = raw.replace("</w:p>", "\n").replace("<w:tab/>", "\t").replace("<w:br/>", "\n")
    return _clean(re.sub(r"<[^>]+>", "", raw))


def _xlsx_text(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            raw = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="replace")
            for chunk in re.findall(r"<si>(.*?)</si>", raw, flags=re.S):
                shared.append(re.sub(r"<[^>]+>", "", chunk))
        sheets: list[str] = []
        for name in sorted(n for n in archive.namelist() if n.startswith("xl/worksheets/sheet")):
            raw = archive.read(name).decode("utf-8", errors="replace")
            for row in re.findall(r"<row[^>]*>(.*?)</row>", raw, flags=re.S):
                cells: list[str] = []
                for cell in re.findall(r"<c[^>]*>(.*?)</c>", row, flags=re.S):
                    inline = re.findall(r"<t[^>]*>(.*?)</t>", cell, flags=re.S)
                    if inline:
                        cells.append("".join(inline))
                    else:
                        index = re.search(r"<v>(\d+)</v>", cell)
                        if index and int(index.group(1)) < len(shared):
                            cells.append(shared[int(index.group(1))])
                if any(cells):
                    sheets.append("\t".join(cells))
    return _clean("\n".join(sheets))


def text_from_attachment(name: str, content_type: str, data: bytes) -> tuple[str, str]:
    """Return (text, method). Empty text with a method label is a normal outcome."""
    suffix = Path(name or "").suffix.lower()
    kind = (content_type or "").lower()
    try:
        if kind in PDF_TYPES or suffix == ".pdf":
            return pdf_text(data)
        if kind in DOCX_TYPES or suffix == ".docx":
            return _docx_text(data), "docx-xml"
        if kind in XLSX_TYPES or suffix == ".xlsx":
            return _xlsx_text(data), "xlsx-xml"
        if suffix in TEXT_SUFFIXES or kind.startswith("text/"):
            return _clean(data.decode("utf-8", errors="replace")), "plain-text"
    except (zipfile.BadZipFile, KeyError, subprocess.TimeoutExpired, OSError, ValueError):
        return "", "extract-failed"
    return "", "unsupported-type"
