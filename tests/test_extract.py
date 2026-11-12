from __future__ import annotations

import io
import zipfile

from outlook_web_mcp.extract import text_from_attachment


def _zip(files: dict[str, str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, body in files.items():
            archive.writestr(name, body)
    return buffer.getvalue()


def test_docx_text_is_extracted_from_the_zip() -> None:
    data = _zip(
        {
            "word/document.xml": (
                "<w:document><w:body><w:p><w:r><w:t>Dear Registry</w:t></w:r></w:p>"
                "<w:p><w:r><w:t>Fees due 12 Nov</w:t></w:r></w:p></w:body></w:document>"
            )
        }
    )
    text, method = text_from_attachment("letter.docx", "", data)
    assert method == "docx-xml"
    assert "Dear Registry" in text and "Fees due 12 Nov" in text
    assert "\n" in text  # paragraphs become line breaks


def test_xlsx_cells_are_joined_per_row() -> None:
    data = _zip(
        {
            "xl/sharedStrings.xml": "<sst><si><t>Tuition</t></si><si><t>21000</t></si></sst>",
            "xl/worksheets/sheet1.xml": (
                "<worksheet><sheetData><row><c><v>0</v></c><c><v>1</v></c></row>"
                "</sheetData></worksheet>"
            ),
        }
    )
    text, method = text_from_attachment("fees.xlsx", "", data)
    assert method == "xlsx-xml"
    assert text == "Tuition\t21000"


def test_plain_text_and_unknown_types() -> None:
    assert text_from_attachment("note.txt", "", b"hello") == ("hello", "plain-text")
    assert text_from_attachment("blob.bin", "application/octet-stream", b"\x00\x01") == (
        "",
        "unsupported-type",
    )


def test_corrupt_office_file_fails_without_raising() -> None:
    text, method = text_from_attachment("broken.docx", "", b"not a zip")
    assert (text, method) == ("", "extract-failed")


def test_js_rendered_hint_separates_markup_from_empty() -> None:
    from outlook_web_mcp.parsers import looks_js_rendered

    js = "<html><body>" + "<script>load()</script>" * 20 + "<div id='app'></div></body></html>"
    assert looks_js_rendered(js, "") is True
    assert looks_js_rendered(js, "x" * 200) is False  # real text present
    assert looks_js_rendered("", "") is False  # empty is not "JS rendered"
    assert looks_js_rendered("<p>hi</p>", "hi") is False
