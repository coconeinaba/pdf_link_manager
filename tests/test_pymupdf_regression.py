from __future__ import annotations

import hashlib
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest

import pymupdf as fitz

import app


def build_overprint_pdf(path: Path) -> None:
    """PyMuPDF Issue #5054と同じ文字重ね描画を持つ最小PDFを作る。"""
    content = (
        b"BT\n"
        b"/F1 24 Tf\n"
        b"1 0 0 1 100 700 Tm\n"
        b"(f)Tj\n"
        b"0 0 TD\n"
        b"(faeces)Tj\n"
        b"ET\n"
    )
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 800] "
            b"/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>"
        ),
        b"<< /Length %d >>\nstream\n%sendstream" % (len(content), content),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]

    pdf = b"%PDF-1.7\n"
    offsets: list[int] = []
    for number, body in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf += b"%d 0 obj\n%s\nendobj\n" % (number, body)
    xref_position = len(pdf)
    pdf += b"xref\n0 %d\n" % (len(objects) + 1)
    pdf += b"0000000000 65535 f \n"
    for offset in offsets:
        pdf += b"%010d 00000 n \n" % offset
    pdf += (
        b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n"
        % (len(objects) + 1, xref_position)
    )
    path.write_bytes(pdf)


def page_content_sha256(path: Path) -> str:
    """1ページ目のコンテンツストリームをSHA-256で比較可能にする。"""
    with fitz.open(path) as document:
        return hashlib.sha256(document[0].read_contents()).hexdigest()


class PyMuPDFSafetyTests(unittest.TestCase):
    def test_required_version_is_active(self) -> None:
        """実行環境が安全確認済み版であり、版検査を通過する。"""
        self.assertEqual(fitz.VersionBind, app.REQUIRED_PYMUPDF_VERSION)
        app.require_pymupdf()

    def test_newer_version_is_rejected(self) -> None:
        """未検証の1.28系を検出した場合はPDFを処理しない。"""
        original = app.fitz
        app.fitz = SimpleNamespace(VersionBind="1.28.2")
        try:
            with self.assertRaisesRegex(app.ValidationError, "1.27.2.3"):
                app.require_pymupdf()
        finally:
            app.fitz = original

    def test_link_save_preserves_overprinted_text_and_content(self) -> None:
        """リンク追加後もIssue #5054の文字と元コンテンツを保持する。"""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "overprint-source.pdf"
            output = root / "overprint-linked.pdf"
            build_overprint_pdf(source)
            source_content_hash = page_content_sha256(source)

            rule = app.LinkRule(
                name="URLリンク",
                page_spec="1",
                x=20,
                y=20,
                width=100,
                height=24,
                link_type="url",
                url="https://example.com/",
            )
            applied, warnings = app.apply_rules_to_pdf(source, output, [rule])

            self.assertEqual(applied, 1)
            self.assertEqual(warnings, [])
            self.assertEqual(page_content_sha256(output), source_content_hash)
            with fitz.open(output) as document:
                page = document[0]
                self.assertEqual(page.get_text().strip(), "faeces")
                links = page.get_links()
                self.assertEqual(len(links), 1)
                self.assertEqual(links[0]["kind"], fitz.LINK_URI)
                self.assertEqual(links[0]["uri"], "https://example.com/")


if __name__ == "__main__":
    unittest.main()
