from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pymupdf as fitz

import app


def build_pdf_with_empty_action(path: Path) -> None:
    """空の /A と有効なURIリンクを同じページへ作る。"""
    with fitz.open() as document:
        page = document.new_page(width=300, height=300)
        for index, uri in enumerate(
            ("https://one.example/", "https://empty.example/", "https://two.example/")
        ):
            page.insert_link(
                {
                    "kind": fitz.LINK_URI,
                    "from": fitz.Rect(10 + index * 80, 10, 60 + index * 80, 30),
                    "uri": uri,
                }
            )
        document.save(path)
    with fitz.open(path) as document:
        xrefs = app._link_annotation_xrefs(document, document[0])
        assert len(xrefs) == 3
        document.xref_set_key(xrefs[1], "A", "<<>>")
        document.saveIncr()


class EmptyLinkActionTests(unittest.TestCase):
    def test_other_links_are_imported_and_replaced_when_one_action_is_empty(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "empty-action.pdf"
            output = Path(directory) / "empty-action-linked.pdf"
            build_pdf_with_empty_action(source)
            with fitz.open(source) as document:
                with patch.object(app, "_safe_page_get_links", return_value=[]):
                    rules = app.existing_links_as_rules(document)
            self.assertEqual(
                [rule.url for rule in rules],
                ["https://one.example/", "https://two.example/"],
            )
            applied, warnings = app.apply_rules_to_pdf(source, output, rules)
            self.assertEqual(applied, 2)
            self.assertFalse(warnings)
            with fitz.open(output) as document:
                records = app._page_link_records(document, 0)
                self.assertEqual(len(records), 3)
                self.assertEqual(
                    sorted(
                        record["link"].get("uri")
                        for record in records
                        if record["link"].get("uri")
                    ),
                    ["https://one.example/", "https://two.example/"],
                )


if __name__ == "__main__":
    unittest.main()
