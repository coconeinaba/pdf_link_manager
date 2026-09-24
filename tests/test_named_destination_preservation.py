from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

import pymupdf as fitz

import app


class NamedDestinationPreservationTests(unittest.TestCase):
    def test_imported_named_destination_is_preserved_when_conversion_is_selected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "named-destination.pdf"
            output = root / "named-destination-linked.pdf"

            with fitz.open() as document:
                document.new_page()
                document.new_page()
                source_page = document[0]
                target_xref = document.page_xref(1)
                names_xref = document.get_new_xref()
                document.update_object(
                    names_xref,
                    f"<</Names[(Chapter2)[{target_xref} 0 R/XYZ 72 700 1.25]]>>",
                )
                document.xref_set_key(
                    document.pdf_catalog(), "Names", f"<</Dests {names_xref} 0 R>>"
                )

                source_page.insert_link(
                    {
                        "kind": fitz.LINK_URI,
                        "from": fitz.Rect(10, 10, 100, 35),
                        "uri": "about:blank",
                    }
                )
                source_page = document.reload_page(source_page)
                link_xref = app._link_annotation_xrefs(document, source_page)[0]
                document.xref_set_key(link_xref, "A", "<</S/GoTo/D(Chapter2)>>")
                document.xref_set_key(link_xref, "Dest", "null")
                document.save(source)

            with fitz.open(source) as document:
                rules = app.existing_links_as_rules(document)

            self.assertEqual(len(rules), 1)
            self.assertEqual(rules[0].target_page, "Chapter2")
            self.assertEqual(rules[0].source_named_destination, "Chapter2")

            applied, warnings = app.apply_rules_to_pdf(
                source,
                output,
                rules,
                named_destination_output="convert",
            )

            self.assertEqual(applied, 1)
            self.assertEqual(warnings, [])
            with fitz.open(output) as document:
                link_xref = app._link_annotation_xrefs(document, document[0])[0]
                self.assertIn("Chapter2", document.xref_get_key(link_xref, "A/D")[1])
                self.assertEqual(document.resolve_names()["Chapter2"]["page"], 1)


if __name__ == "__main__":
    unittest.main()
