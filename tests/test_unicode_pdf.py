"""Regression tests for embedded Unicode fonts in SlideCanon's A4 PDF."""
import tempfile
import unittest
from pathlib import Path

import slidecanon
from reportlab.pdfbase import pdfmetrics


class UnicodePdfTests(unittest.TestCase):
    def test_embedded_font_covers_pl_and_equations(self):
        names = slidecanon.register_pdf_fonts()
        sample = "Zażółć gęślą jaźń ĄĆĘŁŃÓŚŹŻ ąęóśłżźćń H₂O x⁻¹ x² log₂ ∑ ∫ ≤ ≥ ≈ αβ"
        for font in names.values():
            cmap = pdfmetrics.getFont(font).face.charToGlyph
            self.assertTrue(all(char == " " or ord(char) in cmap for char in sample), font)
        self.assertTrue(all(name.startswith("SlideCanonDejaVu") for name in names.values()))

    def test_unavailable_character_fails_not_black_box(self):
        slidecanon.register_pdf_fonts()
        with self.assertRaisesRegex(ValueError, r"U\+1F984"):
            slidecanon.escape("Unsupported 🦄")

    def test_optional_noto_math_fallback(self):
        slidecanon.register_pdf_fonts()
        if slidecanon._PDF_MATH_FONT is None:
            self.skipTest("Noto Sans Math not installed")
        self.assertIn('font name="SlideCanonNotoMath"', slidecanon.escape("𝑥"))

    def test_a4_pdf_compiles_with_unicode(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "unicode.md"
            source.write_text(
                "#slide 1\n## layout\nbullets\n"
                "## slide title\nZażółć gęślą jaźń — x² + H₂O\n"
                "## bullets\n- Równanie: log₂(x), x⁻¹, ∑, √, ≤, α\n"
                "## teleprompter\nPrzykład: H₂O i współczynnik x²; ąęóśłżźćń.\n",
                encoding="utf-8",
            )
            output = Path(temp) / "unicode.pdf"
            slidecanon.build_pdf_document(source, output)
            self.assertGreater(output.stat().st_size, 1000)
            self.assertIn(b"DejaVuSans", output.read_bytes())


if __name__ == "__main__":
    unittest.main()
