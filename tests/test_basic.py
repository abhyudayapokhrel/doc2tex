# Simple test script for doc2tex
# Checks if the converter can load and run basic functions

import os
import unittest
from doc2tex import DocTeXConverter, ConversionOptions

class TestConverter(unittest.TestCase):
    def setUp(self):
        # Setup common objects for tests
        self.options = ConversionOptions()
        self.converter = DocTeXConverter(self.options)
        self.test_dir = os.path.dirname(__file__)
    
    def test_options_init(self):
        # Check if options load correctly
        self.assertEqual(self.options.font_size.value, "12pt")
        self.assertTrue(self.options.standalone_document)

    def test_detect_to_latex(self):
        # Check if it detects docx files correctly
        dir = self.converter._detect_direction("test.docx")
        self.assertEqual(dir, "to_latex")

    def test_detect_to_docx(self):
        # Check if it detects latex files correctly
        dir = self.converter._detect_direction("test.tex")
        self.assertEqual(dir, "to_docx")

if __name__ == '__main__':
    unittest.main()
