import unittest
import os
import shutil
import logging
import sys
from unittest.mock import MagicMock, patch

# Mock win32com before importing converter
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["pythoncom"] = MagicMock()

# Now import the module under test
from converter import document_converter

# Setup test logger
logging.basicConfig(level=logging.INFO)
TEST_DIR = "test_artifacts"

class TestLiteSwitch(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """Create a temporary directory for test artifacts."""
        if not os.path.exists(TEST_DIR):
            os.makedirs(TEST_DIR)
        
        # Create dummy source files
        cls.docx_file = os.path.join(TEST_DIR, "test.docx")
        cls.pdf_file = os.path.join(TEST_DIR, "test.pdf")
        cls.txt_file = os.path.join(TEST_DIR, "test.txt")
        
        with open(cls.txt_file, "w") as f:
            f.write("Hello LiteSwitch")

    @classmethod
    def tearDownClass(cls):
        """Clean up test artifacts."""
        if os.path.exists(TEST_DIR):
            shutil.rmtree(TEST_DIR)

    def test_conversion_map_integrity(self):
        """Ensure CONVERSION_MAP is valid."""
        self.assertIn("docx", document_converter.CONVERSION_MAP)
        self.assertIn("pdf", document_converter.CONVERSION_MAP)
        
    def test_pdf_functions_exist(self):
        """Check if PDF conversion functions are callable."""
        pdf_map = document_converter.CONVERSION_MAP["pdf"]
        self.assertTrue(callable(pdf_map["docx"]))
        self.assertTrue(callable(pdf_map["pptx"]))

    def test_basic_structure(self):
        """Check if critical components are present."""
        self.assertTrue(hasattr(document_converter, "pdf_to_pptx"))
        self.assertTrue(hasattr(document_converter, "logger"))

if __name__ == "__main__":
    unittest.main()
