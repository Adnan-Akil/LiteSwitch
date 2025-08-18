'''Contains logic for all the document conversions.'''

import os
import win32com.client as com
import pymupdf
import pypandoc

CONVERSION_OPTIONS = {"docx": ["pdf"]}
CONVERSION_MAP = {}


def docx_to_pdf(input_path):
    try:
        word = com.Dispatch("Word.Application")
        word.Visible = False  # runs word in bg
        base = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.pdf"
        doc = word.Documents.Open(
            input_path, False, False, False
        )  # do not show dialogue box for conversions; open dialogue box in read/write mode; dont add to recent files
        word_formatpdf = 17  # wdFormatPDF
        doc.SaveAs(output_path, FileFormat=word_formatpdf)
        doc.close()
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to PDF: {e}")
    finally:
        word.Quit()
