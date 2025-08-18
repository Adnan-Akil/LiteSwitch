"""Contains logic for all the document conversions."""

import os
import win32com.client as com
from pdf2docx import parse

# CONVERSION_OPTIONS = {"docx": ["pdf"]}


def docx_to_pdf(input_path):
    try:
        word = com.Dispatch("Word.Application")
        word.Visible = False  # runs word in bg
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.pdf"
        doc = word.Documents.Open(
            input_path, False, False, False
        )  # do not show dialogue box for conversions; open dialogue box in read/write mode; dont add to recent files
        word_formatpdf = 17  # wdFormatPDF
        doc.ExportAsFixedFormat(output_path, word_formatpdf)
        doc.Close(False)  # close the document without saving changes
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to PDF: {e}")
    finally:
        word.Quit()


def pdf_to_docx(input_path):
    try:
        print("creating output path")
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.docx"
        print("converting file to pdf")
        parse(input_path, output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to DOCX: {e}")


CONVERSION_MAP = {"docx": {"pdf": docx_to_pdf}, "pdf": {"docx": pdf_to_docx}}


def testing_block():
    test_docx_path = r"C:\Users\hyped\Desktop\LiteSwitch\SE_laiba.docx"
    test_pdf_path = r"C:\Users\hyped\Desktop\LiteSwitch\DeloitteCertificate.pdf"
    print("Running Docx to PDF conversion test...")
    if os.path.exists(test_docx_path):
        converter = CONVERSION_MAP["docx"].get("pdf")
        if converter:
            converter(test_docx_path)
        else:
            print("Failed")
    else:
        print("File not found")

    print("Running PDF to Docx conversion test...")
    if os.path.exists(test_pdf_path):
        converter = CONVERSION_MAP["pdf"].get("docx")
        if converter:
            converter(test_pdf_path)
        else:
            print("Failed")
    else:
        print("File not found")


if __name__ == "__main__":
    testing_block()
