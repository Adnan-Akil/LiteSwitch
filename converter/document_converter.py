"""Contains logic for all the document conversions."""

import os
import win32com.client as com
from pdf2docx import parse
import pypandoc

def docx_to_pdf(input_path):
    '''Converts a DOCX file to PDF using MS WORD COM interface.'''
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


def docx_to_odt(input_path):    #lossy
    '''Converts DOCX to ODT format using pypandoc'''
    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_LiteSwitch.odt"
    pypandoc.convert_file(input_path, 'odt', outputfile=output_path)
    print(f"Converted {input_path} to {output_path}")



def docx_to_txt(input_path):
    '''Converts DOCX to TXT using pypandoc'''
    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_LiteSwitch.txt"
    pypandoc.convert_file(input_path, 'plain', outputfile=output_path)
    print(f"Converted {input_path} to {output_path}")


def docx_to_md(input_path):
    '''Converts DOCX to Markdown using pypandoc'''
    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_LiteSwitch.md"
    pypandoc.convert_file(input_path, to='markdown', outputfile=output_path)
    print(f"Converted {input_path} to {output_path}")


def docx_to_latex(input_path):
    '''Converts DOCX to LaTeX using pypandoc'''
    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_LiteSwitch.tex"
    pypandoc.convert_file(input_path, to='latex', outputfile=output_path)
    print(f"Converted {input_path} to {output_path}")


def docx_to_html(input_path):
    base, _ = os.path.splitext(input_path)
    output_path = f"{base}_LiteSwitch.html"
    pypandoc.convert_file(input_path, to='html', outputfile=output_path)
    print(f"Converted {input_path} to {output_path}")
    

def pdf_to_docx(input_path):    #lossy
    '''Convert PDF to DOCX using pdf2docx'''
    try:
        print("creating output path")
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.docx"
        print("converting file to pdf")
        parse(input_path, output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to DOCX: {e}")


def pdf_to_jpeg(input_path):
    pass


def pdf_to_txt(input_path):
    pass


def pdf_to_html(input_path):
    pass


def pdf_to_md(input_path):
    pass


CONVERSION_MAP = {
    "docx": {
        "pdf": docx_to_pdf,
        "odt": docx_to_odt,
        "txt": docx_to_txt,
        "md": docx_to_md,
        "latex": docx_to_latex,
        "html": docx_to_html,
    },
    "pdf": {
        "docx": pdf_to_docx,
        "jpeg": pdf_to_jpeg,
        "txt": pdf_to_txt,
        "html": pdf_to_html,
        "md": pdf_to_md,
    },
}
