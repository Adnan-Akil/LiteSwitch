"""Contains logic for all the document conversions."""

import os
import io
import sys

import pypandoc  # most docx conversions
import win32com.client as com
from pdf2docx import parse
from pdfminer.high_level import extract_text  # pdf to txt
from unstructured.partition.pdf import partition_pdf  # pdf to md
import fitz
from PIL import Image  # for image conversions
from pptx import Presentation  
from pptx.util import Inches, Pt

def docx_to_pdf(input_path):
    """Converts a DOCX file to PDF using MS WORD COM interface."""
    try:
        word = com.Dispatch("Word.Application")
        word.Visible = False  # runs word in bg
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.pdf"
        doc = word.Documents.Open(
            input_path, False, False, False
        )  # do not show dialogue box for conversions; 
            #open dialogue box in read/write mode; dont add to recent files
        word_formatpdf = 17  # wdFormatPDF
        doc.ExportAsFixedFormat(output_path, word_formatpdf)
        doc.Close(False)  # close the document without saving changes
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to PDF: {e}")
    finally:
        word.Quit()


def docx_to_odt(input_path):  # lossy
    """Converts DOCX to ODT format using pypandoc"""
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.odt"
        pypandoc.convert_file(input_path, "odt", outputfile=output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to ODT: {e}")


def docx_to_txt(input_path):  # lossy
    """Converts DOCX to TXT using pypandoc"""
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.txt"
        pypandoc.convert_file(input_path, "plain", outputfile=output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to TXT: {e}")


def docx_to_md(input_path):
    """Converts DOCX to Markdown using pypandoc"""
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.md"
        pypandoc.convert_file(input_path, to="markdown", outputfile=output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to Markdown: {e}")


def docx_to_latex(input_path):
    """Converts DOCX to LaTeX using pypandoc"""
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.tex"
        pypandoc.convert_file(input_path, to="latex", outputfile=output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to LaTeX: {e}")


def docx_to_html(input_path):  # lossy
    """Convert DOCX to HTML using pypandoc"""
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.html"
        pypandoc.convert_file(input_path, to="html", outputfile=output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to HTML: {e}")


def pdf_to_docx(input_path):  # lossy
    """Convert PDF to DOCX using pdf2docx"""
    try:
        print("creating output path")
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.docx"
        print("converting file to pdf")
        parse(input_path, output_path)
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to DOCX: {e}")


def pdf_to_txt(input_path):
    '''Convert PDF to TXT using pdfminer.six'''
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.txt"
        text= extract_text(input_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print("Successfully converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to TXT: {e}")

def pdf_to_png(input_path): #implement threadpool for this
    '''Converts PDF to PNG using fitz'''
    try:
        base, _ = os.path.splitext(input_path)
        doc=fitz.open(input_path)
        for page_number,page in enumerate(doc, start=1):
            pix= page.get_pixmap(matrix=fitz.Matrix(3, 3))  # increase resolution
            output_path = f"{base}_LiteSwitch_page_{page_number}.png"
            pix.save(output_path)
        print(f"Converted {input_path} to PNG images.")
        doc.close() 
    except Exception as e:
        print(f"Error converting {input_path} to PNG: {e}")

def pdf_to_html(input_path):    #lossy
    '''Converts PDF to HTML using pymupdf (fitz)'''
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.html"
        doc=fitz.open(input_path)
        with open(output_path, 'w', encoding='utf-8') as f:
            for page in doc:
                f.write(page.get_text("html") + '\n')
        doc.close()
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to HTML: {e}")

def pdf_to_pptx(input_path):    #converts to image and then to pptx => cannot edit text, only for showcase purposes
    try:
        print("Creating a new Presentation")
        prs= Presentation()
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        slide_ratio = slide_width / slide_height
        doc= fitz.open(input_path)

        for page_number in range(doc.page_count):
            page= doc.load_page(page_number)
            pix= page.get_pixmap(matrix=fitz.Matrix(3, 3))
            img_stream= io.BytesIO(pix.tobytes("png"))

            blank_slide= prs.slide_layouts[6]
            slide= prs.slides.add_slide(blank_slide)

            img_width = pix.width
            img_height = pix.height
            img_ratio = img_width / img_height

            if img_ratio > slide_ratio: #image wider than slide
                new_width = slide_width
                new_height = int(new_width / img_ratio)
                left= Inches(0)
                top = (slide_height - new_height) / 2
            else:  #image taller than slide
                new_height = slide_height
                new_width = int(new_height * img_ratio)
                top = Inches(0)
                left = (slide_width - new_width) / 2
            slide.shapes.add_picture(img_stream, left, top, width=new_width, height=new_height)
            print(f"Added page {page_number + 1} to the presentation.")
        prs.save(f"{os.path.splitext(input_path)[0]}_LiteSwitch.pptx")
        print(f"Converted {input_path} to PPTX.")
    except Exception as e:
        print(f"Error converting {input_path} to PPTX: {e}")

def pdf_to_md(input_path):  #implement threadpool for this
    '''Converts PDF to Markdown using unstructured'''
    try:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_LiteSwitch.md"
        elements= partition_pdf(filename=input_path, strategy="auto")
        with open(output_path, 'w', encoding='utf-8') as f:
            for element in elements:
                f.write(str(element) + '\n\n')
        print(f"Converted {input_path} to {output_path}")
    except Exception as e:
        print(f"Error converting {input_path} to Markdown: {e}")


def png_to_pdf(input_path): #will configure this, and other functions to accept multiple files in the future
    try:
        print("opening the first image")
        first_image= Image.open(input_path).convert("RGB")
        first_image.save(f"{os.path.splitext(input_path)[0]}_LiteSwitch.pdf", save_all=True, dpi=(300, 300))
        print(f"Converted {input_path} to PDF.")
    except Exception as e:
        print(f"Error converting {input_path} to PDF: {e}")

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
        "png": pdf_to_png,
        "pptx": pdf_to_pptx,
        "txt": pdf_to_txt,
        "html": pdf_to_html,
        "md": pdf_to_md,
    },
    "png":{
        "pdf": png_to_pdf
    }
}
