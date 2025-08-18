import os
from document_converter import CONVERSION_MAP

def testing_block():
    #paths
    # test_docx_path = r"C:\Users\hyped\Desktop\LiteSwitch\SE_laiba.docx"
    # test_pdf_path = r"C:\Users\hyped\Desktop\LiteSwitch\DeloitteCertificate.pdf"
    # test_odt_path = r"C:\Users\hyped\Desktop\LiteSwitch\SE_laiba.docx"
    
    # #DOCX TO PDF
    # print("Running Docx to PDF conversion test...")
    # if os.path.exists(test_docx_path):
    #     converter = CONVERSION_MAP["docx"].get("pdf")
    #     if converter:
    #         converter(test_docx_path)
    #     else:
    #         print("Failed")
    # else:
    #     print("File not found")

    # #PDF TO DOCX
    # print("Running PDF to Docx conversion test...")
    # if os.path.exists(test_pdf_path):
    #     converter = CONVERSION_MAP["pdf"].get("docx")
    #     if converter:
    #         converter(test_pdf_path)
    #     else:
    #         print("Failed")
    # else:
    #     print("File not found")
    
    #DOCX TO ODT
    # print("Running PDF to Docx conversion test...")
    # if os.path.exists(test_odt_path):
    #     converter = CONVERSION_MAP["docx"].get("odt")
    #     if converter:
    #         converter(test_odt_path)
    #     else:
    #         print("Failed")
    # else:
    #     print("File not found")


if __name__ == "__main__":
    testing_block()