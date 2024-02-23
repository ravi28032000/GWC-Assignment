# import PyPDF2,docx2pdf
# from aspose.pdf.text import TextAbsorber
# import docx2pdf as docx2pdf
# from docx import Document
from pdf2docx import Converter
# import PyPDF2
# import textwrap

pdf_file_path="C:\\Users\\RAVI\\Documents\\MSD_WORD.pdf"     # To Define The PDF File Path Here
doc_file_path="C:\\Users\\RAVI\\Documents\\MSD_WORD.docx"
import pdfplumber

def read_pdf_characters_with_attributes(pdf_file_path):
    characters_with_attributes = []
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            append_data=[]
            char_name=""
            for page in pdf.pages:
                for char in page.chars:
                    text = char['text']
                    font_properties = char['fontname']
                    bold = "Bold" in font_properties
                    italic = "Italic" in font_properties
                    underline = char['height'] < 0  # Approximate way to detect underline
                    page_number = char['page_number']
                    characters_with_attributes.append( {"text": text, "bold": bold, "italic": italic, "underline": underline,
                                                        "no": page_number})
        return characters_with_attributes
    except Exception as error:
        return {"error":str(error),"err_line":str(error.__traceback__.tb_lineno)}

# Example usage
pdf_file_path = pdf_file_path
characters_with_attributes_pdf = read_pdf_characters_with_attributes(pdf_file_path)
if isinstance(characters_with_attributes_pdf,dict):
    print(characters_with_attributes_pdf)
    print("INVALID")

from docx import Document

def read_docx_characters_with_attributes(docx_file_path):
    try:
        doc = Document(docx_file_path)
        characters_with_attributes = []
        line_number = 0
        for paragraph in doc.paragraphs:


            words = paragraph.text.split()  # Split paragraph text into words
            for word in words:
                for char in word:
                    line_number += 1
                    for run in paragraph.runs:
                        bold = run.bold
                        italic = run.italic
                        underline = run.underline
                        font_size = run.font.size.pt if run.font.size else None
                        font_name = run.font.name
                        characters_with_attributes.append(
                            {"text": char, "bold": bold, "italic": italic, "underline": underline, "word": word,
                             "font_size": font_size, "font_name": font_name, "no": line_number})
                        break

        return characters_with_attributes

    except Exception as error:

        return {"error": str(error), "err_line": str(error.__traceback__.tb_lineno)}

# Example usage
docx_file_path = doc_file_path
characters_with_attributes_doc = read_docx_characters_with_attributes(docx_file_path)
if isinstance(characters_with_attributes_doc,dict):
    print(characters_with_attributes_doc)
    print("INVALID_DATA")
print("LEN_PDF="+str(len(characters_with_attributes_pdf)))
print("LEN_DOC="+str(len(characters_with_attributes_doc)))
print("LEN_PDF="+str((characters_with_attributes_pdf[:100])))
print("LEN_DOC="+str((characters_with_attributes_doc[:100])))
disc_data=[]
for i,j in zip(characters_with_attributes_pdf,characters_with_attributes_doc):
    if (i.get('bold') == True and j.get('bold') == None) or (i.get('bold') == False and j.get('bold') == True):###None True
        disc_data.append({"text": i.get('text'), 'docx_bold': i.get('bold'), "pdf_bold": j.get("bold")})
    if (i.get('italic') == True and j.get('bold') == None) or ( i.get('italic') == False and j.get('italic') == True):  # underline
        disc_data.append({"text": i.get('text'), 'docx_italic': i.get('italic'), "pdf_italic": j.get("italic")})

    if (i.get('underline') == True and j.get('underline') == None) or ( i.get('underline') == False and j.get('underline') == True):
        disc_data.append( {"text": i.get('text'), 'docx_underline': i.get('underline'), "pdf_underline": j.get("underline")})
print(disc_data)
print(len(disc_data))