# convert rtf to docx and embed all pictures in the final document
from win32com import client


from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants


def change_word_format(file_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    new_file_abs = os.path.abspath(file_path)

    doc = word.Documents.Open(new_file_abs)
    doc.Activate()

    # Rename path with .doc
    new_file_abs = new_file_abs.split(".")[0] + ".docx"

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatDocument
    )
    doc.Close(False)
# change_word_format('ABLE Account Information.rtf')

def ConvertRtfToDocx(file):
    word = client.Dispatch("Word.Application")
    new_file_abs = os.path.abspath(file)

    doc = word.Documents.Open(new_file_abs)

    doc.SaveAs(new_file_abs.split(".")[0] + ".doc",)
    doc.Close()
    word.Quit()

# ConvertRtfToDocx('ABLE Account Information.rtf')


from docx2pdf import convert

def docx_to_pdf(ori_file, new_file):
     convert(ori_file, new_file)

# docx_to_pdf('ABLE Account Information.doc', 'ABLE Account Information.pdf')


from docx2pdf import convert

def convert_doc_to_pdf(input_file, output_file):
  """Converts a DOC file to PDF using docx2pdf (paid library)."""

  # Convert the DOC file to PDF
  convert(input_file, output_file)

  print("DOC file converted to PDF successfully!")

# Example usage (assuming docx2pdf is installed)
input_file = "ABLE Account Information.docx"
output_file = "output.pdf"

# convert_doc_to_pdf(input_file, output_file)


import pypandoc

def convert_rtf_to_pdf(input_rtf, output_pdf):
    pypandoc.convert_text(source=input_rtf, to='pdf', format='rtf', outputfile=output_pdf)

input_rtf_path = "ABLE Account Information.rtf"
output_pdf_path = "output.pdf"

convert_rtf_to_pdf(input_rtf_path, output_pdf_path)



