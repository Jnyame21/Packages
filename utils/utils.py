import os
import magic
import win32com.client
from docx2pdf import convert
import time
import PyPDF2
import getpass
from gtts import gTTS
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# get user's computer username and set the path to the desktop
username = getpass.getuser()
local_path = f"C:\\Users\\{username}\\Desktop\\"


def text_to_pdf(text_file_name_with_extension, save_file_name, print_output=False):
    text_file_path = f"{local_path}{text_file_name_with_extension}"
    pdf_file_path = f"{local_path}{save_file_name}.pdf"

    # Create a new PDF canvas with the specified output file
    c = canvas.Canvas(pdf_file_path, pagesize=letter)

    # Open the text file and read its content
    with open(text_file_path, 'r') as file:
        content = file.read()

    # Set the font and font size for the PDF
    c.setFont("Helvetica", 12)

    # Specify the position to start drawing the text
    x = 10
    y = letter[1] - 50

    # Split the content into lines and draw each line on the PDF canvas
    for line in content.splitlines():
        c.drawString(x, y, line)
        y -= 15

    # Save the PDF canvas
    c.save()
    if print_output:
        print(f"{text_file_name_with_extension} successfully converted to pdf. It's on your desktop")


def ppt_to_pdf(ppt_file_name_with_extension, save_file_name, print_output=False):
    ppt_file_path = f"{local_path}{ppt_file_name_with_extension}"
    pdf_file_path = f"{local_path}{save_file_name}"
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pdf = powerpoint.Presentations.Open(ppt_file_path, WithWindow=False)
    if print_output:
        print(f"Converting {ppt_file_name_with_extension} to Pdf.....")
    pdf.SaveAs(pdf_file_path, 32)
    pdf.Close()
    powerpoint.Quit()
    if print_output:
        print(f"{ppt_file_name_with_extension} successfully converted to pdf. It's on your desktop")


def doc_to_pdf(file_name_with_extension, save_file_name, print_output=False):
    file_path = f"{local_path}{file_name_with_extension}"
    if print_output:
        print(f"Converting {file_name_with_extension} to Pdf.....")
    convert(f"{file_path}", f"{local_path}{save_file_name}.pdf")
    if print_output:
        print(f"{file_name_with_extension} successfully converted to pdf. Locate it on your desktop")


def file_mime_type(file_name_with_extension, print_output=False):
    file_path = f"{local_path}{file_name_with_extension}"
    file = open(file_path, "rb")
    file_content = file.read()
    file_type = magic.Magic(mime=True).from_buffer(file_content)
    file.close()
    if print_output:
        print(f"File type: {file_type}")
    else:
        return file_type


def file_to_audio(file_name_with_extension, save_file_name, print_output=False):
    file_path = f"{local_path}{save_file_name}"
    # check if the file exists on the desktop
    pdf_reader = ""
    file_type = file_mime_type(file_name_with_extension)
    if file_type == "application/msword":
        doc_to_pdf(file_name_with_extension, save_file_name)
        pdf_file = open(f"{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    elif file_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc_to_pdf(file_name_with_extension, save_file_name)
        pdf_file = open(f"{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    elif file_type == "application/vnd.ms-powerpoint":
        ppt_to_pdf(file_name_with_extension, save_file_name)
        pdf_file = open(f"{local_path}{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    elif file_type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        ppt_to_pdf(file_name_with_extension, save_file_name)
        pdf_file = open(f"{local_path}{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    elif file_type == "text/plain":
        text_to_pdf(file_name_with_extension, save_file_name)
        pdf_file = open(f"{local_path}{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    elif file_type == "application/pdf":
        pdf_file = open(f"{file_path}.pdf", "rb")
        pdf_reader = PyPDF2.PdfReader(pdf_file)

    else:
        raise FileNotFoundError("The file is corrupted or Unsupported file type, only .doc, .docx, .ppt, .pptx, .pdf, "
                                "and .txt file are supported")

    # assign an empty string variable to store the text
    text = ""

    # loop through each page
    for page_num in range(len(pdf_reader.pages)):
        page_text = pdf_reader.pages[page_num].extract_text()
        text += page_text.strip().replace("\n", " ")

    pdf_file.close()
    if print_output:
        print("Generating audio file......")
    audio_file = gTTS(text)
    audio_file.save(f"{local_path}{save_file_name}.mp3")
    if os.path.exists(f"{local_path}{save_file_name}.pdf"):
        os.remove(f"{local_path}{save_file_name}.pdf")
    elif os.path.exists(f"{local_path}Convert.pdf"):
        os.remove(f"{local_path}Convert.pdf")
    if print_output:
        print("The process was successful, the audio file has been saved to your Desktop")


file_to_audio("Inten.docx", "audio", True)