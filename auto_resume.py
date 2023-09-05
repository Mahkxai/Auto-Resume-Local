# Project: Python Script to Automate Resume Downloads
# Description: This script downloads your Google Docs Resume as pdf, docx, and png.
# Author: Hardik Shrestha
# Date: September 20, 2023

# Enable API services from Google Cloud Services and 
# Insert relevant information in <placeholders>.    Lines 24, 27
# Change resume file names as desired.      Lines 43, 61, 91


import os
import io
import fitz  # PyMuPDF
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload
from PIL import Image

# pip install PyMuPDF Pillow

# ID of the Google Docs document you want to convert 
# https://docs.google.com/document/d/<document_id>/edit
document_id = '<document_id>'

# Path to the Google Cloud Console JSON file (<credentials>.json)
credentials_path = 'path/<credentials>.json'

# Define scopes for Google Drive API
scopes = ['https://www.googleapis.com/auth/drive']

# Load the service account credentials with the specified scopes
credentials = service_account.Credentials.from_service_account_file(
    credentials_path,
    scopes=scopes
    )

# Create a service object for interacting with Google Drive API
drive_service = build('drive', 'v3', credentials=credentials)


# Export as PDF and download
pdf_file_path = 'First_Last_Resume.pdf'
pdfRequest = drive_service.files().export_media(
    fileId=document_id, 
    mimeType='application/pdf'
    )

fhPDF = io.FileIO(pdf_file_path, 'wb')
pdfDownloader = MediaIoBaseDownload(fhPDF, pdfRequest)
done = False
while not done:
    status, done = pdfDownloader.next_chunk()
    print(f"Download {int(status.progress() * 100)}%.")
fhPDF.close()

print(f"Downloaded '{pdf_file_path}'.")


# Export as DOCX and download
docx_file_path = 'First_Last_Resume.docx'
docxRequest = drive_service.files().export_media(
    fileId=document_id, 
    mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
fhDOCX = io.FileIO(docx_file_path, 'wb')
docxDownloader = MediaIoBaseDownload(fhDOCX, docxRequest)
done = False
while not done:
    status, done = docxDownloader.next_chunk()
    print(f"Download {int(status.progress() * 100)}%.")
fhDOCX.close()
print(f"Downloaded '{docx_file_path}")


# Export as image and screenshot
# Loading PDF using fitz
page_number = 0
pdf_document = fitz.open(pdf_file_path)
pdf_page = pdf_document.load_page(page_number)
pdf_page_dimensions = pdf_page.rect

# Convert the PDF page to an image 
desired_dpi = 300  
pdf_image = pdf_page.get_pixmap(matrix=fitz.Matrix(desired_dpi/72, desired_dpi/72))

# Create a PIL (Pillow) image from the PyMuPDF image
pil_image = Image.frombytes("RGB", [pdf_image.width, pdf_image.height], pdf_image.samples)

# Save the screenshot as an image file
screenshot_file_path = 'Resume.png'
pil_image.save(screenshot_file_path, dpi = (desired_dpi, desired_dpi))

pdf_document.close()
print(f'Screenshot saved as {screenshot_file_path}')