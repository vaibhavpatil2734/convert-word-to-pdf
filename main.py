# -*- coding: utf-8 -*-

import os
import sys
from tkinter import Tk
from tkinter.filedialog import askopenfilenames, askdirectory
from win32com.client import Dispatch

# Numerical constants for Word Export
wdExportFormatPDF = 17
wdExportDocumentWithMarkup = 1
wdExportCreateHeadingBookmarks = 1

# Function to convert Word document to PDF
def convert_word_to_pdf(docx_path, pdf_path):
    w = Dispatch("Word.Application")
    try:
        # Resolve absolute paths
        docx_path = os.path.abspath(docx_path)
        pdf_path = os.path.abspath(pdf_path)

        # Print paths for debugging
        print(f"Resolved document path: {docx_path}")
        print(f"Resolved PDF path: {pdf_path}")

        # Open the document
        doc = w.Documents.Open(docx_path, ReadOnly=1)
        print(f"Opening document: {docx_path}")

        # Export the document to PDF
        doc.ExportAsFixedFormat(
            pdf_path,
            wdExportFormatPDF,
            Item=wdExportDocumentWithMarkup,
            CreateBookmarks=wdExportCreateHeadingBookmarks
        )
        print(f"Successfully converted {docx_path} to {pdf_path}")
    except Exception as e:
        print(f"Error converting {docx_path} to {pdf_path}: {e}")
    finally:
        try:
            w.Quit()  # Quit Word application
        except Exception as e:
            print(f"Error quitting Word application: {e}")

def main():
    # Initialize Tkinter and hide the main window
    root = Tk()
    root.withdraw()

    # Prompt the user to select Word documents
    print("Select the Word files to convert (Ctrl+Click to select multiple files):")
    file_paths = askopenfilenames(
        title="Select Word Documents",
        filetypes=[("Word Documents", "*.docx *.doc")]
    )

    if not file_paths:
        print("No files selected.")
        return

    # Prompt the user to select the destination directory
    print("Select the destination folder for PDF files:")
    dest_folder = askdirectory(title="Select Destination Folder")

    if not dest_folder:
        print("No destination folder selected.")
        return

    # Ensure destination folder path is valid
    if not os.path.isdir(dest_folder):
        print(f"Invalid destination folder: {dest_folder}")
        return

    # Process each selected file
    for file_path in file_paths:
        identifier = os.path.basename(file_path).split('.')[0]
        file_name = identifier + '.pdf'
        save_path = os.path.join(dest_folder, file_name)
        
        # Check if the file already exists
        if not os.path.exists(save_path):
            print(f"Converting file: {file_path} to {save_path}")
            convert_word_to_pdf(file_path, save_path)
        else:
            print(f"File already exists, no need to convert: {file_path}")

if __name__ == '__main__':
    main()
