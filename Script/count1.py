import os
from PyPDF2 import PdfReader

# Prompt the user for the directory containing the PDF files
directory = input("Please enter the path to your PDF folder: ")

# Check if the directory exists
if not os.path.isdir(directory):
    print("The specified directory does not exist. Please check the path and try again.")
else:
    # Loop through all files in the directory
    for filename in os.listdir(directory):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory, filename)
            try:
                # Open the PDF file
                with open(pdf_path, 'rb') as pdf_file:
                    reader = PdfReader(pdf_file)
                    # Get the number of pages
                    num_pages = len(reader.pages)
		    #print(f"{num_pages}") print(f"{os.path.splitext(filename)[0]}: {num_pages}")
                    print(f"{num_pages}")
            except Exception as e:
                print(f"Could not read {filename}: {e}")