import os
from docx import Document

def replace_text_in_docx(file_path, old_text, new_text):
    try:
        # Load the document
        doc = Document(file_path)
        
        # Iterate through paragraphs in the document
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:  # Iterate through runs in the paragraph
                if old_text in run.text:
                    # Replace the old text with the new text
                    run.text = run.text.replace(old_text, new_text)
        
        # Save the modified document with the same name
        doc.save(file_path)
        print(f"Successfully processed: {file_path}")
    except Exception as e:
        print(f"Error processing {file_path}: {e}")

def process_folder(folder_path, old_text, new_text):
    # Check if the folder path exists
    if not os.path.isdir(folder_path):
        print(f"The folder path '{folder_path}' does not exist.")
        return
    
    # Iterate through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            file_path = os.path.join(folder_path, filename)
            replace_text_in_docx(file_path, old_text, new_text)
        else:
            print(f"Skipped non-docx file: {filename}")

# Ask the user for the folder path
folder_path = input("Please enter the folder path containing the Word files: ").strip()

# Specify the text to replace
old_text = """For appeals related to payment of a medical service/item you already received, we’ll give you a written decision within 30 days. You can’t ask for a fast appeal if you’re asking us to pay you back for a medical service/item you already received."""
new_text = """For appeals related to payment of a medical service/item you already received, we’ll give you a written decision within 60 days. You can’t ask for a fast appeal if you’re asking us to pay you back for a medical service/item you already received."""

# Call the function to process all Word files in the folder
process_folder(folder_path, old_text, new_text)
