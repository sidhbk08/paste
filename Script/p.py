import os
from docx import Document

def replace_text_in_docx(file_path, old_text, new_text):
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

def process_folder(folder_path, old_text, new_text):
    # Iterate through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            file_path = os.path.join(folder_path, filename)
            replace_text_in_docx(file_path, old_text, new_text)
            print(f"Processed: {filename}")

# Ask the user for the folder path
folder_path = input("Please enter the folder path containing the Word files: ")

# Specify the text to replace
old1 = "groupo"
new1 = "grupo"

old2 = "paciente: Un servicio o gasto para el cual usted no tiene cobertura bajo su plan de beneficios de salud."
new2 = "Cantidad sin cobertura que paga el paciente: Un servicio o gasto para el cual usted no tiene cobertura bajo su plan de beneficios de salud."

# Call the function to process all Word files in the folder
process_folder(folder_path, old1, new1)
process_folder(folder_path, old2, new2)