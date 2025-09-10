import os
import fitz  # PyMuPDF
from docx import Document
import re
from openpyxl import Workbook

def extract_words_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    words = []
    for page in doc:
        text = page.get_text()
        words.extend(re.findall(r'\w+', text))
    doc.close()
    return [w.lower() for w in words]

def extract_words_from_docx(docx_path):
    doc = Document(docx_path)
    words = []
    for para in doc.paragraphs:
        words.extend(re.findall(r'\w+', para.text))
    return [w.lower() for w in words]

def compare_words(pdf_words, docx_words):
    pdf_set = set(pdf_words)
    docx_set = set(docx_words)

    missing_in_docx = pdf_set - docx_set
    missing_in_pdf = docx_set - pdf_set

    if not missing_in_docx and not missing_in_pdf:
        return True, [], []
    else:
        return False, missing_in_docx, missing_in_pdf

def process_folder(folder_path):
    files = os.listdir(folder_path)
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]

    # Workbook for missing words in PDF
    wb_pdf = Workbook()
    ws_pdf = wb_pdf.active
    ws_pdf.title = "Missing Words in PDF"
    ws_pdf.append(["DOCX File", "Missing Words in PDF"])

    # Workbook for missing words in DOCX
    wb_docx = Workbook()
    ws_docx = wb_docx.active
    ws_docx.title = "Missing Words in DOCX"
    ws_docx.append(["DOCX File", "Missing Words in DOCX"])

    for pdf_file in pdf_files:
        base_name = pdf_file[:-4]  # Remove .pdf
        docx_name = f"{base_name}_TEMPLATED_REVIEW.docx"
        pdf_path = os.path.join(folder_path, pdf_file)
        docx_path = os.path.join(folder_path, docx_name)

        if os.path.exists(docx_path):
            #print(f"\n🔍 Comparing:\nPDF: {pdf_file}\nDOCX: {docx_name}")
            pdf_words = extract_words_from_pdf(pdf_path)
            docx_words = extract_words_from_docx(docx_path)
            identical, missing_in_docx, missing_in_pdf = compare_words(pdf_words, docx_words)

            if identical:
                print("✅ Files are identical (word-by-word match).")
            else:
                print("❌ Files differ.")
                if missing_in_pdf:
                    missing_pdf_str = ", ".join(sorted(missing_in_pdf))
                    ws_pdf.append([docx_name, missing_pdf_str])
                if missing_in_docx:
                    missing_docx_str = ", ".join(sorted(missing_in_docx))
                    ws_docx.append([docx_name, missing_docx_str])
        else:
            print(f"⚠️ Word file not found for PDF: {pdf_file}")

    # Save both Excel files
    output_pdf_excel = os.path.join(folder_path, "Missing_Words_In_PDF.xlsx")
    output_docx_excel = os.path.join(folder_path, "Missing_Words_In_DOCX.xlsx")
    wb_pdf.save(output_pdf_excel)
    wb_docx.save(output_docx_excel)

    print(f"\n📁 Excel files saved:")
    print(f"   - {output_pdf_excel}")
    print(f"   - {output_docx_excel}")

# Set your folder path here
folder_path = input("Enter the folder path containing .pdf and .docx files: ")
process_folder(folder_path)
