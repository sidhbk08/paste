import os
import fitz  # PyMuPDF

def rotate_pages_in_folder(folder_path):
    """Rotate specific pages (3rd, 4th, 5th) for all PDFs in the folder."""
    # List all PDF files in the folder
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]

    if not pdf_files:
        print("❌ No PDF files found in the folder.")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"\n🔍 Rotating specific pages in: {pdf_file}")
        rotate_specific_pages(pdf_path)

def rotate_specific_pages(pdf_path):
    """Rotate 3rd, 4th, and 5th pages and save with a new name."""
    doc = fitz.open(pdf_path)
    rotated = False  # Flag to check if any page was rotated

    # List of pages to rotate (3rd, 4th, and 5th pages; 0-indexed)
    pages_to_rotate = [2, 3]

    for page_num in pages_to_rotate:
        if page_num < doc.page_count:
            page = doc.load_page(page_num)
            # Rotate the page 90 degrees to portrait
            page.set_rotation(90)  # Rotate page 90 degrees to portrait
            rotated = True
        else:
            print(f"Page {page_num+1} does not exist in this PDF.")

    if rotated:
        # Save the rotated PDF with a new name (appending "_rotated" to the original filename)
        output_path = os.path.splitext(pdf_path)[0] + "_rotated.pdf"
        doc.save(output_path)
        print(f"✅ Rotated PDF saved as: {output_path}")
    else:
        print("No rotation needed for the specified pages.")

if __name__ == "__main__":
    folder = input("📁 Enter the folder path containing PDF files: ").strip()
    
    # Check if the folder exists
    if os.path.isdir(folder):
        rotate_pages_in_folder(folder)
    else:
        print("❌ Invalid folder path.")
