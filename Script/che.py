import fitz  # PyMuPDF
import os

def extract_eob_row_highlight(pdf_path):
    print(f"\n📄 Processing: {os.path.basename(pdf_path)}")
    keyword = "EOB CODE"
    doc = fitz.open(pdf_path)
    folder = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for page_num, page in enumerate(doc, start=1):
        matches = page.search_for(keyword)
        if not matches:
            continue

        for i, rect in enumerate(matches):
            page_width = page.rect.width
            row_top = rect.y0
            row_bottom = rect.y1

            # Define full-width row based on keyword's vertical position
            full_row_rect = fitz.Rect(0, row_top, page_width, row_bottom)

            # Render only the cropped area
            matrix = fitz.Matrix(2, 2)  # Increase for higher resolution
            pix = page.get_pixmap(matrix=matrix, clip=full_row_rect)

            # Save cropped image
            image_path = os.path.join(folder, f"{base_name}_page{page_num}_row{i+1}.png")
            pix.save(image_path)
            print(f"📸 Saved cropped row image: {image_path}")
            break

# Folder containing PDFs
folder = input("Enter the folder path containing .pdf: ")
for file in os.listdir(folder):
    if file.lower().endswith(".pdf"):
        extract_eob_row_highlight(os.path.join(folder, file))
