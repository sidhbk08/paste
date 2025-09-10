import os
import time
import win32com.client
import fitz  
import difflib
from PIL import Image
from pyzbar.pyzbar import decode

# Define the texts to search for
appeal_texts = {
    "United03": """For a Standard Appeal:
Mailing Address:
UnitedHealthcare Appeals & Grievances Department
PO Box 6103
MS: CA124-0157
Cypress, CA 90630-0023
In Person Delivery Address:
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-844-226-0356
For a Fast Appeal: Phone: 1-877-262-9203, TTY users call: 711.
Fax: 1-844-226-0356""",
    
    "United7113": """For a Standard Appeal:
Mailing Address:
UnitedHealthcare Appeals & Grievances Department
PO Box 6106
MS: CA124-0157
Cypress, CA 90630-0016
In Person Delivery Address:
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-888-517-7113
For a Fast Appeal: Phone: 1-877-262-9203, TTY users call: 711.
Fax: 1-866-373-1081""",
     
     "United7103": """For a Standard Appeal:
Mailing Address:
UnitedHealthcare Appeals & Grievances Department
PO Box 6103
MS: CA124-0157
Cypress, CA 90630-0023
In Person Delivery Address:
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-888-517-7113
For a Fast Appeal: Phone: 1-877-262-9203, TTY users call: 711.
Fax: 1-866-373-1081""",
    
    "United1081": """For a Standard Appeal:
Mailing Address:
UnitedHealthcare Appeals & Grievances Department
PO Box 6106
MS: CA124-0157
Cypress, CA 90630-0016
In Person Delivery Address:
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-866-373-1081
For a Fast Appeal: Phone: 1-877-262-9203, TTY users call: 711. Fax: 1-866-373-1081""",
    
    "United06": """For a Standard Appeal:
Mailing Address:
UnitedHealthcare Appeals & Grievances Department
PO Box 6106
MS: CA124-0157
Cypress, CA 90630-0016
In Person Delivery Address:
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-844-226-0356
For a Fast Appeal: Phone: 1-877-262-9203, TTY users call: 711. Fax: 1-866-373-1081""",
    
    "People": """For a Standard Appeal:
Mailing Address:
Appeals and Grievance Department
PO Box 6103
MS CA120-0360
Cypress,CA 90630-0023
In Person Delivery Address:
Peoples Health Medicare Center
3017 Veterans Memorial Blvd
Metairie, LA 70002
Fax: 1-844-226-0356
For a Fast Appeal: Phone: 1-855-409-7041 TTY users call: 711.
Fax: 1-866-373-1081""",

    "Net": "See claim details on following pages or go directly to PCNhealth.com to view.",

    "PrefferedCare": "See claim details on following pages or go directly to myPreferredCare.com to view.",
    
    "UnitedIR": "[IR_170224_155359]"
}

# Define multiple macro maps
macro_maps = {
    "United7113": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United13",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United7103": {
        "[insert_QR": "UHCAdv",
        "You Have The Right To Appeal Our Decision": "United703",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United03": {
        "[insert_QR": "UHCAdv",
        "You Have The Right To Appeal Our Decision": "United23",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United06": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United56",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United1081": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United81",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "People": {
        "[insert_QR": "UHCAdv",
        "You Have The Right To Appeal Our Decision": "People",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "IDN": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "UnitedIDN"
    },
    "PrefferedCare": {
        "[insert_QR": "Preffered",
        "You Have The Right To Appeal Our Decision": "Pref",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "Net": {
        "[insert_QR": "Net",
        "You Have The Right To Appeal Our Decision": "Network",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "UnitedIR": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "IR",
        "[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "Denial"
    } 
}

import fitz  # PyMuPDF
from PIL import Image
import io
import cv2
import numpy as np
import os


def detect_logo_in_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        first_page = doc[0]
        pix = first_page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        width, height = img.size
        footer_pil = img.crop((0, int(height * 0.8), width, height))

        logo = cv2.imread(LOGO_PATH, cv2.IMREAD_COLOR)
        if logo is None:
            return False

        logo = cv2.cvtColor(logo, cv2.COLOR_BGR2RGB)
        footer = np.array(footer_pil)

        akaze = cv2.AKAZE_create()
        kp1, des1 = akaze.detectAndCompute(logo, None)
        kp2, des2 = akaze.detectAndCompute(footer, None)
        if des1 is None or des2 is None:
            return False

        bf = cv2.BFMatcher()
        matches = bf.knnMatch(des1, des2, k=2)
        good_matches = [m for m, n in matches if m.distance < 0.7 * n.distance]

        if len(good_matches) < 13:
            return False

        src_pts = np.float32([kp1[m.queryIdx].pt for m in good_matches]).reshape(-1, 1, 2)
        dst_pts = np.float32([kp2[m.trainIdx].pt for m in good_matches]).reshape(-1, 1, 2)
        M, mask = cv2.findHomography(src_pts, dst_pts, cv2.RANSAC, 5.0)
        inliers = sum(mask.ravel().tolist()) if mask is not None else 0

        return inliers >= 13

    except Exception:
        print(f"⚠️ Error in {pdf_path}: {e}")

def detect_qr_codes_in_page(page, zoom=3):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    decoded_objects = decode(img)
    qr_detected = len(decoded_objects) > 0
    return qr_detected
def find_qr_code(pdf_path):
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        if detect_qr_codes_in_page(page):
            #print(f"QR code found on page {page_num + 1}.")
            return True
    #print("QR code found in the document.")
    return False

def macro(doc_path, keyword_macro_map, blank_pages, found):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    all_success = True

    try:
        doc = word.Documents.Open(doc_path)

        # Check if blank_pages is equal to 2 and run the "section" macro
        if blank_pages == 2:
            try:
                word.Run("Page_break")
            except Exception as e:
                print(f"    [x] Failed to run macro 'section' in {os.path.basename(doc_path)}: {e}")
                all_success = False

        
        
        for keyword, macro in keyword_macro_map.items():
            find = word.Selection.Find
            find.Text = keyword
            find.Forward = True
            find.Wrap = 1
            find.Execute()

            if find.Found:
                try:
                    word.Run(macro)
                except Exception as e:
                    print(f"    [x] Failed to run macro '{macro}' in {os.path.basename(doc_path)}: {e}")
                    all_success = False
            else:
                print(f"    [x] Keyword '{keyword}' not found in {os.path.basename(doc_path)}. Macro '{macro}' skipped.")
                all_success = False
      
        if not found:
            find.Text = "See claim details "
            find.Forward = True
            find.Wrap = 1
            find.Execute()
            if find.Found:
                try:
                    word.Run("Merge")
                except Exception as e:
                    print(f"    [x] Failed to run macro 'Merge' in {os.path.basename(doc_path)}: {e}")
                    all_success = False
            else:
                print(f"    [x] Keyword 'See claim details' not found in {os.path.basename(doc_path)}. Macro 'Merge' skipped.")
                all_success = False
        else: 
            print(f"    [x] QR Code found in {os.path.basename(doc_path)}. Macro 'Merge' skipped.")

        doc.Save()
        doc.Close()
        if all_success:
            print(f"[✓] {os.path.basename(doc_path)} processed successfully.")
    except Exception as e:
        print(f"[!] Error processing {os.path.basename(doc_path)}: {e}")
    finally:
        word.Quit()
        time.sleep(0.5)

def page_break(pdf_path):
    doc = fitz.open(pdf_path)
    #page_breaks = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text("text")

        # Check for large whitespace or specific patterns indicating a page break
        if not text.strip() or text.startswith('\n'):
            return (page_num + 1)  # Store page number (1-indexed)
            continue

        # Analyze the text height
        blocks = page.get_text("blocks")
        if blocks:
            last_block = blocks[-1]
            last_block_height = last_block[3] - last_block[1]  # Calculate height of the last block
            page_height = page.rect.height  # Get the height of the page

            # Check if the last block is near the bottom of the page
            if last_block[1] >= page_height * 0.9:  # Adjust threshold as needed
                return (page_num + 1)

    return None

def find_matching_word_file(pdf_file, word_files):
    pdf_base = os.path.splitext(pdf_file)[0]  # Get base name without extension
    matches = difflib.get_close_matches(pdf_base, [os.path.splitext(word_file)[0] for word_file in word_files], n=1, cutoff=0.6)
    
    if matches:
        return matches[0] + ".docx"  # Return the matched Word file name with extension
    return None

def process_all_files_in_folder(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    word_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
    pdf_files.sort()  # Sort files to maintain order

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Processing PDF: {pdf_file}")

        # Extract text from the PDF to check its content
        pdf_text = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                pdf_text += page.get_text()

        # Normalize the extracted text
        pdf_text = pdf_text.replace('\n', ' ').strip()  # Replace newlines with spaces and strip whitespace

        # Check if the PDF matches any of the defined appeal texts
        pdf_type = None
        for key, text in appeal_texts.items():
            # Normalize the appeal text for comparison
            normalized_text = text.replace('\n', ' ').strip()
            if normalized_text in pdf_text:
                pdf_type = key
                break

        if pdf_type is None:
            print(f"[x] PDF type not recognized for {pdf_file}. Skipping.")
            continue  # Skip if not recognized

        # Check for blank pages
        blank_pages = page_break(pdf_path)  # Assuming you have a function to count blank pages
        qr_found = find_qr_code(pdf_path)
        #print(blank_pages)
        #if blank_pages == 2:
            #pdf_type += "break"
            #print(pdf_type)  # Change this to the desired PDF type

        # Find the corresponding Word file
        word_file = find_matching_word_file(pdf_file, word_files)

        if word_file:
            word_path = os.path.join(folder_path, word_file)
            print(f"Running macro on: {word_file}")
            macro(word_path, macro_maps[pdf_type], blank_pages, qr_found)
        else:
            print(f"[!] Corresponding Word file not found for {pdf_file}")

folder_path = input("Enter the folder path containing .pdf and .docx files: ")
process_all_files_in_folder(folder_path)
