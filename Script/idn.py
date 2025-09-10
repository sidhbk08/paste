import os
import time
import win32com.client
import fitz  
import difflib
from PIL import Image
from pyzbar.pyzbar import decode
import io
import cv2
import numpy as np
import xlsxwriter


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
P.O. Box 6106
MS: CA124-0157
Cypress, CA 90630
In Person Delivery Address:
UnitedHealthcare Appeals & Grievances Department
5701 Katella Avenue
Cypress, CA 90630
Fax: 1-888-517-7113""",
     
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
UnitedHealthcare Appeals & Grievances Department
P.O. Box 6103
MS CA120-0360
Cypress, CA 90630-0023
In Person Delivery Address:
Peoples Health Medicare Center
3017 Veterans Memorial Blvd
Metairie, LA 70002
Fax: 1-844-226-0356
For a Fast Appeal: Phone: 1-877-262-9203""",

    "Net": "See claim details on following pages or go directly to PCNhealth.com to view.",

    "PrefferedCare": "See claim details on following pages or go directly to myPreferredCare.com to view.",
    
    "UnitedIR": "[IR_170224_155359]",

    "LogoRemove": "1-800-496-5841"
}

# Define multiple macro maps
macro_maps = {
    "United7113": {
        "[insert_QR": "AUHC",
        "You Have The Right To Appeal Our Decision": "IDNUHC",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United7103": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "United703",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United03": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "United23",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United06": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "United56",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "United1081": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "United81",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "People": {
        "[insert_QR": "PeopleQR",
        "You Have The Right To Appeal Our Decision": "IDNPep",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "IDN": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "UnitedIDN"
    },
    "PrefferedCare": {
        "[insert_QR": "Preffered",
        "You Have The Right To Appeal Our Decision": "IDNUHC",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "Net": {
        "[insert_QR": "Net",
        "You Have The Right To Appeal Our Decision": "IDNUHC",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "UnitedIR": {
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "IR",
        "[insert_DenialCodeDescription]": "Denial"
    },
    "LogoRemove": {
        "E": "Log",
        "[insert_QR": "UHC",
        "You Have The Right To Appeal Our Decision": "fou",
        "[insert_DenialCodeDescription]": "Denial"
    } 
}


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


def macro(doc_path, keyword_macro_map, found):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    all_success = True

    try:
        doc = word.Documents.Open(doc_path)
 
        try:
            word.Run("footerIDN")
        except Exception as e:
            print(f"    [x] [insert_CustomerServiceNum] found, Macro 'footerIDN' skipped in {os.path.basename(doc_path)}: {e}")
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

        word.ActiveWindow.View.ShowFieldCodes = True

        keyword_found = False

        for section in doc.Sections:
            for header_id in range(1, 4):
                header_range = section.Headers(header_id).Range
                find = header_range.Find
                find.Text = "Claim detail for"
                find.Forward = True
                find.Wrap = 1  # wdFindContinue
                
                if find.Execute():
                    keyword_found = True

                    header_range.Start = header_range.Start + find.Parent.Start - header_range.Start

                    header_range.Select()
                    #print(f"    [✓] Keyword '{keyword}' found in header {header_id} of section. Moved to its start.")
                    break  

            if keyword_found:
                break  

        if keyword_found:
            try:
                print(f"Running macro 'header' on {os.path.basename(doc_path)}")
                word.Application.Run("header")
            except Exception as e:
                print(f"    [x] Failed to run macro '{macro_name}' in {os.path.basename(doc_path)}: {e}")
        else:
            print(f"    [x] Keyword '{keyword}' not found in headers of {os.path.basename(doc_path)}. Macro '{macro_name}' skipped.")

        word.ActiveWindow.View.ShowFieldCodes = False

        doc.Save()
        doc.Close()
        if all_success:
            print(f"[✓] {os.path.basename(doc_path)} processed successfully.")
    except Exception as e:
        print(f"[!] Error processing {os.path.basename(doc_path)}: {e}")
        try:
            doc.Close()
        except:
            pass   

    finally:
        word.Quit()
        time.sleep(0.2)

def page_break(pdf_path):
    doc = fitz.open(pdf_path)
    #page_breaks = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text("text")

        if not text.strip() or text.startswith('\n'):
            return (page_num + 1)  # Store page number (1-indexed)
            continue

        blocks = page.get_text("blocks")
        if blocks:
            last_block = blocks[-1]
            last_block_height = last_block[3] - last_block[1]  # Calculate height of the last block
            page_height = page.rect.height  # Get the height of the page

            if last_block[1] >= page_height * 0.9:  # Adjust threshold as needed
                return (page_num + 1)

    return None

def find_matching_word_file(pdf_file, word_files):
    pdf_base = os.path.splitext(pdf_file)[0]
    for word_file in word_files:
        word_base = os.path.splitext(word_file)[0]
        if word_base.startswith(pdf_base):
            return word_file
    return None


def process_all_files_in_folder(folder_path):

    workbook = xlsxwriter.Workbook(os.path.join(folder_path, 'pdf_type_log.xlsx'))
    worksheet = workbook.add_worksheet()

    bold_format = workbook.add_format({'bold': True})

    worksheet.write('A1', 'Filename', bold_format)
    worksheet.write('B1', 'PDF Type', bold_format)

    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    word_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]
    pdf_files.sort()  # Sort files to maintain order

    row = 1

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        print(f"Processing PDF: {pdf_file}")

        pdf_text = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                pdf_text += page.get_text()

        pdf_text = pdf_text.replace('\n', ' ').strip()

        pdf_type = None
        logo_present = False

        # Check for LogoRemove text
        logo_text = appeal_texts["LogoRemove"].replace('\n', ' ').strip()
        if logo_text in pdf_text:
            logo_present = True

        # Detect PDF type by matching appeal text
        for key, text in appeal_texts.items():
            if key == "LogoRemove":
                continue  # Already handled above
            normalized_text = text.replace('\n', ' ').strip()
            if normalized_text in pdf_text:
                pdf_type = key
                break

        # Override type to LogoRemove if both present
        if pdf_type and logo_present:
            pdf_type = "LogoRemove"

        if pdf_type is None:
            print(f"[x] PDF type not recognized for {pdf_file}. Skipping.")
            continue  # Skip if not recognized

        #blank_pages = page_break(pdf_path)
        qr_found = find_qr_code(pdf_path)
        #key_w = keyword(pdf_path)

        word_file = find_matching_word_file(pdf_file, word_files)

        if word_file:
            word_path = os.path.join(folder_path, word_file)
            print(f"Running macro on: {word_file}")
            #print(pdf_type)
            macro(word_path, macro_maps[pdf_type], qr_found)
        else:
            print(f"[!] Corresponding Word file not found for {pdf_file}")

        worksheet.write(row, 0, pdf_file)
        worksheet.write(row, 1, pdf_type)
        row += 1

    worksheet.autofilter('A1:B1')
    workbook.close()
    print("Excel sheet 'pdf_type_log.xlsx' has been created.")

folder_path = input("Enter the folder path containing .pdf and .docx files: ")
process_all_files_in_folder(folder_path)