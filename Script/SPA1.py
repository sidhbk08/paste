import os
import time
import win32com.client
import fitz  
import difflib
from PIL import Image
from pyzbar.pyzbar import decode

# Define the texts to search for
appeal_texts = {
    
    "SPSmall": """del nivel del plan""",
     
     "SP": """a nivel del plan""",
    
    "ES": """Visite es.Medicare.gov/about-us/accessibility-nondiscrimination-notice""",
    
    "WithoutES": """Visite Medicare.gov/about-us/accessibility-nondiscrimination-notice""",
    
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

    "SPA0356": """Fax: 1-844-226-0356""",

    "SPANet": "PREFERRED CARE NETWORK, INC",

    "SPAPrefferedCare": "PREFERRED CARE PARTNERS INC",
    
    "UnitedIR": "[IR_170224_155359]"
}

# Define multiple macro maps
macro_maps = {
    "SP": {
        "[insert_QR": "SUHC",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPAWes",
        #"[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "DenS"
    },
    "SPSmall": {
        "[insert_QR": "SUHC",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPASmall",
        #"[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "DenS"
    },
    "WithoutES": {
        "[insert_QR": "SUHC",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPANes",
        #"[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "DenS"
    },
    "SPA0356": {
        "[insert_QR": "SUHC",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPA0356",
        #"[insert_keyterms]": "keytable",
        "[insert_DenialCodeDescription]": "DenS"
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
    "SPAPrefferedCare": {
        "Consulte la información": "SPA_Partner",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPA_Preff",
        "[insert_DenialCodeDescription]": "DenS"
    },
    "SPANet": {
        "[insert_QR": "Net",
        "USTED TIENE DERECHO A APELAR NUESTRA DECISIÓN": "SPA_Net",
        "[insert_DenialCodeDescription]": "DenS"
    },
    "UnitedIR": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "IR",
        "[insert_keyterms]": "keytable",
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

def macro(doc_path, keyword_macro_map, blank_pages, found):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    all_success = True

    try:
        doc = word.Documents.Open(doc_path)
        
        try:
            word.Run("footerSPA")
        except Exception as e:
            print(f"    [x] See Claim details found, Macro 'header' skipped in {os.path.basename(doc_path)}: {e}")
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
            find.Text = "Consulte la información"
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


        selection = word.Selection
        selection.HomeKey(Unit=6)
        find.Text = " Cantidad sin"
        find.Forward = True
        find.Wrap = 1  # wdFindContinue
        find.MatchCase = False
        find.MatchWholeWord = False

        if find.Execute():          
            selection.Start = selection.Start
            selection.End = selection.Start
            try:
                word.Run("SPA_Par")
            except Exception as e:
                print(f"    [x] Failed to run macro 'SPA_Par' in {os.path.basename(doc_path)}: {e}")
                all_success = False
        else:
            print(f"    [x] Keyword ' Cantidad sin' not found in {os.path.basename(doc_path)}. Macro 'SPA_Par' skipped.")
            all_success = False

        word.ActiveWindow.View.ShowFieldCodes = True

        keyword_found = False

        for section in doc.Sections:
            for header_id in range(1, 4):
                header_range = section.Headers(header_id).Range
                find = header_range.Find
                find.Text = "Información detallada del"
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
    pdf_base = os.path.splitext(pdf_file)[0]
    for word_file in word_files:
        word_base = os.path.splitext(word_file)[0]
        if word_base.startswith(pdf_base):
            return word_file
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
        #print(pdf_text)
        pdf_type = None
        sp_found = False
        withoutes_found = False
        withes_found = False

        for key in ["SPANet", "SPAPrefferedCare", "SPA0356"]:
            normalized_text = appeal_texts[key].replace('\n', ' ').strip()
            if normalized_text in pdf_text:
                pdf_type = key
                break  

        if not pdf_type:
            sp_text = appeal_texts["SP"].replace('\n', ' ').strip()
            withoutes_text = appeal_texts["WithoutES"].replace('\n', ' ').strip()
            withes_text = appeal_texts["ES"].replace('\n', ' ').strip()
            #print(withoutes_text)
            if sp_text in pdf_text:
                sp_found = True
            if withoutes_text in pdf_text:
                withoutes_found = True
            if withes_text in pdf_text:
                withes_found = True

            if sp_found and withoutes_found:
                pdf_type = "WithoutES"
            elif sp_found and withes_found :
                pdf_type = "SP"
            else:
                for key, text in appeal_texts.items():
                    if key in ("SP", "WithoutES", "SPANet", "SPAPrefferedCare"):
                        continue  # Already handled
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
            #print(pdf_type)
            print(f"Running macro on: {word_file}")
            macro(word_path, macro_maps[pdf_type], blank_pages, qr_found)
        else:
            print(f"[!] Corresponding Word file not found for {pdf_file}")

folder_path = input("Enter the folder path containing .pdf and .docx files: ")
process_all_files_in_folder(folder_path)