import os
import xlsxwriter
import re

folder_path = input("Enter the folder path containing .txt files: ")

# Define the specific text patterns to search for
appeal_texts = {
    "UHCAdvantage": """Consulte la información detallada de los reclamos en las páginas siguientes o visite directamente MyUHCAdvantage.com para verlos.""",
    
    "UHCMedicare": """Consulte la información detallada de los reclamos en las páginas siguientes o visite directamente MyUHCMedicare.com para verlos.""",

     "UHCAdv": """Consulte la información
detallada de los
reclamos en las
páginas siguientes o
visite directamente
MyUHCAdvantage.co
m para verlos.""",
    
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
    
    "United0356": """Fax: 1-844-226-0356""",

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

      "Network": "Consulte la información detallada de los reclamos en las páginas siguientes o visite directamente PCNhealth.com para verlos.",

     "PrefferedCare": "Consulte la información detallada de los reclamos en las páginas siguientes o visite directamente myPreferredCare.com para verlos.",
 
     "SPAIR": "[IR_170224_155359]",

     "LogoRemove": "1-800-496-5841"

}

results = []

# Normalize the appeal texts by removing extra whitespace and line breaks
normalized_appeal_texts = {key: re.sub(r'\s+', ' ', text.strip()) for key, text in appeal_texts.items()}

# Loop through all files in the specified folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        file_path = os.path.join(folder_path, filename)
        
        # Attempt to read the file with different encodings
        content = ""
        for encoding in ['utf-8', 'latin-1', 'windows-1252']:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    content = file.read()
                break  # Exit the loop if reading is successful
            except UnicodeDecodeError:
                continue  # Try the next encoding
            except FileNotFoundError as e:
                print(f"File not found: {e}")
                break
        
        if not content:
            print(f"Could not read {filename} with any of the attempted encodings.")
            continue  # Skip to the next file if content is still empty

        normalized_content = re.sub(r'\s+', ' ', content.strip())  # Normalize the file content
        
        # Check for the presence of the specific text patterns
        for appeal_type, appeal_text in normalized_appeal_texts.items():
            if appeal_text in normalized_content:
                results.append({'Filename': filename, 'Type': appeal_type})
            #else:
                # Debugging: Print the content and the appeal text for comparison
                #print(f"Checking {appeal_type} in {filename}:")
                #print("Appeal Text:")
                #print(repr(appeal_text))  # Use repr to show hidden characters
                #print("File Content:")
                #print(repr(normalized_content))  # Use repr to show hidden characters
                #print("-" * 40)

# Create an Excel file using xlsxwriter
output_file = 'file_types_SPA.xlsx'
workbook = xlsxwriter.Workbook(output_file)
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Filename')
worksheet.write('B1', 'Type')

row = 1
for result in results:
    worksheet.write(row, 0, result['Filename'])
    worksheet.write(row, 1, result['Type'])
    row += 1

workbook.close()

print(f"Results have been written to {output_file}")