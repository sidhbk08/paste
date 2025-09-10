import os
import xlsxwriter

folder_path = input("Enter the folder path containing .txt files: ")

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
Fax: 1-888-517-7113
For a Fast Appeal:
Phone: 1-866-314-8188""",
    
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
    
    "United0356": """For a Standard Appeal:
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
For a Fast Appeal: Phone: 1-877-262-9203"""
    
}

results = []

# Loop through all files in the specified folder
for filename in os.listdir(folder_path):
    if filename.endswith('.txt'):
        file_path = os.path.join(folder_path, filename)
        
        with open(file_path, 'r') as file:
            content = file.read()
            
            # Check for the presence of the specific text patterns
            for appeal_type, appeal_text in appeal_texts.items():
                if appeal_text in content:
                    results.append({'Filename': filename, 'Type': appeal_type})

output_file = 'file_types.xlsx'
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