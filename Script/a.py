import os
import time
import win32com.client

# Define multiple macro maps
macro_maps = {
    "United 7113": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United13",
        "[insert_keyterms]": "keytablebreak"
    },
    "United 0356": {
        "[insert_QR": "UHCAdv",
        "You Have The Right To Appeal Our Decision": "United23",
        "[insert_keyterms]": "keytable"
    },
    "United 06": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United56",
        "[insert_keyterms]": "keytable"
    },
    "United 1081": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "United81",
        "[insert_keyterms]": "keytable"
    },
   "People": {
        "[insert_QR": "UHCAdv",
        "You Have The Right To Appeal Our Decision": "People",
        "[insert_keyterms]": "keytable"
    },
   "IDN": {
        "[insert_QR": "UHCMed",
        "You Have The Right To Appeal Our Decision": "UnitedIDN"
    }
}

def macro(doc_path, keyword_macro_map):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    all_success = True

    try:
        doc = word.Documents.Open(doc_path)
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

        doc.Save()
        doc.Close()
        if all_success:
            print(f"[✓] {os.path.basename(doc_path)} processed successfully.")
    except Exception as e:
        print(f"[!] Error processing {os.path.basename(doc_path)}: {e}")
    finally:
        word.Quit()
        time.sleep(0.5)

def process_all_docs_in_folder(folder_path, keyword_macro_map):
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            macro(file_path, keyword_macro_map)

# ======================== USER INPUT ========================

# Prompt user to select a macro map
print("Available macro maps:")
for i, map_name in enumerate(macro_maps.keys(), 1):
    print(f"{i}. {map_name}")

map_choice = input("\nEnter the number of the macro map you want to run: ").strip()

# Validate the user's choice
if map_choice.isdigit():
    map_index = int(map_choice) - 1
    if 0 <= map_index < len(macro_maps):
        selected_macro_map = list(macro_maps.values())[map_index]
    else:
        print("[x] Invalid macro map number. Exiting.")
        exit()
else:
    print("[x] Invalid input. Exiting.")
    exit()

folder_path = input("Enter the folder path containing .docx files: ")

print("\nAvailable macros to choose from:")
macro_keys = list(selected_macro_map.keys())
for i, (keyword, macro_name) in enumerate(selected_macro_map.items(), 1):
    print(f"{i}. {macro_name} (triggered by keyword: '{keyword}')")

choice = input("\nEnter the number(s) of the macro(s) you want to run (e.g., 1 2 or 1,3). Press Enter to run all: ").strip()

if not choice:
    selected_macro_map = selected_macro_map  # Run all macros
else:
    # Parse comma or space separated input
    choice_list = [c.strip() for c in choice.replace(",", " ").split()]
    selected_macro_map = {}
    for c in choice_list:
        if c.isdigit():
            idx = int(c) - 1
            if 0 <= idx < len(macro_keys):
                keyword = macro_keys[idx]
                selected_macro_map[keyword] = macro_maps[list(macro_maps.keys())[map_index]][keyword]
            else:
                print(f"[x] Invalid macro number: {c}")
        else:
            print(f"[x] Invalid input: {c}")

    if not selected_macro_map:
        print("[x] No valid macros selected. Exiting.")
        exit()

# Run the selected macros
process_all_docs_in_folder(folder_path, selected_macro_map)