import os
import time
import win32com.client
 
macro_map = {
    "You have the right to appeal our decision": "DeENG"
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
 
        word.Run("field")
        doc.Save()
        doc.Close()
        if all_success:
            print(f"[✓] {os.path.basename(doc_path)} processed successfully.")
    except Exception as e:
        print(f"[!] Error processing {os.path.basename(doc_path)}: {e}")
    finally:
        word.Quit()
        time.sleep(0.3)
 
def process_all_docs_in_folder(folder_path, keyword_macro_map):
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            macro(file_path, keyword_macro_map)
 
 
folder_path = input("Enter the folder path containing .docx files: ")
 
process_all_docs_in_folder(folder_path, macro_map)