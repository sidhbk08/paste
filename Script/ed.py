import os
import win32com.client

def run_macro_if_keyword_found(folder_path):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False  # Set to False if you want it hidden

    try:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(('.doc', '.docx')):
                full_path = os.path.join(folder_path, filename)
                print(f"Processing file: {filename}")
                
                try:
                    doc = word.Documents.Open(full_path)
                    
                    # Check if the document contains the keyword (this part can be adjusted as per your needs)
                    # Example: check if "SomeKeyword" is in the document
                    
                    print(f"Running macro 'edit' on {filename}")
                    try:
                        word.Application.Run("edit")  # Pass macro name as a string
                    except Exception as e:
                        print(f"    [x] Failed to run macro 'edit' in {filename}: {e}")
                    
                    doc.Save()
                except Exception as e:
                    print(f"[!] Error processing {filename}: {e}")
                finally:
                    try:
                        doc.Close()
                    except:
                        pass
    finally:
        word.Quit()
        print("All files processed.")

if __name__ == "__main__":
    folder_path = input("Enter the folder path containing .docx files: ")
    run_macro_if_keyword_found(folder_path)
