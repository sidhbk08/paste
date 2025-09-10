import os
import win32com.client

def run_macro_if_keyword_found(folder_path, keyword, macro_name):
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False  # Set to False if you want it hidden

    try:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(('.doc', '.docx')):
                full_path = os.path.join(folder_path, filename)
                print(f"Processing file: {filename}")

                try:
                    doc = word.Documents.Open(full_path)

                    # === IMPORTANT: Show field codes ===
                    word.ActiveWindow.View.ShowFieldCodes = True

                    keyword_found = False

                    for section in doc.Sections:
                        # Header IDs: 1=Primary, 2=FirstPage, 3=EvenPages
                        for header_id in range(1, 4):
                            header_range = section.Headers(header_id).Range
                            find = header_range.Find
                            find.Text = keyword
                            find.Forward = True
                            find.Wrap = 1  # wdFindContinue
                            
                            if find.Execute():
                                keyword_found = True

                                # Move range start to the beginning of the keyword
                                header_range.Start = header_range.Start + find.Parent.Start - header_range.Start

                                # Optional: Select it visibly
                                header_range.Select()
                                print(f"    [✓] Keyword found in header {header_id} of section. Moved to its start.")
                                break

                        if keyword_found:
                            break

                    if keyword_found:
                        try:
                            print(f"Running macro '{macro_name}' on {filename}")
                            word.Application.Run(macro_name)
                        except Exception as e:
                            print(f"    [x] Failed to run macro '{macro_name}' in {filename}: {e}")
                    else:
                        print(f"    [x] Keyword '{keyword}' not found in headers of {filename}. Macro '{macro_name}' skipped.")

                    # Optional: Hide field codes again after processing

                    try:
                        word.Run("changefield")
                    except Exception as e:
                        print(f"    [x] See Claim details found, Macro 'header' skipped in {os.path.basename(doc_path)}: {e}")
                        all_success = False

                    word.ActiveWindow.View.ShowFieldCodes = False

                    doc.Save()
                    doc.Close()
                except Exception as e:
                    print(f"[!] Error processing {filename}: {e}")
                    try:
                        doc.Close()
                    except:
                        pass
    finally:
        word.Quit()
        print("All files processed.")

if __name__ == "__main__":
    folder_path = input("Enter the folder path containing .docx files: ")
    keyword_to_find = "Claim detail for"
    macro_to_run = "header"
  
    run_macro_if_keyword_found(folder_path, keyword_to_find, macro_to_run)

    