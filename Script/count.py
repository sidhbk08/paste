import os
import win32com.client

def count_pages_in_word_files(folder_path):
    # Create a Word application object
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False  # Keep Word hidden

    # List to store the number of pages for each file
    page_counts = {}

    # Iterate through all files in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx') or filename.endswith('.doc'):
            file_path = os.path.join(folder_path, filename)
            try:
                # Open the document
                doc = word.Documents.Open(file_path)
                # Get the number of pages
                page_count = doc.ComputeStatistics(2)  # 2 corresponds to wdStatisticPages
                page_counts[filename] = page_count
                # Close the document
                doc.Close(False)
            except Exception as e:
                print(f"Error processing {filename}: {e}")

    # Quit the Word application
    word.Quit()

    return page_counts

# Get the folder path from the user
folder_path = input("Please enter the full path to the folder containing Word files: ")

# Check if the folder exists
if os.path.isdir(folder_path):
    page_counts = count_pages_in_word_files(folder_path)
    
    # Print the results
    for filename, count in page_counts.items():
        print(f"{count}")
else:
    print("The specified folder does not exist.")