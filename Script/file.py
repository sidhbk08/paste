import os
import shutil

def get_files_without_extension(folder_path):
    """
    Get list of files in the folder without their extensions.
    """
    files = []
    for file_name in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, file_name)):
            name_without_extension = os.path.splitext(file_name)[0]
            files.append(name_without_extension)
    return files

def copy_files(source_folder, target_folder):
    """
    Copy files from the source folder to the target folder, removing the extension.
    """
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # Get the list of files in the source folder without extensions
    files_to_copy = get_files_without_extension(source_folder)
    
    for file_name in files_to_copy:
        source_file = None
        for ext in ['.docx', '.txt', '.jpg', '.pdf', '.png']:  # Add more extensions as needed
            source_file_path = os.path.join(source_folder, file_name + ext)
            if os.path.exists(source_file_path):
                source_file = source_file_path
                break
        
        if source_file:
            target_file_path = os.path.join(target_folder, file_name + os.path.splitext(source_file)[1])
            shutil.copy2(source_file, target_file_path)
            print(f"Copied: {source_file} -> {target_file_path}")
        else:
            print(f"File {file_name} not found in source folder.")

if __name__ == "__main__":
    # Ask user for source and target folder paths
    source_folder = input("Enter the path of the source folder: ").strip()
    target_folder = input("Enter the path of the target folder: ").strip()

    if os.path.exists(source_folder) and os.path.isdir(source_folder):
        if os.path.exists(target_folder) and os.path.isdir(target_folder):
            copy_files(source_folder, target_folder)
        else:
            print(f"Target folder '{target_folder}' does not exist. Please check the path.")
    else:
        print(f"Source folder '{source_folder}' does not exist. Please check the path.")
