import os
import shutil

def create_batches_in_same_folder(source_dir, batch_size=30):
    if not os.path.exists(source_dir):
        print(f"Error: The directory {source_dir} does not exist.")
        return

    all_files = [f for f in os.listdir(source_dir) if os.path.isfile(os.path.join(source_dir, f))]
    
    if not all_files:
        print(f"No files found in {source_dir}.")
        return

    all_files.sort()
    
    batch_num = 1
    for i in range(0, len(all_files), batch_size):
        batch_folder = os.path.join(source_dir, f"batch_{batch_num}")
        os.makedirs(batch_folder, exist_ok=True)
        
        batch_files = all_files[i:i + batch_size]
        
        for file in batch_files:
            source_file = os.path.join(source_dir, file)
            destination_file = os.path.join(batch_folder, file)
            shutil.move(source_file, destination_file)
        
        batch_num += 1
        print(f"Moved batch {batch_num - 1} with {len(batch_files)} files.")

source_directory = input("Please enter the path to the folder with files: ")
create_batches_in_same_folder(source_directory)
