import os
import time

# Define the folder where the .msg files are located
folder_path = r"C:\Users\Public\Downloads"

# Define the number of retries and delay between them
retries = 5
delay = 2  # seconds

def delete_msg_files(folder_path, retries=5, delay=2):
    for filename in os.listdir(folder_path):
        if filename.endswith(".msg"):
            file_path = os.path.join(folder_path, filename)
            for i in range(retries):
                try:
                    os.remove(file_path)
                    print(f"Deleted MSG file: {file_path}")
                    break
                except Exception as e:
                    print(f"Attempt {i+1} failed to delete MSG file {file_path}: {e}")
                    time.sleep(delay)
            else:
                print(f"Failed to delete MSG file after {retries} attempts: {file_path}")

# Run the cleanup function
delete_msg_files(folder_path, retries=retries, delay=delay)
