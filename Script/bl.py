import os
import time
import win32com.client as win32
import pyautogui
import subprocess

folder_path = input("Enter the folder path containing docx files: ").strip()
dbt_path = r"C:\Program Files (x86)\Duxbury\DBT 12.7\dbtw.exe"  # Update if DBT installed elsewhere

def run_word_macro(file_path):
    """Open Word file and run macro to prepare for braille."""
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(file_path)
    word.Application.Run("Braille")
    doc.Close(False)
    word.Quit()

def open_dbt():
    """Launch DBT and wait until it loads."""
    subprocess.Popen([dbt_path])
    time.sleep(5)  # Adjust if DBT takes longer to start

def convert_in_dbt(file_path):
    
    pyautogui.hotkey("ctrl", "o")
    time.sleep(2)
    pyautogui.typewrite(file_path)
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.press("enter")
    time.sleep(12)# Wait for translation/import

    # Save As (F12)
    pyautogui.press("f3")
    time.sleep(2)

    dxp_path = os.path.splitext(file_path)[0] + ".dxp"
    pyautogui.typewrite(dxp_path)
    pyautogui.press("enter")
    time.sleep(2)

    pyautogui.hotkey("ctrl", "f4")
    time.sleep(1)

def batch_process():
    open_dbt()

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")

            run_word_macro(file_path)

            convert_in_dbt(file_path)

    pyautogui.hotkey("alt", "f4")

if __name__ == "__main__":
    batch_process()
