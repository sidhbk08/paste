import os
import time
import subprocess
from pywinauto import Application, Desktop
import win32com.client as win32


folder_path = input("Enter the folder path containing docx files: ").strip()
dbt_path = r"C:\Program Files (x86)\Duxbury\DBT 12.7\dbtw.exe"  # Update if DBT installed elsewhere


def run_word_macro(file_path):
    """Open Word file and run macro to prepare for braille."""
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(file_path)
    word.Application.Run("Braille")  # assumes macro named "Braille"
    doc.Close(False)
    word.Quit()


def open_dbt():
    """Launch DBT and return pywinauto app object."""
    app = Application(backend="uia").start(dbt_path)
    time.sleep(5)  # allow DBT to fully load
    return app


def convert_in_dbt(app, file_path):
    """Automate DBT using pywinauto instead of pyautogui."""

    # Always grab the main DBT window freshly
    main = Desktop(backend="uia").window(title_re=".*Duxbury.*")
    main.wait("visible", timeout=15)

    # Open file dialog
    main.type_keys("^o")  # CTRL+O
    open_dlg = Desktop(backend="uia").window(title_re=".*Select.*file.*")
    open_dlg.wait("visible", timeout=15)

    # File name input
    try:
        file_edit = open_dlg.child_window(auto_id="1148", control_type="Edit")
        file_edit.set_edit_text(file_path)
    except Exception:
        file_edit = open_dlg.child_window(title="File name:", control_type="Edit")
        file_edit.set_edit_text(file_path)

    # Click Open button (handle multiple matches)
    open_buttons = open_dlg.descendants(control_type="Button", title="Open")
    if open_buttons:
        open_buttons[-1].click_input()  # pick last one
    else:
        raise RuntimeError("Open button not found in file dialog")

    # Handle Import File dialog (press Enter to confirm)
    try:
        import_dlg = Desktop(backend="uia").window(title_re=".*Import.*")
        import_dlg.wait("visible", timeout=15)
        print("DEBUG: Import dialog found")
        import_dlg.type_keys("{ENTER}")  # "OK" is default
    except Exception as e:
        print("No Import File dialog detected:", e)

    # Re-acquire main DBT window after import
    main = Desktop(backend="uia").window(title_re=".*The Duxbury Braille Translator*")
    main.wait("visible", timeout=20)

    # Wait for DBT to finish translation
    time.sleep(10)

    # Save as .dxp
    main.type_keys("{F3}")  # F3 = Save As in DBT
    save_dlg = Desktop(backend="uia").window(title_re=".*Save As.*")
    save_dlg.wait("visible", timeout=15)

    dxp_path = os.path.splitext(file_path)[0] + ".dxp"

    try:
        save_edit = save_dlg.child_window(auto_id="1001", control_type="Edit")
        save_edit.set_edit_text(dxp_path)
    except Exception:
        save_edit = save_dlg.child_window(title="File name:", control_type="Edit")
        save_edit.set_edit_text(dxp_path)

    save_btn = save_dlg.child_window(title="Save", control_type="Button")
    save_btn.click_input()
    time.sleep(2)

    # Close the file in DBT
    main.type_keys("^{{F4}}")
    time.sleep(1)


def batch_process():
    app = open_dbt()

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing {filename}...")

            run_word_macro(file_path)
            convert_in_dbt(app, file_path)

    # Close DBT when done
    main = Desktop(backend="uia").window(title_re=".*The Duxbury Braille Translator*")
    main.type_keys("%{F4}")  # Alt+F4 to quit


if __name__ == "__main__":
    batch_process()
