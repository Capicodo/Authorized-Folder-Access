import time
import win32com.client
import os

import subprocess

# Folder path to open
folder_path = r"\\mutest\MuTest"
print(f"Opening folder: --> {folder_path} <--")

# Get Shell Application object
shell = win32com.client.Dispatch("Shell.Application")
# Open the folder
shell.Open(folder_path)


# Normalize path for comparison
def normalize_path(path):
    """
    This function converts any path to a normalized path that will be identical to any other path to the same destination.
    Changes casings, structure, symbols.
    
    Args:
        path (string): The input string.
    
    Returns:
        string: The normalized path.
    """
    return os.path.normcase(os.path.normpath(path))


normalized_target = normalize_path(folder_path)

target_window_found = False


def is_target_window_open():
    windows = shell.Windows()
    print(windows)
    for window in windows:
        try:
            if window and window.Document and hasattr(window.Document, "Folder"):
                folder = window.Document.Folder
                # Try to get path via Self first
                try:
                    current_path = folder.Self.Path
                except Exception:
                    # Fallback: get path from the first item in the folder
                    current_path = folder.Items().Item().Path
                print(current_path)
                if normalize_path(current_path) == normalized_target:
                    return True
        except Exception:
            pass
    return False


# Wait until window is closed
print("Checking for Explorer window to close...")


def disconnect():
    try:
        subprocess.run(["net", "use", "*", "/delete", "/y"], check=True, shell=True)

        message = "Die Verbindung Personalwesen wird getrennt"
        subprocess.run(["msg", "*", "/time:4", message], check=True, shell=True)

        print("Drives disconnected and message sent.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


while True:
    if is_target_window_open():
        print("Target window found")
        target_window_found = True
        time.sleep(0.5)
    else:
        print("Target window not found")
        if target_window_found:
            print("Explorer window was closed.")
            target_window_found = False
            disconnect()
            break

input("Press Enter to exit...")
