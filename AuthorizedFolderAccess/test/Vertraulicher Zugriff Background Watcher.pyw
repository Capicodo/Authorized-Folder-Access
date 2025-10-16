import time
import win32com.client
import os

import subprocess

# Folder path to watch
import configparser

# ---------------------------
# Configuration (INI)
# ---------------------------
CONFIG_FILE = "config.ini"


def readConfig():

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    folder_path = config["Settings"]["folder_path"]

    # Get Shell Application object
    shell = win32com.client.Dispatch("Shell.Application")
    return folder_path, shell


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
                if normalize_path(current_path).startswith(normalized_target):
                    return True
        except Exception:
            pass
    return False


try:
    folder_path, shell = readConfig()
except Exception as e:
    message = (
        "❌ FEHLER – Prozess: 'Vertraulicher Zugriff Background Watcher' "
        "konnte den zu überwachenden Pfad nicht bestimmen. \n"
        "❗Bitte umgehend bei Mu melden: calvin.delloro@piluweri.de"
    )
    subprocess.run(["msg", "*", message], check=True)


normalized_target = normalize_path(folder_path)

# Wait until window is closed
print(f"Checking for Explorer window {folder_path} behaviour")


def disconnect():

    try:
        subprocess.run(["net", "use", "*", "/delete", "/y"], check=False, shell=True)

        message = "Die Netzwerkverbindungen werden getrennt"
        subprocess.run(["msg", "*", "/time:4", message], check=True, shell=True)

        print("Drives disconnected and message sent.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


while True:
    if is_target_window_open():
        print("Target window found")
        target_window_found = True

    else:
        print("Target window not found")
        if target_window_found:
            print("Explorer window was closed.")
            target_window_found = False
            disconnect()
    time.sleep(0.33333333333333333)
