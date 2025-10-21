"""
Authorized Folder Access – Background Watcher
=============================================

Monitors Windows Explorer windows for access to a confidential network folder.
When the target folder is closed, this script automatically disconnects all
active network drives and user sessions to prevent unauthorized access.

Background
----------
This utility supports PCs used by multiple people sharing one Windows account.
Certain users have specific network credentials granting access to confidential
directories. Normally, these sessions persist even after the folder is closed,
leaving the folder accessible to others until the Windows user logs off.

This script ensures that once the confidential folder is no longer open in
Windows Explorer, all related network sessions are terminated automatically.

Features
--------
- Continuously monitors open Windows Explorer windows.
- Detects if a confidential folder (defined in `config.ini`) is open.
- Disconnects network drives and user sessions upon folder closure.
- Displays Windows message notifications when disconnections occur.
- Runs silently in the background via Windows AutoStart.

Configuration
-------------
The configuration file `config.ini` must be located in the same directory as
this script. Example structure:

    [Settings]
    folder_path = \\mutest\MuTest

Usage
-----
1. Place the folder `Vertraulicher Zugriff Background Watcher` in:
       C:\Program Files
2. Edit the `config.ini` file to set the correct `folder_path`.
3. Create a shortcut to this script and move it to:
       C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup
4. Restart the computer — the script will run automatically at login.

Exceptions
----------
If the configuration or shell access fails, a Windows message is displayed:
“❌ FEHLER – Prozess: 'Vertraulicher Zugriff Background Watcher' konnte den zu
überwachenden Pfad nicht bestimmen. Bitte umgehend bei Mu melden.”

Functions
---------
- `readConfig()`: Reads folder path from `config.ini`.
- `getShell()`: Returns the Shell.Application COM object.
- `normalize_path(path)`: Normalizes file paths for comparison.
- `is_target_window_open()`: Checks whether the target folder is currently open.
- `disconnect()`: Disconnects network drives and active sessions.

Author
------
Mu Dell'Oro
Version: 1.0.0
Date: 16.10.2025
GitHub: https://github.com/Capicodo/Authorized-Folder-Access

"""

import configparser
import os
import subprocess
import sys
import time
import win32com.client


# ---------------------------
# Configuration (INI)
# ---------------------------

"""The Path of the python file gets joined, so config.ini can be found relative to the python file"""
if getattr(sys, "frozen", False):
    # Running as PyInstaller .exe
    base_dir = os.path.dirname(sys.executable)
else:
    # Running as .py script
    base_dir = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE_PATH = os.path.join(base_dir, "config.ini")

CHECK_INTERVAL = 0.333
# ---------------------------
# Functions
# ---------------------------


def read_config():
    """Reads the configuration file from CONFIG_FILE_PATH.

    Returns:
        folder_path (str): The normalized folder path.
    """
    config = configparser.ConfigParser()

    if not os.path.exists(CONFIG_FILE_PATH):
        message = (
            "❌ FEHLER – 'config.ini' fehlt. "
            "Bitte den Pfad zum zu überwachenden Ordner in config.ini setzen."
        )
        subprocess.run(["msg", "*", message], check=True)
        sys.exit(1)

    config.read(CONFIG_FILE_PATH)
    folder_path = config["Settings"]["folder_path"]

    return folder_path


def get_shell():
    """Gets the Shell Application

    Returns:
        shell (win32com.client.CDispatch): The Shell Application object.
    """

    shell = win32com.client.Dispatch("Shell.Application")
    return shell


def normalize_path(path):
    """Converts a file path to a normalized, consistent format.

    Normalization makes equivalent paths identical by standardizing
    casing, structure, and symbols.

    Args:
        path (str): The input file path.

    Returns:
        str: The normalized path.
    """

    return os.path.normcase(os.path.normpath(path))


def is_target_window_open(shell, normalized_path):
    """Checks whether a target folder window is currently open.

    Iterates through all open Windows Explorer windows and determines
    if any of them correspond to a folder whose path starts with the
    specified target path.

    Args:
        shell (win32com.client.CDispatch): The Shell.Application COM object.
        normalized_path (str): The normalized path of the target folder.

    Returns:
        bool: True if a folder window with a path starting with the
        target path is open, otherwise False.
    """
    windows = shell.Windows()
    for window in windows:
        try:
            if window and window.Document and hasattr(window.Document, "Folder"):
                folder = window.Document.Folder
                try:
                    current_path = folder.Self.Path
                except Exception:
                    current_path = folder.Items().Item().Path
                if normalize_path(current_path).startswith(normalized_path):
                    return True
        except Exception:
            pass
    return False


def disconnect():
    """Disconnects all network drives and active user sessions.

    Runs batch commands in a subprocess to terminate all active
    network connections and end the current authorization session.
    After disconnection, users will need to re-enter their credentials.

    Warning:
        The disconnection process can take up to 30 seconds.

    Raises:
        subprocess.CalledProcessError: If an error occurs while
        running the subprocess commands.
    """
    try:
        subprocess.run(["net", "use", "*", "/delete", "/y"], check=False, shell=True)

        message = "Die Netzwerkverbindungen werden getrennt"
        subprocess.run(["msg", "*", "/time:4", message], check=True, shell=True)

        print("Drives disconnected and message sent.")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")


# ---------------------------
# Main Loop
# ---------------------------


def main():
    """Main monitoring loop that watches the target folder three times per second"""
    try:
        folder_path = read_config()
        shell = get_shell()
    except Exception:
        message = (
            "❌ FEHLER – Prozess: 'Vertraulicher Zugriff Background Watcher' "
            "konnte den zu überwachenden Pfad nicht bestimmen. \n"
            "❗Bitte umgehend bei Mu melden: calvin.delloro@piluweri.de"
        )
        subprocess.run(["msg", "*", message], check=True)
        return

    normalized_path = normalize_path(folder_path)
    target_window_is_open = False

    print(f"Checking for Explorer window {folder_path} behaviour")

    while True:
        if is_target_window_open(shell, normalized_path):
            print("Target window found")
            target_window_is_open = True
        else:
            print("Target window not found")
            if target_window_is_open:
                print("Explorer window was closed.")
                target_window_is_open = False
                disconnect()
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
