from pywinauto import Desktop
import os
import time

target_path = r"C:\Windows"
normalized_target = os.path.normcase(os.path.normpath(target_path))

"""This function converts any path to a normalized path that will be identical to any other path to the same destination"""
def normalize_path(path):
    norm = os.path.normcase(os.path.normpath(path))
    if norm.endswith('\\'):
        norm = norm[:-1]
    return norm

def get_explorer_folder_paths():
    windows = Desktop(backend="uia").windows(class_name="CabinetWClass")
    paths = []
    for win in windows:
        try:
            addr_bar = win.child_window(title="Address", control_type="ToolBar")
            edit = addr_bar.child_window(control_type="Edit")
            if edit.exists():
                folder_path = edit.get_value()
                if folder_path:
                    paths.append(folder_path)
            else:
                # Fallback: combine breadcrumb buttons
                btns = addr_bar.children(control_type="Button")
                breadcrumb = "\\".join(btn.window_text() for btn in btns)
                if breadcrumb:
                    paths.append(breadcrumb)
        except Exception:
            continue
    return paths

def is_target_window_open():
    for path in get_explorer_folder_paths():
        if normalize_path(path) == normalized_target:
            return True
    return False

print(f"Opening folder: {target_path}")
os.startfile(target_path)

time.sleep(2)
print("Waiting for Explorer window to open the folder...")
for _ in range(30):  # wait max 30 seconds
    if is_target_window_open():
        print("Explorer window detected!")
        break
    time.sleep(1)
else:
    print("Explorer window NOT detected!")

input("Press Enter to exit...")
