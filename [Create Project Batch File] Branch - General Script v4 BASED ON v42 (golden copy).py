import os
import sys
import tkinter as tk
from tkinter import filedialog
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QPushButton, QVBoxLayout, QWidget, QCheckBox
from PyQt5.QtCore import Qt, QMimeData
import pythoncom
import win32com.client
import ctypes
import re
from PyQt5.QtCore import Qt
from win10toast import ToastNotifier

# Get the absolute path of the current script
script_path = os.path.realpath(__file__)

# Define the path to the icon.ico file in the util folder
icon_path = os.path.join(os.path.dirname(script_path), "[Create Project Batch File] Util", "icon.ico")

# Define the path to the toast_notification.py file
util_folder_path = os.path.join(os.path.dirname(script_path), "[Create Project Batch File] Util")
toast_notification_path = os.path.join(util_folder_path, "toast_notification.py")
config_ini_path = os.path.join(util_folder_path, "config.ini")

# Replace backslashes with double backslashes for Python string format
icon_path_formatted = icon_path.replace('\\', '\\\\')

# Enclose toast_notification_path in double quotes
toast_notification_path_formatted = f'"{toast_notification_path}"'

# Content for toast_notification.py
content = f"""
from win10toast import ToastNotifier

def show_notification(title, message, icon_path=None):
    toast = ToastNotifier()
    toast.show_toast(title, message, icon_path=icon_path, duration=10)

if __name__ == "__main__":
    # Change these values as needed
    notification_title = "Everything you selected is now open"
    notification_message = "Good luck with your project!"
    notification_icon_path = "{icon_path_formatted}"  # Dynamic path to icon.ico

    show_notification(notification_title, notification_message, notification_icon_path)
"""

# Write the content to toast_notification.py
with open(toast_notification_path, "w") as file:
    file.write(content)

# Write the content to toast_notification.py
with open("[Create Project Batch File] Util/toast_notification.py", "w") as file:
    file.write(content)



def extract_path_from_lnk(lnk_path):
    try:
        if not lnk_path.lower().endswith(('.lnk', '.url')):
            return None
        else:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(lnk_path)
            target_path = shortcut.TargetPath
            return target_path
    except Exception as e:
        print(f"Error: {e}")
        return None

class DragDropWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.file_paths = []
        self.urls = []
        self.dark_mode = self.read_dark_mode_setting()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Drag and Drop Files')
        self.setGeometry(300, 300, 400, 400)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.dark_mode_checkbox = QCheckBox('Dark Mode')
        self.dark_mode_checkbox.stateChanged.connect(self.toggle_dark_mode)
        self.dark_mode_checkbox.setStyleSheet("color: black; border: 1px solid grey;")
        self.layout.addWidget(self.dark_mode_checkbox)

        self.text_edit = QTextEdit()
        self.text_edit.setAcceptDrops(True)
        self.text_edit.textChanged.connect(self.update_file_paths)  # Update file_paths on text change
        self.layout.addWidget(self.text_edit)

        self.write_exit_button = QPushButton('Write to file and exit')
        self.write_exit_button.clicked.connect(self.write_to_file_and_exit)
        self.write_exit_button.setStyleSheet("border: 1px solid grey;")
        self.layout.addWidget(self.write_exit_button)

        self.select_files_button = QPushButton('Select Files/Folders')
        self.select_files_button.clicked.connect(self.select_files)
        self.select_files_button.setStyleSheet("border: 1px solid grey;")
        self.layout.addWidget(self.select_files_button)

        self.set_checkbox_state()
        self.text_edit.dropEvent = self.drop_event

        # Set window flag to stay on top
        self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)

        self.show()

    def set_checkbox_state(self):
        self.dark_mode_checkbox.setChecked(self.dark_mode)

    def toggle_dark_mode(self, state):
        self.dark_mode = state == Qt.Checked
        self.save_dark_mode_setting()  # Save the new state to config.ini
        self.update_checkbox_text()
        self.apply_styles()

    def update_checkbox_text(self):
        self.dark_mode_checkbox.setText('Light Mode' if self.dark_mode else 'Dark Mode')

    def apply_styles(self):
        if self.dark_mode:
            self.setStyleSheet("background-color: black; color: white;")
            self.dark_mode_checkbox.setStyleSheet("color: white; border: 1px solid grey;")
        else:
            self.setStyleSheet("background-color: white; color: black;")
            self.dark_mode_checkbox.setStyleSheet("color: black; border: 1px solid grey;")

    def read_dark_mode_setting(self):
        if not os.path.exists(config_ini_path):
            # If the config.ini doesn't exist, create it and set the default dark mode as False
            self.save_dark_mode_setting(False)
            return False
        
        with open(config_ini_path, 'r') as file:
            return file.read() == 'True'

    def save_dark_mode_setting(self, value=None):
        if value is None:
            value = self.dark_mode
        
        with open(config_ini_path, 'w') as file:
            file.write(str(value))

    def drop_event(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]

        for file in files:
            if file.lower().endswith(('.lnk', '.url')):
                actual_path = extract_path_from_lnk(file)
                if actual_path:
                    self.file_paths.append(actual_path)
                    self.text_edit.append(actual_path)
            else:
                self.file_paths.append(file)
                self.text_edit.append(file)

        # Check for dropped URLs that don't have a file extension
        text = event.mimeData().text()
        urls = re.findall(r'(https?://\S+)', text)
        for url in urls:
            self.urls.append(url)
            self.text_edit.append(url)

    def select_files(self):
        root = tk.Tk()
        root.withdraw()

        file_paths = filedialog.askopenfilenames(title="Select Files/Folders")

        if file_paths:
            file_paths = root.tk.splitlist(file_paths)
            self.file_paths.extend(file_paths)

            for file_path in file_paths:
                self.text_edit.append(file_path)

    def update_file_paths(self):
        text = self.text_edit.toPlainText()
        current_paths = text.split('\n')
        self.file_paths = [path for path in current_paths if os.path.exists(path)]

    def write_to_file_and_exit(self):
        if self.file_paths or self.urls:
            with open('open_prepared_items.bat', 'w') as batch_file:
                batch_file.write('@echo off\n')

                largest_pur_delay = 1  # To keep track of the delay for the largest ".pur" file

                for path in self.file_paths:
                    if os.path.exists(path):
                        if path.lower().endswith('.pur'):
                            file_size = get_file_size(path)
                            delay = calculate_delay(file_size)
                            largest_pur_delay = max(largest_pur_delay, delay)
                            batch_file.write(f'start "" "{path}"\n')
                            batch_file.write(f'timeout /t 1\n')  # Default 1-second delay after each ".pur" file
                        else:
                            batch_file.write(f'start "" "{path}"\n')
                            batch_file.write('timeout /t 1\n')

                for url in self.urls:
                    batch_file.write(f'start "" "{url}"\n')
                    batch_file.write('timeout /t 1\n')

                batch_file.write(f'timeout /t {largest_pur_delay}\n')  # Large delay after the largest ".pur" file
                batch_file.write(f"start pythonw {toast_notification_path_formatted}\n")  # Run PowerShell script
                batch_file.write('taskkill /f /im cmd.exe\n')  # Close the command prompt window

            print("Batch file 'open_prepared_items.bat' created successfully!")
        else:
            print("No files or URLs to write.")

        self.close()


def calculate_delay(file_size):
    if file_size > 500 * 1024 * 1024:  # Check if size is greater than 500MB
        extra_delay = ((file_size - 500 * 1024 * 1024) // (500 * 1024 * 1024)) * 32 + 32
        return extra_delay
    return 1  # Default delay is 1 second


def get_target_file_extension(path):
    if path.lower().endswith('.lnk'):
        pythoncom.CoInitialize()  # Initialize the COM library
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        target_path = shortcut.Targetpath
        return os.path.splitext(target_path)[1].lower() if os.path.exists(target_path) else ""
    return os.path.splitext(path)[1].lower() if os.path.exists(path) else ""


def get_file_size(file_path):
    if file_path.lower().endswith('.lnk'):
        pythoncom.CoInitialize()  # Initialize the COM library
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(file_path)
        target_path = shortcut.Targetpath
        if os.path.exists(target_path):
            return os.path.getsize(target_path)
    elif os.path.exists(file_path):
        return os.path.getsize(file_path)
    return 0


def create_batch_file(file_paths):
    with open('open_prepared_items.bat', 'w') as batch_file:
        batch_file.write('@echo off\n')

        last_pur_index = -1  # To keep track of the last ".pur" file index
        largest_pur_delay = 1  # To keep track of the delay for the largest ".pur" file
        for i, path in enumerate(file_paths):
            if os.path.exists(path):
                if os.path.isfile(path):
                    file_size = get_file_size(path)
                    delay = calculate_delay(file_size)
                    if path.lower().endswith('.pur') or get_target_file_extension(path).endswith('.pur'):
                        last_pur_index = i
                        largest_pur_delay = max(largest_pur_delay, delay)
                        batch_file.write(f'start "" "{path}"\n')
                        batch_file.write(f'timeout /t 1\n')  # Default 1-second delay after each ".pur" file
                    else:
                        batch_file.write(f'start "" "{path}"\n')
                        batch_file.write(f'timeout /t {delay}\n')
                elif os.path.isdir(path):
                    batch_file.write(f'start "" "{path}"\n')
                    batch_file.write(f'timeout /t 1\n')  # Default 1-second delay after each folder

        batch_file.write(f'timeout /t 1\n')
        batch_file.write(f'timeout /t {largest_pur_delay}\n')  # Large delay after the largest ".pur" file
        batch_file.write(f"start pythonw {toast_notification_path_formatted}\n")  # Run PowerShell script
        batch_file.write('taskkill /f /im cmd.exe\n')  # Close the command prompt window


def minimize_cmd_window():
    # Get the handle for the console window and minimize it
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 6)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DragDropWindow()
    minimize_cmd_window()  # Minimize the command line window

    if window.dark_mode:
        window.setStyleSheet("background-color: black; color: white;")
    else:
        window.setStyleSheet("background-color: white; color: black;")

    window.show()
    sys.exit(app.exec_())
