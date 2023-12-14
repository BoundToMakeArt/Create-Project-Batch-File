import subprocess
import sys
import time
import ctypes

# Function to change console text color on Windows
def set_color(color):
    STD_OUTPUT_HANDLE = -11
    handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)

# Windows color codes
FOREGROUND_GREEN = 0x0A
FOREGROUND_RED = 0x0C
FOREGROUND_RESET = 0x07

def update_pip():
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])

def print_with_color(text, color):
    if color:
        set_color(color)
        print(text)
        set_color(FOREGROUND_RESET)  # Reset color to default
    else:
        print(text)

def install_dependencies():
    dependencies = [
        'pyqt5',
        'pywin32',
        'pywin32-ctypes',
        'win10toast',
    ]
    for dependency in dependencies:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', dependency])

if __name__ == "__main__":
    try:
        update_pip()
        install_dependencies()
        print_with_color("\n" * 4 + "Dependencies installed successfully.", FOREGROUND_GREEN)  # Display in green
        time.sleep(5)  # Add a 5-second delay
    except subprocess.CalledProcessError as e:
        print_with_color("\n" * 4 + f"Error occurred: {e}", FOREGROUND_RED)  # Display in red
        time.sleep(5)  # Add a 5-second delay
