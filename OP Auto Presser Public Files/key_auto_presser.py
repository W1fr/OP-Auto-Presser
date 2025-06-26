import time
import threading
import tkinter as tk
from tkinter import ttk
from pynput.keyboard import Controller
import os # Import the os module to help with path handling
import sys # Import sys module to check if running from PyInstaller
import ctypes # Import ctypes for direct Windows API calls

# Initialize the keyboard controller
keyboard = Controller()

# Global control for the key pressing loop
running = False

# --- Functions ---

def press_key(key, press_duration):
    """
    Simulates a key press, holds it for a specified duration, then releases it.
    """
    keyboard.press(key)
    # Hold the key down for the specified duration
    time.sleep(press_duration)
    keyboard.release(key)

def run_independent(w_delay, s_delay, press_duration):
    """
    Starts two separate daemon threads to press 'w' and 's' keys independently.
    Each key has its own delay and press duration.
    Daemon threads automatically close when the main program exits.
    """
    def press_w():
        """Thread function to continuously press 'w' key."""
        while running:
            press_key('w', press_duration)
            time.sleep(w_delay)

    def press_s():
        """Thread function to continuously press 's' key."""
        while running:
            press_key('s', press_duration)
            time.sleep(s_delay)

    threading.Thread(target=press_w, daemon=True).start()
    threading.Thread(target=press_s, daemon=True).start()

def run_alternate(w_delay, s_delay, press_duration):
    """
    Starts a single daemon thread to press 'w' then 's' alternately,
    with their respective delays between presses and a shared press duration.
    """
    def run():
        """Thread function to press 'w' and 's' in alternation."""
        while running:
            press_key('w', press_duration)
            time.sleep(w_delay)
            # Check if 'running' is still True after the first sleep.
            # This is important in case 'Stop' was pressed during the delay.
            if not running:
                break
            press_key('s', press_duration)
            time.sleep(s_delay)

    threading.Thread(target=run, daemon=True).start()

def start_pressing():
    """
    Initiates the key pressing based on the selected mode and delay values.
    Handles input validation for delays and updates the status label.
    """
    global running
    try:
        # Get delay values and the new press length from entry widgets
        w_delay = float(w_entry.get())
        s_delay = float(s_entry.get())
        key_press_length = float(press_length_entry.get()) # Get the new press length

        # Basic validation for positive values
        if w_delay < 0 or s_delay < 0 or key_press_length < 0:
            status_label.config(text="Delays and press length must be positive numbers!", foreground="red")
            return

    except ValueError:
        # Display an error message if the input is not a valid number
        status_label.config(text="Invalid number! Please enter numbers for delays and press length.", foreground="red")
        return

    # Only start the pressing action if it's not already running
    if not running:
        running = True
        status_label.config(text="Running...", foreground="green") # Update status label
        # Call the appropriate run function based on the selected mode, passing press_duration
        if mode.get() == "independent":
            run_independent(w_delay, s_delay, key_press_length)
        else:
            run_alternate(w_delay, s_delay, key_press_length)

def stop_pressing():
    """
    Stops all active key pressing threads by setting the global 'running' flag to False.
    This flag is checked by the while loops in the thread functions.
    Updates the status label.
    """
    global running
    running = False
    status_label.config(text="Stopped", foreground="gray") # Update status label

# --- GUI Setup ---
# Create the main application window
root = tk.Tk()
root.title("OP Auto Presser") # Set the title of the window to "OP Auto Presser"
root.geometry("600x250") # Make the window wider and a bit shorter
root.resizable(False, False) # Prevent the user from resizing the window

# --- Icon Setup (using ctypes for robustness) ---
# This function helps find the correct path for resources (like the icon)
# whether the script is run normally or as a PyInstaller executable.
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError: # Happens when not running as a frozen executable
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Try to set the icon using ctypes for robustness in bundled executables
try:
    # Use root.winfo_id() to get the window handle (HWND) for the Tkinter window
    hwnd = root.winfo_id()

    # Define Windows API constants (values taken from winuser.h)
    WM_SETICON = 0x0080
    ICON_SMALL = 0
    ICON_LARGE = 1
    LR_LOADFROMFILE = 0x00000010 # Load from file
    IMAGE_ICON = 1 # Type of image to load (icon)

    icon_path = resource_path("op_logo.ico")

    if os.path.exists(icon_path):
        # Load large icon (typically 32x32 for taskbar/alt-tab)
        hicon_large = ctypes.windll.user32.LoadImageW(
            0, # hinst (handle to instance, 0 for file)
            icon_path, # name (path to file)
            IMAGE_ICON, # type (icon)
            0, 0, # cx, cy (desired width, height; 0,0 uses original size)
            LR_LOADFROMFILE # flags
        )
        # Load small icon (typically 16x16 for title bar)
        hicon_small = ctypes.windll.user32.LoadImageW(
            0,
            icon_path,
            IMAGE_ICON,
            0, 0,
            LR_LOADFROMFILE
        )

        # Send WM_SETICON messages to the window to set the icons
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
        ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_LARGE, hicon_large)

        # Optional: Set App User Model ID for better taskbar grouping on Windows 7+
        # This helps Windows recognize your app and group it correctly.
        myappid = u'OPAutoPresser.MyCompany.MyProduct.1.0' # Arbitrary unique string for your app
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except AttributeError:
            pass # Not available on older Windows versions, ignore error
    else:
        print(f"Warning: Icon file '{icon_path}' not found for ctypes. Window will use default icon.")

except Exception as e:
    print(f"Error setting icon with ctypes: {e}")


# Define Tkinter variables AFTER the root window is created to avoid "Too early to create variable" error
mode = tk.StringVar(value="independent") # Variable to hold the selected mode (Independent or Alternate)
always_on_top = tk.BooleanVar(value=True) # Variable to control the "Always on top" feature

def toggle_always_on_top():
    """
    Toggles the 'always on top' attribute of the main window based on the checkbox state.
    """
    root.attributes("-topmost", always_on_top.get())

# Set the initial 'always on top' state of the window when it starts
root.attributes("-topmost", always_on_top.get())

# --- UI Elements ---

# Frame for Delay and Press Length inputs
input_frame = ttk.Frame(root)
input_frame.pack(pady=10)

# W key delay
w_delay_frame = ttk.Frame(input_frame)
w_delay_frame.pack(side='left', padx=10)
ttk.Label(w_delay_frame, text="W key delay (sec):").pack()
w_entry = ttk.Entry(w_delay_frame, width=8)
w_entry.insert(0, "3") # Default value
w_entry.pack()

# S key delay
s_delay_frame = ttk.Frame(input_frame)
s_delay_frame.pack(side='left', padx=10)
ttk.Label(s_delay_frame, text="S key delay (sec):").pack()
s_entry = ttk.Entry(s_delay_frame, width=8)
s_entry.insert(0, "3") # Default value
s_entry.pack()

# Key Press Length
press_length_frame = ttk.Frame(input_frame)
press_length_frame.pack(side='left', padx=10)
ttk.Label(press_length_frame, text="Key press length (sec):").pack()
press_length_entry = ttk.Entry(press_length_frame, width=8)
press_length_entry.insert(0, "0.1") # Default value for key press duration (e.g., 0.1 seconds)
press_length_entry.pack()

# Frame for Mode selection
mode_frame = ttk.Frame(root)
mode_frame.pack(pady=5)
ttk.Label(mode_frame, text="Mode:").pack(side='left', padx=10)
ttk.Radiobutton(mode_frame, text="Independent", variable=mode, value="independent").pack(side='left', padx=5)
ttk.Radiobutton(mode_frame, text="Alternate (W then S)", variable=mode, value="alternate").pack(side='left', padx=5)

# Always on top checkbox
ttk.Checkbutton(root, text="Always on top", variable=always_on_top, command=toggle_always_on_top).pack(pady=5)

# Frame for buttons
button_frame = ttk.Frame(root)
button_frame.pack(pady=5)
ttk.Button(button_frame, text="Start", command=start_pressing).pack(side='left', padx=10)
ttk.Button(button_frame, text="Stop", command=stop_pressing).pack(side='left', padx=10)

# Status label to display the current status of the auto-presser
status_label = ttk.Label(root, text="Stopped", foreground="gray")
status_label.pack(pady=10)

# Start the Tkinter event loop. This makes the GUI responsive and keeps it running
# until the window is closed.
root.mainloop()
