import sys
import os
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext
import openpyxl


from PIL import Image, ImageTk

# Import the run_main_process function from main.py
from main import run_main_process

# Global flag to indicate if the stop button has been pressed
stop_flag = False

class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, message)
        self.text_widget.config(state=tk.DISABLED)
        self.text_widget.see(tk.END)  # Scroll to the bottom

    def flush(self):
        pass  # This method is needed for Python 3 compatibility

# Function to get the correct path for resources
def resource_path(relative_path):
    """ Get the absolute path to the resource, works for both development and PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Function to handle search process
def search_process():
    global stop_flag, entry_input, entry_output, log_window
    stop_flag = False  # Reset the stop flag when starting a new search

    input_file = entry_input.get().strip()
    output_file = entry_output.get().strip()

    # Perform validation on input and output paths
    if not input_file:
        log_output("Please enter a valid input CSV file path.\n")
        return

    if not output_file:
        log_output("Please enter a valid output CSV file path.\n")
        return

    try:
        # Execute the main process in a separate daemon thread to avoid freezing the GUI
        thread = threading.Thread(target=run_main_process, args=(input_file, output_file, lambda: stop_flag), daemon=True)
        thread.start()

    except Exception as e:
        log_output(f"An unexpected error occurred: {e}\n")

# Function to open file dialog for input file path (both CSV and Excel)
def browse_input_file():
    filename = filedialog.askopenfilename(filetypes=[("CSV and Excel files", "*.csv *.xlsx"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    entry_input.delete(0, tk.END)
    entry_input.insert(0, filename)


# Function to open file dialog for output CSV file path
def browse_output_file():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("CSV and Excel files", "*.csv *.xlsx"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    entry_output.delete(0, tk.END)
    entry_output.insert(0, filename)


# Function to output logs to the log window
def log_output(message):
    global log_window
    log_window.config(state=tk.NORMAL)
    log_window.insert(tk.END, message)
    log_window.config(state=tk.DISABLED)
    log_window.see(tk.END)  # Scroll to the bottom

# Function to clear the activity log
def clear_log():
    global log_window
    log_window.config(state=tk.NORMAL)
    log_window.delete(1.0, tk.END)
    log_window.config(state=tk.DISABLED)

# Function to stop the search process
def stop_process():
    global stop_flag
    stop_flag = True
    log_output("Stop signal sent. Please wait for subscribers to be processed.\n")

def main():
    global entry_input, entry_output, log_window

    # Create the main window
    root = tk.Tk()
    root.title("OJRC Search")

    # Load and set the window and taskbar icon using a properly formatted .ico file
    try:
        root.iconbitmap("bbcatpfp2.ico")  # Ensure this is the correct path to your .ico file
    except Exception as e:
        print(f"Error loading icon: {e}")

    # Set padding
    padding_y = 15
    padding_x = 10

    # Create and pack input file path widgets with padding
    label_input = tk.Label(root, text="Input CSV File:")
    label_input.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    entry_input = tk.Entry(root, width=50)
    entry_input.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    button_browse_input = tk.Button(root, text="Browse", command=browse_input_file)
    button_browse_input.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Create and pack output file path widgets with padding
    label_output = tk.Label(root, text="Output CSV File:")
    label_output.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    entry_output = tk.Entry(root, width=50)
    entry_output.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    button_browse_output = tk.Button(root, text="Browse", command=browse_output_file)
    button_browse_output.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Create and pack search button widget with padding
    button_search = tk.Button(root, text="Search", command=search_process)
    button_search.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Add a title above the log window
    label_log = tk.Label(root, text="Activity Log", font=("Arial", 14, "bold"))
    label_log.pack(pady=(padding_y, 0), padx=padding_x, anchor=tk.W)

    # Create a scrolled text widget for logging
    log_window = scrolledtext.ScrolledText(root, width=80, height=10, state=tk.DISABLED)
    log_window.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Create and pack clear button widget with padding
    button_clear = tk.Button(root, text="Clear", command=clear_log)
    button_clear.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Create and pack stop button widget with padding
    button_stop = tk.Button(root, text="Stop", command=stop_process)
    button_stop.pack(pady=padding_y, padx=padding_x, anchor=tk.W)

    # Load and place the background image properly in the top-right corner
    try:
        # Use resource_path to get the correct path for the image
        bg_image_path = resource_path("bbcatprofile.jpg")
        bg_image = Image.open(bg_image_path)
        bg_photo = ImageTk.PhotoImage(bg_image)

        # Create a label with the image
        bg_label = tk.Label(root, image=bg_photo)
        bg_label.image = bg_photo  # Keep a reference to prevent garbage collection

        # Use relative positioning to place the image slightly more to the right within the top-right corner
        bg_label.place(relx=0.95, rely=0.1, anchor='ne')
    except Exception as e:
        print(f"Error loading background image: {e}")

    # Redirect stdout and stderr to the log window
    sys.stdout = TextRedirector(log_window)
    sys.stderr = TextRedirector(log_window)

    # Run the Tkinter main loop
    root.mainloop()

if __name__ == "__main__":
    main()
