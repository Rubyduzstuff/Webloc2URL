import tkinter as tk
from tkinter import filedialog, messagebox
import os
import plistlib
import sys
import platform

#def open_url_file(file_path):
#    """Opens the URL file using the default system application."""
#    try:
#       if platform.system() == "Windows":
#            os.startfile(file_path)
#        elif platform.system() == "Darwin":  # macOS
#            os.system(f"open '{file_path}'")
#        elif platform.system() == "Linux":
#            os.system(f"xdg-open '{file_path}'")
#        else:
#            messagebox.showerror("Error", f"Cannot open URL file on this operating system.")
#    except Exception as e:
#        messagebox.showerror("Error", f"Error opening URL file: {e}")

def create_usable_url_file(url, output_path):
    """Creates a .url file with the correct format."""
    try:
        with open(output_path, 'w') as outfile:
            outfile.write("[InternetShortcut]\n")
            outfile.write(f"URL={url}\n")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Error creating .url file: {e}")
        return False

def convert_webloc_to_url_file(webloc_path, output_folder=".", delete_original=False):
    """
    Converts a single .webloc file to a usable .url file.

    Args:
        webloc_path (str): The path to the .webloc file.
        output_folder (str, optional): The folder to save the .url file.
                                       Defaults to the same directory as the .webloc.
        delete_original (bool, optional): Whether to delete the original .webloc file
                                          after successful conversion. Defaults to False.
    """
    if not webloc_path.lower().endswith(".webloc"):
        messagebox.showerror("Error", f"Invalid file type: {webloc_path}. Must be a .webloc file.")
        return None

    try:
        with open(webloc_path, 'rb') as fp:
            plist = plistlib.load(fp)
            url = plist.get('URL')
            if not url:
                messagebox.showerror("Error", f"'URL' key not found in {webloc_path}")
                return None

            base, ext = os.path.splitext(os.path.basename(webloc_path))
            output_filename = f"{base}.url"
            output_path = os.path.join(output_folder, output_filename)

            os.makedirs(output_folder, exist_ok=True)

            if create_usable_url_file(url, output_path):
                messagebox.showinfo("Conversion Complete", f"Successfully converted '{os.path.basename(webloc_path)}' to '{output_filename}'.")
                #open_url_file(output_path)  # Open the .url file

                if delete_original:
                    try:
                        os.remove(webloc_path)
                        print(f"Deleted: {webloc_path}")
                    except OSError as e:
                        messagebox.showerror("Error", f"Error deleting '{webloc_path}': {e}")

                return output_path
            else:
                return None

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while processing {webloc_path}: {e}")
        return None

def convert_folder():
    """Handles the conversion of all .webloc files in a selected folder."""
    folder_selected = filedialog.askdirectory(title="Select Folder with .webloc Files")
    if folder_selected:
        output_folder = filedialog.askdirectory(title="Select Output Folder (optional)") or folder_selected
        delete_originals = delete_var.get()

        webloc_files_found = False
        for filename in os.listdir(folder_selected):
            if filename.lower().endswith(".webloc"):
                webloc_files_found = True
                webloc_path = os.path.join(folder_selected, filename)
                convert_webloc_to_url_file(webloc_path, output_folder, delete_originals)

        if not webloc_files_found:
            messagebox.showinfo("Info", "No .webloc files found in the selected folder.")
        else:
            messagebox.showinfo("Conversion Batch Complete", "Finished processing all .webloc files in the folder.")

def browse_webloc_file():
    """Opens a dialog for the user to select a single .webloc file."""
    file_selected = filedialog.askopenfilename(
        title="Select .webloc file",
        filetypes=[("Webloc files", "*.webloc")]
    )
    if file_selected:
        single_webloc_path_entry.delete(0, tk.END)
        single_webloc_path_entry.insert(0, file_selected)

def browse_output_single_file():
    """Opens a dialog for the user to select an output folder for a single file."""
    folder_selected = filedialog.askdirectory(title="Select Output Folder (optional)")
    if folder_selected:
        output_single_path_entry.delete(0, tk.END)
        output_single_path_entry.insert(0, folder_selected)

def convert_single_file():
    """Handles the conversion of a single selected .webloc file."""
    webloc_path = single_webloc_path_entry.get()
    output_path = output_single_path_entry.get() or os.path.dirname(webloc_path) if webloc_path else "."
    delete_original = delete_single_var.get()

    if not webloc_path:
        messagebox.showerror("Error", "Please select a .webloc file to convert.")
        return

    convert_webloc_to_url_file(webloc_path, output_path, delete_original)

# --- UI Setup ---
root = tk.Tk()
root.title(".webloc to Usable URL Converter")

# --- Single File Conversion ---
single_file_frame = tk.LabelFrame(root, text="Convert Single .webloc File")
single_file_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

webloc_label = tk.Label(single_file_frame, text="Select .webloc File:")
webloc_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

single_webloc_path_entry = tk.Entry(single_file_frame, width=50)
single_webloc_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

browse_webloc_button = tk.Button(single_file_frame, text="Browse", command=browse_webloc_file)
browse_webloc_button.grid(row=0, column=2, padx=5, pady=5)

output_single_label = tk.Label(single_file_frame, text="Output Folder (optional):")
output_single_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

output_single_path_entry = tk.Entry(single_file_frame, width=50)
output_single_path_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

browse_output_single_button = tk.Button(single_file_frame, text="Browse", command=browse_output_single_file)
browse_output_single_button.grid(row=1, column=2, padx=5, pady=5)

delete_single_var = tk.BooleanVar()
delete_single_check = tk.Checkbutton(single_file_frame, text="Delete original .webloc file", variable=delete_single_var)
delete_single_check.grid(row=2, column=0, columnspan=3, padx=5, pady=5, sticky="w")

convert_single_button = tk.Button(single_file_frame, text="Convert Single File", command=convert_single_file)
convert_single_button.grid(row=3, column=0, columnspan=3, padx=5, pady=10)

# --- Folder Conversion ---
folder_frame = tk.LabelFrame(root, text="Convert Folder of .webloc Files")
folder_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

convert_folder_button = tk.Button(folder_frame, text="Select Folder and Convert All", command=convert_folder)
convert_folder_button.grid(row=0, column=0, padx=5, pady=5)

delete_var = tk.BooleanVar()
delete_check = tk.Checkbutton(folder_frame, text="Delete original .webloc files after conversion", variable=delete_var)
delete_check.grid(row=1, column=0, padx=5, pady=5, sticky="w")

# Configure grid column weights to make the entry fields expand
root.grid_columnconfigure(0, weight=1)
single_file_frame.grid_columnconfigure(1, weight=1)

root.mainloop()