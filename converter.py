import tkinter as tk
from tkinter import filedialog
from docx import Document
import os
import re
import win32com.client

def convert_doc_to_docx(doc_file):
    try:
        # Initialize Word application
        word = win32com.client.Dispatch("Word.Application")
        
        # Open .doc file
        print("Opening .doc file:", doc_file)
        
        doc = word.Documents.Open(doc_file)
        print("Opened the file:", doc_file)
        # Output file name without whitespace
        output_filename = os.path.splitext(doc_file)[0] + '.docx'
        docx_file = output_filename.replace(" ", "_")  # Replace whitespace with underscore

        # Save as .docx
        doc.SaveAs2(docx_file, FileFormat=16)

        # Close the document and quit Word application
        doc.Close()
        word.Quit()

        # Delete the original .doc file
        os.remove(doc_file)

        return docx_file
    except Exception as e:
        print("Error:", e)
        return None
    
def rename_files(files):
    renamed_files = []
    original_filenames = []
    for file in files:
        # Get the directory and filename
        directory, filename = os.path.split(file)
        # Store the original filename
        original_filenames.append(filename)
        # Remove whitespace from the filename
        new_filename = os.path.join(directory, filename.replace(" ", "_"))
        # Rename the file
        os.rename(file, new_filename)
        renamed_files.append(new_filename)
    return renamed_files, original_filenames

def revert_filenames(renamed_files, original_filenames):
    for renamed_file, original_filename in zip(renamed_files, original_filenames):
        try:
            # Revert back to original filename
            os.rename(renamed_file, os.path.join(os.path.dirname(renamed_file), original_filename))
        except Exception as e:
            print("Error renaming file back to original name:", e)

def browse_files():
    try:
        # Open file dialog to select .doc files
        files = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.doc")])
        if files:
            # Rename files before conversion
            renamed_files, original_filenames = rename_files(files)
            for file in renamed_files:
                # Convert each selected file
                docx_file = convert_doc_to_docx(file)
                if docx_file:
                    print("Conversion complete. Output file:", docx_file)
            # Revert filenames back to original
            revert_filenames(renamed_files, original_filenames)
    except Exception as e:
        print("Error:", e)

# Create the GUI
root = tk.Tk()
root.title("File Converter")

# Customizing the UI
root.geometry("720x480")  # Setting initial window size

# Create a frame for better organization
frame = tk.Frame(root)
frame.pack(pady=100)


# Button to browse and select files
browse_button = tk.Button(root, text="Browse Files", command=browse_files, width=12, height=1, bg="darkblue", fg="white")
browse_button.pack()

root.mainloop()
