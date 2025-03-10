from pdf2docx import Converter
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to handle file conversion
def convert_pdf_to_word():
    # Get the input PDF file path
    pdf_file = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not pdf_file:
        messagebox.showerror("Error", "No PDF file selected!")
        return

    # Get the output Word file path
    docx_file = filedialog.asksaveasfilename(
        title="Select where you want to save as Word document",
        defaultextension=".docx",
        filetypes=[("Word Files", "*.docx")]
    )
    if not docx_file:
        messagebox.showerror("Error", "No output file specified!")
        return

    try:
        # Create a Converter object
        cv = Converter(pdf_file)

        # Convert the PDF to a Word document
        cv.convert(docx_file, start=0, end=None)

        # Close the Converter
        cv.close()

        # Show success message
        messagebox.showinfo("Awesome!", f"Your PDF was saved as a word document at this path: {docx_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred, please make sure you are saving in a valid path.: {e}")

# Create the main application window
root = tk.Tk()
root.title("ScrewYouAdobe")

# Set window size
root.geometry("400x200")

# Add a label
label = tk.Label(root, text="Convert your PDF to Word", font=("Arial", 16))
label.pack(pady=20)

# Add a button to start the conversion process
convert_button = tk.Button(root, text="Select the file", command=convert_pdf_to_word)
convert_button.pack(pady=20)

# Run the application
root.mainloop()