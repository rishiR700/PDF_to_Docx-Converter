import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import PyPDF2
from docx import Document


def convert_pdf_to_docx(pdf_file, docx_file):
    try:
        
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            document = Document()

            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                if text:
                    document.add_paragraph(text)

            document.save(docx_file)
        messagebox.showinfo("Success", "PDF successfully converted to DOCX!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def open_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        save_docx_file(file_path)


def save_docx_file(pdf_file):
    save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if save_path:
        convert_pdf_to_docx(pdf_file, save_path)


root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("400x200")

convert_button = tk.Button(root, text="Convert PDF to DOCX", command=open_pdf_file, width=20, height=2)
convert_button.pack(pady=50)

root.mainloop()
