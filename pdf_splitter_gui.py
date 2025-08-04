import os
import platform
import subprocess
import sys
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageTk
from pdf2image import convert_from_path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

def get_ghostscript_cmd():
    """
    Returns the Ghostscript command based on the operating system.
    """
    system = platform.system()
    return "gswin64c" if system == "Windows" else "gs"

def compress_pdf_ghostscript(input_path, output_path):
    """
    Compresses a PDF using Ghostscript.
    Args:
        input_path (str): Path to the input PDF file.
        output_path (str): Path to save the compressed PDF file.
    """
    gs_cmd = get_ghostscript_cmd()
    gs_command = [
        gs_cmd,
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        '-dPDFSETTINGS=/screen',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        f'-sOutputFile={output_path}',
        input_path
    ]

    try:
        subprocess.run(gs_command, check=True)
        final_size_kb = os.path.getsize(output_path) / 1024
        if final_size_kb > 500:
            logging.warning(f"{output_path} is still {final_size_kb:.1f}KB, above 500KB limit.")
        else:
            logging.info(f"Compressed {output_path} to {final_size_kb:.1f}KB with Ghostscript.")
    except Exception as e:
        logging.error(f"Ghostscript compression failed: {e}")

class PDFSplitterGUI:
    def __init__(self, root):
        """
        Initialize the PDF Splitter GUI.
        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.title("PDF Splitter Preview & Compression")

        # Prompt user to select a PDF file
        self.file_path = filedialog.askopenfilename(
            title="Select a PDF file", filetypes=[("PDF files", "*.pdf")])
        if not self.file_path:
            messagebox.showerror("No File", "No file was selected. Exiting.")
            root.destroy()
            return

        self.base_name = os.path.splitext(os.path.basename(self.file_path))[0]

        # Convert PDF pages to images for preview
        try:
            self.pages_images = convert_from_path(
                self.file_path,
                dpi=100,
                poppler_path=r'C:\Program Files\poppler-24.08.0\Library\bin'
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF pages: {e}")
            logging.error(f"Failed to load PDF pages: {e}")
            root.destroy()
            return

        self.total_pages = len(self.pages_images)
        self.current_index = 0
        self.filename_inputs = [None] * self.total_pages

        # Setup GUI widgets
        self.image_label = tk.Label(root)
        self.image_label.pack(padx=10, pady=10)

        self.filename_label = tk.Label(root, text="ENTER FILE NAME:")
        self.filename_label.pack()

        self.filename_entry = tk.Entry(root)
        self.filename_entry.pack()
        self.filename_entry.focus_set()
        self.filename_entry.bind("<Return>", self.handle_enter_key)

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        self.prev_btn = tk.Button(btn_frame, text="Previous", command=self.prev_page)
        self.prev_btn.grid(row=0, column=0, padx=5)

        self.next_btn = tk.Button(btn_frame, text="Next", command=self.next_page)
        self.next_btn.grid(row=0, column=1, padx=5)

        self.finish_btn = tk.Button(btn_frame, text="Finish", command=self.finish, state="disabled")
        self.finish_btn.grid(row=0, column=2, padx=5)

        self.update_page()
    
    def update_page(self):
        """
        Update the preview image and filename entry for the current page.
        """
        img = self.pages_images[self.current_index]
        img.thumbnail((600, 800))
        self.photo_img = ImageTk.PhotoImage(img)
        self.image_label.config(image=self.photo_img)
        self.filename_label.config(text="ENTER FILE NAME:")

        # Restore filename input if previously entered
        if self.filename_inputs[self.current_index] is not None:
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, self.filename_inputs[self.current_index])
        else:
            self.filename_entry.delete(0, tk.END)

        self.filename_entry.focus_set()

        # Update button states
        self.prev_btn.config(state="normal" if self.current_index > 0 else "disabled")
        if self.current_index == self.total_pages - 1:
            self.next_btn.config(state="disabled")
            self.finish_btn.config(state="normal")
        else:
            self.next_btn.config(state="normal")
            self.finish_btn.config(state="disabled")

    def save_current_input(self):
        """
        Save the filename input for the current page.
        Returns:
            bool: True if input is valid, False otherwise.
        """
        val = self.filename_entry.get().strip()
        if not val:
            messagebox.showwarning("Input Required", "Please enter a file name before proceeding.")
            logging.warning("Filename input required before proceeding.")
            return False
        self.filename_inputs[self.current_index] = val
        return True

    def handle_enter_key(self, event):
        """
        Handle Enter key event in filename entry.
        """
        if not self.save_current_input():
            return "break"
        if self.current_index == self.total_pages - 1:
            self.finish()
        else:
            self.next_page()
        return "break"

    def next_page(self):
        """
        Go to the next page in the preview.
        """
        if not self.save_current_input():
            return
        if self.current_index < self.total_pages - 1:
            self.current_index += 1
            self.update_page()

    def prev_page(self):
        """
        Go to the previous page in the preview.
        """
        if self.current_index > 0:
            self.current_index -= 1
            self.update_page()

    def finish(self):
        """
        Finish input and process the PDF splitting and compression.
        """
        if not self.save_current_input():
            return
        self.root.destroy()
        self.process_files()

    def process_files(self):
        """
        Split the PDF into separate files and compress each using Ghostscript.
        """
        reader = PdfReader(self.file_path)
        output_folder = os.path.join(os.path.expanduser("~"), "Documents", "PDFSplitterOut", self.base_name)
        os.makedirs(output_folder, exist_ok=True)

        # Map filenames to page indices
        files_dict = {}
        for page_idx, fname in enumerate(self.filename_inputs):
            files_dict.setdefault(fname, []).append(page_idx)

        for fname, pages in files_dict.items():
            writer = PdfWriter()
            for p in pages:
                if p < len(reader.pages):
                    writer.add_page(reader.pages[p])

            output_pdf_path = os.path.join(output_folder, f"{self.base_name}_{fname}.pdf")
            temp_path = output_pdf_path + ".tmp"

            # Write split PDF to temp file
            with open(temp_path, 'wb') as f_out:
                writer.write(f_out)

            # Compress PDF
            compress_pdf_ghostscript(temp_path, output_pdf_path)
            os.remove(temp_path)

            logging.info(f"Saved: {output_pdf_path}")

        self.prompt_rerun()

    def prompt_rerun(self):
        """
        Prompt user to run the tool again or exit.
        """
        answer = messagebox.askyesno("Run Again?", "Do you want to split/compress another PDF?")
        if answer:
            python_executable = sys.executable
            logging.info("Restarting application for another run.")
            os.execv(python_executable, [python_executable] + sys.argv)
        else:
            logging.info("Exiting application.")

def main():
    """
    Main entry point for the PDF Splitter GUI application.
    """
    root = tk.Tk()
    app = PDFSplitterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
