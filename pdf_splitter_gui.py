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

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

POPPLER_BIN = os.path.join(BASE_DIR, "poppler", "bin")
GHOSTSCRIPT_EXE = os.path.join(BASE_DIR, "ghostscript", "bin", "gswin64c.exe")
# Use POPPLER_BIN for pdf2image, and GHOSTSCRIPT_EXE for Ghostscript calls

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

def get_ghostscript_cmd():
    return GHOSTSCRIPT_EXE

def compress_pdf_ghostscript(input_path, output_path):
    """
    Compresses a PDF using Ghostscript.
    Args:
        input_path (str): Path to the input PDF file.
        output_path (str): Path to save the compressed PDF file.
    """
    input_size_kb = os.path.getsize(input_path) / 1024
    if input_size_kb <= 500:
        # If already under 500KB, just copy
        import shutil
        shutil.copy2(input_path, output_path)
        logging.info(f"Skipped compression for {output_path} (already {input_size_kb:.1f}KB)")
        return

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

        # Apple-like color palette
        BG_COLOR = "#f5f5f7"
        FG_COLOR = "#1d1d1f"
        ACCENT_COLOR = "#0071e3"
        ENTRY_BG = "#fff"
        BTN_BG = "#e0e0e5"
        BTN_ACTIVE_BG = "#0071e3"
        BTN_ACTIVE_FG = "#fff"

        # Set window size: fit PDF width, full height
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        # Try to get the first page width (in pixels at 100 dpi)
        try:
            test_img = convert_from_path(self.file_path, dpi=100, poppler_path=POPPLER_BIN, first_page=1, last_page=1)[0]
            pdf_width, pdf_height = test_img.size
        except Exception:
            pdf_width, pdf_height = 800, 1000
        win_width = min(pdf_width + 120, screen_width - 40)
        # Subtract taskbar height (approx 60px) from screen height
        taskbar_height = 80
        win_height = min(int(screen_height * 0.90), screen_height - taskbar_height)
        x = (screen_width - win_width) // 2
        y = (screen_height - win_height) // 2
        root.geometry(f"{win_width}x{win_height}+{x}+{y}")
        root.configure(bg=BG_COLOR)

        # Prompt user to select a PDF file
        self.file_path = filedialog.askopenfilename(
            title="Select a PDF file", filetypes=[("PDF files", "*.pdf")])
        if not self.file_path:
            messagebox.showerror("No File", "No file was selected. Exiting.")
            root.destroy()
            return

        self.base_name = os.path.splitext(os.path.basename(self.file_path))[0]

        # Ensure poppler bin is in PATH for DLL loading
        os.environ["PATH"] = POPPLER_BIN + os.pathsep + os.environ.get("PATH", "")
        # Convert PDF pages to images for preview
        try:
            self.pages_images = convert_from_path(
                self.file_path,
                dpi=100,
                poppler_path=POPPLER_BIN
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF pages: {e}")
            logging.error(f"Failed to load PDF pages: {e}")
            root.destroy()
            return

        self.total_pages = len(self.pages_images)
        self.current_index = 0
        self.filename_inputs = [None] * self.total_pages

        # --- Scrollable main frame ---
        container = tk.Frame(root, bg=BG_COLOR)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=BG_COLOR, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Centering frame inside canvas (vertical and horizontal)
        self.centering_frame = tk.Frame(canvas, bg=BG_COLOR)
        self.centering_frame.pack(expand=True, fill="both")

        self.scrollable_frame = tk.Frame(self.centering_frame, bg=BG_COLOR)
        self.scrollable_frame.pack(expand=True, anchor="center")

        window_id = canvas.create_window((0, 0), window=self.centering_frame, anchor="center")
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_configure(event):
            # Center the centering_frame horizontally and vertically in the canvas
            canvas_width = event.width
            canvas_height = event.height
            self.centering_frame.update_idletasks()
            frame_width = self.centering_frame.winfo_reqwidth()
            frame_height = self.centering_frame.winfo_reqheight()
            x = max((canvas_width - frame_width) // 2, 0)
            y = max((canvas_height - frame_height) // 2, 0)
            canvas.coords(window_id, x, y)

        canvas.bind('<Configure>', on_configure)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Enable mouse wheel scrolling (Windows, Mac, Linux)
        def _on_mousewheel(event):
            if event.num == 5 or event.delta == -120:
                canvas.yview_scroll(1, "units")
            elif event.num == 4 or event.delta == 120:
                canvas.yview_scroll(-1, "units")

        # Windows and MacOS
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        # Linux (event.num 4/5)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        # --- Image preview ---
        self.image_label = tk.Label(self.scrollable_frame, bg=BG_COLOR)
        self.image_label.pack(padx=20, pady=(30, 10), anchor="center")

        # --- Filename input ---
        self.filename_label = tk.Label(self.scrollable_frame, text="Enter file name:",
                                       font=("San Francisco", 12, "bold"), bg=BG_COLOR, fg=FG_COLOR)
        self.filename_label.pack(pady=(10, 2), anchor="center")

        self.filename_entry = tk.Entry(self.scrollable_frame, font=("San Francisco", 12),
                                       bg=ENTRY_BG, fg=FG_COLOR, relief="flat", highlightthickness=1, highlightbackground="#ccc")
        self.filename_entry.pack(ipady=6, ipadx=4, pady=(0, 20), fill=None, padx=60, anchor="center")
        self.filename_entry.bind("<Return>", self.handle_enter_key)

        # Always focus the filename entry after widgets are created
        self.root.after(100, self.filename_entry.focus_set)

        # --- Navigation buttons ---
        btn_frame = tk.Frame(self.scrollable_frame, bg=BG_COLOR)
        btn_frame.pack(pady=10, anchor="center")

        self.prev_btn = tk.Button(btn_frame, text="Previous", command=self.prev_page,
                                  font=("San Francisco", 12), bg=BTN_BG, fg=FG_COLOR, relief="flat",
                                  activebackground=BTN_ACTIVE_BG, activeforeground=BTN_ACTIVE_FG, width=10)
        self.prev_btn.grid(row=0, column=0, padx=8)

        self.next_btn = tk.Button(btn_frame, text="Next", command=self.next_page,
                                  font=("San Francisco", 12), bg=BTN_BG, fg=FG_COLOR, relief="flat",
                                  activebackground=BTN_ACTIVE_BG, activeforeground=BTN_ACTIVE_FG, width=10)
        self.next_btn.grid(row=0, column=1, padx=8)

        self.finish_btn = tk.Button(btn_frame, text="Finish", command=self.finish, state="disabled",
                                    font=("San Francisco", 12, "bold"), bg=ACCENT_COLOR, fg="#fff", relief="flat",
                                    activebackground="#005bb5", activeforeground="#fff", width=10)
        self.finish_btn.grid(row=0, column=2, padx=8)

        # --- Footer ---
        footer = tk.Label(self.scrollable_frame, text="PDF Splitter", font=("San Francisco", 11),
                         bg=BG_COLOR, fg="#888", pady=20)
        footer.pack(side="bottom", fill="x", anchor="center")

        self.update_page()
    
    def update_page(self):
        """
        Update the preview image and filename entry for the current page.
        """
        # Dynamically scale image to fit window width, max 90% of window height
        win_width = self.root.winfo_width() or 800
        win_height = self.root.winfo_height() or 600
        max_img_width = int(win_width * 0.85)
        max_img_height = int(win_height * 0.9)
        img = self.pages_images[self.current_index].copy()
        img.thumbnail((max_img_width, max_img_height))
        self.photo_img = ImageTk.PhotoImage(img)
        self.image_label.config(image=self.photo_img, bg="#f5f5f7")
        self.filename_label.config(text="Enter file name for this page:")

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
        # Use the directory where the input PDF file is located
        output_base = os.path.dirname(self.file_path)
        output_folder = os.path.join(output_base, self.base_name)
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
