import tkinter as tk
import tkinter.font as font
from tkinter import filedialog
from pdf_to_word import PDFtoWordConverter
from tkinter import ttk
import subprocess


class PDFtoWordGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Docx")
        self.root.configure(bg="#FFFFFF")  # Set background color for the root window

        style = ttk.Style(root)
        style.theme_use('clam')  # Set the ttk theme to 'clam' (or any other theme of your choice)

        self.canvas = tk.Canvas(self.root, height=500, width=700, bg="#FFFFFF")  # Set background color for the canvas
        self.canvas.pack()

        self.frame = tk.Frame(self.root, bg="#FFFFFF")
        self.frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

        button_font = font.Font(size=30)
        self.button = tk.Button(
            self.frame,
            text="OPEN PDF FILE",
            padx=20,
            pady=10,
            bg="#Aa0808",
            fg="#fff",
            command=self.add_file,
            font=button_font,
        )
        self.button.pack()

    def add_file(self):
        if file_path := filedialog.askopenfilename(
            initialdir="/", title="Select file"
        ):
            self.convert_file(file_path)

    def convert_file(self, file_path):
        if save_path := filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Word Files", "*.docx")]
        ):
            converter = PDFtoWordConverter()
            converter.convert_pdf_to_word(file_path, save_path)
            self.display_success_message(save_path)

    def display_success_message(self, save_path):
        success_text = tk.Label(
            self.frame,
            text="PDF has been successfully converted",
            font=("Arial", 16),
            bg="#FFFFFF",
            fg="#000000",
        )
        success_text.pack()

        open_button = tk.Button(
            self.frame,
            text="OPEN CONVERTED WORD FILE",
            padx=20,
            pady=10,
            bg="#075896",
            fg="#fff",
            command=lambda: self.open_file(save_path),
            font=font.Font(size=12),
        )
        open_button.pack()

    def open_file(self, file_path):
        if file_path:
            try:
                subprocess.run(["start", "", file_path], shell=True)  # Open the file with the default associated application
            except Exception as e:
                print("Failed to open the file:", e)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoWordGUI(root)
    root.mainloop()
