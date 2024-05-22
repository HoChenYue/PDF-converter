import os
from tkinter import Tk, Label, filedialog, Text, END
from tkinter import ttk
from tkinter import font
from ttkthemes import ThemedTk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
from pdf2docx import Converter

class ImageToPDFConverter:
    def __init__(self, parent, pdf_location_text, info_label):
        self.parent = parent
        self.pdf_location_text = pdf_location_text
        self.info_label = info_label
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill="both", expand=True)

        subtitle = ttk.Label(self.frame, text="PDF Converter", font=('Verdana', 12, 'bold'))
        subtitle.pack(pady=5)

        self.file_paths = []
        self_font = font.Font(family="Verdana", size=12, weight="bold")
        self.select_button = ttk.Button(self.frame, text="Select Images", command=self.select_images)
        self.select_button.pack(pady=10)

        self.convert_button = ttk.Button(self.frame, text="Convert to PDF", command=self.convert_to_pdf)
        self.convert_button.pack(pady=10)

    def select_images(self):
        self.file_paths = filedialog.askopenfilenames(title="Select Image Files", filetypes=[("Image files", "*.jpg *.jpeg *.png *.gif")])
        self.info_label.config(text=f"{len(self.file_paths)} image(s) selected.")

    def convert_to_pdf(self):
        if not self.file_paths:
            self.info_label.config(text="No images selected.")
            return

        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_pdf:
            c = canvas.Canvas(output_pdf, pagesize=letter)
            page_width, page_height = letter
            margin = 50

            for image_path in self.file_paths:
                img = Image.open(image_path)
                width, height = img.size
                aspect_ratio = width / height

                max_width = page_width - 2 * margin
                max_height = page_height - 2 * margin

                if width > height:
                    width = min(max_width, width)
                    height = width / aspect_ratio
                else:
                    height = min(max_height, height)
                    width = height * aspect_ratio

                x = (page_width - width) / 2
                y = (page_height - height) / 2

                c.drawImage(image_path, x, y, width=width, height=height)
                c.showPage()

            c.save()
            self.info_label.config(text="PDF file created successfully.")
            self.pdf_location_text.config(state='normal')
            self.pdf_location_text.delete(1.0, END)
            self.pdf_location_text.insert(END, output_pdf)
            self.pdf_location_text.config(state='disabled')
            os.startfile(os.path.dirname(output_pdf))

class PDFToWordExcelConverter:
    def __init__(self, parent, pdf_location_text, info_label):
        self.parent = parent
        self.pdf_location_text = pdf_location_text
        self.info_label = info_label
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill="both", expand=True)

        subtitle = ttk.Label(self.frame, text="From PDF to Word/Excel", font=('Verdana', 12, 'bold'))
        subtitle.pack(pady=5)

        self.select_pdf_button = ttk.Button(self.frame, text="Select PDF", command=self.select_pdf)
        self.select_pdf_button.pack(pady=10)

        self.convert_to_word_button = ttk.Button(self.frame, text="Convert to Word", command=self.convert_to_word)
        self.convert_to_word_button.pack(pady=10)

        self.convert_to_excel_button = ttk.Button(self.frame, text="Convert to Excel", command=self.convert_to_excel)
        self.convert_to_excel_button.pack(pady=10)

        self.pdf_path = ""

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF files", "*.pdf")])
        self.info_label.config(text=f"Selected PDF: {self.pdf_path}")

    def convert_to_word(self):
        if not self.pdf_path:
            self.info_label.config(text="No PDF selected.")
            return

        output_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if output_docx:
            cv = Converter(self.pdf_path)
            cv.convert(output_docx)
            cv.close()
            self.info_label.config(text=f"PDF converted to Word: {output_docx}")
            self.pdf_location_text.config(state='normal')
            self.pdf_location_text.delete(1.0, END)
            self.pdf_location_text.insert(END, output_docx)
            self.pdf_location_text.config(state='disabled')
            os.startfile(os.path.dirname(output_docx))

    def convert_to_excel(self):
        if not self.pdf_path:
            self.info_label.config(text="No PDF selected.")
            return

        output_excel = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_excel:
            # Placeholder for actual PDF to Excel conversion logic
            self.info_label.config(text="PDF to Excel conversion is not yet implemented.")
            self.pdf_location_text.config(state='normal')
            self.pdf_location_text.delete(1.0, END)
            self.pdf_location_text.insert(END, output_excel)
            self.pdf_location_text.config(state='disabled')
            os.startfile(os.path.dirname(output_excel))

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Conversion Tool")
        self.root.geometry("800x350")

        self.info_label = ttk.Label(root, text="")
        self.info_label.pack(pady=10)

        # Shared text box for displaying PDF file path
        shared_font = font.Font(family="Verdana", size=12, weight="bold")
        self.pdf_location_text = Text(root, height=2, wrap='word', state='disabled', font=shared_font)
        self.pdf_location_text.pack(pady=10)

        self.copy_button = ttk.Button(root, text="Copy File Path", command=self.copy_to_clipboard)
        self.copy_button.pack(pady=10)

        # Use PanedWindow to split the main window into two sections
        paned_window = ttk.PanedWindow(root, orient="horizontal")
        paned_window.pack(fill="both", expand=True)

        # Create left and right sections
        left_frame = ttk.Frame(paned_window)
        right_frame = ttk.Frame(paned_window)

        paned_window.add(left_frame, weight=1)
        paned_window.add(right_frame, weight=1)

        # Initialize the converters
        self.app_left = ImageToPDFConverter(left_frame, self.pdf_location_text, self.info_label)
        self.app_right = PDFToWordExcelConverter(right_frame, self.pdf_location_text, self.info_label)

    def copy_to_clipboard(self):
        self.root.clipboard_clear()
        text = self.pdf_location_text.get("1.0", END).strip()
        self.root.clipboard_append(text)
        self.info_label.config(text="File path copied to clipboard!")

# Create the main window with a theme
root = ThemedTk(theme="yaru")
app = Application(root)
root.mainloop()
