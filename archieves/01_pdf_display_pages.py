import os
import re
import sys
import tkinter as tk
import warnings
from multiprocessing import Process
from tkinter import messagebox

import fitz
from PyPDF2 import PdfReader
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QVBoxLayout, QWidget, QPushButton, QScrollArea

warnings.filterwarnings("ignore")


def find_pdf_in_folder(folder_path, bsn_input):
    """
    Finds a PDF file in a folder that matches a BSN number.

    Args:
        folder_path (str): Path to the folder containing PDF files.
        bsn_input (str): Input string containing the BSN number (e.g., "BSN1234567" or "1234567").

    Returns:
        str: Path to the matching PDF file, or None if no match is found.
    """

    folder_path = os.path.join(folder_path, "Database")

    # Validate and extract the numeric code
    match = re.match(r'BSN(\d{7})', bsn_input, re.IGNORECASE)  # Match "BSNxxxxxxx"

    if match:
        bsn_number = match.group(1)
    elif re.match(r'\d{7}', bsn_input):  # Match "xxxxxxx"
        bsn_number = bsn_input
    else:
        print("\nInvalid input. Please provide BSNxxxxxxx or xxxxxxx.")
        return

    # Search for matching PDF files
    for filename in os.listdir(folder_path):
        if bsn_number in filename and filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)

            # Count pages in the matching PDF
            reader = PdfReader(pdf_path)
            # total_pages = len(reader.pages)

            print(f"\nPDF Match found for the BSN Number ({bsn_input}): {filename}")
            # print(f"\nPDF path: {pdf_path}")
            return pdf_path

    # If no match is found
    print(f"\nNo matching PDF file found for code: {bsn_number}")


def find_initial_pages_to_skip(pdf_path):
    """
    Identifies the number of initial pages to skip in a PDF, stopping at the 'List of Tables' section.

    Args:
        pdf_path (str): The file path to the PDF document.

    Returns:
        int: The page number from which to start reading the main content,
             or 0 if 'List of Tables' is not found.
    """

    # Open the PDF
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    # Initialize variables
    skip_until_page = None

    # Regex patterns
    list_of_tables_pattern = re.compile(r'List of Tables', re.IGNORECASE)

    # Detect and skip the 'List of Tables' page
    for page_num in range(total_pages):
        page_text = reader.pages[page_num].extract_text()
        if list_of_tables_pattern.search(page_text):
            skip_until_page = page_num + 1  # Start search after this page
            break

    if skip_until_page is None:
        print("\nCould not find the 'List of Tables' page. Searching from the beginning.")
        skip_until_page = 0  # Default to start from the first page if not found

    # # Debug: Confirm the starting page for the main content
    # print(f"\nSkipping initial pages: {skip_until_page + 1}")

    return (skip_until_page + 1)


def get_toc_page_num(pdf_path):
    """
    Identifies the page number where the 'List of Tables and Figures' appears in a PDF.

    Args:
        pdf_path (str): The file path to the PDF document.

    Returns:
        int or None: The page number where 'List of Tables and Figures' is found,
                     or None if not found.
    """

    # Open the PDF
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    toc_page_num = None

    # Regex patterns
    toc_pattern = re.compile(r'List of Tables', re.IGNORECASE)

    # Detect and skip the 'List of Tables' page
    for page_num in range(total_pages):
        page_text = reader.pages[page_num].extract_text()
        if toc_pattern.search(page_text):
            toc_page_num = page_num + 1  # Start search after this page
            break

    return toc_page_num


def extract_main_text(pdf_path, page_num, header_height=60, footer_height=60):
    """
    Extracts the main text from a page, excluding headers and footers.

    Args:
        pdf_path (str): Path to the PDF file.
        page_num (int): Page number to extract text from (1-based indexing).
        header_height (int): Height of the header area to exclude (in points).
        footer_height (int): Height of the footer area to exclude (in points).

    Returns:
        str: Extracted main text from the page.
    """
    with fitz.open(pdf_path) as pdf:
        page = pdf[page_num - 1]  # Convert to 0-based indexing
        page_height = page.rect.height
        page_width = page.rect.width

        # Define the rectangle for the main content
        main_content_rect = fitz.Rect(0, header_height, page_width, page_height - footer_height)

        # Extract text within the defined rectangle
        main_text = page.get_text("text", clip=main_content_rect)
        return main_text.strip()


def get_section_pages_from_toc(pdf_path, section_input, toc_page_num, skip_pages):
    """
    Identifies the start and end pages for a given section (main or subsection)
    from the Table of Contents (TOC).

    Args:
        pdf_path (str): Path to the PDF file.
        section_input (str): Section number to search for (e.g., "2", "2.10").
        toc_page_num (int): Page number of the Table of Contents (1-based indexing).

    Returns:
        tuple: Start and end page numbers for the section, or None if not found.
    """

    # Open the PDF
    reader = PdfReader(pdf_path)

    # Extract text from the TOC page (convert 1-based index to 0-based)
    toc_text = reader.pages[toc_page_num - 1].extract_text()

    # Normalize the extracted text (remove extra spaces)
    normalized_toc = "\n".join([" ".join(line.split()) for line in toc_text.splitlines()])

    # Split the normalized text into lines
    toc_lines = normalized_toc.splitlines()

    # Initialize variables
    start_page = None
    end_page = None
    temp_end_page = None

    # next_subsection_pattern = re.compile(r'^\d+\.\d+')
    next_subsection_pattern = re.compile(r'^\s*\d{1}(\.\d+)?\s+[A-Za-z]')

    # Determine if the input is a main section or subsection
    is_main_section = re.match(r'^\d+$', section_input)  # Matches single digits (e.g., "1", "2")

    # If main section, calculate the next main section
    next_section = str(int(section_input) + 1) if is_main_section else None

    # Iterate through the TOC lines to find the section
    for i, line in enumerate(toc_lines):
        # Match the current section's line (strict match for main or subsection)
        if re.match(rf'^{re.escape(section_input)}\s', line):
            # Extract the start page number from the line
            start_page_match = re.search(r'(\d+)$', line)
            if start_page_match:
                start_page = int(start_page_match.group(1))

            # Find the end page
            if is_main_section:
                # For main sections, find the next main section
                for j in range(i + 1, len(toc_lines)):
                    next_line = toc_lines[j]
                    if re.match(rf'^{re.escape(next_section)}\s', next_line):  # Match the next main section
                        next_section_match = re.search(r'(\d+)$', next_line)
                        if next_section_match:
                            end_page = (int(next_section_match.group(1)) - 1)
                        break
            else:
                # For subsections
                if i + 1 < len(toc_lines):
                    next_line_match = re.search(r'(\d+)$', toc_lines[i + 1])
                    if next_line_match:
                        temp_end_page = int(next_line_match.group(1))
                        # print(f"\nTemporary end page: {temp_end_page}")

                        # Extract the content of the next page
                        next_page_text = extract_main_text(pdf_path, temp_end_page + skip_pages, header_height=60,
                                                           footer_height=60)
                        # print(f"\nNext Page Text: {next_page_text[:200]}...")

                        # Validation of end page
                        if temp_end_page == len(reader.pages):
                            end_page = temp_end_page
                        elif start_page == temp_end_page:
                            end_page = temp_end_page
                        elif next_subsection_pattern.match(next_page_text):
                            end_page = temp_end_page - 1
                            # print("\nYes! Text matched")
                        else:
                            # print("\nNo! Text not matched")
                            end_page = temp_end_page
                    break
            break

    # print(f"\nDeciding end page: start_page={start_page}, temp_end_page={temp_end_page}, end_page={end_page}")

    # Handle cases where the next section or page is not found
    if start_page is None:
        raise ValueError(f"Section '{section_input}' not found in the PDF.")
    if end_page is None:
        end_page = len(reader.pages) - skip_pages

    # Apply skipped pages adjustment
    start_page += skip_pages
    end_page += skip_pages

    # Return the result
    return start_page, end_page


class PDFViewer(QMainWindow):
    def __init__(self, pdf_path, page_list):
        super().__init__()
        self.setWindowTitle("PDF Viewer")
        self.setGeometry(100, 100, 900, 700)

        # Main layout
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Scrollable area for the PDF pages
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_area.setWidget(scroll_content)
        self.layout.addWidget(scroll_area)

        # Load and display each page
        pdf_document = fitz.open(pdf_path)
        # print(f"Total Pages in PDF: {len(pdf_document)}")
        for page_num in page_list:
            if page_num < 1 or page_num > len(pdf_document):
                print(f"Page {page_num} is out of range. Skipping...")
                continue

            print(f"\nRendering page {page_num}...")
            page = pdf_document.load_page(page_num - 1)
            pix = page.get_pixmap(dpi=100)
            # print(f"Page {page_num} rendered. Dimensions: {pix.width}x{pix.height}")

            # Convert pixmap to QImage
            if pix.alpha:
                qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGBA8888)
            else:
                qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimage)

            # Add page as QLabel to the layout
            label = QLabel(self)
            label.setPixmap(pixmap)
            label.setAlignment(Qt.AlignCenter)
            scroll_layout.addWidget(label)

        scroll_content.setLayout(scroll_layout)

        # Add a close button
        close_button = QPushButton("Close Viewer", self)
        close_button.clicked.connect(self.close)
        self.layout.addWidget(close_button)


def run_pyqt_viewer(pdf_path, page_list):
    """Runs the PyQt5 viewer in a separate process."""
    app = QApplication(sys.argv)
    viewer = PDFViewer(pdf_path, page_list)
    viewer.show()
    sys.exit(app.exec_())


def display_pdf_pages(pdf_path, page_list):
    """Launch the PyQt5 viewer in a separate process."""
    viewer_process = Process(target=run_pyqt_viewer, args=(pdf_path, page_list))
    viewer_process.start()
    viewer_process.join()


def reset_fields():
    """Clears all input fields."""
    bsn_entry.delete(0, tk.END)
    section_entry.delete(0, tk.END)


def close_app():
    """Closes the application."""
    root.destroy()


def process_inputs():

    bsn_input = bsn_entry.get().strip()
    section_input = section_entry.get().strip()

    if not bsn_input or not section_input:
        messagebox.showerror("Error", "Please provide both BSN and section input.")
        return

    # Get the folder where the script is located
    folder_path = os.path.dirname(os.path.abspath(__file__))

    pdf_path = find_pdf_in_folder(folder_path, bsn_input)
    if not pdf_path:
        return

    toc_page_num = get_toc_page_num(pdf_path)
    if toc_page_num is None:
        messagebox.showerror("Error", "Table of Contents not found.")
        return

    skip_pages = find_initial_pages_to_skip(pdf_path)
    try:
        section_start_page, section_end_page = get_section_pages_from_toc(pdf_path, section_input, toc_page_num,
                                                                          skip_pages)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
        return

    snapshot_pages = [section_start_page]
    if section_start_page != section_end_page:
        snapshot_pages = list(range(section_start_page, section_end_page + 1))

    display_pdf_pages(pdf_path, snapshot_pages)


if __name__ == "__main__":
    # GUI Setup
    root = tk.Tk()
    root.title("PDF Section Snapshot Tool")

    tk.Label(root, text="Enter BSN Number:").grid(row=0, column=0, padx=10, pady=10)
    bsn_entry = tk.Entry(root, width=30)
    bsn_entry.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(root, text="Enter Section Number:").grid(row=1, column=0, padx=10, pady=10)
    section_entry = tk.Entry(root, width=30)
    section_entry.grid(row=1, column=1, padx=10, pady=10)

    # Button Row
    button_frame = tk.Frame(root)
    button_frame.grid(row=2, column=0, columnspan=2, pady=10)

    apply_button = tk.Button(button_frame, text="Apply", command=process_inputs, width=10)
    apply_button.grid(row=0, column=0, padx=5)

    reset_button = tk.Button(button_frame, text="Reset", command=reset_fields, width=10)
    reset_button.grid(row=0, column=1, padx=5)

    cancel_button = tk.Button(button_frame, text="Cancel", command=close_app, width=10)
    cancel_button.grid(row=0, column=2, padx=5)

    root.mainloop()