import os
import fitz  # PyMuPDF
from docx import Document
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askopenfilenames, askdirectory

def pdf_to_word(pdf_path):
    """
    Convert a PDF file to a Word document.
    """
    doc = Document()
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text()
        doc.add_paragraph(text)
    
    output_word_path = os.path.splitext(pdf_path)[0] + '.docx'
    doc.save(output_word_path)
    print(f"PDF converted to Word and saved as {output_word_path}")

def word_to_pdf(word_path):
    """
    Convert a Word document to a PDF file, preserving colors and formatting.
    """
    # Load the Word document
    doc = Document(word_path)
    
    # Create a new PDF document
    pdf_document = fitz.open()
    
    # Define page dimensions (A4 size)
    page_width = 595  # A4 width in points (1 point = 1/72 inch)
    page_height = 842  # A4 height in points
    margin = 50  # Margin from the top and left
    
    # Initialize variables for text placement
    y_position = margin
    current_page = pdf_document.new_page(width=page_width, height=page_height)
    
    # Process each paragraph in the Word document
    for para in doc.paragraphs:
        for run in para.runs:
            # Extract text and formatting from the run
            text = run.text.strip()
            if not text:
                continue
            
            # Get font color (default to black if not specified)
            if run.font.color.rgb:
                color = (
                    run.font.color.rgb[0] / 255,
                    run.font.color.rgb[1] / 255,
                    run.font.color.rgb[2] / 255,
                )
            else:
                color = (0, 0, 0)  # Black
            
            # Get font size (default to 12 if not specified)
            font_size = run.font.size.pt if run.font.size else 12
            
            # Split text into lines based on page width and font size
            approx_char_width = 0.6 * font_size  # Approximate width per character
            max_line_width = page_width - 2 * margin
            max_chars_per_line = int(max_line_width / approx_char_width)
            
            # Split text into words and build lines
            words = text.split()
            lines = []
            current_line = []
            current_length = 0
            
            for word in words:
                word_length = len(word) * approx_char_width
                if current_length + word_length > max_line_width:
                    lines.append(" ".join(current_line))
                    current_line = [word]
                    current_length = word_length
                else:
                    current_line.append(word)
                    current_length += word_length + approx_char_width  # Add space
            
            if current_line:
                lines.append(" ".join(current_line))
            
            # Insert each line into the PDF
            for line in lines:
                line_height = font_size * 1.2  # Line spacing
                
                # Check if a new page is needed
                if y_position + line_height > page_height - margin:
                    current_page = pdf_document.new_page(width=page_width, height=page_height)
                    y_position = margin
                
                # Insert text with formatting
                current_page.insert_text(
                    (margin, y_position),
                    line,
                    fontsize=font_size,
                    color=color,
                    fontname="helv"  # Use Helvetica font (default)
                )
                
                # Update y_position for the next line
                y_position += line_height
    
    # Save the PDF document
    output_pdf_path = os.path.splitext(word_path)[0] + ".pdf"
    pdf_document.save(output_pdf_path)
    print(f"Word document converted to PDF and saved as {output_pdf_path}")
def merge_ppt(ppt_files):
    """
    Merge multiple PPT files into a single PPT file.
    """
    merged_presentation = Presentation()
    
    for ppt_file in ppt_files:
        presentation = Presentation(ppt_file)
        for slide in presentation.slides:
            merged_slide = merged_presentation.slides.add_slide(merged_presentation.slide_layouts[slide.slide_layout.slide_master.index])
            for shape in slide.shapes:
                new_shape = merged_slide.shapes.add_shape(shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height)
                new_shape.text = shape.text
    
    output_ppt_path = 'merged_presentation.pptx'
    merged_presentation.save(output_ppt_path)
    print(f"PPT files merged and saved as {output_ppt_path}")

def ppt_to_jpg(ppt_path):
    """
    Convert each slide of a PPT file to JPG images.
    """
    output_folder = os.path.join(os.getcwd(), 'ppt_images')
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    presentation = Presentation(ppt_path)
    
    for i, slide in enumerate(presentation.slides):
        slide_image_path = os.path.join(output_folder, f"slide_{i + 1}.jpg")
        slide_image = slide.export()
        Image.open(slide_image).save(slide_image_path)
        print(f"Slide {i + 1} saved as {slide_image_path}")

def main():
    root = Tk()
    root.withdraw()  # Hide the root window
    root.attributes('-topmost', True)  # Bring the dialog to the front
    
    print("Choose an option:")
    print("1. Convert PDF to Word")
    print("2. Convert Word to PDF")
    print("3. Merge PPT files")
    print("4. Convert PPT to JPG")
    
    choice = input("Enter your choice (1/2/3/4): ")
    
    if choice == "1":
        pdf_path = askopenfilename(title="Select PDF file", filetypes=[("PDF files", "*.pdf")])
        if pdf_path:
            pdf_to_word(pdf_path)
    
    elif choice == "2":
        word_path = askopenfilename(title="Select Word document", filetypes=[("Word files", "*.docx")])
        if word_path:
            word_to_pdf(word_path)
    
    elif choice == "3":
        ppt_files = askopenfilenames(title="Select PPT files to merge", filetypes=[("PPT files", "*.pptx")])
        if ppt_files:
            merge_ppt(ppt_files)
    
    elif choice == "4":
        ppt_path = askopenfilename(title="Select PPT file", filetypes=[("PPT files", "*.pptx")])
        if ppt_path:
            ppt_to_jpg(ppt_path)
    
    else:
        print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()