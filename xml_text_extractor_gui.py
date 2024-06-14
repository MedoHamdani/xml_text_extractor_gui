import xml.etree.ElementTree as ET
import os
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to parse XML and extract text content
def parse_xml(file_path):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        ns = {'page': 'http://schema.primaresearch.org/PAGE/gts/pagecontent/2019-07-15'}

        text_contents = []
        for text_equiv in root.findall('.//page:TextEquiv', ns):
            unicode_tag = text_equiv.find('page:Unicode', ns)
            if unicode_tag is not None and unicode_tag.text:
                text_contents.append(unicode_tag.text.strip())

        return '\n'.join(text_contents)
    except Exception as e:
        print(f"Error parsing {file_path}: {e}")
        return ''  # Return empty string or handle the error as needed

# Function to convert text content to Word document
def text_to_word(text_content, output_path):
    document = Document()
    for paragraph in text_content.split('\n'):
        document.add_paragraph(paragraph, style='Normal')
    document.save(output_path)

# Function to handle file selection and conversion
def convert_files(input_dir, output_dir):
    for root_dir, _, files in os.walk(input_dir):
        for file_name in files:
            if file_name.endswith('.xml'):
                file_path = os.path.join(root_dir, file_name)
                text_content = parse_xml(file_path)
                if text_content:
                    output_path = os.path.join(output_dir, f"{os.path.splitext(file_name)[0]}.docx")
                    text_to_word(text_content, output_path)
    messagebox.showinfo("Conversion Complete", "All files have been converted successfully!")

# Function to handle directory selection for conversion
def convert_directory():
    input_dir = filedialog.askdirectory(title="Select Directory with XML Files")
    if not input_dir:
        return

    output_dir = filedialog.askdirectory(title="Select Output Directory")
    if not output_dir:
        return

    convert_files(input_dir, output_dir)

# Setting up the GUI
root = tk.Tk()
root.title("XML to Word Converter")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="Click the button below to convert XML files to Word documents.")
label.pack(pady=10)

convert_button = tk.Button(frame, text="Convert XML to Word", command=convert_directory)
convert_button.pack(pady=10)

root.mainloop()
