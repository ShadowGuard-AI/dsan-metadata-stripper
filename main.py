```python
import argparse
import os
import logging
from pathlib import Path
from faker import Faker
import chardet
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
from docx import Document
from docx.oxml import OxmlElement

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize Faker for anonymization
faker = Faker()

def setup_argparse():
    """
    Sets up the command-line argument parser.
    """
    parser = argparse.ArgumentParser(
        description="A tool to remove metadata from files for data sanitization."
    )
    parser.add_argument(
        "input_path",
        type=str,
        help="Path to the file or directory to process."
    )
    parser.add_argument(
        "--output",
        type=str,
        default="sanitized_output",
        help="Path to the output directory for sanitized files. Defaults to 'sanitized_output'."
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively process files in directories."
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging."
    )
    parser.add_argument(
        "--version",
        action="version",
        version="dsan-metadata-stripper 1.0"
    )
    return parser

def strip_pdf_metadata(file_path, output_path):
    """
    Removes metadata from a PDF file.
    """
    try:
        reader = PdfReader(file_path)
        writer = PdfWriter()

        # Copy content without metadata
        for page in reader.pages:
            writer.add_page(page)

        # Remove metadata
        writer._info = None

        with open(output_path, "wb") as f:
            writer.write(f)

        logging.info(f"Sanitized PDF: {file_path} -> {output_path}")
    except Exception as e:
        logging.error(f"Error processing PDF file {file_path}: {e}")

def strip_docx_metadata(file_path, output_path):
    """
    Removes metadata from a DOCX file.
    """
    try:
        doc = Document(file_path)

        # Remove custom properties
        custom_props = doc.core_properties
        custom_props.author = None
        custom_props.title = None
        custom_props.subject = None
        custom_props.keywords = None
        custom_props.comments = None
        custom_props.last_modified_by = None

        # Remove comments
        for comment in doc.element.xpath("//w:comment"):
            parent = comment.getparent()
            parent.remove(comment)

        doc.save(output_path)
        logging.info(f"Sanitized DOCX: {file_path} -> {output_path}")
    except Exception as e:
        logging.error(f"Error processing DOCX file {file_path}: {e}")

def strip_image_metadata(file_path, output_path):
    """
    Removes metadata from an image file.
    """
    try:
        image = Image.open(file_path)
        data = list(image.getdata())
        image_without_metadata = Image.new(image.mode, image.size)
        image_without_metadata.putdata(data)
        image_without_metadata.save(output_path)

        logging.info(f"Sanitized Image: {file_path} -> {output_path}")
    except Exception as e:
        logging.error(f"Error processing image file {file_path}: {e}")

def process_file(file_path, output_dir):
    """
    Processes a single file to remove metadata.
    """
    file_extension = file_path.suffix.lower()
    output_path = output_dir / file_path.name

    if file_extension == ".pdf":
        strip_pdf_metadata(file_path, output_path)
    elif file_extension == ".docx":
        strip_docx_metadata(file_path, output_path)
    elif file_extension in [".jpg", ".jpeg", ".png"]:
        strip_image_metadata(file_path, output_path)
    else:
        logging.warning(f"Unsupported file type: {file_path}")

def process_directory(input_dir, output_dir, recursive):
    """
    Processes a directory to remove metadata from all supported files.
    """
    for root, _, files in os.walk(input_dir):
        for file in files:
            file_path = Path(root) / file
            process_file(file_path, output_dir)

        if not recursive:
            break

def main():
    """
    Main entry point for the script.
    """
    parser = setup_argparse()
    args = parser.parse_args()

    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    input_path = Path(args.input_path)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    if input_path.is_file():
        process_file(input_path, output_dir)
    elif input_path.is_dir():