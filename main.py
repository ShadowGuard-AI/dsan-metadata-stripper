import argparse
import os
import logging
import re
import sys
import zipfile
import shutil
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
    Removes metadata from a PDF file. Returns True on success, False on error.
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
        return True
    except Exception as e:
        logging.error(f"Error processing PDF file {file_path}: {e}")
        return False

def _clean_docx_package(path):
    """Remove custom properties, comments, and all core property values
    from a saved DOCX package at the zip level.

    Core property elements are removed from docProps/core.xml so no
    fixture value and no placeholder text survives. The custom
    properties part and the comments part are dropped together with
    their relationship and content type entries, and comment
    references are removed from the document body.
    """
    tmp_path = str(path) + ".tmpclean"
    with zipfile.ZipFile(path, "r") as zin:
        items = {name: zin.read(name) for name in zin.namelist()}

    # Clear every core property element, keeping only the root element.
    if "docProps/core.xml" in items:
        core = items["docProps/core.xml"].decode("utf-8")
        core = re.sub(r"(<cp:coreProperties\b[^>]*>).*?(</cp:coreProperties>)",
                      r"\1\2", core, flags=re.DOTALL)
        items["docProps/core.xml"] = core.encode("utf-8")

    # Remove the custom properties part and its references.
    items.pop("docProps/custom.xml", None)
    if "_rels/.rels" in items:
        rels = items["_rels/.rels"].decode("utf-8")
        rels = re.sub(r"<Relationship\b[^>]*custom-properties[^>]*/>", "", rels)
        rels = re.sub(r"<Relationship\b[^>]*Target=\"docProps/custom\.xml\"[^>]*/>", "", rels)
        items["_rels/.rels"] = rels.encode("utf-8")

    # Remove the comments part and its references.
    items.pop("word/comments.xml", None)
    if "word/_rels/document.xml.rels" in items:
        rels = items["word/_rels/document.xml.rels"].decode("utf-8")
        rels = re.sub(r"<Relationship\b[^>]*relationships/comments[^>]*/>", "", rels)
        rels = re.sub(r"<Relationship\b[^>]*Target=\"comments\.xml\"[^>]*/>", "", rels)
        items["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    if "word/document.xml" in items:
        doc_xml = items["word/document.xml"].decode("utf-8")
        doc_xml = re.sub(r"<w:commentRangeStart\b[^>]*/>", "", doc_xml)
        doc_xml = re.sub(r"<w:commentRangeEnd\b[^>]*/>", "", doc_xml)
        doc_xml = re.sub(r"<w:r\b[^>]*>\s*<w:commentReference\b[^>]*/>\s*</w:r>", "", doc_xml)
        doc_xml = re.sub(r"<w:commentReference\b[^>]*/>", "", doc_xml)
        items["word/document.xml"] = doc_xml.encode("utf-8")

    if "[Content_Types].xml" in items:
        ct = items["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r"<Override\b[^>]*PartName=\"/docProps/custom\.xml\"[^>]*/>", "", ct)
        ct = re.sub(r"<Override\b[^>]*PartName=\"/word/comments\.xml\"[^>]*/>", "", ct)
        items["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    shutil.move(tmp_path, path)

def strip_docx_metadata(file_path, output_path):
    """
    Removes metadata from a DOCX file. Returns True on success, False on error.
    """
    try:
        doc = Document(file_path)

        # Clear core properties with empty values. The package level
        # cleanup below then removes the elements themselves, along
        # with the fields python-docx does not expose for clearing
        # (category, created, modified, revision, identifier, version).
        props = doc.core_properties
        for field in ("author", "title", "subject", "keywords", "comments",
                      "last_modified_by", "category", "identifier", "version"):
            try:
                setattr(props, field, "")
            except Exception:
                pass

        doc.save(output_path)
        _clean_docx_package(output_path)
        logging.info(f"Sanitized DOCX: {file_path} -> {output_path}")
        return True
    except Exception as e:
        logging.error(f"Error processing DOCX file {file_path}: {e}")
        return False

def strip_image_metadata(file_path, output_path):
    """
    Removes metadata from an image file. Returns True on success, False on error.
    """
    try:
        image = Image.open(file_path)
        data = list(image.getdata())
        image_without_metadata = Image.new(image.mode, image.size)
        image_without_metadata.putdata(data)
        image_without_metadata.save(output_path)

        logging.info(f"Sanitized Image: {file_path} -> {output_path}")
        return True
    except Exception as e:
        logging.error(f"Error processing image file {file_path}: {e}")
        return False

SUPPORTED = {".pdf", ".docx", ".jpg", ".jpeg", ".png"}

def process_file(file_path, output_path):
    """
    Processes a single file to remove metadata.
    Returns one of "ok", "unsupported", or "error".
    """
    file_extension = file_path.suffix.lower()

    if file_extension == ".pdf":
        return "ok" if strip_pdf_metadata(file_path, output_path) else "error"
    elif file_extension == ".docx":
        return "ok" if strip_docx_metadata(file_path, output_path) else "error"
    elif file_extension in [".jpg", ".jpeg", ".png"]:
        return "ok" if strip_image_metadata(file_path, output_path) else "error"
    else:
        logging.warning(f"Unsupported file type: {file_path}")
        return "unsupported"

def process_directory(input_dir, output_dir, recursive):
    """
    Processes a directory, mirroring the subdirectory structure in the
    output directory so same named files from different subdirectories
    never overwrite each other.
    Returns (processed, failed) counts.
    """
    processed = 0
    failed = 0
    for root, _, files in os.walk(input_dir):
        for file in files:
            file_path = Path(root) / file
            rel = file_path.relative_to(input_dir)
            output_path = output_dir / rel
            output_path.parent.mkdir(parents=True, exist_ok=True)
            result = process_file(file_path, output_path)
            if result == "ok":
                processed += 1
            else:
                failed += 1

        if not recursive:
            break
    return processed, failed

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

    if not input_path.exists():
        logging.error(f"Input path does not exist: {input_path}")
        sys.exit(1)

    if input_path.is_file():
        output_path = output_dir / input_path.name
        if output_path.resolve() == input_path.resolve():
            logging.error("Refusing to run: output path resolves to the input file.")
            sys.exit(1)
        output_dir.mkdir(parents=True, exist_ok=True)
        result = process_file(input_path, output_path)
        if result != "ok":
            sys.exit(1)
    elif input_path.is_dir():
        if output_dir.resolve() == input_path.resolve():
            logging.error("Refusing to run: output directory resolves to the input directory.")
            sys.exit(1)
        output_dir.mkdir(parents=True, exist_ok=True)
        processed, failed = process_directory(input_path, output_dir, args.recursive)
        if processed == 0 or failed > 0:
            if processed == 0:
                logging.error("No files were processed.")
            sys.exit(1)
    else:
        logging.error(f"Input path does not exist: {input_path}")
        sys.exit(1)

if __name__ == "__main__":
    main()
