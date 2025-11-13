#!/usr/bin/env python3
"""
DOCX to PDF Converter
Converts DOCX files to PDF while preserving original formatting.
"""

import os
import sys
from pathlib import Path
from typing import Optional
import argparse


def install_dependencies():
    """Install required packages using pip."""
    import subprocess
    packages = [
        "docx2pdf",
        "python-docx",
        "reportlab"
    ]

    print("Installing required packages...")
    for package in packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"✓ {package} installed successfully")
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {package}: {e}")
            return False
    return True


def check_dependencies():
    """Check if required packages are installed."""
    required_packages = ['docx2pdf']
    missing_packages = []

    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            missing_packages.append(package)

    if missing_packages:
        print(f"Missing packages: {', '.join(missing_packages)}")
        print("Please install them using: pip install " + " ".join(missing_packages))
        return False
    return True


def convert_docx_to_pdf(input_path: str, output_path: Optional[str] = None) -> bool:
    """
    Convert a DOCX file to PDF while preserving formatting.

    Args:
        input_path: Path to the input DOCX file
        output_path: Optional path for the output PDF file

    Returns:
        bool: True if conversion successful, False otherwise
    """
    try:
        from docx2pdf import convert

        input_file = Path(input_path)

        # Validate input file
        if not input_file.exists():
            print(f"Error: Input file '{input_path}' does not exist.")
            return False

        if not input_file.suffix.lower() == '.docx':
            print(f"Error: Input file '{input_path}' is not a DOCX file.")
            return False

        # Determine output path
        if output_path is None:
            output_path = input_file.with_suffix('.pdf')
        else:
            output_path = Path(output_path)
            if not output_path.suffix.lower() == '.pdf':
                output_path = output_path.with_suffix('.pdf')

        # Create output directory if it doesn't exist
        output_path.parent.mkdir(parents=True, exist_ok=True)

        print(f"Converting '{input_path}' to '{output_path}'...")

        # Convert using docx2pdf (preserves formatting)
        convert(str(input_file), str(output_path))

        if output_path.exists():
            print(f"✓ Conversion successful! PDF saved to: {output_path}")
            return True
        else:
            print("✗ Conversion failed - output file not created.")
            return False

    except ImportError:
        print("Error: docx2pdf package not found. Please install it using:")
        print("pip install docx2pdf")
        return False
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return False


def convert_multiple_files(input_dir: str, output_dir: Optional[str] = None) -> int:
    """
    Convert all DOCX files in a directory to PDF.

    Args:
        input_dir: Directory containing DOCX files
        output_dir: Optional directory for output PDF files

    Returns:
        int: Number of successfully converted files
    """
    input_path = Path(input_dir)

    if not input_path.exists():
        print(f"Error: Directory '{input_dir}' does not exist.")
        return 0

    if not input_path.is_dir():
        print(f"Error: '{input_dir}' is not a directory.")
        return 0

    # Set output directory
    if output_dir is None:
        output_path = input_path
    else:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

    # Find all DOCX files
    docx_files = list(input_path.glob("*.docx"))

    if not docx_files:
        print(f"No DOCX files found in '{input_dir}'.")
        return 0

    print(f"Found {len(docx_files)} DOCX files to convert...")

    successful_conversions = 0

    for docx_file in docx_files:
        pdf_file = output_path / docx_file.with_suffix('.pdf').name
        if convert_docx_to_pdf(str(docx_file), str(pdf_file)):
            successful_conversions += 1

    print(f"Successfully converted {successful_conversions}/{len(docx_files)} files.")
    return successful_conversions


def main():
    """Main function to handle command line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert DOCX files to PDF while preserving original formatting.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert single file
  python docx_to_pdf_converter.py document.docx

  # Convert single file with custom output path
  python docx_to_pdf_converter.py document.docx -o output.pdf

  # Convert all DOCX files in a directory
  python docx_to_pdf_converter.py -d ./documents

  # Convert all DOCX files to a specific output directory
  python docx_to_pdf_converter.py -d ./documents -o ./pdfs
        """
    )

    parser.add_argument("input", nargs="?", help="Input DOCX file or directory")
    parser.add_argument("-o", "--output", help="Output PDF file or directory")
    parser.add_argument("-d", "--directory", action="store_true",
                       help="Treat input as directory containing DOCX files")
    parser.add_argument("--install", action="store_true",
                       help="Install required dependencies")

    args = parser.parse_args()

    # Install dependencies if requested
    if args.install:
        if install_dependencies():
            print("✓ All dependencies installed successfully!")
        else:
            print("✗ Failed to install some dependencies.")
        return

    # Check if dependencies are available
    if not check_dependencies():
        return

    # Handle directory conversion
    if args.directory:
        if not args.input:
            print("Error: Please specify a directory when using -d flag.")
            return

        convert_multiple_files(args.input, args.output)

    # Handle single file conversion
    else:
        if not args.input:
            print("Error: Please specify an input DOCX file.")
            parser.print_help()
            return

        convert_docx_to_pdf(args.input, args.output)


if __name__ == "__main__":
    main()