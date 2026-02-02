"""
Execution Version Creator

Merges signed pages (from DocuSign) back into the original document
to create the final execution version.

Handles DocuSign's PDF protection/locking by attempting to open
with various methods.
"""

import fitz  # PyMuPDF
import os
import tempfile
from typing import Callable, Dict, Any


def unlock_pdf(pdf_path: str) -> fitz.Document:
    """
    Open and unlock a DocuSign PDF.

    DocuSign PDFs are typically protected with permission restrictions
    (not password encryption). This function attempts to open them.

    Args:
        pdf_path: Path to the PDF file

    Returns:
        Opened fitz.Document

    Raises:
        ValueError: If PDF cannot be opened or unlocked
    """
    doc = fitz.open(pdf_path)

    if doc.is_encrypted:
        # Try empty password first (most common for DocuSign)
        if doc.authenticate(""):
            return doc

        # Try other common passwords
        common_passwords = ["", "docusign", "1234", "password"]
        for pwd in common_passwords:
            if doc.authenticate(pwd):
                return doc

        raise ValueError(
            "The PDF is password protected and cannot be unlocked. "
            "Please contact the sender for the password."
        )

    return doc


def create_execution_version(
    original_path: str,
    signed_path: str,
    insert_after: int,
    progress_callback: Callable[[str, int, str], None]
) -> Dict[str, Any]:
    """
    Create an execution version by merging signed pages into the original.

    Args:
        original_path: Path to original PDF (without signature pages)
        signed_path: Path to signed PDF from DocuSign
        insert_after: Page number after which to insert signed pages
                     (0 = beginning, -1 or > page_count = end)
        progress_callback: Function to report progress

    Returns:
        Dictionary with output path and details
    """
    progress_callback("loading", 10, "Loading original document...")

    # Validate inputs
    if not os.path.exists(original_path):
        raise ValueError(f"Original PDF not found: {original_path}")
    if not os.path.exists(signed_path):
        raise ValueError(f"Signed PDF not found: {signed_path}")

    # Open original document
    original = fitz.open(original_path)
    original_page_count = original.page_count

    progress_callback("loading", 20, "Unlocking signed document...")

    # Open and unlock signed document
    try:
        signed = unlock_pdf(signed_path)
    except Exception as e:
        original.close()
        raise ValueError(f"Could not open signed PDF: {str(e)}")

    signed_page_count = signed.page_count

    progress_callback("merging", 40, f"Merging {signed_page_count} signed pages...")

    # Determine insertion point
    if insert_after < 0 or insert_after >= original_page_count:
        # Insert at end
        insert_after = original_page_count

    # Create result document
    result = fitz.open()

    try:
        # Copy pages before insertion point
        if insert_after > 0:
            progress_callback("merging", 50, "Adding pages before signature pages...")
            result.insert_pdf(original, from_page=0, to_page=insert_after - 1)

        # Insert all signed pages
        progress_callback("merging", 70, "Inserting signed pages...")
        result.insert_pdf(signed)

        # Copy remaining pages from original
        if insert_after < original_page_count:
            progress_callback("merging", 85, "Adding remaining pages...")
            result.insert_pdf(original, from_page=insert_after)

        # Generate output filename
        original_basename = os.path.basename(original_path)
        name_without_ext = os.path.splitext(original_basename)[0]

        # Remove common suffixes like "_Clean", "_Without_Sigs", etc.
        for suffix in ["_Clean", "_Without_Sigs", "_Unsigned", "_Draft"]:
            if name_without_ext.endswith(suffix):
                name_without_ext = name_without_ext[:-len(suffix)]

        output_filename = f"{name_without_ext} (Execution Version).pdf"

        # Save to temp directory
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, output_filename)

        progress_callback("saving", 95, "Saving execution version...")
        result.save(output_path)

        final_page_count = result.page_count

    finally:
        result.close()
        original.close()
        signed.close()

    progress_callback("complete", 100, "Execution version created successfully!")

    return {
        "success": True,
        "outputPath": output_path,
        "outputFilename": output_filename,
        "originalPages": original_page_count,
        "signedPages": signed_page_count,
        "totalPages": final_page_count
    }
