#!/usr/bin/env python3
"""
EmmaNeigh - Signature Block Generator
Generates and inserts signature blocks into documents.
"""

import fitz
import os
import re
import sys
import json
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def generate_signature_block(entity_name, role=None, signer_name=None, signer_title=None):
    """
    Generate a formatted signature block for an entity.

    Args:
        entity_name: Legal name of the entity (e.g., "ABC Holdings, LLC")
        role: Role in the transaction (e.g., "Borrower", "Administrative Agent")
        signer_name: Name of the authorized signer
        signer_title: Title of the signer

    Returns:
        str: Formatted signature block text
    """
    lines = []

    # Entity name line
    lines.append(entity_name.upper() + ",")

    # Role line (if provided)
    if role:
        lines.append(f"as {role}")

    lines.append("")  # Blank line

    # Signature line
    lines.append("")
    lines.append("By: ________________________")

    # Name line
    if signer_name:
        lines.append(f"Name: {signer_name}")
    else:
        lines.append("Name: ________________________")

    # Title line
    if signer_title:
        lines.append(f"Title: {signer_title}")
    else:
        lines.append("Title: ________________________")

    return "\n".join(lines)


def generate_individual_signature_block(person_name):
    """
    Generate a signature block for an individual (not an entity).

    Args:
        person_name: Name of the individual

    Returns:
        str: Formatted signature block text
    """
    lines = [
        "",
        "________________________",
        person_name
    ]
    return "\n".join(lines)


def generate_signature_page(doc_name, signatories, entity_signer_map):
    """
    Generate a complete signature page for a document.

    Args:
        doc_name: Name of the document
        signatories: List of signatory roles for this document
        entity_signer_map: Dict mapping role -> {entity_name, signer_name, signer_title}

    Returns:
        str: Complete signature page text
    """
    blocks = []

    # Header
    blocks.append(f"SIGNATURE PAGE TO {doc_name.upper()}")
    blocks.append("")
    blocks.append("[Signature Page Follows]")
    blocks.append("")
    blocks.append("")

    # Generate block for each signatory
    for role in signatories:
        if role in entity_signer_map:
            entity_data = entity_signer_map[role]
            block = generate_signature_block(
                entity_name=entity_data.get('entity_name', role),
                role=role,
                signer_name=entity_data.get('signer_name'),
                signer_title=entity_data.get('signer_title')
            )
            blocks.append(block)
            blocks.append("")
            blocks.append("")
        else:
            # Unknown entity - create placeholder
            block = generate_signature_block(
                entity_name=f"[{role}]",
                role=role
            )
            blocks.append(block)
            blocks.append("")
            blocks.append("")

    return "\n".join(blocks)


def insert_signature_page_docx(docx_path, signature_page_text, output_path):
    """
    Insert a signature page into a Word document.

    Args:
        docx_path: Path to the input Word document
        signature_page_text: Text of the signature page to insert
        output_path: Path for the output document
    """
    doc = Document(docx_path)

    # Add page break before signature page
    doc.add_page_break()

    # Add signature page content
    lines = signature_page_text.split('\n')

    for i, line in enumerate(lines):
        para = doc.add_paragraph(line)

        # Style the header line
        if i == 0 and "SIGNATURE PAGE TO" in line:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True

    doc.save(output_path)


def insert_signature_page_pdf(pdf_path, signature_page_text, output_path):
    """
    Insert a signature page into a PDF document.

    Args:
        pdf_path: Path to the input PDF
        signature_page_text: Text of the signature page to insert
        output_path: Path for the output document
    """
    doc = fitz.open(pdf_path)

    # Create a new page at the end
    # Use the size of the last page as reference
    last_page = doc[-1]
    width = last_page.rect.width
    height = last_page.rect.height

    new_page = doc.new_page(-1, width=width, height=height)

    # Insert text
    # Start 72 points (1 inch) from top
    text_point = fitz.Point(72, 72)

    # Insert the signature page text
    new_page.insert_text(
        text_point,
        signature_page_text,
        fontsize=11,
        fontname="helv"
    )

    doc.save(output_path)
    doc.close()


def process_documents(documents_folder, checklist_data, entity_signer_map, output_folder):
    """
    Process all documents and insert signature pages.

    Args:
        documents_folder: Path to folder containing unsigned documents
        checklist_data: Parsed checklist data {doc_name: [roles]}
        entity_signer_map: Mapping of role -> entity/signer data
        output_folder: Path for output documents

    Returns:
        dict: Results with processed documents
    """
    os.makedirs(output_folder, exist_ok=True)

    processed = []
    errors = []

    # Get list of document files
    doc_files = [f for f in os.listdir(documents_folder)
                 if f.lower().endswith(('.pdf', '.docx'))]

    total = len(doc_files)

    for idx, filename in enumerate(doc_files):
        percent = int((idx / total) * 100) if total > 0 else 0
        emit("progress", percent=percent, message=f"Processing {filename}")

        filepath = os.path.join(documents_folder, filename)
        output_path = os.path.join(output_folder, filename)

        # Try to match document to checklist
        doc_name_base = os.path.splitext(filename)[0]
        signatories = None

        # Look for matching document in checklist
        for checklist_doc, roles in checklist_data.items():
            # Fuzzy match document names
            if (doc_name_base.lower() in checklist_doc.lower() or
                checklist_doc.lower() in doc_name_base.lower()):
                signatories = roles
                break

        if not signatories:
            # No match found - copy file as-is
            errors.append({
                "file": filename,
                "error": "No matching entry in checklist"
            })
            continue

        try:
            # Generate signature page
            sig_page = generate_signature_page(
                doc_name=doc_name_base,
                signatories=signatories,
                entity_signer_map=entity_signer_map
            )

            # Insert into document
            if filename.lower().endswith('.docx'):
                insert_signature_page_docx(filepath, sig_page, output_path)
            elif filename.lower().endswith('.pdf'):
                insert_signature_page_pdf(filepath, sig_page, output_path)

            processed.append({
                "file": filename,
                "signatories": signatories,
                "blocks_added": len(signatories)
            })

        except Exception as e:
            errors.append({
                "file": filename,
                "error": str(e)
            })

    return {
        "processed": processed,
        "errors": errors,
        "total_processed": len(processed),
        "total_errors": len(errors)
    }


def main():
    """CLI entry point for testing."""
    # For testing, generate a sample signature block
    if len(sys.argv) < 2:
        # Demo mode - print sample signature block
        sample = generate_signature_block(
            entity_name="ABC Holdings, LLC",
            role="Borrower",
            signer_name="John Smith",
            signer_title="Chief Executive Officer"
        )
        print("Sample Signature Block:")
        print("=" * 40)
        print(sample)
        print("=" * 40)
        sys.exit(0)

    emit("error", message="This module is meant to be imported, not run directly.")
    sys.exit(1)


if __name__ == "__main__":
    main()
