#!/usr/bin/env python3
"""
EmmaNeigh - Signature Packet Processor
v5.1.1: Added footer extraction, extended signature detection, output format support.
"""

import fitz
import os
import pandas as pd
import re
import sys
import json
import shutil
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


# Column header patterns for signature table detection
NAME_HEADERS = ["NAME", "PRINTED NAME", "SIGNATORY", "SIGNER", "PRINT NAME", "TITLE"]
SIGNATURE_HEADERS = ["SIGNATURE", "SIGN", "BY", "SIGN HERE"]

# Trigger phrases that indicate signature pages without BY:
SIGNATURE_TRIGGER_PHRASES = [
    "AGREED", "ACKNOWLEDGED", "UNDERSIGNED", "WITNESS", "ATTEST",
    "IN WITNESS WHEREOF", "EXECUTED", "CERTIFIED", "AUTHORIZED",
    "THE PARTIES HERETO", "DULY AUTHORIZED"
]


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def normalize_name(name):
    """Normalize signer name: uppercase, remove punctuation, collapse spaces."""
    if not name:
        return ""
    name = name.upper()
    name = re.sub(r"[.,]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def is_probable_person(name):
    """Check if name is likely a person (not an entity)."""
    if not name:
        return False
    entity_terms = ["LLC", "INC", "CORP", "CORPORATION", "LP", "LLP", "TRUST", "COMPANY", "LTD", "LIMITED"]
    name_upper = name.upper()
    if any(term in name_upper for term in entity_terms):
        return False
    # Check for reasonable word count (2-4 words typical for person names)
    word_count = len(name.split())
    if word_count < 1 or word_count > 5:
        return False
    # Check that it's not just numbers or special characters
    if not re.search(r'[A-Za-z]{2,}', name):
        return False
    return True


# ========== FOOTER EXTRACTION ==========

def extract_footer_from_pdf_page(page):
    """
    Extract footer text from the bottom of a PDF page.
    Looks for 'Signature Page to X' pattern or returns last meaningful line.
    """
    text = page.get_text()
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    if not lines:
        return ""

    # Check last 5 lines for footer patterns
    footer_lines = lines[-5:] if len(lines) >= 5 else lines

    # Priority 1: Look for "SIGNATURE PAGE TO X" pattern
    for line in footer_lines:
        line_upper = line.upper()
        if 'SIGNATURE PAGE' in line_upper:
            return line.strip()

    # Priority 2: Look for page numbers or document identifiers at bottom
    for line in reversed(footer_lines):
        # Skip very short lines (likely just page numbers)
        if len(line) < 3:
            continue
        # Skip lines that are just numbers
        if re.match(r'^[\d\s\-\.]+$', line):
            continue
        return line.strip()

    return ""


def extract_footer_from_docx(docx_path):
    """
    Extract footer from DOCX document sections.
    Falls back to last paragraph if no explicit footer.
    """
    try:
        doc = Document(docx_path)

        # Try to get explicit footer from sections
        for section in doc.sections:
            try:
                footer = section.footer
                if footer and footer.paragraphs:
                    footer_text = ' '.join(
                        p.text.strip() for p in footer.paragraphs if p.text.strip()
                    )
                    if footer_text:
                        return footer_text
            except Exception:
                pass

        # Fallback: check last few paragraphs for footer-like content
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        if paragraphs:
            for para in reversed(paragraphs[-3:]):
                if 'SIGNATURE PAGE' in para.upper():
                    return para
            # Return last non-empty paragraph
            return paragraphs[-1] if paragraphs else ""
    except Exception:
        pass

    return ""


# ========== EXTENDED SIGNATURE DETECTION ==========

def extract_signers_underscore_name(text):
    """
    Pattern: Name followed by underscore line (resolution style)
    Example: John Smith ________________________
    """
    signers = set()

    # Pattern: Name followed by 4+ underscores
    pattern = r'([A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+){0,3})\s*_{4,}'
    matches = re.findall(pattern, text)

    for name in matches:
        name = name.strip()
        if is_probable_person(name):
            signers.add(normalize_name(name))

    return signers


def extract_signers_underscore_label(text):
    """
    Pattern: Underscore line followed by Name:/Title: label
    Example:
    _______________________________
    Name: John Smith
    Title: President
    """
    signers = set()
    lines = text.splitlines()

    for i, line in enumerate(lines):
        # Look for line with 10+ underscores
        if re.search(r'_{10,}', line):
            # Check next 3 lines for Name: label
            for j in range(1, 4):
                if i + j < len(lines):
                    next_line = lines[i + j]
                    name_match = re.search(r'Name:\s*(.+)', next_line, re.IGNORECASE)
                    if name_match:
                        name = name_match.group(1).strip()
                        # Clean up the name (remove trailing underscores, etc.)
                        name = re.sub(r'_{2,}.*$', '', name).strip()
                        if name and is_probable_person(name):
                            signers.add(normalize_name(name))
                        break

    return signers


def extract_signers_trigger_phrase(text):
    """
    Pattern: Trigger phrases followed by names in subsequent lines
    Example:
    THE UNDERSIGNED HEREBY AGREE:

    Person A
    Person B
    Person C
    """
    signers = set()
    text_upper = text.upper()

    for phrase in SIGNATURE_TRIGGER_PHRASES:
        if phrase in text_upper:
            # Find where the phrase occurs
            phrase_idx = text_upper.find(phrase)
            subsequent_text = text[phrase_idx:]
            lines = subsequent_text.split('\n')[1:15]  # Check next 15 lines

            for line in lines:
                line = line.strip()
                # Skip empty lines, short lines, and lines that are just underscores
                if not line or len(line) < 3 or re.match(r'^[_\-\s]+$', line):
                    continue
                # Skip lines that look like instructions
                if any(word in line.upper() for word in ['PLEASE', 'SIGN', 'DATE', 'PRINT', 'BELOW']):
                    continue

                # Check if line looks like a name (2-4 words, starts with capital)
                words = line.split()
                if 1 <= len(words) <= 5:
                    # Remove trailing underscores or colons
                    candidate = re.sub(r'[_:]+$', '', ' '.join(words)).strip()
                    if is_probable_person(candidate):
                        signers.add(normalize_name(candidate))

    return signers


def extract_signers_horizontal_table(table_data):
    """
    Pattern: Horizontal signature tables (common in incumbency certs)
    Example:
    | Name        | Title       | Signature    |
    |-------------|-------------|--------------|
    | John Smith  | CEO         | ____________ |
    """
    if not table_data or len(table_data) < 2:
        return set()

    headers = table_data[0]
    headers_upper = [(h.upper().strip() if h else "") for h in headers]

    # Check if this looks like a signature/incumbency table
    has_name = any(any(nh in h for nh in NAME_HEADERS) for h in headers_upper)
    has_title = any('TITLE' in h for h in headers_upper)
    has_sig_or_empty = any(
        any(sh in h for sh in SIGNATURE_HEADERS) or h == "" or '___' in h
        for h in headers_upper
    )

    if not (has_name or has_title):
        return set()

    signers = set()
    name_col_idx = find_name_column(headers)

    # If no explicit name column, try first column
    if name_col_idx is None:
        name_col_idx = 0

    for row in table_data[1:]:
        if name_col_idx < len(row) and row[name_col_idx]:
            cell_text = row[name_col_idx]
            if isinstance(cell_text, str):
                # Handle multi-line cells
                for line in cell_text.split('\n'):
                    name = normalize_name(line.strip())
                    if name and is_probable_person(name):
                        signers.add(name)

    return signers


def detect_signature_page_extended(text, tables=None):
    """
    Extended signature page detection using multiple patterns.

    Returns:
        tuple: (is_signature_page: bool, signers: set, detection_method: str)
    """
    all_signers = set()
    methods_used = []

    # Method 1: Traditional BY: blocks
    by_signers = extract_signers_from_by_blocks(text)
    if by_signers:
        all_signers.update(by_signers)
        methods_used.append("BY_BLOCK")

    # Method 2: Standard signature tables
    if tables:
        for table_data in tables:
            if table_data and len(table_data) > 0:
                if is_signature_table(table_data[0]):
                    table_signers = extract_signers_from_table(table_data)
                    if table_signers:
                        all_signers.update(table_signers)
                        methods_used.append("TABLE")
                else:
                    # Try horizontal table detection
                    horiz_signers = extract_signers_horizontal_table(table_data)
                    if horiz_signers:
                        all_signers.update(horiz_signers)
                        methods_used.append("HORIZ_TABLE")

    # Method 3: Underscore + Name pattern
    underscore_signers = extract_signers_underscore_name(text)
    if underscore_signers:
        all_signers.update(underscore_signers)
        methods_used.append("UNDERSCORE_NAME")

    # Method 4: Underscore followed by Name: label
    label_signers = extract_signers_underscore_label(text)
    if label_signers:
        all_signers.update(label_signers)
        methods_used.append("UNDERSCORE_LABEL")

    # Method 5: Trigger phrases followed by names
    trigger_signers = extract_signers_trigger_phrase(text)
    if trigger_signers:
        all_signers.update(trigger_signers)
        methods_used.append("TRIGGER_PHRASE")

    # If we found signers, it's a signature page
    if all_signers:
        return (True, all_signers, ",".join(methods_used))

    # Additional heuristic: Check for signature indicators without detected names
    text_upper = text.upper()
    signature_indicators = [
        'SIGNATURE PAGE', 'EXECUTION PAGE', 'COUNTERPART SIGNATURE',
        'WITNESS WHEREOF', 'DULY AUTHORIZED', 'AUTHORIZED SIGNATORY',
        'NOTARY PUBLIC', 'ACKNOWLEDGED BEFORE ME'
    ]

    has_indicator = any(ind in text_upper for ind in signature_indicators)
    has_underscore_line = bool(re.search(r'_{10,}', text))

    if has_indicator or (has_underscore_line and any(phrase in text_upper for phrase in SIGNATURE_TRIGGER_PHRASES)):
        # This looks like a signature page but we couldn't detect names
        # Return with UNKNOWN SIGNER
        return (True, {"UNKNOWN SIGNER"}, "UNKNOWN")

    return (False, set(), None)


def extract_signers_from_by_blocks(text):
    """Extract signer names from traditional BY:/Name: blocks."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    signers = set()

    for i, line in enumerate(lines):
        if "BY:" in line.upper():
            # Tier 1: Prefer explicit Name: field
            for j in range(1, 7):
                if i + j >= len(lines):
                    break
                cand = lines[i + j]
                if cand.upper().startswith("NAME:"):
                    signers.add(normalize_name(cand.split(":", 1)[1]))
                    break
            else:
                # Tier 2: Look for probable person name nearby
                for j in range(1, 7):
                    if i + j >= len(lines):
                        break
                    cand = normalize_name(lines[i + j])
                    if is_probable_person(cand):
                        signers.add(cand)
                        break
    return signers


def find_name_column(headers):
    """Find the index of the name column in table headers."""
    if not headers:
        return None
    headers_upper = [(h.upper().strip() if h else "") for h in headers]
    for i, h in enumerate(headers_upper):
        for name_header in NAME_HEADERS:
            if name_header in h:
                return i
    return None


def is_signature_table(headers):
    """Check if table headers indicate a signature table."""
    if not headers:
        return False

    headers_upper = [(h.upper().strip() if h else "") for h in headers]

    # Must have a name-like column
    has_name = any(
        any(nh in h for nh in NAME_HEADERS)
        for h in headers_upper
    )

    # Should have signature-like column (or empty column which often is signature)
    has_sig = any(
        any(sh in h for sh in SIGNATURE_HEADERS) or h == ""
        for h in headers_upper
    )

    return has_name and has_sig


def extract_signers_from_table(table_data):
    """Extract signer names from a signature table."""
    if not table_data or len(table_data) < 2:
        return set()

    headers = table_data[0]
    name_col_idx = find_name_column(headers)

    if name_col_idx is None:
        return set()

    signers = set()
    for row in table_data[1:]:  # Skip header row
        if name_col_idx < len(row) and row[name_col_idx]:
            # Handle multi-line cell content
            cell_text = row[name_col_idx]
            if isinstance(cell_text, str):
                cell_text = " ".join(cell_text.split())  # Normalize whitespace
            name = normalize_name(cell_text)
            if name and is_probable_person(name):
                signers.add(name)

    return signers


def extract_person_signers(page):
    """
    Extract signer names from PDF page using extended detection.
    Returns a tuple: (signers: set, detection_method: str)
    """
    text = page.get_text()

    # Get tables from page
    tables_data = []
    try:
        tables = page.find_tables()
        for table in tables.tables:
            data = table.extract()
            if data:
                tables_data.append(data)
    except Exception:
        pass

    # Use extended detection
    is_sig_page, signers, method = detect_signature_page_extended(text, tables_data)

    return signers, method if method else ""


# ========== DOCX SUPPORT ==========

def extract_text_from_docx(docx_path):
    """Extract all text from DOCX document (headers, body)."""
    doc = Document(docx_path)
    text_parts = []

    # Get header text
    for section in doc.sections:
        try:
            for para in section.header.paragraphs:
                if para.text.strip():
                    text_parts.append(para.text)
        except Exception:
            pass

    # Get body text
    for para in doc.paragraphs:
        if para.text.strip():
            text_parts.append(para.text)

    return '\n'.join(text_parts)


def extract_tables_from_docx(docx_path):
    """Extract all tables from DOCX as list of rows."""
    doc = Document(docx_path)
    tables_list = []

    for table in doc.tables:
        rows = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            rows.append(row_data)
        if rows:
            tables_list.append(rows)

    return tables_list


def extract_signers_from_docx(docx_path):
    """
    Extract signers from DOCX file using extended detection.
    Returns a tuple: (signers: set, detection_method: str)
    """
    text = extract_text_from_docx(docx_path)
    tables_data = extract_tables_from_docx(docx_path)

    # Use extended detection
    is_sig_page, signers, method = detect_signature_page_extended(text, tables_data)

    return signers, method if method else ""


def create_docx_packet(signer_name, docs_for_signer, output_folder):
    """
    Create a DOCX signature packet for a signer from DOCX source documents.

    Args:
        signer_name: Name of the signer
        docs_for_signer: List of (filename, filepath) tuples for this signer
        output_folder: Where to save the packet

    Returns:
        Path to created packet, or None if failed
    """
    try:
        # Create new document for the packet
        packet_doc = Document()

        # Add header
        header_para = packet_doc.add_paragraph()
        header_run = header_para.add_run(f"SIGNATURE PACKET FOR {signer_name}")
        header_run.bold = True
        header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        packet_doc.add_paragraph()  # Blank line

        docs_added = 0

        for filename, filepath in docs_for_signer:
            if not filepath.lower().endswith('.docx'):
                continue

            try:
                # Open source document
                source_doc = Document(filepath)

                # Add document separator
                sep_para = packet_doc.add_paragraph()
                sep_run = sep_para.add_run(f"─" * 50)
                packet_doc.add_paragraph()

                doc_title = packet_doc.add_paragraph()
                title_run = doc_title.add_run(f"Document: {filename}")
                title_run.bold = True

                packet_doc.add_paragraph()  # Blank line

                # Copy content from source document
                for para in source_doc.paragraphs:
                    new_para = packet_doc.add_paragraph()
                    new_para.style = para.style

                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        new_run.bold = run.bold
                        new_run.italic = run.italic
                        if run.font.size:
                            new_run.font.size = run.font.size

                # Copy tables
                for table in source_doc.tables:
                    # Create table with same dimensions
                    new_table = packet_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                    new_table.style = 'Table Grid'

                    for i, row in enumerate(table.rows):
                        for j, cell in enumerate(row.cells):
                            new_table.rows[i].cells[j].text = cell.text

                docs_added += 1

                # Add page break between documents
                packet_doc.add_page_break()

            except Exception as e:
                # Skip problematic documents
                continue

        if docs_added > 0:
            packet_path = os.path.join(output_folder, f"signature_packet - {signer_name}.docx")
            packet_doc.save(packet_path)
            return packet_path

    except Exception as e:
        pass

    return None


def create_pdf_packet(signer_name, docs_for_signer, output_folder, filepath_lookup):
    """
    Create a PDF signature packet for a signer from PDF source documents.

    Args:
        signer_name: Name of the signer
        docs_for_signer: DataFrame rows for this signer
        output_folder: Where to save the packet
        filepath_lookup: Dict mapping filename -> filepath

    Returns:
        Tuple of (path, page_count) or (None, 0) if failed
    """
    try:
        packet = fitz.open()

        for _, r in docs_for_signer.iterrows():
            if r["Document"].lower().endswith('.pdf'):
                try:
                    doc_path = filepath_lookup.get(r["Document"], r["Document"])
                    src = fitz.open(doc_path)
                    packet.insert_pdf(src, from_page=r["Page"] - 1, to_page=r["Page"] - 1)
                    src.close()
                except Exception:
                    pass

        if packet.page_count > 0:
            pdf_path = os.path.join(output_folder, f"signature_packet - {signer_name}.pdf")
            packet.save(pdf_path)
            page_count = packet.page_count
            packet.close()
            return pdf_path, page_count

        packet.close()
    except Exception:
        pass

    return None, 0


# ========== FORMAT CONVERSION ==========

def convert_pdf_to_docx(pdf_path, docx_path):
    """
    Convert PDF to DOCX using pdf2docx for high fidelity conversion.
    Returns True on success, False on failure.
    """
    try:
        from pdf2docx import Converter
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        return True
    except ImportError:
        emit("progress", percent=0, message="Warning: pdf2docx not installed, skipping conversion")
        return False
    except Exception as e:
        emit("progress", percent=0, message=f"Warning: PDF to DOCX conversion failed: {str(e)}")
        return False


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert DOCX to PDF.
    Uses python-docx to read and PyMuPDF to create PDF.
    Note: This is a basic conversion - complex formatting may not be preserved perfectly.
    """
    try:
        # Read DOCX content
        doc = Document(docx_path)

        # Create new PDF
        pdf_doc = fitz.open()
        page = pdf_doc.new_page()

        # Calculate page dimensions
        rect = page.rect
        margin = 72  # 1 inch margins
        y_position = margin

        # Process paragraphs
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                y_position += 12  # Blank line
                continue

            # Determine font size based on style
            font_size = 11
            if para.style and 'Heading' in para.style.name:
                font_size = 14
            elif para.style and 'Title' in para.style.name:
                font_size = 16

            # Check if we need a new page
            if y_position > rect.height - margin:
                page = pdf_doc.new_page()
                y_position = margin

            # Insert text
            text_rect = fitz.Rect(margin, y_position, rect.width - margin, y_position + font_size + 4)
            page.insert_textbox(text_rect, text, fontsize=font_size)
            y_position += font_size + 6

        # Process tables
        for table in doc.tables:
            # Simple table representation
            for row in table.rows:
                row_text = " | ".join(cell.text.strip() for cell in row.cells)
                if y_position > rect.height - margin:
                    page = pdf_doc.new_page()
                    y_position = margin

                text_rect = fitz.Rect(margin, y_position, rect.width - margin, y_position + 14)
                page.insert_textbox(text_rect, row_text, fontsize=10)
                y_position += 16

        if pdf_doc.page_count > 0:
            pdf_doc.save(pdf_path)
            pdf_doc.close()
            return True

        pdf_doc.close()
    except Exception as e:
        emit("progress", percent=0, message=f"Warning: DOCX to PDF conversion failed: {str(e)}")

    return False


def create_packet_with_format(signer_name, docs_for_signer, output_folder, filepath_lookup, output_format='preserve'):
    """
    Create signature packet(s) with specified output format.

    Args:
        signer_name: Name of the signer
        docs_for_signer: DataFrame rows for this signer
        output_folder: Where to save packets
        filepath_lookup: Dict mapping filename -> filepath
        output_format: 'preserve', 'pdf', 'docx', or 'both'

    Returns:
        List of created packet info dicts
    """
    packets = []

    # Separate by source format
    pdf_docs = docs_for_signer[docs_for_signer["Document"].str.lower().str.endswith('.pdf')]
    docx_docs = docs_for_signer[docs_for_signer["Document"].str.lower().str.endswith('.docx')]

    if output_format == 'preserve':
        # Original behavior: output matches input
        if len(pdf_docs) > 0:
            pdf_path, page_count = create_pdf_packet(signer_name, pdf_docs, output_folder, filepath_lookup)
            if pdf_path:
                packets.append({"name": signer_name, "pages": page_count, "format": "pdf", "path": pdf_path})

        if len(docx_docs) > 0:
            docx_files = [(r["Document"], filepath_lookup.get(r["Document"])) for _, r in docx_docs.iterrows()]
            docx_path = create_docx_packet(signer_name, docx_files, output_folder)
            if docx_path:
                packets.append({"name": signer_name, "pages": len(docx_files), "format": "docx", "path": docx_path})

    elif output_format == 'pdf':
        # Convert everything to PDF
        if len(pdf_docs) > 0:
            pdf_path, page_count = create_pdf_packet(signer_name, pdf_docs, output_folder, filepath_lookup)
            if pdf_path:
                packets.append({"name": signer_name, "pages": page_count, "format": "pdf", "path": pdf_path})

        if len(docx_docs) > 0:
            # First create DOCX packet, then convert to PDF
            docx_files = [(r["Document"], filepath_lookup.get(r["Document"])) for _, r in docx_docs.iterrows()]
            temp_docx_path = create_docx_packet(signer_name + "_temp", docx_files, output_folder)
            if temp_docx_path:
                pdf_path = os.path.join(output_folder, f"signature_packet - {signer_name} (from docx).pdf")
                if convert_docx_to_pdf(temp_docx_path, pdf_path):
                    packets.append({"name": signer_name, "pages": len(docx_files), "format": "pdf", "path": pdf_path})
                # Clean up temp file
                try:
                    os.remove(temp_docx_path)
                except:
                    pass

    elif output_format == 'docx':
        # Convert everything to DOCX
        if len(docx_docs) > 0:
            docx_files = [(r["Document"], filepath_lookup.get(r["Document"])) for _, r in docx_docs.iterrows()]
            docx_path = create_docx_packet(signer_name, docx_files, output_folder)
            if docx_path:
                packets.append({"name": signer_name, "pages": len(docx_files), "format": "docx", "path": docx_path})

        if len(pdf_docs) > 0:
            # First create PDF packet, then convert to DOCX
            pdf_path, page_count = create_pdf_packet(signer_name + "_temp", pdf_docs, output_folder, filepath_lookup)
            if pdf_path:
                docx_path = os.path.join(output_folder, f"signature_packet - {signer_name} (from pdf).docx")
                if convert_pdf_to_docx(pdf_path, docx_path):
                    packets.append({"name": signer_name, "pages": page_count, "format": "docx", "path": docx_path})
                # Clean up temp file
                try:
                    os.remove(pdf_path)
                except:
                    pass

    elif output_format == 'both':
        # Create both formats
        # First, create native format packets
        if len(pdf_docs) > 0:
            pdf_path, page_count = create_pdf_packet(signer_name, pdf_docs, output_folder, filepath_lookup)
            if pdf_path:
                packets.append({"name": signer_name, "pages": page_count, "format": "pdf", "path": pdf_path})
                # Also create DOCX version
                docx_path = os.path.join(output_folder, f"signature_packet - {signer_name} (from pdf).docx")
                if convert_pdf_to_docx(pdf_path, docx_path):
                    packets.append({"name": signer_name, "pages": page_count, "format": "docx", "path": docx_path})

        if len(docx_docs) > 0:
            docx_files = [(r["Document"], filepath_lookup.get(r["Document"])) for _, r in docx_docs.iterrows()]
            docx_path = create_docx_packet(signer_name, docx_files, output_folder)
            if docx_path:
                packets.append({"name": signer_name, "pages": len(docx_files), "format": "docx", "path": docx_path})
                # Also create PDF version
                pdf_path = os.path.join(output_folder, f"signature_packet - {signer_name} (from docx).pdf")
                if convert_docx_to_pdf(docx_path, pdf_path):
                    packets.append({"name": signer_name, "pages": len(docx_files), "format": "pdf", "path": pdf_path})

    return packets


# ========== MAIN ==========

def main():
    if len(sys.argv) < 2:
        emit("error", message="No input provided.")
        sys.exit(1)

    input_dir = None
    file_paths = None
    output_format = 'preserve'  # Default: output format matches input format

    # Check if we have --config argument (for file list) or folder path
    if sys.argv[1] == '--config':
        if len(sys.argv) < 3:
            emit("error", message="No config file provided.")
            sys.exit(1)
        config_path = sys.argv[2]
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
            file_paths = config.get('files', [])
            output_format = config.get('output_format', 'preserve')
            if not file_paths:
                emit("error", message="No files in config.")
                sys.exit(1)
            # Use temp directory for output
            import tempfile
            input_dir = tempfile.mkdtemp(prefix='emmaneigh_packets_')
        except Exception as e:
            emit("error", message=f"Failed to read config: {str(e)}")
            sys.exit(1)
    else:
        input_dir = sys.argv[1]
        if not os.path.isdir(input_dir):
            emit("error", message=f"Invalid folder: {input_dir}")
            sys.exit(1)

    # Set up output directories
    output_base = os.path.join(input_dir, "signature_packets_output")
    output_pdf_dir = os.path.join(output_base, "packets")
    output_table_dir = os.path.join(output_base, "tables")

    os.makedirs(output_pdf_dir, exist_ok=True)
    os.makedirs(output_table_dir, exist_ok=True)

    # Get document files - either from file_paths list or scan directory
    if file_paths:
        # Filter to valid PDF/DOCX files
        document_files = [(os.path.basename(f), f) for f in file_paths
                          if os.path.isfile(f) and f.lower().endswith((".pdf", ".docx"))]
    else:
        # Scan directory
        document_files = [(f, os.path.join(input_dir, f)) for f in os.listdir(input_dir)
                          if f.lower().endswith((".pdf", ".docx"))]

    if not document_files:
        emit("error", message="No PDF or Word files found.")
        sys.exit(1)

    total = len(document_files)
    emit("progress", percent=0, message=f"Found {total} documents")

    # Scan all documents for signature pages
    rows = []
    # Build filepath lookup for later use
    filepath_lookup = {filename: filepath for filename, filepath in document_files}

    for idx, (filename, filepath) in enumerate(document_files):
        percent = int((idx / total) * 50)
        emit("progress", percent=percent, message=f"Scanning {filename}")

        try:
            if filename.lower().endswith('.pdf'):
                # PDF processing
                doc = fitz.open(filepath)
                for page_num, page in enumerate(doc, start=1):
                    signers, detection_method = extract_person_signers(page)
                    if signers:
                        # Extract footer for this page
                        footer = extract_footer_from_pdf_page(page)
                        for signer in signers:
                            rows.append({
                                "Signer Name": signer,
                                "Document": filename,
                                "Page": page_num,
                                "Footer": footer,
                                "Detection Method": detection_method
                            })
                doc.close()
            elif filename.lower().endswith('.docx'):
                # DOCX processing
                signers, detection_method = extract_signers_from_docx(filepath)
                if signers:
                    # Extract footer for DOCX
                    footer = extract_footer_from_docx(filepath)
                    for signer in signers:
                        rows.append({
                            "Signer Name": signer,
                            "Document": filename,
                            "Page": 1,  # DOCX doesn't have pages
                            "Footer": footer,
                            "Detection Method": detection_method
                        })
        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: {filename} - {str(e)}")

    if not rows:
        emit("error", message="No signature pages detected in any documents.")
        sys.exit(1)

    # Create DataFrame and sort
    # Columns: Signer Name, Document, Page, Footer, Detection Method
    df = pd.DataFrame(rows)
    # Reorder columns for cleaner output
    column_order = ["Signer Name", "Document", "Page", "Footer", "Detection Method"]
    df = df[[col for col in column_order if col in df.columns]]
    df = df.sort_values(["Signer Name", "Document", "Page"])

    # Save master index
    emit("progress", percent=55, message="Creating master index...")
    df.to_excel(os.path.join(output_table_dir, "MASTER_SIGNATURE_INDEX.xlsx"), index=False)

    # Create individual packets with specified output format
    signers = df.groupby("Signer Name")
    total_signers = len(signers)
    packets_created = []

    emit("progress", percent=55, message=f"Creating packets (format: {output_format})...")

    for idx, (signer, group) in enumerate(signers):
        percent = 55 + int((idx / total_signers) * 45)
        emit("progress", percent=percent, message=f"Creating packet for {signer}")

        # Save signer's Excel file
        group.to_excel(
            os.path.join(output_table_dir, f"signature_packet - {signer}.xlsx"),
            index=False
        )

        # Create packets using the new format-aware function
        signer_packets = create_packet_with_format(
            signer, group, output_pdf_dir, filepath_lookup, output_format
        )
        packets_created.extend(signer_packets)

    emit("progress", percent=100, message="Complete!")
    emit("result",
         success=True,
         outputPath=output_base,
         packetsCreated=len(packets_created),
         packets=packets_created)


if __name__ == "__main__":
    main()
