#!/usr/bin/env python3
"""
EmmaNeigh - Signature Packet Processor
v3.2.1: Added MS Word (.docx) support.
"""

import fitz
import os
import pandas as pd
import re
import sys
import json
from docx import Document


# Column header patterns for signature table detection
NAME_HEADERS = ["NAME", "PRINTED NAME", "SIGNATORY", "SIGNER", "PRINT NAME"]
SIGNATURE_HEADERS = ["SIGNATURE", "SIGN", "BY", "SIGN HERE"]


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
    entity_terms = ["LLC", "INC", "CORP", "CORPORATION", "LP", "LLP", "TRUST"]
    if any(term in name for term in entity_terms):
        return False
    return 2 <= len(name.split()) <= 4


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
    Extract signer names from PDF page (BY: blocks AND signature tables).
    Returns a set of normalized signer names.
    """
    signers = set()

    # Method 1: Traditional BY:/NAME: detection
    text = page.get_text()
    signers.update(extract_signers_from_by_blocks(text))

    # Method 2: Table-based signature detection
    try:
        tables = page.find_tables()
        for table in tables.tables:
            data = table.extract()
            if data and len(data) > 0:
                if is_signature_table(data[0]):
                    signers.update(extract_signers_from_table(data))
    except Exception:
        pass

    return signers


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
    """Extract signers from DOCX file (BY: blocks AND tables)."""
    signers = set()

    # Method 1: BY:/NAME: blocks
    text = extract_text_from_docx(docx_path)
    signers.update(extract_signers_from_by_blocks(text))

    # Method 2: Tables
    for table_data in extract_tables_from_docx(docx_path):
        if table_data and is_signature_table(table_data[0]):
            signers.update(extract_signers_from_table(table_data))

    return signers


# ========== MAIN ==========

def main():
    if len(sys.argv) < 2:
        emit("error", message="No folder provided.")
        sys.exit(1)

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

    # Find PDF and DOCX files
    document_files = [f for f in os.listdir(input_dir)
                      if f.lower().endswith((".pdf", ".docx"))]

    if not document_files:
        emit("error", message="No PDF or Word files found in the folder.")
        sys.exit(1)

    total = len(document_files)
    emit("progress", percent=0, message=f"Found {total} documents")

    # Scan all documents for signature pages
    rows = []

    for idx, filename in enumerate(document_files):
        percent = int((idx / total) * 50)
        emit("progress", percent=percent, message=f"Scanning {filename}")
        filepath = os.path.join(input_dir, filename)

        try:
            if filename.lower().endswith('.pdf'):
                # PDF processing
                doc = fitz.open(filepath)
                for page_num, page in enumerate(doc, start=1):
                    signers = extract_person_signers(page)
                    for signer in signers:
                        rows.append({
                            "Signer Name": signer,
                            "Document": filename,
                            "Page": page_num
                        })
                doc.close()
            elif filename.lower().endswith('.docx'):
                # DOCX processing
                signers = extract_signers_from_docx(filepath)
                for signer in signers:
                    rows.append({
                        "Signer Name": signer,
                        "Document": filename,
                        "Page": 1  # DOCX doesn't have pages
                    })
        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: {filename} - {str(e)}")

    if not rows:
        emit("error", message="No signature pages detected in any documents.")
        sys.exit(1)

    # Create DataFrame and sort
    df = pd.DataFrame(rows).sort_values(["Signer Name", "Document", "Page"])

    # Save master index
    emit("progress", percent=55, message="Creating master index...")
    df.to_excel(os.path.join(output_table_dir, "MASTER_SIGNATURE_INDEX.xlsx"), index=False)

    # Create individual packets
    signers = df.groupby("Signer Name")
    total_signers = len(signers)
    packets_created = []

    for idx, (signer, group) in enumerate(signers):
        percent = 55 + int((idx / total_signers) * 45)
        emit("progress", percent=percent, message=f"Creating packet for {signer}")

        # Save signer's Excel file
        group.to_excel(
            os.path.join(output_table_dir, f"signature_packet - {signer}.xlsx"),
            index=False
        )

        # Create PDF packet (only from PDF sources)
        packet = fitz.open()
        for _, r in group.iterrows():
            if r["Document"].lower().endswith('.pdf'):
                try:
                    src = fitz.open(os.path.join(input_dir, r["Document"]))
                    packet.insert_pdf(src, from_page=r["Page"] - 1, to_page=r["Page"] - 1)
                    src.close()
                except Exception:
                    pass

        if packet.page_count > 0:
            pdf_path = os.path.join(output_pdf_dir, f"signature_packet - {signer}.pdf")
            packet.save(pdf_path)
            packets_created.append({
                "name": signer,
                "pages": packet.page_count
            })
        packet.close()

    emit("progress", percent=100, message="Complete!")
    emit("result",
         success=True,
         outputPath=output_base,
         packetsCreated=len(packets_created),
         packets=packets_created)


if __name__ == "__main__":
    main()
