#!/usr/bin/env python3
"""
EmmaNeigh - Signature Packet Processor
This is the EXACT v1 logic with JSON output for the UI.
"""

import fitz
import os
import pandas as pd
import re
import sys
import json


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def normalize_name(name):
    """Normalize signer name: uppercase, remove punctuation, collapse spaces."""
    name = name.upper()
    name = re.sub(r"[.,]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def is_probable_person(name):
    """Check if name is likely a person (not an entity)."""
    entity_terms = ["LLC", "INC", "CORP", "CORPORATION", "LP", "LLP", "TRUST"]
    if any(term in name for term in entity_terms):
        return False
    return 2 <= len(name.split()) <= 4


def extract_person_signers(text):
    """Extract signer names from page text. Look for BY: then Name: field."""
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

    # Find PDF files
    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]

    if not pdf_files:
        emit("error", message="No PDF files found in the folder.")
        sys.exit(1)

    total = len(pdf_files)
    emit("progress", percent=0, message=f"Found {total} PDF files")

    # Scan all PDFs for signature pages
    rows = []

    for idx, filename in enumerate(pdf_files):
        percent = int((idx / total) * 50)
        emit("progress", percent=percent, message=f"Scanning {filename}")

        try:
            doc = fitz.open(os.path.join(input_dir, filename))
            for page_num, page in enumerate(doc, start=1):
                signers = extract_person_signers(page.get_text())
                for signer in signers:
                    rows.append({
                        "Signer Name": signer,
                        "Document": filename,
                        "Page": page_num
                    })
            doc.close()
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

        # Create PDF packet
        packet = fitz.open()
        for _, r in group.iterrows():
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
