"""
Signature Packets Processor

Extracts signature pages from multiple PDFs and organizes them by signer.
Creates individual signature packets (PDFs) and Excel tracking sheets.
"""

import fitz  # PyMuPDF
import os
import re
from typing import Callable, Set, List, Dict, Any


def normalize_name(name: str) -> str:
    """
    Normalize a signer name for consistent grouping.

    - Convert to uppercase
    - Remove punctuation
    - Collapse whitespace
    """
    name = name.upper()
    name = re.sub(r"[.,]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def is_probable_person(name: str) -> bool:
    """
    Determine if a name is likely a person (not an entity).

    Entity indicators: LLC, INC, CORP, LP, LLP, TRUST
    Person heuristic: 2-4 words in name
    """
    entity_terms = [
        "LLC", "INC", "CORP", "CORPORATION",
        "LP", "LLP", "TRUST", "HOLDINGS",
        "PARTNERS", "FUND", "CAPITAL"
    ]
    name_upper = name.upper()
    if any(term in name_upper for term in entity_terms):
        return False
    return 2 <= len(name.split()) <= 4


def extract_signers_from_page(text: str) -> Set[str]:
    """
    Extract signer names from a page's text content.

    Strategy:
    1. Look for "BY:" markers
    2. Prefer explicit "Name:" field (Tier 1)
    3. Fall back to probable person name nearby (Tier 2)
    """
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    signers = set()

    for i, line in enumerate(lines):
        if "BY:" in line.upper():
            # Tier 1: Look for explicit Name: field
            found_name = False
            for j in range(1, 7):
                if i + j >= len(lines):
                    break
                candidate = lines[i + j]
                if candidate.upper().startswith("NAME:"):
                    name = candidate.split(":", 1)[1].strip()
                    if name:
                        signers.add(normalize_name(name))
                        found_name = True
                    break

            # Tier 2: Look for probable person name
            if not found_name:
                for j in range(1, 7):
                    if i + j >= len(lines):
                        break
                    candidate = normalize_name(lines[i + j])
                    if candidate and is_probable_person(candidate):
                        signers.add(candidate)
                        break

    return signers


def process_signature_packets(
    input_folder: str,
    progress_callback: Callable[[str, int, str], None]
) -> Dict[str, Any]:
    """
    Process all PDFs in a folder and create signature packets.

    Args:
        input_folder: Path to folder containing PDF files
        progress_callback: Function to report progress (stage, percent, message)

    Returns:
        Dictionary with processing results
    """
    # Validate input
    if not os.path.isdir(input_folder):
        raise ValueError(f"Invalid folder: {input_folder}")

    # Set up output directories
    output_base = os.path.join(input_folder, "signature_packets_output")
    output_pdf_dir = os.path.join(output_base, "packets")
    output_table_dir = os.path.join(output_base, "tables")

    os.makedirs(output_pdf_dir, exist_ok=True)
    os.makedirs(output_table_dir, exist_ok=True)

    # Find all PDF files
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

    if not pdf_files:
        raise ValueError("No PDF files found in the specified folder.")

    total_files = len(pdf_files)
    progress_callback("scanning", 0, f"Found {total_files} PDF files")

    # Collect all signature data
    rows: List[Dict[str, Any]] = []

    for idx, filename in enumerate(pdf_files):
        percent = int((idx / total_files) * 50)  # First 50% is scanning
        progress_callback("scanning", percent, f"Scanning {filename}")

        try:
            doc = fitz.open(os.path.join(input_folder, filename))

            for page_num, page in enumerate(doc, start=1):
                text = page.get_text()
                signers = extract_signers_from_page(text)

                for signer in signers:
                    rows.append({
                        "Signer Name": signer,
                        "Document": filename,
                        "Page": page_num
                    })

            doc.close()
        except Exception as e:
            # Log error but continue processing
            progress_callback("warning", percent, f"Error processing {filename}: {str(e)}")

    if not rows:
        raise ValueError("No signature pages detected in any documents.")

    # Sort by signer name, then document, then page
    rows.sort(key=lambda r: (r["Signer Name"], r["Document"], r["Page"]))

    # Group by signer
    signers: Dict[str, List[Dict[str, Any]]] = {}
    for row in rows:
        signer_name = row["Signer Name"]
        if signer_name not in signers:
            signers[signer_name] = []
        signers[signer_name].append(row)

    total_signers = len(signers)
    progress_callback("extracting", 50, f"Creating packets for {total_signers} signers")

    # Create signature packets and Excel files
    try:
        import pandas as pd
        has_pandas = True
    except ImportError:
        has_pandas = False

    # Create master index
    if has_pandas:
        import pandas as pd
        df = pd.DataFrame(rows)
        df.to_excel(os.path.join(output_table_dir, "MASTER_SIGNATURE_INDEX.xlsx"), index=False)

    packets_created = []

    for idx, (signer_name, signer_rows) in enumerate(signers.items()):
        percent = 50 + int((idx / total_signers) * 50)  # Second 50% is creating packets
        progress_callback("extracting", percent, f"Producing signature packet for {signer_name}")

        # Create individual Excel file
        if has_pandas:
            signer_df = pd.DataFrame(signer_rows)
            excel_path = os.path.join(output_table_dir, f"signature_packet - {signer_name}.xlsx")
            signer_df.to_excel(excel_path, index=False)

        # Create signature packet PDF
        packet = fitz.open()

        for row in signer_rows:
            try:
                src_path = os.path.join(input_folder, row["Document"])
                src = fitz.open(src_path)
                packet.insert_pdf(
                    src,
                    from_page=row["Page"] - 1,
                    to_page=row["Page"] - 1
                )
                src.close()
            except Exception as e:
                progress_callback("warning", percent, f"Error extracting page from {row['Document']}: {str(e)}")

        if packet.page_count > 0:
            pdf_path = os.path.join(output_pdf_dir, f"signature_packet - {signer_name}.pdf")
            packet.save(pdf_path)
            packets_created.append({
                "name": signer_name,
                "pages": packet.page_count,
                "path": pdf_path
            })

        packet.close()

    progress_callback("complete", 100, f"Created {len(packets_created)} signature packets")

    return {
        "success": True,
        "packetsCreated": len(packets_created),
        "packets": packets_created,
        "outputPath": output_base
    }
