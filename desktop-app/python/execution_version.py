#!/usr/bin/env python3
"""
EmmaNeigh - Execution Version Processor
Merges signed DocuSign pages back into original agreements, replacing blank signature pages in-place.
"""

import fitz
import os
import re
import sys
import json
from difflib import SequenceMatcher


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def normalize_name(name):
    """Normalize signer name: uppercase, remove punctuation, collapse spaces."""
    name = name.upper()
    name = re.sub(r"[.,]", "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def sanitize_output_name(filename):
    """
    Remove existing parenthetical suffixes and add (executed).
    Example: 'Credit Agreement (execution version).pdf' -> 'Credit Agreement (executed).pdf'
    """
    # Remove .pdf extension
    name = filename
    if name.lower().endswith('.pdf'):
        name = name[:-4]

    # Remove any parenthetical suffixes
    name = re.sub(r'\s*\([^)]*\)\s*$', '', name).strip()

    return f"{name} (executed).pdf"


def detect_document_name_from_footer(text):
    """
    Look for footer pattern: 'SIGNATURE PAGE TO [DOCUMENT NAME]' or '[DOCUMENT NAME] SIGNATURE PAGE'
    Returns the document name or None.
    """
    text_upper = text.upper()

    # Pattern 1: "SIGNATURE PAGE TO [X]"
    match = re.search(r'SIGNATURE\s+PAGE\s+TO\s+(?:THE\s+)?(.+?)(?:\n|$)', text_upper)
    if match:
        doc_name = match.group(1).strip()
        # Clean up - remove trailing punctuation
        doc_name = re.sub(r'[.\-–—]+$', '', doc_name).strip()
        if len(doc_name) > 3:
            return doc_name

    # Pattern 2: "[X] SIGNATURE PAGE"
    match = re.search(r'(.+?)\s+SIGNATURE\s+PAGE(?:\s|$)', text_upper)
    if match:
        doc_name = match.group(1).strip()
        # Exclude if it starts with generic words
        if not doc_name.startswith(('THIS', 'THE', 'A ', 'AN ')):
            doc_name = re.sub(r'[.\-–—]+$', '', doc_name).strip()
            if len(doc_name) > 3:
                return doc_name

    return None


def detect_document_name_from_title(text):
    """
    Look for common document titles in the header/text.
    Returns the document name or None.
    """
    text_upper = text.upper()

    # Common M&A document types
    doc_types = [
        'CREDIT AGREEMENT',
        'GUARANTEE',
        'GUARANTY',
        'PLEDGE AGREEMENT',
        'SECURITY AGREEMENT',
        'COLLATERAL AGREEMENT',
        'INTERCREDITOR AGREEMENT',
        'SUBORDINATION AGREEMENT',
        'LOAN AGREEMENT',
        'NOTE PURCHASE AGREEMENT',
        'INDENTURE',
        'AMENDMENT',
        'CONSENT',
        'JOINDER',
        'ASSIGNMENT',
    ]

    for doc_type in doc_types:
        if doc_type in text_upper:
            # Try to get more context around the doc type
            match = re.search(rf'({doc_type}[A-Z\s]*?)(?:\n|BY:|DATED)', text_upper)
            if match:
                return match.group(1).strip()
            return doc_type

    return None


def extract_signer_names(text):
    """Extract signer names from page text. Look for BY: then Name: field."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    signers = set()

    for i, line in enumerate(lines):
        if "BY:" in line.upper():
            # Prefer explicit Name: field
            for j in range(1, 7):
                if i + j >= len(lines):
                    break
                cand = lines[i + j]
                if cand.upper().startswith("NAME:"):
                    signers.add(normalize_name(cand.split(":", 1)[1]))
                    break

    return signers


def fuzzy_match_score(s1, s2):
    """Return similarity ratio between two strings (0-1)."""
    s1 = s1.upper()
    s2 = s2.upper()
    return SequenceMatcher(None, s1, s2).ratio()


def find_signature_pages_in_document(doc):
    """
    Find pages with signature blocks (BY: fields) in a document.
    Returns list of {page_num, signers} dicts.
    """
    sig_pages = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()

        if "BY:" in text.upper():
            signers = extract_signer_names(text)
            sig_pages.append({
                'page_num': page_num,
                'signers': signers,
                'text': text
            })

    return sig_pages


def process_execution_version(originals_folder, signed_pdf_path):
    """
    Main processing function.

    Args:
        originals_folder: Path to folder containing original agreement PDFs
        signed_pdf_path: Path to the signed PDF from DocuSign
    """

    # Validate inputs
    if not os.path.isdir(originals_folder):
        emit("error", message=f"Invalid originals folder: {originals_folder}")
        sys.exit(1)

    if not os.path.isfile(signed_pdf_path):
        emit("error", message=f"Invalid signed PDF: {signed_pdf_path}")
        sys.exit(1)

    # Set up output directories
    output_base = os.path.join(originals_folder, "execution_version_output")
    executed_dir = os.path.join(output_base, "executed")
    unmatched_agreements_dir = os.path.join(output_base, "unmatched", "agreements")
    unmatched_pages_dir = os.path.join(output_base, "unmatched", "pages")

    os.makedirs(executed_dir, exist_ok=True)
    os.makedirs(unmatched_agreements_dir, exist_ok=True)
    os.makedirs(unmatched_pages_dir, exist_ok=True)

    # Find original PDF files
    original_files = [f for f in os.listdir(originals_folder) if f.lower().endswith(".pdf")]

    if not original_files:
        emit("error", message="No PDF files found in the originals folder.")
        sys.exit(1)

    emit("progress", percent=5, message=f"Found {len(original_files)} original documents")

    # Load and analyze original documents
    original_docs = {}  # filename -> {doc, sig_pages, doc_name_variants}

    for idx, filename in enumerate(original_files):
        percent = 5 + int((idx / len(original_files)) * 20)
        emit("progress", percent=percent, message=f"Analyzing {filename}")

        try:
            filepath = os.path.join(originals_folder, filename)
            doc = fitz.open(filepath)
            sig_pages = find_signature_pages_in_document(doc)

            # Extract possible document name variants for matching
            base_name = filename[:-4] if filename.lower().endswith('.pdf') else filename
            # Remove parentheticals for matching
            clean_name = re.sub(r'\s*\([^)]*\)', '', base_name).strip()

            original_docs[filename] = {
                'doc': doc,
                'filepath': filepath,
                'sig_pages': sig_pages,
                'base_name': base_name,
                'clean_name': clean_name.upper(),
                'matched_signed_pages': {}  # page_num -> signed_page_data
            }
        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: Could not open {filename} - {str(e)}")

    if not original_docs:
        emit("error", message="Could not open any original documents.")
        sys.exit(1)

    # Load and analyze signed PDF
    emit("progress", percent=30, message="Analyzing signed document...")

    try:
        signed_doc = fitz.open(signed_pdf_path)
    except Exception as e:
        emit("error", message=f"Could not open signed PDF: {str(e)}")
        sys.exit(1)

    # Extract signed pages and their document references
    signed_pages = []  # list of {page_num, doc_name, signers, text}

    for page_num in range(len(signed_doc)):
        page = signed_doc[page_num]
        text = page.get_text()

        # Skip non-signature pages
        if "BY:" not in text.upper():
            continue

        # Detect document name
        doc_name = detect_document_name_from_footer(text)
        if not doc_name:
            doc_name = detect_document_name_from_title(text)

        signers = extract_signer_names(text)

        signed_pages.append({
            'page_num': page_num,
            'doc_name': doc_name,
            'signers': signers,
            'text': text,
            'matched': False
        })

    emit("progress", percent=40, message=f"Found {len(signed_pages)} signed pages")

    if not signed_pages:
        emit("error", message="No signature pages found in the signed PDF.")
        sys.exit(1)

    # Match signed pages to original documents
    emit("progress", percent=45, message="Matching signed pages to originals...")

    for signed_page in signed_pages:
        best_match = None
        best_score = 0
        best_orig_page = None

        for filename, orig_data in original_docs.items():
            # Try to match by document name
            if signed_page['doc_name']:
                score = fuzzy_match_score(signed_page['doc_name'], orig_data['clean_name'])

                if score > best_score and score > 0.5:  # Minimum 50% match
                    # Find the best matching signature page in this document
                    for orig_sig_page in orig_data['sig_pages']:
                        # Check if signers overlap
                        signer_overlap = signed_page['signers'] & orig_sig_page['signers']

                        # If signers match, or if we have high doc name match
                        if signer_overlap or score > 0.7:
                            # Check if this page hasn't been matched yet
                            if orig_sig_page['page_num'] not in orig_data['matched_signed_pages']:
                                best_match = filename
                                best_score = score
                                best_orig_page = orig_sig_page['page_num']
                                break

            # Also try matching by signer names if no doc name detected
            if not signed_page['doc_name'] or best_score < 0.6:
                for orig_sig_page in orig_data['sig_pages']:
                    signer_overlap = signed_page['signers'] & orig_sig_page['signers']
                    if signer_overlap and orig_sig_page['page_num'] not in orig_data['matched_signed_pages']:
                        # Use signer match as tiebreaker
                        if not best_match or len(signer_overlap) > 0:
                            best_match = filename
                            best_orig_page = orig_sig_page['page_num']
                            best_score = 0.6  # Default score for signer-only match

        # Record the match
        if best_match and best_orig_page is not None:
            original_docs[best_match]['matched_signed_pages'][best_orig_page] = signed_page
            signed_page['matched'] = True

    # Count matches
    matched_count = sum(1 for sp in signed_pages if sp['matched'])
    emit("progress", percent=55, message=f"Matched {matched_count} of {len(signed_pages)} signed pages")

    # Create executed documents
    emit("progress", percent=60, message="Creating executed documents...")

    executed_count = 0
    unmatched_agreements = []

    for idx, (filename, orig_data) in enumerate(original_docs.items()):
        percent = 60 + int((idx / len(original_docs)) * 30)

        if not orig_data['matched_signed_pages']:
            # No matches - copy to unmatched
            unmatched_agreements.append(filename)
            try:
                orig_data['doc'].save(os.path.join(unmatched_agreements_dir, filename))
            except Exception:
                pass
            orig_data['doc'].close()
            continue

        emit("progress", percent=percent, message=f"Creating executed version of {filename}")

        # Create new document with signed pages swapped in
        try:
            new_doc = fitz.open()

            for page_num in range(len(orig_data['doc'])):
                if page_num in orig_data['matched_signed_pages']:
                    # Insert signed page instead
                    signed_page_data = orig_data['matched_signed_pages'][page_num]
                    new_doc.insert_pdf(
                        signed_doc,
                        from_page=signed_page_data['page_num'],
                        to_page=signed_page_data['page_num']
                    )
                else:
                    # Copy original page
                    new_doc.insert_pdf(
                        orig_data['doc'],
                        from_page=page_num,
                        to_page=page_num
                    )

            # Save with new name
            output_name = sanitize_output_name(filename)
            new_doc.save(os.path.join(executed_dir, output_name))
            new_doc.close()
            executed_count += 1
        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: Failed to create executed version of {filename} - {str(e)}")

        orig_data['doc'].close()

    # Save unmatched signed pages
    emit("progress", percent=92, message="Saving unmatched pages...")

    unmatched_signed_pages = [sp for sp in signed_pages if not sp['matched']]

    if unmatched_signed_pages:
        for idx, sp in enumerate(unmatched_signed_pages):
            try:
                unmatched_doc = fitz.open()
                unmatched_doc.insert_pdf(signed_doc, from_page=sp['page_num'], to_page=sp['page_num'])

                # Try to name it meaningfully
                doc_name_part = sp['doc_name'][:50] if sp['doc_name'] else f"page_{sp['page_num'] + 1}"
                doc_name_part = re.sub(r'[<>:"/\\|?*]', '_', doc_name_part)  # Remove invalid chars

                unmatched_doc.save(os.path.join(unmatched_pages_dir, f"unmatched_{doc_name_part}.pdf"))
                unmatched_doc.close()
            except Exception:
                pass

    signed_doc.close()

    # Cleanup empty directories
    try:
        if not os.listdir(unmatched_agreements_dir):
            os.rmdir(unmatched_agreements_dir)
        if not os.listdir(unmatched_pages_dir):
            os.rmdir(unmatched_pages_dir)
        unmatched_dir = os.path.join(output_base, "unmatched")
        if os.path.exists(unmatched_dir) and not os.listdir(unmatched_dir):
            os.rmdir(unmatched_dir)
    except Exception:
        pass

    emit("progress", percent=100, message="Complete!")
    emit("result",
         success=True,
         outputPath=output_base,
         executedCount=executed_count,
         matchedPages=matched_count,
         totalSignedPages=len(signed_pages),
         unmatchedAgreements=len(unmatched_agreements),
         unmatchedPages=len(unmatched_signed_pages))


def main():
    if len(sys.argv) < 3:
        emit("error", message="Usage: execution_version.py <originals_folder> <signed_pdf_path>")
        sys.exit(1)

    originals_folder = sys.argv[1]
    signed_pdf_path = sys.argv[2]

    process_execution_version(originals_folder, signed_pdf_path)


if __name__ == "__main__":
    main()
