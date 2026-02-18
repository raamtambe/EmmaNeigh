#!/usr/bin/env python3
"""
EmmaNeigh - Execution Version Processor
v5.1.3: Two-folder workflow rewrite.

Takes a folder of original agreements and a folder of signed pages/schedules,
then creates executed versions by replacing signature pages in-place.

New features:
- Two folder inputs: originals folder + signed pages folder
- Higher match threshold (70%) to reduce wrong matches
- Schedule/exhibit detection and appending
- Better footer-based document matching
"""

import fitz
import os
import re
import sys
import json
import shutil
from difflib import SequenceMatcher
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def normalize_text(text):
    """Normalize text for comparison: uppercase, remove extra whitespace."""
    if not text:
        return ""
    text = text.upper()
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def extract_document_name_from_footer(text):
    """
    Extract document name from footer pattern.
    Common patterns:
    - "SIGNATURE PAGE TO [DOCUMENT NAME]"
    - "[DOCUMENT NAME] SIGNATURE PAGE"
    - "Signature Page - [Document Name]"
    """
    text_upper = text.upper()

    # Pattern 1: "SIGNATURE PAGE TO [X]"
    patterns = [
        r'SIGNATURE\s+PAGE\s+TO\s+(?:THE\s+)?(.+?)(?:\n|$)',
        r'SIGNATURE\s+PAGE\s*[-–—]\s*(.+?)(?:\n|$)',
        r'(.+?)\s+SIGNATURE\s+PAGE(?:\s|$)',
        r'COUNTERPART\s+SIGNATURE\s+PAGE\s+TO\s+(?:THE\s+)?(.+?)(?:\n|$)',
    ]

    for pattern in patterns:
        match = re.search(pattern, text_upper)
        if match:
            doc_name = match.group(1).strip()
            # Clean up - remove trailing punctuation and common suffixes
            doc_name = re.sub(r'[.\-–—]+$', '', doc_name).strip()
            doc_name = re.sub(r'\s*\(CONTINUED\)$', '', doc_name).strip()
            if len(doc_name) > 3 and not doc_name.startswith(('THIS', 'THE ', 'A ', 'AN ')):
                return doc_name

    return None


def extract_document_name_from_title(text):
    """
    Extract document name from common document types in text.
    """
    text_upper = text.upper()

    # Common M&A and finance document types
    doc_types = [
        'CREDIT AGREEMENT',
        'GUARANTEE', 'GUARANTY',
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
        'PROMISSORY NOTE',
        "OFFICER'S CERTIFICATE",
        'INCUMBENCY CERTIFICATE',
        'SECRETARY CERTIFICATE',
        'SOLVENCY CERTIFICATE',
    ]

    for doc_type in doc_types:
        if doc_type in text_upper:
            # Try to get more context
            match = re.search(rf'({doc_type}[A-Z\s]*?)(?:\n|BY:|DATED|$)', text_upper)
            if match:
                return match.group(1).strip()
            return doc_type

    return None


def is_schedule_or_exhibit(text, filename):
    """
    Check if a page/file is a schedule or exhibit (should be appended, not matched).
    """
    text_upper = text.upper() if text else ""
    filename_upper = filename.upper() if filename else ""

    schedule_patterns = [
        r'\bSCHEDULE\s*[A-Z0-9]+',
        r'\bEXHIBIT\s*[A-Z0-9]+',
        r'\bANNEX\s*[A-Z0-9]+',
        r'\bAPPENDIX\s*[A-Z0-9]+',
        r'\bATTACHMENT\s*[A-Z0-9]+',
    ]

    for pattern in schedule_patterns:
        if re.search(pattern, text_upper) or re.search(pattern, filename_upper):
            return True

    return False


def is_signature_page(page):
    """
    Check if a PDF page contains signature content.
    Looks for BY: blocks or signature-related patterns.
    """
    text = page.get_text().upper()

    # Must have actual content (not blank)
    content_text = re.sub(r'[_\s\-\=]+', '', text)
    if len(content_text) < 30:
        return False

    # Check for signature indicators
    has_by = bool(re.search(r'\bBY\s*:', text))
    has_name_label = bool(re.search(r'\bNAME\s*:', text))
    has_title_label = bool(re.search(r'\bTITLE\s*:', text))
    has_signature_page = 'SIGNATURE PAGE' in text
    has_underscore_line = bool(re.search(r'_{10,}', text))

    # Strong indicator
    if has_signature_page:
        return True

    # BY: with supporting labels
    if has_by and (has_name_label or has_title_label):
        return True

    # Underscore lines with labels
    if has_underscore_line and (has_name_label or has_title_label):
        return True

    return False


def fuzzy_match_score(s1, s2):
    """Return similarity ratio between two normalized strings (0-1)."""
    s1 = normalize_text(s1)
    s2 = normalize_text(s2)
    return SequenceMatcher(None, s1, s2).ratio()


def find_best_match(signed_doc_name, signed_text, original_docs, min_threshold=0.70):
    """
    Find the best matching original document for a signed page.

    Args:
        signed_doc_name: Document name extracted from signed page footer
        signed_text: Full text of signed page
        original_docs: Dict of {filename: {clean_name, sig_pages, etc}}
        min_threshold: Minimum match score (default 0.70 = 70%)

    Returns:
        (best_match_filename, best_score, best_orig_page_num) or (None, 0, None)
    """
    best_match = None
    best_score = 0
    best_orig_page = None

    for filename, orig_data in original_docs.items():
        # Calculate match score
        if signed_doc_name:
            # Primary: match by document name from footer
            name_score = fuzzy_match_score(signed_doc_name, orig_data['clean_name'])

            # Also try matching against document name extracted from original
            if orig_data.get('detected_name'):
                alt_score = fuzzy_match_score(signed_doc_name, orig_data['detected_name'])
                name_score = max(name_score, alt_score)
        else:
            # No footer name - try text-based matching
            name_score = 0

        if name_score >= min_threshold and name_score > best_score:
            # Find the matching signature page in original
            for orig_sig_page in orig_data.get('sig_pages', []):
                if orig_sig_page['page_num'] not in orig_data.get('matched_pages', set()):
                    best_match = filename
                    best_score = name_score
                    best_orig_page = orig_sig_page['page_num']
                    break  # Take first unmatched signature page

    return best_match, best_score, best_orig_page


def process_execution_version_two_folders(originals_folder, signed_folder, output_folder=None):
    """
    Main processing function for two-folder workflow.

    Args:
        originals_folder: Path to folder containing original agreement PDFs
        signed_folder: Path to folder containing signed pages/schedules
        output_folder: Optional output folder (defaults to originals_folder/execution_output)
    """

    # Validate inputs
    if not os.path.isdir(originals_folder):
        emit("error", message=f"Invalid originals folder: {originals_folder}")
        sys.exit(1)

    if not os.path.isdir(signed_folder):
        emit("error", message=f"Invalid signed pages folder: {signed_folder}")
        sys.exit(1)

    # Set up output directories
    if not output_folder:
        output_folder = os.path.join(originals_folder, "execution_output")

    executed_dir = os.path.join(output_folder, "executed")
    unmatched_dir = os.path.join(output_folder, "unmatched")
    schedules_dir = os.path.join(output_folder, "schedules")

    os.makedirs(executed_dir, exist_ok=True)
    os.makedirs(unmatched_dir, exist_ok=True)

    # Find all PDF files
    original_files = [f for f in os.listdir(originals_folder)
                      if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(originals_folder, f))]
    signed_files = [f for f in os.listdir(signed_folder)
                    if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(signed_folder, f))]

    if not original_files:
        emit("error", message="No PDF files found in the originals folder.")
        sys.exit(1)

    if not signed_files:
        emit("error", message="No PDF files found in the signed pages folder.")
        sys.exit(1)

    emit("progress", percent=5, message=f"Found {len(original_files)} originals, {len(signed_files)} signed pages")

    # ========== PHASE 1: Analyze Original Documents ==========
    emit("progress", percent=10, message="Analyzing original documents...")

    original_docs = {}  # filename -> {doc, filepath, sig_pages, clean_name, detected_name, matched_pages}

    for idx, filename in enumerate(original_files):
        percent = 10 + int((idx / len(original_files)) * 15)
        emit("progress", percent=percent, message=f"Analyzing {filename}")

        try:
            filepath = os.path.join(originals_folder, filename)
            doc = fitz.open(filepath)

            # Find signature pages in original
            sig_pages = []
            detected_name = None

            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text()

                if is_signature_page(page):
                    # Try to extract document name from signature page
                    page_doc_name = extract_document_name_from_footer(text)
                    if not page_doc_name:
                        page_doc_name = extract_document_name_from_title(text)

                    if page_doc_name and not detected_name:
                        detected_name = page_doc_name

                    sig_pages.append({
                        'page_num': page_num,
                        'text': text[:500]
                    })

            # Clean filename for matching
            base_name = filename[:-4] if filename.lower().endswith('.pdf') else filename
            clean_name = re.sub(r'\s*\([^)]*\)', '', base_name).strip()  # Remove parentheticals

            original_docs[filename] = {
                'doc': doc,
                'filepath': filepath,
                'sig_pages': sig_pages,
                'clean_name': clean_name.upper(),
                'detected_name': detected_name,
                'matched_pages': set(),  # Track which pages have been matched
                'signed_replacements': {},  # page_num -> signed_page_data
                'schedules': []  # Schedules to append
            }

        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: Could not open {filename} - {str(e)}")

    if not original_docs:
        emit("error", message="Could not open any original documents.")
        sys.exit(1)

    # ========== PHASE 2: Analyze Signed Pages ==========
    emit("progress", percent=30, message="Analyzing signed pages...")

    signed_pages = []  # List of {filepath, filename, doc_name, is_schedule, text, matched}
    schedules = []  # Separate list for schedules/exhibits

    for idx, filename in enumerate(signed_files):
        percent = 30 + int((idx / len(signed_files)) * 15)

        try:
            filepath = os.path.join(signed_folder, filename)
            doc = fitz.open(filepath)

            # For multi-page signed PDFs, check each page
            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text()

                # Extract document name from footer/title
                doc_name = extract_document_name_from_footer(text)
                if not doc_name:
                    doc_name = extract_document_name_from_title(text)

                # Check if this is a schedule/exhibit
                is_schedule = is_schedule_or_exhibit(text, filename)

                page_data = {
                    'filepath': filepath,
                    'filename': filename,
                    'page_num': page_num,
                    'doc_name': doc_name,
                    'is_schedule': is_schedule,
                    'text': text[:500],
                    'matched': False
                }

                if is_schedule:
                    schedules.append(page_data)
                elif is_signature_page(page):
                    signed_pages.append(page_data)
                # Skip pages that are neither signature pages nor schedules

            doc.close()

        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: Could not read {filename} - {str(e)}")

    emit("progress", percent=50, message=f"Found {len(signed_pages)} signature pages, {len(schedules)} schedules")

    # ========== PHASE 3: Match Signed Pages to Originals ==========
    emit("progress", percent=55, message="Matching signed pages to originals...")

    # First pass: Match by document name (high confidence)
    for signed_page in signed_pages:
        if signed_page['matched']:
            continue

        best_match, best_score, best_orig_page = find_best_match(
            signed_page['doc_name'],
            signed_page['text'],
            original_docs,
            min_threshold=0.70  # 70% threshold
        )

        if best_match and best_orig_page is not None:
            # Record the match
            original_docs[best_match]['matched_pages'].add(best_orig_page)
            original_docs[best_match]['signed_replacements'][best_orig_page] = signed_page
            signed_page['matched'] = True
            signed_page['matched_to'] = best_match
            signed_page['match_score'] = best_score

    # Match schedules to documents
    for schedule in schedules:
        if schedule['doc_name']:
            for filename, orig_data in original_docs.items():
                score = fuzzy_match_score(schedule['doc_name'], orig_data['clean_name'])
                if score >= 0.60:  # Lower threshold for schedules
                    orig_data['schedules'].append(schedule)
                    schedule['matched'] = True
                    schedule['matched_to'] = filename
                    break

    # Count matches
    matched_sig_count = sum(1 for sp in signed_pages if sp['matched'])
    matched_sched_count = sum(1 for s in schedules if s['matched'])

    emit("progress", percent=65, message=f"Matched {matched_sig_count}/{len(signed_pages)} signature pages, {matched_sched_count}/{len(schedules)} schedules")

    # ========== PHASE 4: Create Executed Documents ==========
    emit("progress", percent=70, message="Creating executed documents...")

    executed_count = 0
    unmatched_agreements = []

    for idx, (filename, orig_data) in enumerate(original_docs.items()):
        percent = 70 + int((idx / len(original_docs)) * 25)

        has_replacements = len(orig_data['signed_replacements']) > 0
        has_schedules = len(orig_data['schedules']) > 0

        if not has_replacements and not has_schedules:
            # No matches - this agreement is unmatched
            unmatched_agreements.append(filename)
            orig_data['doc'].close()
            continue

        emit("progress", percent=percent, message=f"Creating executed version of {filename}")

        try:
            # Create new document with replacements
            new_doc = fitz.open()

            for page_num in range(len(orig_data['doc'])):
                if page_num in orig_data['signed_replacements']:
                    # Insert signed page instead of original
                    signed_page = orig_data['signed_replacements'][page_num]
                    signed_doc = fitz.open(signed_page['filepath'])
                    new_doc.insert_pdf(
                        signed_doc,
                        from_page=signed_page['page_num'],
                        to_page=signed_page['page_num']
                    )
                    signed_doc.close()
                else:
                    # Copy original page
                    new_doc.insert_pdf(
                        orig_data['doc'],
                        from_page=page_num,
                        to_page=page_num
                    )

            # Append schedules at the end
            for schedule in orig_data['schedules']:
                sched_doc = fitz.open(schedule['filepath'])
                new_doc.insert_pdf(
                    sched_doc,
                    from_page=schedule['page_num'],
                    to_page=schedule['page_num']
                )
                sched_doc.close()

            # Save with (executed) suffix
            base_name = filename[:-4] if filename.lower().endswith('.pdf') else filename
            base_name = re.sub(r'\s*\([^)]*\)\s*$', '', base_name).strip()  # Remove existing parentheticals
            output_name = f"{base_name} (executed).pdf"
            output_path = os.path.join(executed_dir, output_name)

            new_doc.save(output_path)
            new_doc.close()
            executed_count += 1

        except Exception as e:
            emit("progress", percent=percent, message=f"Warning: Failed to create {filename} - {str(e)}")

        orig_data['doc'].close()

    # ========== PHASE 5: Save Unmatched Pages ==========
    emit("progress", percent=96, message="Saving unmatched pages...")

    unmatched_sig_pages = [sp for sp in signed_pages if not sp['matched']]
    unmatched_schedules = [s for s in schedules if not s['matched']]

    if unmatched_sig_pages or unmatched_schedules:
        for item in unmatched_sig_pages + unmatched_schedules:
            try:
                src_doc = fitz.open(item['filepath'])
                out_doc = fitz.open()
                out_doc.insert_pdf(src_doc, from_page=item['page_num'], to_page=item['page_num'])

                # Create meaningful name
                name_part = item['doc_name'][:40] if item['doc_name'] else item['filename'][:-4]
                name_part = re.sub(r'[<>:"/\\|?*]', '_', name_part)
                page_suffix = f"_p{item['page_num']+1}" if item['page_num'] > 0 else ""

                out_name = f"unmatched_{name_part}{page_suffix}.pdf"
                out_doc.save(os.path.join(unmatched_dir, out_name))
                out_doc.close()
                src_doc.close()
            except Exception:
                pass

    # Cleanup empty directories
    try:
        if os.path.exists(unmatched_dir) and not os.listdir(unmatched_dir):
            os.rmdir(unmatched_dir)
    except Exception:
        pass

    # ========== COMPLETE ==========
    emit("progress", percent=100, message="Complete!")
    emit("result",
         success=True,
         outputPath=output_folder,
         executedCount=executed_count,
         matchedPages=matched_sig_count,
         totalSignedPages=len(signed_pages),
         matchedSchedules=matched_sched_count,
         totalSchedules=len(schedules),
         unmatchedAgreements=len(unmatched_agreements),
         unmatchedPages=len(unmatched_sig_pages) + len(unmatched_schedules))


def process_legacy_single_pdf(originals_folder, signed_pdf_path):
    """
    Legacy support: Process with single signed PDF.
    Extracts pages from the PDF and processes as if they were in a folder.
    """
    import tempfile

    # Create temp folder for signed pages
    temp_signed_folder = tempfile.mkdtemp(prefix='emmaneigh_signed_')

    try:
        # Extract each page from signed PDF to individual files
        signed_doc = fitz.open(signed_pdf_path)

        for page_num in range(len(signed_doc)):
            page = signed_doc[page_num]
            text = page.get_text()

            # Get document name for naming
            doc_name = extract_document_name_from_footer(text)
            if not doc_name:
                doc_name = extract_document_name_from_title(text)

            # Create individual PDF for this page
            page_doc = fitz.open()
            page_doc.insert_pdf(signed_doc, from_page=page_num, to_page=page_num)

            name_part = doc_name[:40] if doc_name else f"page_{page_num+1}"
            name_part = re.sub(r'[<>:"/\\|?*]', '_', name_part)
            page_doc.save(os.path.join(temp_signed_folder, f"{name_part}_p{page_num+1}.pdf"))
            page_doc.close()

        signed_doc.close()

        # Process with two-folder workflow
        process_execution_version_two_folders(originals_folder, temp_signed_folder)

    finally:
        # Cleanup temp folder
        try:
            shutil.rmtree(temp_signed_folder)
        except Exception:
            pass


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: execution_version.py <originals_folder> <signed_folder|signed_pdf>")
        sys.exit(1)

    # Check for config file mode
    if sys.argv[1] == '--config':
        if len(sys.argv) < 3:
            emit("error", message="No config file provided.")
            sys.exit(1)

        config_path = sys.argv[2]
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)

            originals = config.get('originals_folder') or config.get('files')
            signed = config.get('signed_folder') or config.get('signed_pdf')
            output = config.get('output_folder')

            if not originals or not signed:
                emit("error", message="Config must have originals and signed paths.")
                sys.exit(1)

            # Determine if signed is folder or file
            if os.path.isdir(signed):
                # Handle file list for originals
                if isinstance(originals, list):
                    # Create temp folder with symlinks/copies
                    import tempfile
                    temp_orig_folder = tempfile.mkdtemp(prefix='emmaneigh_orig_')
                    for f in originals:
                        if os.path.isfile(f):
                            shutil.copy(f, temp_orig_folder)
                    process_execution_version_two_folders(temp_orig_folder, signed, output)
                    shutil.rmtree(temp_orig_folder, ignore_errors=True)
                else:
                    process_execution_version_two_folders(originals, signed, output)
            else:
                # Legacy: single signed PDF
                if isinstance(originals, list):
                    import tempfile
                    temp_orig_folder = tempfile.mkdtemp(prefix='emmaneigh_orig_')
                    for f in originals:
                        if os.path.isfile(f):
                            shutil.copy(f, temp_orig_folder)
                    process_legacy_single_pdf(temp_orig_folder, signed)
                    shutil.rmtree(temp_orig_folder, ignore_errors=True)
                else:
                    process_legacy_single_pdf(originals, signed)

        except Exception as e:
            import traceback
            emit("error", message=f"Config error: {str(e)}\n{traceback.format_exc()}")
            sys.exit(1)
    else:
        # Direct CLI: originals_folder signed_folder|signed_pdf
        if len(sys.argv) < 3:
            emit("error", message="Usage: execution_version.py <originals_folder> <signed_folder|signed_pdf>")
            sys.exit(1)

        originals_folder = sys.argv[1]
        signed_path = sys.argv[2]

        if os.path.isdir(signed_path):
            process_execution_version_two_folders(originals_folder, signed_path)
        elif os.path.isfile(signed_path):
            process_legacy_single_pdf(originals_folder, signed_path)
        else:
            emit("error", message=f"Signed path not found: {signed_path}")
            sys.exit(1)


if __name__ == "__main__":
    main()
