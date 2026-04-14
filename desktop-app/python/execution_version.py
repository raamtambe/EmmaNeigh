#!/usr/bin/env python3
"""
EmmaNeigh - Execution Version Processor
v5.4.17: Page-level execution workflow for bulk signature packets.

Takes original agreements plus a signed packet/folder, then creates executed
versions by matching signed pages back to the exact original signature page.

Key behaviors:
- accepts many originals + one large DocuSign packet or a folder of signed PDFs
- matches at the page level, not just the document-name level
- uses footer/title hints, signature-block text, and visual page fingerprints
- greedily assigns the best unique matches across the full packet
- appends matched schedules/exhibits and writes a JSON match report
"""

import fitz
import json
import os
import re
import shutil
import sys
import tempfile
from difflib import SequenceMatcher

MATCH_SCORE_THRESHOLD = 0.58
SCHEDULE_SCORE_THRESHOLD = 0.55
VISUAL_STRONG_THRESHOLD = 0.84
MAX_MATCH_TEXT_LINES = 28
DOC_TYPE_HINTS = [
    'AGREEMENT',
    'AMENDMENT',
    'CERTIFICATE',
    'CONSENT',
    'GUARANT',
    'GUARANTEE',
    'GUARANTY',
    'INCUMBENCY',
    'INDENTURE',
    'INTERCREDITOR',
    'JOINDER',
    'LOAN',
    'NOTE',
    'OFFICER',
    'PLEDGE',
    'PROMISSORY',
    'SECURITY',
    'SECRETARY',
    'SIGNATURE PAGE',
    'SOLVENCY',
    'SUBORDINATION',
]
NOISE_PATTERNS = [
    r'^DOCUSIGNED\s+BY\b',
    r'^ENVELOPE\s+ID\b',
    r'^SIGNED\s+BY\b',
    r'^SENT\s+BY\b',
    r'^DOCUSIGN\s+ENVELOPE\b',
    r'^DOCUSIGN\b',
    r'^PLEASE\s+REVIEW\b',
    r'^COMPLETED\s+BY\s+DOCUSIGN\b',
    r'^PLEASE\s+SIGN\b',
    r'^SIGN\s+HERE\b',
    r'^CLICK\s+TO\s+SIGN\b',
    r'^PAGE\s+\d+\s+OF\s+\d+\b',
    r'^\d+\s*/\s*\d+\s*$',
    r'^\d{1,2}/\d{1,2}/\d{2,4}\b',
    r'^[0-9A-F]{8,}$',
]
SCHEDULE_PATTERNS = [
    r'\bSCHEDULE\s*[A-Z0-9]+',
    r'\bEXHIBIT\s*[A-Z0-9]+',
    r'\bANNEX\s*[A-Z0-9]+',
    r'\bAPPENDIX\s*[A-Z0-9]+',
    r'\bATTACHMENT\s*[A-Z0-9]+',
]


def emit(msg_type, **kwargs):
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)



def normalize_text(text):
    if not text:
        return ""
    text = str(text).replace('\x00', ' ')
    text = text.upper()
    text = re.sub(r'\s+', ' ', text).strip()
    return text



def normalize_line(text):
    line = normalize_text(text)
    if not line:
        return ""
    line = re.sub(r'^[\s\-–—•·\*]+', '', line)
    line = re.sub(r'[\s\-–—•·\*]+$', '', line)
    return line.strip()



def clean_filename_stem(filename):
    base_name = filename[:-4] if filename.lower().endswith('.pdf') else filename
    base_name = re.sub(r'\s*\([^)]*\)', '', base_name).strip()
    return normalize_text(base_name)



def safe_filename(value, fallback='document'):
    cleaned = re.sub(r'[<>:"/\\|?*]', '_', StringLike(value))
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(' ._')
    return cleaned or fallback



def StringLike(value):
    return str(value or '')



def extract_document_name_from_footer(text):
    text_upper = normalize_text(text)
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
            doc_name = re.sub(r'[.\-–—]+$', '', doc_name).strip()
            doc_name = re.sub(r'\s*\(CONTINUED\)$', '', doc_name).strip()
            if len(doc_name) > 3 and not doc_name.startswith(('THIS', 'THE ', 'A ', 'AN ')):
                return doc_name
    return None



def extract_document_name_from_title(text):
    text_upper = normalize_text(text)
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
            match = re.search(rf'({doc_type}[A-Z\s]*?)(?:\n|BY:|DATED|$)', text_upper)
            if match:
                return match.group(1).strip()
            return doc_type
    return None



def is_schedule_or_exhibit(text, filename):
    text_upper = normalize_text(text)
    filename_upper = normalize_text(filename)
    return any(re.search(pattern, text_upper) or re.search(pattern, filename_upper) for pattern in SCHEDULE_PATTERNS)



def is_signature_page(page):
    text = normalize_text(page.get_text())
    content_text = re.sub(r'[_\s\-\=]+', '', text)
    if len(content_text) < 30:
        return False

    has_by = bool(re.search(r'\bBY\s*:', text))
    has_name_label = bool(re.search(r'\bNAME\s*:', text))
    has_title_label = bool(re.search(r'\bTITLE\s*:', text))
    has_date_label = bool(re.search(r'\bDATE\s*:', text))
    has_signature_page = 'SIGNATURE PAGE' in text
    has_underscore_line = bool(re.search(r'_{8,}', text))
    has_signatory_title = bool(re.search(r'\bITS\s*:', text))

    if has_signature_page:
        return True
    if has_by and (has_name_label or has_title_label or has_date_label or has_signatory_title):
        return True
    if has_underscore_line and (has_name_label or has_title_label or has_date_label):
        return True
    return False



def fuzzy_match_score(left, right):
    a = normalize_text(left)
    b = normalize_text(right)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()



def extract_meaningful_lines(text):
    lines = []
    seen = set()
    for raw_line in StringLike(text).splitlines():
        line = normalize_line(raw_line)
        if not line:
            continue
        if len(line) < 2:
            continue
        if any(re.search(pattern, line) for pattern in NOISE_PATTERNS):
            continue
        if re.fullmatch(r'[_\-–—=\s]{6,}', line):
            continue
        if line not in seen:
            seen.add(line)
            lines.append(line)
    return lines



def compress_lines(lines, max_lines=12, max_chars=900):
    parts = []
    total_chars = 0
    for line in lines[:max_lines]:
        projected = total_chars + len(line) + (3 if parts else 0)
        if projected > max_chars:
            break
        parts.append(line)
        total_chars = projected
    return ' | '.join(parts)



def build_match_line_sample(lines):
    if not lines:
        return ''
    head = lines[:14]
    tail = lines[-14:] if len(lines) > 14 else []
    combined = []
    seen = set()
    for line in head + tail:
        if line and line not in seen:
            seen.add(line)
            combined.append(line)
    return compress_lines(combined, max_lines=MAX_MATCH_TEXT_LINES, max_chars=1400)



def extract_anchor_lines(lines):
    anchors = []
    for index, line in enumerate(lines):
        keep = (
            'SIGNATURE PAGE' in line or
            any(token in line for token in DOC_TYPE_HINTS) or
            line.startswith(('BY:', 'NAME:', 'TITLE:', 'DATE:', 'ITS:'))
        )
        if keep:
            window = lines[max(0, index - 1):min(len(lines), index + 2)]
            for item in window:
                if item not in anchors:
                    anchors.append(item)
    if not anchors:
        for item in (lines[:8] + lines[-8:]):
            if len(item) >= 8 and item not in anchors:
                anchors.append(item)
    return anchors[:18]



def build_visual_fingerprint(page, grid_width=24, grid_height=32):
    rect = page.rect
    return build_region_fingerprint(
        page,
        rect.x0 + rect.width * 0.05,
        rect.y0 + rect.height * 0.03,
        rect.x1 - rect.width * 0.05,
        rect.y0 + rect.height * 0.82,
        grid_width=grid_width,
        grid_height=grid_height,
    )



def build_region_fingerprint(page, x0, y0, x1, y1, grid_width=24, grid_height=32):
    clip = fitz.Rect(
        max(page.rect.x0, x0),
        max(page.rect.y0, y0),
        min(page.rect.x1, x1),
        min(page.rect.y1, y1),
    )
    if clip.width <= 1 or clip.height <= 1:
        return []
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(0.32, 0.32), colorspace=fitz.csGRAY, alpha=False, clip=clip)
    except Exception:
        return []

    if pix.width < 4 or pix.height < 4:
        return []

    samples = pix.samples
    fingerprint = []
    for row in range(grid_height):
        y0 = int(row * pix.height / grid_height)
        y1 = max(y0 + 1, int((row + 1) * pix.height / grid_height))
        for col in range(grid_width):
            x0 = int(col * pix.width / grid_width)
            x1 = max(x0 + 1, int((col + 1) * pix.width / grid_width))
            total = 0
            count = 0
            for y in range(y0, y1):
                row_offset = y * pix.width
                for x in range(x0, x1):
                    total += samples[row_offset + x]
                    count += 1
            fingerprint.append(round(total / count, 2) if count else 255.0)
    return fingerprint



def visual_similarity(left_fp, right_fp):
    if not left_fp or not right_fp or len(left_fp) != len(right_fp):
        return 0.0
    mean_diff = sum(abs(a - b) for a, b in zip(left_fp, right_fp)) / len(left_fp)
    similarity = 1.0 - (mean_diff / 255.0)
    return max(0.0, min(1.0, similarity))



def line_overlap_score(left_lines, right_lines):
    left = {line for line in left_lines if len(line) >= 4}
    right = {line for line in right_lines if len(line) >= 4}
    if not left or not right:
        return 0.0
    union = left | right
    if not union:
        return 0.0
    return len(left & right) / len(union)



def extract_footer_lines(lines):
    return [line for line in lines[-8:] if len(line) >= 4]



def extract_signature_zone_lines(lines):
    signature_lines = []
    for index, line in enumerate(lines):
        if re.search(r'\b(SIGNATURE PAGE|BY:|NAME:|TITLE:|ITS:|AUTHORIZED SIGNATORY|DULY AUTHORIZED)\b', line):
            window = lines[index:min(len(lines), index + 4)]
            for item in window:
                if item and item not in signature_lines:
                    signature_lines.append(item)

    if not signature_lines:
        signature_lines = [line for line in lines[-16:] if len(line) >= 4]
    return signature_lines[:16]



def build_page_features(page, filename, page_num):
    text = page.get_text('text') or ''
    lines = extract_meaningful_lines(text)
    footer_name = extract_document_name_from_footer(text)
    title_name = extract_document_name_from_title(text)
    doc_name = footer_name or title_name
    header_lines = lines[:12]
    tail_lines = lines[-18:] if lines else []
    anchor_lines = extract_anchor_lines(lines)
    footer_lines = extract_footer_lines(lines)
    signature_lines = extract_signature_zone_lines(lines)
    rect = page.rect

    return {
        'filepath': '',
        'filename': filename,
        'page_num': page_num,
        'text': text,
        'lines': lines,
        'doc_name': doc_name,
        'footer_name': footer_name,
        'title_name': title_name,
        'header_text': compress_lines(header_lines, max_lines=12, max_chars=800),
        'tail_text': compress_lines(tail_lines, max_lines=18, max_chars=1000),
        'footer_text': compress_lines(footer_lines, max_lines=8, max_chars=500),
        'signature_block_text': compress_lines(signature_lines, max_lines=14, max_chars=900),
        'match_text': build_match_line_sample(lines),
        'anchor_lines': anchor_lines,
        'visual_fp': build_visual_fingerprint(page),
        'signature_visual_fp': build_region_fingerprint(
            page,
            rect.x0 + rect.width * 0.05,
            rect.y0 + rect.height * 0.48,
            rect.x1 - rect.width * 0.05,
            rect.y0 + rect.height * 0.92,
            grid_width=18,
            grid_height=20,
        ),
        'footer_visual_fp': build_region_fingerprint(
            page,
            rect.x0 + rect.width * 0.05,
            rect.y0 + rect.height * 0.80,
            rect.x1 - rect.width * 0.05,
            rect.y0 + rect.height * 0.97,
            grid_width=18,
            grid_height=10,
        ),
        'matched': False,
    }



def score_signature_page_match(signed_page, original_page, original_doc):
    signed_doc_name = signed_page.get('doc_name') or ''
    original_name_candidates = [
        original_doc.get('clean_name', ''),
        original_doc.get('detected_name', ''),
        original_page.get('doc_name', ''),
        original_page.get('footer_name', ''),
        original_page.get('title_name', ''),
    ]
    doc_name_score = max(
        [fuzzy_match_score(signed_doc_name, candidate) for candidate in original_name_candidates if signed_doc_name and candidate] or [0.0]
    )
    text_score = fuzzy_match_score(signed_page.get('match_text', ''), original_page.get('match_text', ''))
    header_score = fuzzy_match_score(signed_page.get('header_text', ''), original_page.get('header_text', ''))
    tail_score = fuzzy_match_score(signed_page.get('tail_text', ''), original_page.get('tail_text', ''))
    footer_score = fuzzy_match_score(signed_page.get('footer_text', ''), original_page.get('footer_text', ''))
    signature_block_score = fuzzy_match_score(signed_page.get('signature_block_text', ''), original_page.get('signature_block_text', ''))
    anchor_score = line_overlap_score(signed_page.get('anchor_lines', []), original_page.get('anchor_lines', []))
    visual_score = visual_similarity(signed_page.get('visual_fp', []), original_page.get('visual_fp', []))
    signature_visual_score = visual_similarity(signed_page.get('signature_visual_fp', []), original_page.get('signature_visual_fp', []))
    footer_visual_score = visual_similarity(signed_page.get('footer_visual_fp', []), original_page.get('footer_visual_fp', []))

    weighted_components = []
    if doc_name_score > 0:
        weighted_components.append((0.24, doc_name_score))
    if text_score > 0:
        weighted_components.append((0.16, text_score))
    if header_score > 0:
        weighted_components.append((0.07, header_score))
    if tail_score > 0:
        weighted_components.append((0.12, tail_score))
    if footer_score > 0:
        weighted_components.append((0.11, footer_score))
    if signature_block_score > 0:
        weighted_components.append((0.12, signature_block_score))
    if anchor_score > 0:
        weighted_components.append((0.08, anchor_score))
    if visual_score > 0:
        weighted_components.append((0.04, visual_score))
    if signature_visual_score > 0:
        weighted_components.append((0.04, signature_visual_score))
    if footer_visual_score > 0:
        weighted_components.append((0.02, footer_visual_score))

    if not weighted_components:
        return 0.0, {
            'doc_name_score': 0.0,
            'text_score': 0.0,
            'header_score': 0.0,
            'tail_score': 0.0,
            'footer_score': 0.0,
            'signature_block_score': 0.0,
            'anchor_score': 0.0,
            'visual_score': 0.0,
            'signature_visual_score': 0.0,
            'footer_visual_score': 0.0,
        }, False

    total_weight = sum(weight for weight, _ in weighted_components)
    score = sum(weight * value for weight, value in weighted_components) / total_weight

    if doc_name_score >= 0.80 and visual_score >= 0.88:
        score = min(0.99, score + 0.05)
    if tail_score >= 0.85:
        score = min(0.99, score + 0.03)
    if signature_block_score >= 0.72 and signature_visual_score >= 0.78:
        score = min(0.99, score + 0.04)

    plausible = (
        doc_name_score >= 0.45 or
        text_score >= 0.52 or
        tail_score >= 0.55 or
        footer_score >= 0.62 or
        signature_block_score >= 0.55 or
        anchor_score >= 0.35 or
        visual_score >= VISUAL_STRONG_THRESHOLD or
        signature_visual_score >= 0.82
    )

    if signed_doc_name and doc_name_score < 0.25 and text_score < 0.60 and tail_score < 0.60 and footer_score < 0.55 and signature_block_score < 0.48 and visual_score < 0.92 and signature_visual_score < 0.84:
        plausible = False

    details = {
        'doc_name_score': round(doc_name_score, 4),
        'text_score': round(text_score, 4),
        'header_score': round(header_score, 4),
        'tail_score': round(tail_score, 4),
        'footer_score': round(footer_score, 4),
        'signature_block_score': round(signature_block_score, 4),
        'anchor_score': round(anchor_score, 4),
        'visual_score': round(visual_score, 4),
        'signature_visual_score': round(signature_visual_score, 4),
        'footer_visual_score': round(footer_visual_score, 4),
    }
    return round(score, 4), details, plausible



def assign_signature_pages(signed_pages, original_docs):
    candidates = []
    for signed_page in signed_pages:
        for filename, original_doc in original_docs.items():
            for original_page in original_doc.get('sig_pages', []):
                score, details, plausible = score_signature_page_match(signed_page, original_page, original_doc)
                if not plausible or score < MATCH_SCORE_THRESHOLD:
                    continue
                candidates.append({
                    'signed_page': signed_page,
                    'original_filename': filename,
                    'original_page': original_page,
                    'score': score,
                    'details': details,
                })

    candidates.sort(
        key=lambda item: (
            item['score'],
            item['details']['doc_name_score'],
            item['details']['tail_score'],
            item['details']['visual_score'],
            item['details']['text_score'],
        ),
        reverse=True,
    )

    used_signed = set()
    used_original = set()
    matches = []

    for candidate in candidates:
        signed_key = (candidate['signed_page']['filepath'], candidate['signed_page']['page_num'])
        original_key = (candidate['original_filename'], candidate['original_page']['page_num'])
        if signed_key in used_signed or original_key in used_original:
            continue
        used_signed.add(signed_key)
        used_original.add(original_key)
        matches.append(candidate)

    return matches



def match_schedule_to_document(schedule, original_docs):
    best_filename = None
    best_score = 0.0
    schedule_name = schedule.get('doc_name') or ''
    if not schedule_name:
        return None, 0.0

    for filename, original_doc in original_docs.items():
        candidates = [original_doc.get('clean_name', ''), original_doc.get('detected_name', '')]
        score = max([fuzzy_match_score(schedule_name, candidate) for candidate in candidates if candidate] or [0.0])
        if score > best_score:
            best_score = score
            best_filename = filename

    if best_score >= SCHEDULE_SCORE_THRESHOLD:
        return best_filename, best_score
    return None, best_score



def build_unmatched_page_name(item, prefix='unmatched'):
    name_part = item.get('doc_name') or clean_filename_stem(item.get('filename', '')) or 'page'
    name_part = safe_filename(name_part, 'page')[:50]
    page_suffix = f"_p{int(item.get('page_num', 0)) + 1}"
    return f"{prefix}_{name_part}{page_suffix}.pdf"



def write_match_report(output_folder, original_docs, signed_pages, schedules, unmatched_agreements):
    report = {
        'matched_signature_pages': [],
        'unmatched_signature_pages': [],
        'matched_schedules': [],
        'unmatched_schedules': [],
        'documents': [],
        'unmatched_agreements': unmatched_agreements,
    }

    for filename, original_doc in original_docs.items():
        document_entry = {
            'filename': filename,
            'clean_name': original_doc.get('clean_name', ''),
            'detected_name': original_doc.get('detected_name', ''),
            'signature_pages_detected': len(original_doc.get('sig_pages', [])),
            'matched_signature_pages': [],
            'appended_schedules': [],
        }

        for page_num, detail in sorted(original_doc.get('match_details', {}).items()):
            entry = {
                'original_filename': filename,
                'original_page_num': page_num + 1,
                'signed_filename': detail['signed_page']['filename'],
                'signed_page_num': detail['signed_page']['page_num'] + 1,
                'score': detail['score'],
                'details': detail['details'],
                'signed_doc_name': detail['signed_page'].get('doc_name') or '',
            }
            report['matched_signature_pages'].append(entry)
            document_entry['matched_signature_pages'].append(entry)

        for schedule in original_doc.get('schedules', []):
            entry = {
                'original_filename': filename,
                'schedule_filename': schedule['filename'],
                'schedule_page_num': schedule['page_num'] + 1,
                'schedule_doc_name': schedule.get('doc_name') or '',
            }
            report['matched_schedules'].append(entry)
            document_entry['appended_schedules'].append(entry)

        report['documents'].append(document_entry)

    for signed_page in signed_pages:
        if not signed_page.get('matched'):
            report['unmatched_signature_pages'].append({
                'filename': signed_page['filename'],
                'page_num': signed_page['page_num'] + 1,
                'doc_name': signed_page.get('doc_name') or '',
            })

    for schedule in schedules:
        if not schedule.get('matched'):
            report['unmatched_schedules'].append({
                'filename': schedule['filename'],
                'page_num': schedule['page_num'] + 1,
                'doc_name': schedule.get('doc_name') or '',
            })

    report_path = os.path.join(output_folder, 'execution_match_report.json')
    with open(report_path, 'w', encoding='utf-8') as handle:
        json.dump(report, handle, indent=2)
    return report_path



def process_execution_version_two_folders(originals_folder, signed_folder, output_folder=None):
    if not os.path.isdir(originals_folder):
        emit('error', message=f'Invalid originals folder: {originals_folder}')
        sys.exit(1)
    if not os.path.isdir(signed_folder):
        emit('error', message=f'Invalid signed pages folder: {signed_folder}')
        sys.exit(1)

    if not output_folder:
        output_folder = os.path.join(originals_folder, 'execution_output')

    executed_dir = os.path.join(output_folder, 'executed')
    unmatched_dir = os.path.join(output_folder, 'unmatched')
    os.makedirs(executed_dir, exist_ok=True)
    os.makedirs(unmatched_dir, exist_ok=True)

    original_files = sorted(
        f for f in os.listdir(originals_folder)
        if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(originals_folder, f))
    )
    signed_files = sorted(
        f for f in os.listdir(signed_folder)
        if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(signed_folder, f))
    )

    if not original_files:
        emit('error', message='No PDF files found in the originals folder.')
        sys.exit(1)
    if not signed_files:
        emit('error', message='No PDF files found in the signed pages folder.')
        sys.exit(1)

    emit('progress', percent=5, message=f'Found {len(original_files)} originals and {len(signed_files)} signed PDF files')
    emit('progress', percent=10, message='Analyzing original agreements...')

    original_docs = {}
    total_original_sig_pages = 0

    for index, filename in enumerate(original_files):
        percent = 10 + int((index / max(len(original_files), 1)) * 18)
        emit('progress', percent=percent, message=f'Analyzing original {filename}')
        filepath = os.path.join(originals_folder, filename)
        try:
            document = fitz.open(filepath)
        except Exception as error:
            emit('progress', percent=percent, message=f'Warning: Could not open {filename} - {error}')
            continue

        sig_pages = []
        detected_name = None
        for page_num in range(len(document)):
            page = document[page_num]
            if not is_signature_page(page):
                continue
            features = build_page_features(page, filename, page_num)
            features['filepath'] = filepath
            sig_pages.append(features)
            if features.get('doc_name') and not detected_name:
                detected_name = features['doc_name']

        total_original_sig_pages += len(sig_pages)
        original_docs[filename] = {
            'doc': document,
            'filepath': filepath,
            'sig_pages': sig_pages,
            'clean_name': clean_filename_stem(filename),
            'detected_name': detected_name,
            'matched_pages': set(),
            'signed_replacements': {},
            'match_details': {},
            'schedules': [],
        }

    if not original_docs:
        emit('error', message='Could not open any original documents.')
        sys.exit(1)

    emit('progress', percent=30, message='Analyzing signed packet pages...')

    signed_pages = []
    schedules = []
    for index, filename in enumerate(signed_files):
        percent = 30 + int((index / max(len(signed_files), 1)) * 20)
        filepath = os.path.join(signed_folder, filename)
        try:
            document = fitz.open(filepath)
        except Exception as error:
            emit('progress', percent=percent, message=f'Warning: Could not open {filename} - {error}')
            continue

        for page_num in range(len(document)):
            page = document[page_num]
            text = page.get_text('text') or ''
            features = build_page_features(page, filename, page_num)
            features['filepath'] = filepath
            if is_schedule_or_exhibit(text, filename):
                schedules.append(features)
            elif is_signature_page(page):
                signed_pages.append(features)
        document.close()

    emit(
        'progress',
        percent=52,
        message=f'Found {len(signed_pages)} signed page(s), {len(schedules)} schedule/exhibit page(s), and {total_original_sig_pages} original signature page candidate(s)',
    )

    emit('progress', percent=58, message='Matching signed pages back to original agreements...')
    signature_matches = assign_signature_pages(signed_pages, original_docs)
    for match in signature_matches:
        signed_page = match['signed_page']
        original_filename = match['original_filename']
        original_page_num = match['original_page']['page_num']
        signed_page['matched'] = True
        signed_page['matched_to'] = original_filename
        signed_page['match_score'] = match['score']
        original_docs[original_filename]['matched_pages'].add(original_page_num)
        original_docs[original_filename]['signed_replacements'][original_page_num] = signed_page
        original_docs[original_filename]['match_details'][original_page_num] = match

    for schedule in schedules:
        matched_filename, schedule_score = match_schedule_to_document(schedule, original_docs)
        if matched_filename:
            schedule['matched'] = True
            schedule['matched_to'] = matched_filename
            schedule['match_score'] = round(schedule_score, 4)
            original_docs[matched_filename]['schedules'].append(schedule)

    matched_sig_count = sum(1 for item in signed_pages if item.get('matched'))
    matched_sched_count = sum(1 for item in schedules if item.get('matched'))
    emit(
        'progress',
        percent=68,
        message=f'Matched {matched_sig_count}/{len(signed_pages)} signed page(s) and {matched_sched_count}/{len(schedules)} schedule/exhibit page(s)',
    )

    emit('progress', percent=72, message='Creating executed versions...')
    executed_count = 0
    unmatched_agreements = []

    for index, (filename, original_doc) in enumerate(original_docs.items()):
        percent = 72 + int((index / max(len(original_docs), 1)) * 22)
        has_replacements = bool(original_doc['signed_replacements'])
        has_schedules = bool(original_doc['schedules'])
        if not has_replacements and not has_schedules:
            unmatched_agreements.append(filename)
            original_doc['doc'].close()
            continue

        emit('progress', percent=percent, message=f'Creating executed version of {filename}')
        try:
            new_doc = fitz.open()
            for page_num in range(len(original_doc['doc'])):
                signed_page = original_doc['signed_replacements'].get(page_num)
                if signed_page:
                    signed_doc = fitz.open(signed_page['filepath'])
                    new_doc.insert_pdf(signed_doc, from_page=signed_page['page_num'], to_page=signed_page['page_num'])
                    signed_doc.close()
                else:
                    new_doc.insert_pdf(original_doc['doc'], from_page=page_num, to_page=page_num)

            for schedule in original_doc['schedules']:
                schedule_doc = fitz.open(schedule['filepath'])
                new_doc.insert_pdf(schedule_doc, from_page=schedule['page_num'], to_page=schedule['page_num'])
                schedule_doc.close()

            output_name = f"{safe_filename(clean_filename_stem(filename), 'agreement')} (executed).pdf"
            output_path = os.path.join(executed_dir, output_name)
            new_doc.save(output_path)
            new_doc.close()
            executed_count += 1
        except Exception as error:
            emit('progress', percent=percent, message=f'Warning: Failed to create {filename} - {error}')
        finally:
            original_doc['doc'].close()

    emit('progress', percent=96, message='Saving unmatched pages and writing match report...')

    unmatched_sig_pages = [item for item in signed_pages if not item.get('matched')]
    unmatched_schedules = [item for item in schedules if not item.get('matched')]

    for item in unmatched_sig_pages + unmatched_schedules:
        try:
            source_doc = fitz.open(item['filepath'])
            output_doc = fitz.open()
            output_doc.insert_pdf(source_doc, from_page=item['page_num'], to_page=item['page_num'])
            output_doc.save(os.path.join(unmatched_dir, build_unmatched_page_name(item)))
            output_doc.close()
            source_doc.close()
        except Exception:
            pass

    try:
        if os.path.exists(unmatched_dir) and not os.listdir(unmatched_dir):
            os.rmdir(unmatched_dir)
    except Exception:
        pass

    matched_documents = sum(1 for original_doc in original_docs.values() if original_doc['signed_replacements'])
    report_path = write_match_report(output_folder, original_docs, signed_pages, schedules, unmatched_agreements)

    emit('progress', percent=100, message='Complete!')
    emit(
        'result',
        success=True,
        outputPath=output_folder,
        reportPath=report_path,
        executedCount=executed_count,
        matchedDocuments=matched_documents,
        totalOriginalDocuments=len(original_docs),
        matchedPages=matched_sig_count,
        totalSignedPages=len(signed_pages),
        matchedSchedules=matched_sched_count,
        totalSchedules=len(schedules),
        unmatchedAgreements=len(unmatched_agreements),
        unmatchedAgreementNames=unmatched_agreements,
        unmatchedPages=len(unmatched_sig_pages) + len(unmatched_schedules),
    )



def process_legacy_single_pdf(originals_folder, signed_pdf_path, output_folder=None):
    temp_signed_folder = tempfile.mkdtemp(prefix='emmaneigh_signed_')
    try:
        signed_doc = fitz.open(signed_pdf_path)
        for page_num in range(len(signed_doc)):
            page = signed_doc[page_num]
            text = page.get_text('text') or ''
            doc_name = extract_document_name_from_footer(text) or extract_document_name_from_title(text)
            page_doc = fitz.open()
            page_doc.insert_pdf(signed_doc, from_page=page_num, to_page=page_num)
            name_part = safe_filename(doc_name or f'page_{page_num + 1}', f'page_{page_num + 1}')[:50]
            page_doc.save(os.path.join(temp_signed_folder, f'{name_part}_p{page_num + 1}.pdf'))
            page_doc.close()
        signed_doc.close()
        process_execution_version_two_folders(originals_folder, temp_signed_folder, output_folder)
    finally:
        try:
            shutil.rmtree(temp_signed_folder)
        except Exception:
            pass



def copy_pdf_files_to_temp_folder(files, prefix):
    temp_folder = tempfile.mkdtemp(prefix=prefix)
    for index, item in enumerate(files):
        if not os.path.isfile(item):
            continue

        original_name = os.path.basename(item)
        stem, ext = os.path.splitext(original_name)
        stem = safe_filename(stem or f'document_{index + 1}', f'document_{index + 1}')
        ext = ext or '.pdf'
        candidate_name = f'{stem}{ext}'
        destination = os.path.join(temp_folder, candidate_name)
        suffix = 2

        while os.path.exists(destination):
            destination = os.path.join(temp_folder, f'{stem}_{suffix}{ext}')
            suffix += 1

        shutil.copy(item, destination)
    return temp_folder



def main():
    if len(sys.argv) < 2:
        emit('error', message='Usage: execution_version.py <originals_folder> <signed_folder|signed_pdf>')
        sys.exit(1)

    if sys.argv[1] == '--config':
        if len(sys.argv) < 3:
            emit('error', message='No config file provided.')
            sys.exit(1)

        config_path = sys.argv[2]
        try:
            with open(config_path, 'r', encoding='utf-8') as handle:
                config = json.load(handle)

            originals = config.get('originals_folder') or config.get('files')
            signed = config.get('signed_folder') or config.get('signed_pdf') or config.get('signed_files')
            output = config.get('output_folder')

            if not originals or not signed:
                emit('error', message='Config must have originals and signed paths.')
                sys.exit(1)

            if not output and (isinstance(originals, list) or isinstance(signed, list)):
                output = tempfile.mkdtemp(prefix='emmaneigh_exec_output_')

            temp_folders = []
            originals_input = originals
            signed_input = signed

            try:
                if isinstance(originals, list):
                    originals_input = copy_pdf_files_to_temp_folder(originals, 'emmaneigh_orig_')
                    temp_folders.append(originals_input)

                if isinstance(signed, list):
                    signed_input = copy_pdf_files_to_temp_folder(signed, 'emmaneigh_signed_')
                    temp_folders.append(signed_input)

                if os.path.isdir(signed_input):
                    process_execution_version_two_folders(originals_input, signed_input, output)
                else:
                    process_legacy_single_pdf(originals_input, signed_input, output)
            finally:
                for temp_folder in temp_folders:
                    shutil.rmtree(temp_folder, ignore_errors=True)
        except Exception as error:
            import traceback
            emit('error', message=f'Config error: {error}\n{traceback.format_exc()}')
            sys.exit(1)
    else:
        if len(sys.argv) < 3:
            emit('error', message='Usage: execution_version.py <originals_folder> <signed_folder|signed_pdf>')
            sys.exit(1)

        originals_folder = sys.argv[1]
        signed_path = sys.argv[2]
        if os.path.isdir(signed_path):
            process_execution_version_two_folders(originals_folder, signed_path)
        elif os.path.isfile(signed_path):
            process_legacy_single_pdf(originals_folder, signed_path)
        else:
            emit('error', message=f'Signed path not found: {signed_path}')
            sys.exit(1)


if __name__ == '__main__':
    main()
