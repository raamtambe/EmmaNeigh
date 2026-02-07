#!/usr/bin/env python3
"""
EmmaNeigh - Document Redline
Compares two documents and creates a redlined output showing differences.
- Deleted text: Red with strikethrough
- Added text: Blue
- Moved text: Green (detected when same text appears in different locations)
"""

import os
import sys
import json
import re
import difflib
from collections import defaultdict
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def get_paragraph_texts(docx_path):
    """Extract all paragraph texts from a docx file."""
    doc = Document(docx_path)
    return [p.text for p in doc.paragraphs if p.text.strip()]


def detect_moved_paragraphs(doc1_paras, doc2_paras):
    """
    Detect paragraphs that were moved (exist in both but at different positions).
    Returns sets of indices that represent moved content.
    """
    # Find paragraphs that exist in both but at different positions
    doc1_set = {(i, text) for i, text in enumerate(doc1_paras)}
    doc2_set = {(i, text) for i, text in enumerate(doc2_paras)}

    # Get texts that appear in both documents
    doc1_texts = {text for _, text in doc1_set}
    doc2_texts = {text for _, text in doc2_set}
    common_texts = doc1_texts & doc2_texts

    # Find paragraphs where the same text exists but at different indices
    moved_in_doc1 = set()
    moved_in_doc2 = set()

    for text in common_texts:
        doc1_indices = [i for i, t in enumerate(doc1_paras) if t == text]
        doc2_indices = [i for i, t in enumerate(doc2_paras) if t == text]

        # If indices differ, it's potentially moved
        if doc1_indices != doc2_indices:
            # Simple heuristic: if relative position changed significantly
            for i1 in doc1_indices:
                rel_pos1 = i1 / max(len(doc1_paras), 1)
                for i2 in doc2_indices:
                    rel_pos2 = i2 / max(len(doc2_paras), 1)
                    # If relative position changed by more than 10%, consider it moved
                    if abs(rel_pos1 - rel_pos2) > 0.1:
                        moved_in_doc1.add(i1)
                        moved_in_doc2.add(i2)

    return moved_in_doc1, moved_in_doc2


def apply_word_diff(output_doc, text1, text2, is_moved=False):
    """
    Apply word-level diff between two texts.
    Creates a paragraph with proper coloring.
    """
    para = output_doc.add_paragraph()

    # Split into words preserving whitespace
    words1 = re.findall(r'\S+|\s+', text1)
    words2 = re.findall(r'\S+|\s+', text2)

    matcher = difflib.SequenceMatcher(None, words1, words2)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            text = ''.join(words1[i1:i2])
            if text:
                run = para.add_run(text)
                if is_moved:
                    run.font.color.rgb = RGBColor(0, 128, 0)  # Green for moved

        elif tag == 'delete':
            text = ''.join(words1[i1:i2])
            if text.strip():
                run = para.add_run(text)
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red for deleted
                run.font.strike = True

        elif tag == 'insert':
            text = ''.join(words2[j1:j2])
            if text.strip():
                run = para.add_run(text)
                run.font.color.rgb = RGBColor(0, 0, 255)  # Blue for added

        elif tag == 'replace':
            # Show deleted then inserted
            del_text = ''.join(words1[i1:i2])
            ins_text = ''.join(words2[j1:j2])

            if del_text.strip():
                run = para.add_run(del_text)
                run.font.color.rgb = RGBColor(255, 0, 0)
                run.font.strike = True

            if ins_text.strip():
                run = para.add_run(ins_text)
                run.font.color.rgb = RGBColor(0, 0, 255)

    return para


def compare_documents(doc1_path, doc2_path, output_path):
    """
    Compare two documents and create a redlined output.

    Args:
        doc1_path: Path to original document
        doc2_path: Path to modified document
        output_path: Path for output redlined document

    Returns:
        dict: Statistics about the comparison
    """
    emit("progress", percent=10, message="Reading documents...")

    # Read both documents
    doc1_paras = get_paragraph_texts(doc1_path)
    doc2_paras = get_paragraph_texts(doc2_path)

    emit("progress", percent=20, message="Detecting moved paragraphs...")

    # Detect moved paragraphs
    moved_in_doc1, moved_in_doc2 = detect_moved_paragraphs(doc1_paras, doc2_paras)

    emit("progress", percent=30, message="Comparing documents...")

    # Create output document
    output_doc = Document()

    # Add header with legend
    title = output_doc.add_paragraph()
    title_run = title.add_run("DOCUMENT REDLINE COMPARISON")
    title_run.bold = True
    title_run.font.size = Pt(16)

    output_doc.add_paragraph(f"Original: {os.path.basename(doc1_path)}")
    output_doc.add_paragraph(f"Modified: {os.path.basename(doc2_path)}")
    output_doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    output_doc.add_paragraph("")

    # Legend
    legend = output_doc.add_paragraph()
    legend.add_run("LEGEND: ").bold = True

    del_run = legend.add_run("Deleted text")
    del_run.font.color.rgb = RGBColor(255, 0, 0)
    del_run.font.strike = True

    legend.add_run(" | ")

    add_run = legend.add_run("Added text")
    add_run.font.color.rgb = RGBColor(0, 0, 255)

    legend.add_run(" | ")

    move_run = legend.add_run("Moved text")
    move_run.font.color.rgb = RGBColor(0, 128, 0)

    output_doc.add_paragraph("")
    output_doc.add_paragraph("─" * 50)
    output_doc.add_paragraph("")

    # Statistics
    stats = {
        'paragraphs_deleted': 0,
        'paragraphs_added': 0,
        'paragraphs_modified': 0,
        'paragraphs_unchanged': 0,
        'paragraphs_moved': 0,
        'words_deleted': 0,
        'words_added': 0
    }

    # Compare paragraph by paragraph
    matcher = difflib.SequenceMatcher(None, doc1_paras, doc2_paras)
    opcodes = matcher.get_opcodes()
    total_ops = len(opcodes)

    for op_idx, (tag, i1, i2, j1, j2) in enumerate(opcodes):
        percent = 30 + int((op_idx / max(total_ops, 1)) * 60)
        emit("progress", percent=percent, message="Building redline document...")

        if tag == 'equal':
            # Unchanged paragraphs - but check if any are "moved"
            for idx, para_idx in enumerate(range(i1, i2)):
                para_text = doc1_paras[para_idx]
                if para_idx in moved_in_doc1:
                    # Show as moved (green)
                    para = output_doc.add_paragraph()
                    run = para.add_run(para_text)
                    run.font.color.rgb = RGBColor(0, 128, 0)
                    stats['paragraphs_moved'] += 1
                else:
                    output_doc.add_paragraph(para_text)
                    stats['paragraphs_unchanged'] += 1

        elif tag == 'delete':
            # Deleted paragraphs
            for para_idx in range(i1, i2):
                para_text = doc1_paras[para_idx]

                if para_idx in moved_in_doc1:
                    # This paragraph moved elsewhere, show as moved-from
                    para = output_doc.add_paragraph()
                    run = para.add_run(f"[Moved from here] {para_text}")
                    run.font.color.rgb = RGBColor(0, 128, 0)
                    run.font.italic = True
                    stats['paragraphs_moved'] += 1
                else:
                    # Actually deleted
                    para = output_doc.add_paragraph()
                    run = para.add_run(para_text)
                    run.font.color.rgb = RGBColor(255, 0, 0)
                    run.font.strike = True
                    stats['paragraphs_deleted'] += 1
                    stats['words_deleted'] += len(para_text.split())

        elif tag == 'insert':
            # Added paragraphs
            for para_idx in range(j1, j2):
                para_text = doc2_paras[para_idx]

                if para_idx in moved_in_doc2:
                    # This paragraph moved from elsewhere
                    para = output_doc.add_paragraph()
                    run = para.add_run(f"[Moved to here] {para_text}")
                    run.font.color.rgb = RGBColor(0, 128, 0)
                    run.font.italic = True
                    stats['paragraphs_moved'] += 1
                else:
                    # Actually added
                    para = output_doc.add_paragraph()
                    run = para.add_run(para_text)
                    run.font.color.rgb = RGBColor(0, 0, 255)
                    stats['paragraphs_added'] += 1
                    stats['words_added'] += len(para_text.split())

        elif tag == 'replace':
            # Modified paragraphs - show word-level diff
            # Pair up paragraphs where possible
            for idx in range(max(i2 - i1, j2 - j1)):
                base_idx = i1 + idx if i1 + idx < i2 else None
                mod_idx = j1 + idx if j1 + idx < j2 else None

                if base_idx is not None and mod_idx is not None:
                    # Both exist - show word diff
                    base_text = doc1_paras[base_idx]
                    mod_text = doc2_paras[mod_idx]
                    apply_word_diff(output_doc, base_text, mod_text)
                    stats['paragraphs_modified'] += 1

                    # Count word-level changes
                    base_words = set(base_text.split())
                    mod_words = set(mod_text.split())
                    stats['words_deleted'] += len(base_words - mod_words)
                    stats['words_added'] += len(mod_words - base_words)

                elif base_idx is not None:
                    # Only base exists - deleted
                    para = output_doc.add_paragraph()
                    run = para.add_run(doc1_paras[base_idx])
                    run.font.color.rgb = RGBColor(255, 0, 0)
                    run.font.strike = True
                    stats['paragraphs_deleted'] += 1

                elif mod_idx is not None:
                    # Only modified exists - added
                    para = output_doc.add_paragraph()
                    run = para.add_run(doc2_paras[mod_idx])
                    run.font.color.rgb = RGBColor(0, 0, 255)
                    stats['paragraphs_added'] += 1

    emit("progress", percent=95, message="Adding summary...")

    # Add summary page
    output_doc.add_page_break()

    summary_title = output_doc.add_paragraph()
    summary_run = summary_title.add_run("COMPARISON SUMMARY")
    summary_run.bold = True
    summary_run.font.size = Pt(14)

    output_doc.add_paragraph("")
    output_doc.add_paragraph(f"Paragraphs unchanged: {stats['paragraphs_unchanged']}")
    output_doc.add_paragraph(f"Paragraphs modified: {stats['paragraphs_modified']}")
    output_doc.add_paragraph(f"Paragraphs deleted: {stats['paragraphs_deleted']}")
    output_doc.add_paragraph(f"Paragraphs added: {stats['paragraphs_added']}")
    output_doc.add_paragraph(f"Paragraphs moved: {stats['paragraphs_moved']}")
    output_doc.add_paragraph("")
    output_doc.add_paragraph(f"Words deleted: {stats['words_deleted']}")
    output_doc.add_paragraph(f"Words added: {stats['words_added']}")

    # Save output
    output_doc.save(output_path)

    emit("progress", percent=100, message="Complete!")

    return stats


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: document_redline.py <config_json_path>")
        sys.exit(1)

    config_path = sys.argv[1]

    if not os.path.isfile(config_path):
        emit("error", message=f"Config file not found: {config_path}")
        sys.exit(1)

    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        emit("error", message=f"Invalid JSON config: {str(e)}")
        sys.exit(1)

    doc1_path = config.get('original_document')
    doc2_path = config.get('modified_document')
    output_folder = config.get('output_folder')

    if not doc1_path or not os.path.isfile(doc1_path):
        emit("error", message="Original document not found")
        sys.exit(1)

    if not doc2_path or not os.path.isfile(doc2_path):
        emit("error", message="Modified document not found")
        sys.exit(1)

    if not output_folder:
        output_folder = os.path.dirname(doc1_path)

    os.makedirs(output_folder, exist_ok=True)

    # Generate output filename
    base_name = os.path.splitext(os.path.basename(doc1_path))[0]
    output_path = os.path.join(output_folder, f"{base_name}_Redline.docx")

    try:
        stats = compare_documents(doc1_path, doc2_path, output_path)

        emit("result",
             success=True,
             output_folder=output_folder,
             output_document=output_path,
             paragraphs_unchanged=stats['paragraphs_unchanged'],
             paragraphs_modified=stats['paragraphs_modified'],
             paragraphs_deleted=stats['paragraphs_deleted'],
             paragraphs_added=stats['paragraphs_added'],
             paragraphs_moved=stats['paragraphs_moved'],
             words_deleted=stats['words_deleted'],
             words_added=stats['words_added'])

    except Exception as e:
        emit("error", message=str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
