#!/usr/bin/env python3
"""
EmmaNeigh - Document Collator
v5.1.2: Simplified merge of track changes from multiple Word documents.

Takes a precedent (base) document and multiple modified versions,
then merges all edits into one combined document showing all changes.
"""

import os
import sys
import json
import zipfile
import difflib
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def get_author_from_docx(docx_path):
    """Try to extract author from document properties or use filename."""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'docProps/core.xml' in zf.namelist():
                import re
                core_xml = zf.read('docProps/core.xml').decode('utf-8')
                match = re.search(r'<dc:creator[^>]*>([^<]+)</dc:creator>', core_xml)
                if match:
                    return match.group(1).strip()
    except:
        pass
    # Fall back to filename without extension
    return os.path.splitext(os.path.basename(docx_path))[0]


def get_paragraphs_text(docx_path):
    """Extract all paragraph texts from a Word document."""
    try:
        doc = Document(docx_path)
        return [p.text for p in doc.paragraphs]
    except Exception as e:
        emit("progress", percent=0, message=f"Warning: Could not read {docx_path}: {str(e)}")
        return []


def create_merged_document(base_path, modified_paths, output_path):
    """
    Create a merged document showing all changes from all modified versions.

    - Deletions shown in red with strikethrough
    - Insertions shown in blue
    - Each change attributed to its author
    """
    # Read base document
    base_paragraphs = get_paragraphs_text(base_path)

    # Collect all modifications from all sources
    all_modifications = []  # List of (author, base_text, modified_text, para_index)

    for mod_path in modified_paths:
        author = get_author_from_docx(mod_path)
        mod_paragraphs = get_paragraphs_text(mod_path)

        # Use difflib to align paragraphs
        matcher = difflib.SequenceMatcher(None, base_paragraphs, mod_paragraphs)

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'replace':
                # Paragraphs were modified
                for bi, mi in zip(range(i1, i2), range(j1, j2)):
                    if bi < len(base_paragraphs) and mi < len(mod_paragraphs):
                        if base_paragraphs[bi] != mod_paragraphs[mi]:
                            all_modifications.append({
                                'author': author,
                                'base_text': base_paragraphs[bi],
                                'modified_text': mod_paragraphs[mi],
                                'para_index': bi,
                                'type': 'modified'
                            })
            elif tag == 'delete':
                # Paragraphs deleted
                for bi in range(i1, i2):
                    all_modifications.append({
                        'author': author,
                        'base_text': base_paragraphs[bi],
                        'modified_text': '',
                        'para_index': bi,
                        'type': 'deleted'
                    })
            elif tag == 'insert':
                # New paragraphs added
                insert_after = i1 - 1 if i1 > 0 else 0
                for mi in range(j1, j2):
                    all_modifications.append({
                        'author': author,
                        'base_text': '',
                        'modified_text': mod_paragraphs[mi],
                        'para_index': insert_after,
                        'type': 'inserted'
                    })

    # Group modifications by paragraph index
    mods_by_para = {}
    for mod in all_modifications:
        idx = mod['para_index']
        if idx not in mods_by_para:
            mods_by_para[idx] = []
        mods_by_para[idx].append(mod)

    # Create output document
    output_doc = Document()

    # Add legend at the top
    legend = output_doc.add_paragraph()
    legend.add_run("MERGED DOCUMENT - ").bold = True
    legend.add_run("Legend: ")
    del_run = legend.add_run("Deleted")
    del_run.font.color.rgb = RGBColor(255, 0, 0)
    del_run.font.strike = True
    legend.add_run(" | ")
    ins_run = legend.add_run("Added")
    ins_run.font.color.rgb = RGBColor(0, 0, 255)
    output_doc.add_paragraph("")

    changes_count = 0
    insertions_to_add = {}  # para_index -> list of insertions to add after

    # Collect insertions separately
    for idx, mods in mods_by_para.items():
        for mod in mods:
            if mod['type'] == 'inserted':
                if idx not in insertions_to_add:
                    insertions_to_add[idx] = []
                insertions_to_add[idx].append(mod)

    # Build the document paragraph by paragraph
    for i, base_text in enumerate(base_paragraphs):
        mods = mods_by_para.get(i, [])

        # Filter to non-insertions for this paragraph
        para_mods = [m for m in mods if m['type'] != 'inserted']

        if not para_mods:
            # No changes to this paragraph - keep as is
            output_doc.add_paragraph(base_text)
        else:
            # Apply changes
            # If multiple authors modified same paragraph, show all their changes
            para = output_doc.add_paragraph()

            # Check if any author deleted this paragraph entirely
            deletions = [m for m in para_mods if m['type'] == 'deleted']
            modifications = [m for m in para_mods if m['type'] == 'modified']

            if deletions and not modifications:
                # Paragraph was deleted - show as strikethrough red
                if base_text.strip():
                    run = para.add_run(base_text)
                    run.font.color.rgb = RGBColor(255, 0, 0)
                    run.font.strike = True
                    # Add author attribution
                    attr = para.add_run(f" [{deletions[0]['author']}]")
                    attr.font.size = Pt(8)
                    attr.font.color.rgb = RGBColor(128, 128, 128)
                    changes_count += 1
            elif modifications:
                # Use the first modification and show word-level diff
                mod = modifications[0]
                add_word_diff(para, base_text, mod['modified_text'], mod['author'])
                changes_count += 1
            else:
                # Keep original
                output_doc.add_paragraph(base_text)

        # Add any insertions that come after this paragraph
        if i in insertions_to_add:
            for ins in insertions_to_add[i]:
                if ins['modified_text'].strip():
                    ins_para = output_doc.add_paragraph()
                    run = ins_para.add_run(ins['modified_text'])
                    run.font.color.rgb = RGBColor(0, 0, 255)
                    # Add author attribution
                    attr = ins_para.add_run(f" [Added by {ins['author']}]")
                    attr.font.size = Pt(8)
                    attr.font.color.rgb = RGBColor(128, 128, 128)
                    changes_count += 1

    # Save the merged document
    output_doc.save(output_path)
    return changes_count


def add_word_diff(para, base_text, modified_text, author):
    """Add word-level diff to a paragraph, showing deletions and insertions."""
    import re

    # Split into words preserving whitespace
    base_words = re.findall(r'\S+|\s+', base_text)
    mod_words = re.findall(r'\S+|\s+', modified_text)

    matcher = difflib.SequenceMatcher(None, base_words, mod_words)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            text = ''.join(base_words[i1:i2])
            if text:
                para.add_run(text)

        elif tag == 'delete':
            text = ''.join(base_words[i1:i2])
            if text.strip():
                run = para.add_run(text)
                run.font.color.rgb = RGBColor(255, 0, 0)
                run.font.strike = True

        elif tag == 'insert':
            text = ''.join(mod_words[j1:j2])
            if text.strip():
                run = para.add_run(text)
                run.font.color.rgb = RGBColor(0, 0, 255)

        elif tag == 'replace':
            # Show deletion then insertion
            del_text = ''.join(base_words[i1:i2])
            ins_text = ''.join(mod_words[j1:j2])

            if del_text.strip():
                run = para.add_run(del_text)
                run.font.color.rgb = RGBColor(255, 0, 0)
                run.font.strike = True

            if ins_text.strip():
                run = para.add_run(ins_text)
                run.font.color.rgb = RGBColor(0, 0, 255)

    # Add author attribution at end
    attr = para.add_run(f" [{author}]")
    attr.font.size = Pt(8)
    attr.font.color.rgb = RGBColor(128, 128, 128)


def create_summary(base_path, modified_paths, changes_count, output_folder):
    """Create a simple summary document."""
    doc = Document()

    # Title
    title = doc.add_paragraph()
    title_run = title.add_run("Document Collation Summary")
    title_run.bold = True
    title_run.font.size = Pt(16)

    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph("")

    doc.add_paragraph(f"Base Document: {os.path.basename(base_path)}")
    doc.add_paragraph(f"Modified Versions: {len(modified_paths)}")
    doc.add_paragraph("")

    # List modified docs with authors
    sources_heading = doc.add_paragraph()
    sources_heading.add_run("Sources:").bold = True

    for path in modified_paths:
        author = get_author_from_docx(path)
        doc.add_paragraph(f"  • {os.path.basename(path)} (Author: {author})")

    doc.add_paragraph("")
    doc.add_paragraph(f"Total Changes Applied: {changes_count}")

    summary_path = os.path.join(output_folder, "Collation_Summary.docx")
    doc.save(summary_path)
    return summary_path


def collate_documents(base_path, modified_paths, output_folder):
    """
    Main collation function.

    Args:
        base_path: Path to the precedent/base document
        modified_paths: List of paths to modified versions
        output_folder: Where to save output

    Returns:
        dict with results
    """
    os.makedirs(output_folder, exist_ok=True)

    emit("progress", percent=10, message="Reading base document...")

    # Validate inputs
    if not os.path.isfile(base_path):
        raise ValueError(f"Base document not found: {base_path}")

    valid_modified = []
    for p in modified_paths:
        if os.path.isfile(p):
            valid_modified.append(p)
        else:
            emit("progress", percent=0, message=f"Warning: Skipping missing file {p}")

    if not valid_modified:
        raise ValueError("No valid modified documents found")

    emit("progress", percent=30, message=f"Processing {len(valid_modified)} modified documents...")

    # Create merged document
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    output_path = os.path.join(output_folder, f"{base_name}_Collated.docx")

    emit("progress", percent=50, message="Merging changes...")

    changes_count = create_merged_document(base_path, valid_modified, output_path)

    emit("progress", percent=80, message="Creating summary...")

    # Create summary
    summary_path = create_summary(base_path, valid_modified, changes_count, output_folder)

    emit("progress", percent=100, message="Complete!")

    return {
        'output_folder': output_folder,
        'output_document': output_path,
        'summary_document': summary_path,
        'total_changes': changes_count,
        'documents_processed': len(valid_modified)
    }


def main():
    """CLI entry point."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: document_collator.py <config_json_path>")
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

    base_path = config.get('base_document')
    modified_paths = config.get('commented_documents', [])
    output_folder = config.get('output_folder')

    if not base_path:
        emit("error", message="No base document specified")
        sys.exit(1)

    if not modified_paths:
        emit("error", message="No modified documents specified")
        sys.exit(1)

    if not output_folder:
        output_folder = os.path.join(os.path.dirname(base_path), "collated_output")

    try:
        results = collate_documents(base_path, modified_paths, output_folder)

        emit("result",
             success=True,
             output_folder=results['output_folder'],
             output_document=results['output_document'],
             summary_document=results['summary_document'],
             total_changes=results['total_changes'],
             total_comments=0,
             conflicts=0,
             unresolved_conflicts=0,
             changes_by_source={},
             comments_by_source={})

    except Exception as e:
        import traceback
        emit("error", message=f"{str(e)}\n{traceback.format_exc()}")
        sys.exit(1)


if __name__ == "__main__":
    main()
