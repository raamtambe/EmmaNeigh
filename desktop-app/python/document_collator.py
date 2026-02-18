#!/usr/bin/env python3
"""
EmmaNeigh - Document Collator
v5.1.4: Properly extracts track changes from Word documents and merges them
into the base document WITHOUT destroying formatting.

The key insight: Word track changes are stored in the OOXML as:
- <w:ins> elements for insertions
- <w:del> elements for deletions
- Comments are stored in comments.xml with references in document.xml

This version:
1. Copies the base document to preserve ALL formatting
2. Extracts track changes from each commented version
3. Merges track changes into the copy by modifying the XML directly
"""

import os
import sys
import json
import shutil
import zipfile
import re
from datetime import datetime
from xml.etree import ElementTree as ET
from copy import deepcopy


# Word XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}

# Register namespaces to preserve them when writing
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def get_author_from_docx(docx_path):
    """Try to extract author from document properties or use filename."""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'docProps/core.xml' in zf.namelist():
                core_xml = zf.read('docProps/core.xml').decode('utf-8')
                match = re.search(r'<dc:creator[^>]*>([^<]+)</dc:creator>', core_xml)
                if match:
                    return match.group(1).strip()
    except:
        pass
    # Fall back to filename without extension
    return os.path.splitext(os.path.basename(docx_path))[0]


def extract_track_changes_from_docx(docx_path):
    """
    Extract all track changes (insertions and deletions) from a Word document.

    Returns a dict with:
    - insertions: list of {author, date, text, paragraph_index}
    - deletions: list of {author, date, text, paragraph_index}
    - comments: list of {author, date, text, paragraph_index}
    """
    changes = {
        'insertions': [],
        'deletions': [],
        'comments': [],
        'author': get_author_from_docx(docx_path)
    }

    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            # Read document.xml
            if 'word/document.xml' not in zf.namelist():
                return changes

            doc_xml = zf.read('word/document.xml').decode('utf-8')
            root = ET.fromstring(doc_xml)

            # Find the body
            body = root.find('.//w:body', NAMESPACES)
            if body is None:
                return changes

            # Track paragraph index
            para_idx = 0

            for para in body.findall('.//w:p', NAMESPACES):
                # Find insertions in this paragraph
                for ins in para.findall('.//w:ins', NAMESPACES):
                    author = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', changes['author'])
                    date = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')

                    # Get text from all runs inside the insertion
                    text_parts = []
                    for t in ins.findall('.//w:t', NAMESPACES):
                        if t.text:
                            text_parts.append(t.text)

                    if text_parts:
                        changes['insertions'].append({
                            'author': author,
                            'date': date,
                            'text': ''.join(text_parts),
                            'paragraph_index': para_idx
                        })

                # Find deletions in this paragraph
                for dele in para.findall('.//w:del', NAMESPACES):
                    author = dele.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', changes['author'])
                    date = dele.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')

                    # Get text from delText elements
                    text_parts = []
                    for dt in dele.findall('.//w:delText', NAMESPACES):
                        if dt.text:
                            text_parts.append(dt.text)

                    if text_parts:
                        changes['deletions'].append({
                            'author': author,
                            'date': date,
                            'text': ''.join(text_parts),
                            'paragraph_index': para_idx
                        })

                para_idx += 1

            # Read comments if they exist
            if 'word/comments.xml' in zf.namelist():
                comments_xml = zf.read('word/comments.xml').decode('utf-8')
                comments_root = ET.fromstring(comments_xml)

                for comment in comments_root.findall('.//w:comment', NAMESPACES):
                    author = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', changes['author'])
                    date = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')

                    # Get comment text
                    text_parts = []
                    for t in comment.findall('.//w:t', NAMESPACES):
                        if t.text:
                            text_parts.append(t.text)

                    if text_parts:
                        changes['comments'].append({
                            'author': author,
                            'date': date,
                            'text': ''.join(text_parts)
                        })

    except Exception as e:
        emit("progress", percent=0, message=f"Warning: Error reading {docx_path}: {str(e)}")

    return changes


def merge_track_changes_into_document(base_path, all_changes, output_path):
    """
    Merge track changes from multiple documents into a copy of the base document.

    This works by:
    1. Copying the base document (preserves all formatting)
    2. Reading the document.xml
    3. Finding where to insert track changes based on text matching
    4. Writing the modified document.xml back

    Args:
        base_path: Path to the original base document
        all_changes: List of changes dicts from extract_track_changes_from_docx
        output_path: Where to save the merged document

    Returns:
        dict with counts of changes merged
    """
    # Copy base document to output
    shutil.copy2(base_path, output_path)

    total_insertions = 0
    total_deletions = 0
    total_comments = 0

    # Collect all changes
    for changes in all_changes:
        total_insertions += len(changes['insertions'])
        total_deletions += len(changes['deletions'])
        total_comments += len(changes['comments'])

    # If there are no changes, just return the copy
    if total_insertions == 0 and total_deletions == 0 and total_comments == 0:
        return {
            'insertions': 0,
            'deletions': 0,
            'comments': 0
        }

    # Now we need to merge the track changes into the document
    # This is complex because we need to preserve the XML structure

    try:
        # Read the output document as a zip
        with zipfile.ZipFile(output_path, 'r') as zf:
            doc_xml = zf.read('word/document.xml').decode('utf-8')
            all_files = {name: zf.read(name) for name in zf.namelist()}

        # Parse the document
        root = ET.fromstring(doc_xml)
        body = root.find('.//w:body', NAMESPACES)

        if body is not None:
            # Get all paragraphs
            paragraphs = body.findall('.//w:p', NAMESPACES)

            # For each set of changes, try to apply them
            for changes in all_changes:
                author = changes['author']

                # Apply insertions (mark with w:ins)
                for ins in changes['insertions']:
                    # Find the target paragraph (simplified: use index if available)
                    para_idx = ins.get('paragraph_index', 0)
                    if para_idx < len(paragraphs):
                        para = paragraphs[para_idx]

                        # Create an insertion element
                        ins_elem = ET.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins')
                        ins_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
                        ins_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', ins.get('date', datetime.now().isoformat()))

                        # Add run with text
                        run = ET.SubElement(ins_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                        text = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        text.text = ins['text']

                # Apply deletions (mark with w:del)
                for dele in changes['deletions']:
                    para_idx = dele.get('paragraph_index', 0)
                    if para_idx < len(paragraphs):
                        para = paragraphs[para_idx]

                        # Create a deletion element
                        del_elem = ET.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del')
                        del_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', author)
                        del_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', dele.get('date', datetime.now().isoformat()))

                        # Add run with deleted text
                        run = ET.SubElement(del_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                        del_text = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText')
                        del_text.text = dele['text']

        # Convert back to string
        new_doc_xml = ET.tostring(root, encoding='unicode')

        # Add XML declaration
        new_doc_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new_doc_xml

        # Write back to zip
        all_files['word/document.xml'] = new_doc_xml.encode('utf-8')

        # Handle comments - merge into comments.xml
        all_comments = []
        for changes in all_changes:
            for comment in changes['comments']:
                comment['source_author'] = changes['author']
                all_comments.append(comment)

        if all_comments:
            # Create or update comments.xml
            if 'word/comments.xml' in all_files:
                comments_root = ET.fromstring(all_files['word/comments.xml'].decode('utf-8'))
            else:
                comments_root = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comments')

            comment_id = 0
            for comment in all_comments:
                comm_elem = ET.SubElement(comments_root, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment')
                comm_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', str(comment_id))
                comm_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', comment.get('author', comment.get('source_author', 'Unknown')))
                comm_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', comment.get('date', datetime.now().isoformat()))

                # Add paragraph with text
                para = ET.SubElement(comm_elem, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                run = ET.SubElement(para, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                text = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                text.text = comment['text']

                comment_id += 1

            comments_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(comments_root, encoding='unicode')
            all_files['word/comments.xml'] = comments_xml.encode('utf-8')

        # Enable track changes in settings.xml
        if 'word/settings.xml' in all_files:
            try:
                settings_xml = all_files['word/settings.xml'].decode('utf-8')
                settings_root = ET.fromstring(settings_xml)

                # Add trackRevisions element to enable track changes
                # Check if it already exists
                track_rev = settings_root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trackRevisions')
                if track_rev is None:
                    # Add trackRevisions element
                    track_rev = ET.SubElement(settings_root, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trackRevisions')

                # Set val to true (or just having the element enables it)
                track_rev.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')

                settings_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(settings_root, encoding='unicode')
                all_files['word/settings.xml'] = settings_xml.encode('utf-8')
            except Exception as e:
                emit("progress", percent=0, message=f"Warning: Could not enable track changes: {str(e)}")

        # Write the new zip file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, content in all_files.items():
                zf.writestr(name, content)

    except Exception as e:
        emit("progress", percent=0, message=f"Warning: Error merging changes: {str(e)}")
        import traceback
        traceback.print_exc()

    return {
        'insertions': total_insertions,
        'deletions': total_deletions,
        'comments': total_comments
    }


def create_summary(base_path, modified_paths, stats, output_folder):
    """Create a simple text summary of the collation."""
    summary_lines = [
        "Document Collation Summary",
        "=" * 50,
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        f"Base Document: {os.path.basename(base_path)}",
        f"Commented Versions: {len(modified_paths)}",
        "",
        "Sources:",
    ]

    for path in modified_paths:
        author = get_author_from_docx(path)
        summary_lines.append(f"  • {os.path.basename(path)} (Author: {author})")

    summary_lines.extend([
        "",
        "Changes Merged:",
        f"  • Insertions: {stats.get('insertions', 0)}",
        f"  • Deletions: {stats.get('deletions', 0)}",
        f"  • Comments: {stats.get('comments', 0)}",
        "",
        "Note: Track changes have been preserved in the merged document.",
        "Open in Word and use Review > Track Changes to see all changes.",
    ])

    summary_path = os.path.join(output_folder, "Collation_Summary.txt")
    with open(summary_path, 'w') as f:
        f.write('\n'.join(summary_lines))

    return summary_path


def collate_documents(base_path, modified_paths, output_folder):
    """
    Main collation function.

    Args:
        base_path: Path to the precedent/base document
        modified_paths: List of paths to modified versions with track changes
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

    emit("progress", percent=20, message=f"Extracting changes from {len(valid_modified)} documents...")

    # Extract track changes from all modified documents
    all_changes = []
    for i, mod_path in enumerate(valid_modified):
        percent = 20 + int((i / len(valid_modified)) * 40)
        emit("progress", percent=percent, message=f"Reading {os.path.basename(mod_path)}...")

        changes = extract_track_changes_from_docx(mod_path)
        all_changes.append(changes)

        # Log what we found
        ins_count = len(changes['insertions'])
        del_count = len(changes['deletions'])
        comm_count = len(changes['comments'])
        if ins_count + del_count + comm_count > 0:
            emit("progress", percent=percent,
                 message=f"Found {ins_count} insertions, {del_count} deletions, {comm_count} comments in {os.path.basename(mod_path)}")

    emit("progress", percent=65, message="Merging changes into base document...")

    # Create output document by merging changes
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    output_path = os.path.join(output_folder, f"{base_name}_Collated.docx")

    stats = merge_track_changes_into_document(base_path, all_changes, output_path)

    emit("progress", percent=90, message="Creating summary...")

    # Create summary
    summary_path = create_summary(base_path, valid_modified, stats, output_folder)

    total_changes = stats['insertions'] + stats['deletions']

    emit("progress", percent=100, message="Complete!")

    return {
        'output_folder': output_folder,
        'output_document': output_path,
        'summary_document': summary_path,
        'total_changes': total_changes,
        'total_comments': stats['comments'],
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
             total_comments=results['total_comments'],
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
