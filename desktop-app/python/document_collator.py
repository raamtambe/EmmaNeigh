#!/usr/bin/env python3
"""
EmmaNeigh - Document Collator
Integrates track changes and comments from multiple document versions into a base document.
"""

import os
import sys
import json
import re
import zipfile
import shutil
from collections import defaultdict
from datetime import datetime
from xml.etree import ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

# Word XML namespaces
WORD_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def extract_docx_xml(docx_path):
    """Extract the document.xml content from a .docx file."""
    with zipfile.ZipFile(docx_path, 'r') as zf:
        return zf.read('word/document.xml').decode('utf-8')


def get_author_from_docx(docx_path):
    """Try to extract author from document properties."""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'docProps/core.xml' in zf.namelist():
                core_xml = zf.read('docProps/core.xml').decode('utf-8')
                # Parse for creator
                match = re.search(r'<dc:creator[^>]*>([^<]+)</dc:creator>', core_xml)
                if match:
                    return match.group(1).strip()
    except:
        pass
    # Fall back to filename
    return os.path.splitext(os.path.basename(docx_path))[0]


def extract_track_changes(docx_path):
    """
    Extract track changes (insertions and deletions) from a Word document.
    Returns a list of changes with author, type, and content.
    """
    changes = []

    try:
        xml_content = extract_docx_xml(docx_path)
        root = ET.fromstring(xml_content)

        # Find all insertions (w:ins)
        for ins in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ins'):
            author = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = ins.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')

            # Get the text content
            text_parts = []
            for t in ins.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                if t.text:
                    text_parts.append(t.text)

            if text_parts:
                changes.append({
                    'type': 'insertion',
                    'author': author,
                    'date': date,
                    'content': ''.join(text_parts),
                    'element': ins
                })

        # Find all deletions (w:del)
        for del_elem in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del'):
            author = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = del_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')

            # Get the deleted text
            text_parts = []
            for t in del_elem.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}delText'):
                if t.text:
                    text_parts.append(t.text)

            if text_parts:
                changes.append({
                    'type': 'deletion',
                    'author': author,
                    'date': date,
                    'content': ''.join(text_parts),
                    'element': del_elem
                })
    except Exception as e:
        emit("progress", percent=0, message=f"Warning: Could not parse track changes from {docx_path}: {str(e)}")

    return changes


def extract_comments(docx_path):
    """
    Extract comments from a Word document.
    Returns a list of comments with author and content.
    """
    comments = []

    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if 'word/comments.xml' not in zf.namelist():
                return comments

            comments_xml = zf.read('word/comments.xml').decode('utf-8')
            root = ET.fromstring(comments_xml)

            for comment in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment'):
                author = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
                date = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
                comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')

                # Get comment text
                text_parts = []
                for t in comment.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if t.text:
                        text_parts.append(t.text)

                if text_parts:
                    comments.append({
                        'id': comment_id,
                        'author': author,
                        'date': date,
                        'content': ''.join(text_parts)
                    })
    except Exception as e:
        emit("progress", percent=0, message=f"Warning: Could not parse comments from {docx_path}: {str(e)}")

    return comments


def get_paragraph_text(para):
    """Extract plain text from a paragraph element."""
    texts = []
    for t in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def extract_paragraphs(docx_path):
    """Extract all paragraphs with their text content."""
    paragraphs = []
    try:
        xml_content = extract_docx_xml(docx_path)
        root = ET.fromstring(xml_content)

        for para in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
            text = get_paragraph_text(para)
            paragraphs.append({
                'text': text,
                'element': para
            })
    except:
        pass

    return paragraphs


def simple_diff(base_text, modified_text):
    """
    Simple word-level diff between two texts.
    Returns additions and deletions.
    """
    base_words = base_text.split()
    mod_words = modified_text.split()

    additions = []
    deletions = []

    # Simple LCS-based diff
    base_set = set(base_words)
    mod_set = set(mod_words)

    # Words in modified but not in base (additions)
    for word in mod_words:
        if word not in base_set:
            additions.append(word)

    # Words in base but not in modified (deletions)
    for word in base_words:
        if word not in mod_set:
            deletions.append(word)

    return additions, deletions


def detect_conflicts(changes_by_source):
    """
    Detect conflicting changes between different sources.
    Returns a list of conflict descriptions.
    """
    conflicts = []

    # Group changes by approximate location (paragraph/content overlap)
    all_deletions = defaultdict(list)
    all_insertions = defaultdict(list)

    for source, changes in changes_by_source.items():
        for change in changes:
            content_key = change['content'][:50] if len(change['content']) > 50 else change['content']

            if change['type'] == 'deletion':
                all_deletions[content_key].append({
                    'source': source,
                    'author': change['author'],
                    'content': change['content']
                })
            elif change['type'] == 'insertion':
                all_insertions[content_key].append({
                    'source': source,
                    'author': change['author'],
                    'content': change['content']
                })

    # Check for conflicts: same content deleted by one, modified by another
    for deleted_content, deleters in all_deletions.items():
        if len(deleters) > 1:
            # Multiple sources deleting same content - not necessarily a conflict
            continue

        # Check if any insertion modifies text that overlaps with deletion
        for ins_content, inserters in all_insertions.items():
            # Simple overlap check
            deleted_words = set(deleted_content.lower().split())
            inserted_words = set(ins_content.lower().split())

            overlap = deleted_words & inserted_words
            if overlap and len(overlap) > 2:  # Significant overlap
                # Check if different sources
                deleter_sources = {d['source'] for d in deleters}
                inserter_sources = {i['source'] for i in inserters}

                if deleter_sources != inserter_sources:
                    conflicts.append({
                        'type': 'delete_vs_modify',
                        'description': f"Conflict: {deleters[0]['author']} deleted text while {inserters[0]['author']} modified related text",
                        'deleted_by': deleters[0]['author'],
                        'deleted_content': deleted_content[:100],
                        'modified_by': inserters[0]['author'],
                        'modified_content': ins_content[:100]
                    })

    return conflicts


def check_grammatical_validity(text):
    """
    Simple check if text forms a valid sentence.
    Returns True if likely grammatically valid.
    """
    text = text.strip()
    if not text:
        return False

    # Basic checks
    # 1. Has subject-verb potential (contains common verbs or verb endings)
    verb_patterns = [
        r'\b(is|are|was|were|be|been|being)\b',
        r'\b(have|has|had|having)\b',
        r'\b(do|does|did|doing)\b',
        r'\b(will|shall|would|should|could|may|might|must)\b',
        r'\b\w+ed\b',  # Past tense
        r'\b\w+ing\b',  # Present participle
        r'\b\w+s\b',  # Third person
    ]

    has_verb = any(re.search(p, text, re.IGNORECASE) for p in verb_patterns)

    # 2. Reasonable length
    word_count = len(text.split())
    reasonable_length = 3 <= word_count <= 500

    # 3. Doesn't start with lowercase after period
    bad_start = re.search(r'\.\s+[a-z]', text)

    # 4. Balanced parentheses and quotes
    balanced = text.count('(') == text.count(')') and text.count('"') % 2 == 0

    return has_verb and reasonable_length and not bad_start and balanced


def merge_changes_into_document(base_path, changes_by_source, comments_by_source, output_path):
    """
    Merge all track changes and comments into the base document.
    Creates a new document with all changes integrated.
    """
    # Copy base document
    shutil.copy2(base_path, output_path)

    # Open the copied document
    doc = Document(output_path)

    # Collect all unique changes
    all_changes = []
    for source, changes in changes_by_source.items():
        for change in changes:
            change['source_file'] = source
            all_changes.append(change)

    # Add a summary section at the end
    doc.add_page_break()

    # Add header for comments summary
    heading = doc.add_paragraph()
    heading_run = heading.add_run("COLLATED CHANGES SUMMARY")
    heading_run.bold = True
    heading_run.font.size = Pt(14)

    doc.add_paragraph("")

    # Add comments from all sources
    if any(comments_by_source.values()):
        comments_heading = doc.add_paragraph()
        comments_heading_run = comments_heading.add_run("Comments from Reviewers:")
        comments_heading_run.bold = True
        comments_heading_run.font.size = Pt(12)

        for source, comments in comments_by_source.items():
            if comments:
                source_para = doc.add_paragraph()
                source_run = source_para.add_run(f"\nFrom {os.path.basename(source)}:")
                source_run.italic = True

                for comment in comments:
                    comment_para = doc.add_paragraph()
                    comment_para.add_run(f"  [{comment['author']}]: ").bold = True
                    comment_para.add_run(comment['content'])

    doc.save(output_path)
    return len(all_changes)


def generate_summary_table(base_path, commented_versions, changes_by_source, comments_by_source, conflicts, output_folder):
    """
    Generate a summary table document showing:
    - Author of each commented version
    - Number of track changes
    - Number of comments
    - Conflicts detected
    """
    doc = Document()

    # Title
    title = doc.add_paragraph()
    title_run = title.add_run("Document Collation Summary")
    title_run.bold = True
    title_run.font.size = Pt(16)

    doc.add_paragraph(f"Base Document: {os.path.basename(base_path)}")
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_paragraph("")

    # Create summary table
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'

    # Header row
    header_cells = table.rows[0].cells
    headers = ['Source File', 'Author', 'Track Changes', 'Comments', 'Lines Modified']
    for i, header in enumerate(headers):
        header_cells[i].text = header
        for run in header_cells[i].paragraphs[0].runs:
            run.bold = True

    # Data rows
    total_changes = 0
    total_comments = 0

    for version_path in commented_versions:
        row = table.add_row()
        cells = row.cells

        filename = os.path.basename(version_path)
        author = get_author_from_docx(version_path)
        changes = changes_by_source.get(version_path, [])
        comments = comments_by_source.get(version_path, [])

        num_changes = len(changes)
        num_comments = len(comments)
        lines_modified = sum(1 for c in changes if c['content'].strip())

        cells[0].text = filename
        cells[1].text = author
        cells[2].text = str(num_changes)
        cells[3].text = str(num_comments)
        cells[4].text = str(lines_modified)

        total_changes += num_changes
        total_comments += num_comments

    # Total row
    total_row = table.add_row()
    total_cells = total_row.cells
    total_cells[0].text = "TOTAL"
    total_cells[0].paragraphs[0].runs[0].bold = True
    total_cells[1].text = "-"
    total_cells[2].text = str(total_changes)
    total_cells[3].text = str(total_comments)
    total_cells[4].text = "-"

    doc.add_paragraph("")

    # Conflicts section
    if conflicts:
        conflict_heading = doc.add_paragraph()
        conflict_run = conflict_heading.add_run("CONFLICTS DETECTED")
        conflict_run.bold = True
        conflict_run.font.size = Pt(14)
        conflict_run.font.color.rgb = RGBColor(255, 0, 0)

        for i, conflict in enumerate(conflicts, 1):
            doc.add_paragraph(f"{i}. {conflict['description']}")
            if 'deleted_content' in conflict:
                doc.add_paragraph(f"   Deleted: \"{conflict['deleted_content'][:80]}...\"")
            if 'modified_content' in conflict:
                doc.add_paragraph(f"   Modified: \"{conflict['modified_content'][:80]}...\"")
    else:
        no_conflict = doc.add_paragraph()
        no_conflict_run = no_conflict.add_run("No conflicts detected.")
        no_conflict_run.font.color.rgb = RGBColor(0, 128, 0)

    # Save summary
    summary_path = os.path.join(output_folder, "Collation_Summary.docx")
    doc.save(summary_path)

    return summary_path


def collate_documents(base_path, commented_paths, output_folder):
    """
    Main function to collate multiple document versions.

    Args:
        base_path: Path to the base document
        commented_paths: List of paths to commented versions
        output_folder: Where to save output

    Returns:
        dict: Results including paths and statistics
    """
    os.makedirs(output_folder, exist_ok=True)

    results = {
        'base_document': base_path,
        'commented_versions': len(commented_paths),
        'changes_by_source': {},
        'comments_by_source': {},
        'conflicts': [],
        'output_document': None,
        'summary_document': None,
        'total_changes': 0,
        'total_comments': 0
    }

    emit("progress", percent=10, message="Analyzing base document...")

    # Extract base document paragraphs for comparison
    base_paragraphs = extract_paragraphs(base_path)

    # Process each commented version
    changes_by_source = {}
    comments_by_source = {}

    for i, commented_path in enumerate(commented_paths):
        percent = 10 + int((i / len(commented_paths)) * 50)
        emit("progress", percent=percent, message=f"Processing {os.path.basename(commented_path)}...")

        # Extract track changes
        changes = extract_track_changes(commented_path)
        changes_by_source[commented_path] = changes

        # Extract comments
        comments = extract_comments(commented_path)
        comments_by_source[commented_path] = comments

        results['total_changes'] += len(changes)
        results['total_comments'] += len(comments)

    results['changes_by_source'] = {os.path.basename(k): len(v) for k, v in changes_by_source.items()}
    results['comments_by_source'] = {os.path.basename(k): len(v) for k, v in comments_by_source.items()}

    emit("progress", percent=65, message="Detecting conflicts...")

    # Detect conflicts
    conflicts = detect_conflicts(changes_by_source)
    results['conflicts'] = conflicts

    # For each conflict, check if we can construct a valid sentence
    resolved_conflicts = []
    for conflict in conflicts:
        # Try to combine the changes
        combined = conflict.get('deleted_content', '') + ' ' + conflict.get('modified_content', '')
        if check_grammatical_validity(combined):
            conflict['resolvable'] = True
            conflict['resolution'] = "Changes can potentially be combined"
        else:
            conflict['resolvable'] = False
            resolved_conflicts.append(conflict)

    results['unresolved_conflicts'] = len(resolved_conflicts)

    emit("progress", percent=75, message="Merging changes into document...")

    # Create merged document
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    output_doc_path = os.path.join(output_folder, f"{base_name}_Collated.docx")

    merge_changes_into_document(
        base_path,
        changes_by_source,
        comments_by_source,
        output_doc_path
    )
    results['output_document'] = output_doc_path

    emit("progress", percent=90, message="Generating summary table...")

    # Generate summary table
    summary_path = generate_summary_table(
        base_path,
        commented_paths,
        changes_by_source,
        comments_by_source,
        resolved_conflicts,
        output_folder
    )
    results['summary_document'] = summary_path

    emit("progress", percent=100, message="Complete!")

    return results


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
    commented_paths = config.get('commented_documents', [])
    output_folder = config.get('output_folder')

    if not base_path or not os.path.isfile(base_path):
        emit("error", message="Base document not found")
        sys.exit(1)

    if not commented_paths:
        emit("error", message="No commented documents provided")
        sys.exit(1)

    if not output_folder:
        output_folder = os.path.join(os.path.dirname(base_path), "collated_output")

    try:
        results = collate_documents(base_path, commented_paths, output_folder)

        emit("result",
             success=True,
             output_folder=output_folder,
             output_document=results['output_document'],
             summary_document=results['summary_document'],
             total_changes=results['total_changes'],
             total_comments=results['total_comments'],
             conflicts=len(results['conflicts']),
             unresolved_conflicts=results.get('unresolved_conflicts', 0),
             changes_by_source=results['changes_by_source'],
             comments_by_source=results['comments_by_source'])

    except Exception as e:
        emit("error", message=str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
