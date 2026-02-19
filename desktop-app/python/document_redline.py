#!/usr/bin/env python3
"""
EmmaNeigh - Document Redline
v5.1.7: Compare documents and generate redline markup

Supports: PDF, DOCX, PPTX, XLSX
Features:
- Fast paragraph comparison with hash-based pre-filtering
- Superior table handling (row additions, column reordering, re-sorting)
- Batch processing with parallel execution
- Multi-format support
"""

import os
import sys
import json
import hashlib
import difflib
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass, field
from concurrent.futures import ProcessPoolExecutor, as_completed
import multiprocessing as mp

# Document processing libraries
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from pptx import Presentation
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# =============================================================================
# Progress Emission
# =============================================================================

def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class Cell:
    """Table cell with position and content."""
    row: int
    col: int
    text: str

    @property
    def fingerprint(self) -> str:
        """Content-based fingerprint for matching."""
        normalized = self.text.strip().lower()
        return hashlib.md5(normalized.encode()).hexdigest()[:8]


@dataclass
class TableRow:
    """Table row with cells."""
    index: int
    cells: List[Cell]
    is_header: bool = False

    @property
    def fingerprint(self) -> str:
        """Row fingerprint for move/reorder detection."""
        cell_fps = [c.fingerprint for c in self.cells]
        return hashlib.md5('|'.join(cell_fps).encode()).hexdigest()[:12]

    @property
    def key_values(self) -> List[str]:
        """First 2 columns often contain unique identifiers."""
        return [c.text.strip() for c in self.cells[:2]]


@dataclass
class Table:
    """Table structure with rows."""
    id: str
    rows: List[TableRow]
    position: Tuple[int, int]  # (page/slide/sheet, index)

    @property
    def header_row(self) -> Optional[TableRow]:
        if self.rows and self.rows[0].is_header:
            return self.rows[0]
        return None

    @property
    def header_fingerprint(self) -> str:
        if self.header_row:
            return self.header_row.fingerprint
        return ""

    @property
    def column_count(self) -> int:
        return max(len(row.cells) for row in self.rows) if self.rows else 0


@dataclass
class DiffBlock:
    """Represents a difference between documents."""
    operation: str  # 'equal', 'insert', 'delete', 'replace', 'move'
    original_text: str = ""
    modified_text: str = ""
    original_index: int = -1
    modified_index: int = -1


@dataclass
class CellChange:
    """Change in a table cell."""
    change_type: str  # 'unchanged', 'added', 'deleted', 'modified'
    original_value: Optional[str] = None
    modified_value: Optional[str] = None
    row: int = 0
    col: int = 0


@dataclass
class RowChange:
    """Change in a table row."""
    change_type: str  # 'unchanged', 'added', 'deleted', 'modified', 'moved'
    original_index: Optional[int] = None
    modified_index: Optional[int] = None
    cells: List[CellChange] = field(default_factory=list)
    moved_from: Optional[int] = None


@dataclass
class TableDiff:
    """Complete diff between two tables."""
    row_changes: List[RowChange]
    column_mapping: Dict[int, int]  # {orig_col: mod_col}
    is_resorted: bool
    added_columns: List[int]
    deleted_columns: List[int]


@dataclass
class DocumentContent:
    """Extracted content from a document."""
    paragraphs: List[Dict[str, Any]]
    tables: List[Table]
    file_path: str
    format: str


# =============================================================================
# Content Extractors
# =============================================================================

def extract_from_pdf(file_path: str) -> DocumentContent:
    """Extract content from PDF using PyMuPDF."""
    if not HAS_FITZ:
        raise ImportError("PyMuPDF (fitz) not installed")

    paragraphs = []
    tables = []

    doc = fitz.open(file_path)

    for page_num, page in enumerate(doc):
        # Extract text blocks as paragraphs
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if block["type"] == 0:  # Text block
                para_text = ""
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        para_text += span["text"]

                if para_text.strip():
                    paragraphs.append({
                        "text": para_text.strip(),
                        "page": page_num,
                        "index": len(paragraphs)
                    })

        # Extract tables
        try:
            page_tables = page.find_tables()
            for table_idx, table_data in enumerate(page_tables):
                rows = []
                extracted = table_data.extract()

                for row_idx, row_data in enumerate(extracted):
                    cells = []
                    for col_idx, cell_text in enumerate(row_data):
                        cells.append(Cell(
                            row=row_idx,
                            col=col_idx,
                            text=cell_text or ""
                        ))

                    rows.append(TableRow(
                        index=row_idx,
                        cells=cells,
                        is_header=(row_idx == 0)
                    ))

                if rows:
                    tables.append(Table(
                        id=f"p{page_num}_t{table_idx}",
                        rows=rows,
                        position=(page_num, table_idx)
                    ))
        except Exception:
            pass  # Table detection may fail on some pages

    doc.close()

    return DocumentContent(
        paragraphs=paragraphs,
        tables=tables,
        file_path=file_path,
        format='pdf'
    )


def extract_from_docx(file_path: str) -> DocumentContent:
    """Extract content from Word document."""
    if not HAS_DOCX:
        raise ImportError("python-docx not installed")

    paragraphs = []
    tables = []

    doc = Document(file_path)

    # Extract paragraphs
    for para_idx, para in enumerate(doc.paragraphs):
        if para.text.strip():
            paragraphs.append({
                "text": para.text.strip(),
                "style": para.style.name if para.style else "Normal",
                "index": para_idx
            })

    # Extract tables
    for table_idx, docx_table in enumerate(doc.tables):
        rows = []

        for row_idx, docx_row in enumerate(docx_table.rows):
            cells = []

            for col_idx, docx_cell in enumerate(docx_row.cells):
                cells.append(Cell(
                    row=row_idx,
                    col=col_idx,
                    text=docx_cell.text
                ))

            rows.append(TableRow(
                index=row_idx,
                cells=cells,
                is_header=(row_idx == 0)
            ))

        if rows:
            tables.append(Table(
                id=f"table_{table_idx}",
                rows=rows,
                position=(0, table_idx)
            ))

    return DocumentContent(
        paragraphs=paragraphs,
        tables=tables,
        file_path=file_path,
        format='docx'
    )


def extract_from_pptx(file_path: str) -> DocumentContent:
    """Extract content from PowerPoint."""
    if not HAS_PPTX:
        raise ImportError("python-pptx not installed")

    paragraphs = []
    tables = []

    prs = Presentation(file_path)

    for slide_idx, slide in enumerate(prs.slides):
        # Extract text from shapes
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                paragraphs.append({
                    "text": shape.text.strip(),
                    "slide": slide_idx,
                    "index": len(paragraphs)
                })

            # Extract tables from shapes
            if shape.has_table:
                pptx_table = shape.table
                rows = []

                for row_idx in range(len(pptx_table.rows)):
                    cells = []
                    for col_idx in range(len(pptx_table.columns)):
                        cell = pptx_table.cell(row_idx, col_idx)
                        cells.append(Cell(
                            row=row_idx,
                            col=col_idx,
                            text=cell.text
                        ))

                    rows.append(TableRow(
                        index=row_idx,
                        cells=cells,
                        is_header=(row_idx == 0)
                    ))

                if rows:
                    tables.append(Table(
                        id=f"s{slide_idx}_t{len(tables)}",
                        rows=rows,
                        position=(slide_idx, len(tables))
                    ))

    return DocumentContent(
        paragraphs=paragraphs,
        tables=tables,
        file_path=file_path,
        format='pptx'
    )


def extract_from_xlsx(file_path: str) -> DocumentContent:
    """Extract content from Excel. Each sheet becomes a table."""
    if not HAS_OPENPYXL:
        raise ImportError("openpyxl not installed")

    paragraphs = []  # Excel doesn't have paragraphs
    tables = []

    wb = load_workbook(file_path, data_only=True)

    for sheet_idx, sheet_name in enumerate(wb.sheetnames):
        sheet = wb[sheet_name]
        rows = []

        for row_idx, row in enumerate(sheet.iter_rows()):
            cells = []
            has_content = False

            for col_idx, cell in enumerate(row):
                cell_value = str(cell.value) if cell.value is not None else ""
                if cell_value:
                    has_content = True

                cells.append(Cell(
                    row=row_idx,
                    col=col_idx,
                    text=cell_value
                ))

            if has_content:
                rows.append(TableRow(
                    index=row_idx,
                    cells=cells,
                    is_header=(row_idx == 0)
                ))

        if rows:
            tables.append(Table(
                id=f"sheet_{sheet_name}",
                rows=rows,
                position=(sheet_idx, 0)
            ))

    wb.close()

    return DocumentContent(
        paragraphs=paragraphs,
        tables=tables,
        file_path=file_path,
        format='xlsx'
    )


def extract_content(file_path: str) -> DocumentContent:
    """Extract content based on file type."""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.pdf':
        return extract_from_pdf(file_path)
    elif ext in ['.docx', '.doc']:
        return extract_from_docx(file_path)
    elif ext in ['.pptx', '.ppt']:
        return extract_from_pptx(file_path)
    elif ext in ['.xlsx', '.xls']:
        return extract_from_xlsx(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")


# =============================================================================
# Paragraph Comparison (with hash-based pre-filtering)
# =============================================================================

def compute_paragraph_hash(text: str) -> str:
    """Hash normalized paragraph for fast comparison."""
    normalized = ' '.join(text.lower().split())
    return hashlib.md5(normalized.encode()).hexdigest()


def compare_paragraphs(orig_paras: List[Dict], mod_paras: List[Dict]) -> List[DiffBlock]:
    """Compare paragraphs using difflib with hash pre-filtering."""

    # Extract text and compute hashes
    orig_texts = [p['text'] for p in orig_paras]
    mod_texts = [p['text'] for p in mod_paras]

    orig_hashes = [compute_paragraph_hash(t) for t in orig_texts]
    mod_hashes = [compute_paragraph_hash(t) for t in mod_texts]

    # Use SequenceMatcher on hashes for speed
    matcher = difflib.SequenceMatcher(None, orig_hashes, mod_hashes)

    diffs = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            for k in range(i2 - i1):
                diffs.append(DiffBlock(
                    operation='equal',
                    original_text=orig_texts[i1 + k],
                    modified_text=mod_texts[j1 + k],
                    original_index=i1 + k,
                    modified_index=j1 + k
                ))

        elif tag == 'delete':
            for k in range(i1, i2):
                diffs.append(DiffBlock(
                    operation='delete',
                    original_text=orig_texts[k],
                    original_index=k
                ))

        elif tag == 'insert':
            for k in range(j1, j2):
                diffs.append(DiffBlock(
                    operation='insert',
                    modified_text=mod_texts[k],
                    modified_index=k
                ))

        elif tag == 'replace':
            # Do word-level diff for replace blocks
            for k in range(max(i2 - i1, j2 - j1)):
                orig_text = orig_texts[i1 + k] if i1 + k < i2 else ""
                mod_text = mod_texts[j1 + k] if j1 + k < j2 else ""

                diffs.append(DiffBlock(
                    operation='replace',
                    original_text=orig_text,
                    modified_text=mod_text,
                    original_index=i1 + k if i1 + k < i2 else -1,
                    modified_index=j1 + k if j1 + k < j2 else -1
                ))

    return diffs


# =============================================================================
# Table Comparison
# =============================================================================

def match_tables(orig_tables: List[Table], mod_tables: List[Table]) -> List[Tuple[Optional[Table], Optional[Table]]]:
    """Match original tables to modified tables."""
    matches = []
    used_modified = set()

    for orig in orig_tables:
        best_match = None
        best_score = 0.5  # Minimum threshold

        for i, mod in enumerate(mod_tables):
            if i in used_modified:
                continue

            score = compute_table_similarity(orig, mod)
            if score > best_score:
                best_score = score
                best_match = (i, mod)

        if best_match:
            used_modified.add(best_match[0])
            matches.append((orig, best_match[1]))
        else:
            matches.append((orig, None))  # Deleted table

    # Remaining modified tables are additions
    for i, mod in enumerate(mod_tables):
        if i not in used_modified:
            matches.append((None, mod))

    return matches


def compute_table_similarity(t1: Table, t2: Table) -> float:
    """Compute similarity score between two tables (0-1)."""
    scores = []

    # Header match (40% weight)
    if t1.header_fingerprint and t2.header_fingerprint:
        header_match = 1.0 if t1.header_fingerprint == t2.header_fingerprint else 0.0
        scores.append(('header', header_match, 0.4))

    # Position match (30% weight)
    if t1.position[0] == t2.position[0]:  # Same page/slide
        scores.append(('position', 1.0, 0.3))
    else:
        scores.append(('position', 0.0, 0.3))

    # Content overlap (30% weight)
    fp1 = set(row.fingerprint for row in t1.rows)
    fp2 = set(row.fingerprint for row in t2.rows)
    intersection = len(fp1 & fp2)
    union = len(fp1 | fp2)
    content_score = intersection / union if union > 0 else 0
    scores.append(('content', content_score, 0.3))

    # Weighted average
    total_weight = sum(s[2] for s in scores)
    if total_weight == 0:
        return 0
    return sum(s[1] * s[2] for s in scores) / total_weight


def detect_column_reorder(orig_table: Table, mod_table: Table) -> Dict[int, int]:
    """Detect column reordering by matching header cells."""
    if not orig_table.header_row or not mod_table.header_row:
        # No headers, assume columns unchanged
        return {i: i for i in range(orig_table.column_count)}

    orig_headers = [c.text.strip().lower() for c in orig_table.header_row.cells]
    mod_headers = [c.text.strip().lower() for c in mod_table.header_row.cells]

    mapping = {}
    for i, orig_h in enumerate(orig_headers):
        if orig_h in mod_headers:
            mapping[i] = mod_headers.index(orig_h)

    return mapping


def reorder_row_cells(row: TableRow, col_mapping: Dict[int, int]) -> List[str]:
    """Reorder row cells according to column mapping."""
    result = [''] * len(col_mapping)
    for orig_col, mod_col in col_mapping.items():
        if orig_col < len(row.cells) and mod_col < len(result):
            result[mod_col] = row.cells[orig_col].text.strip().lower()
    return result


def compute_reordered_fingerprint(row: TableRow, col_mapping: Dict[int, int]) -> str:
    """Compute row fingerprint accounting for column reorder."""
    reordered = reorder_row_cells(row, col_mapping)
    return hashlib.md5('|'.join(reordered).encode()).hexdigest()[:12]


def match_rows(orig_table: Table, mod_table: Table, col_mapping: Dict[int, int]) -> Dict[int, Tuple[int, float]]:
    """Match rows between tables, accounting for column reorder."""
    matches = {}
    used_modified = set()

    # Skip header row
    orig_data_rows = [(i, r) for i, r in enumerate(orig_table.rows) if not r.is_header]
    mod_data_rows = [(i, r) for i, r in enumerate(mod_table.rows) if not r.is_header]

    # Pass 1: Exact fingerprint matches
    for orig_idx, orig_row in orig_data_rows:
        orig_fp = compute_reordered_fingerprint(orig_row, col_mapping)

        for mod_idx, mod_row in mod_data_rows:
            if mod_idx in used_modified:
                continue

            mod_fp = hashlib.md5('|'.join(c.text.strip().lower() for c in mod_row.cells).encode()).hexdigest()[:12]

            if orig_fp == mod_fp:
                matches[orig_idx] = (mod_idx, 1.0)
                used_modified.add(mod_idx)
                break

    # Pass 2: Key column matching (first 2 columns)
    for orig_idx, orig_row in orig_data_rows:
        if orig_idx in matches:
            continue

        orig_keys = orig_row.key_values
        best_match = None
        best_score = 0.5

        for mod_idx, mod_row in mod_data_rows:
            if mod_idx in used_modified:
                continue

            mod_keys = mod_row.key_values

            # Count matching keys
            match_count = sum(1 for k1, k2 in zip(orig_keys, mod_keys)
                            if k1.strip().lower() == k2.strip().lower())
            score = match_count / max(len(orig_keys), 1)

            if score > best_score:
                best_score = score
                best_match = mod_idx

        if best_match is not None:
            matches[orig_idx] = (best_match, best_score)
            used_modified.add(best_match)

    return matches


def is_table_resorted(orig_table: Table, mod_table: Table, row_matches: Dict[int, Tuple[int, float]]) -> bool:
    """Check if tables have same rows but in different order."""
    # Get data row indices (excluding header)
    orig_data_indices = [i for i, r in enumerate(orig_table.rows) if not r.is_header]

    # All rows must be matched
    if len(row_matches) != len(orig_data_indices):
        return False

    # Check if order changed
    mod_indices = [row_matches[i][0] for i in orig_data_indices if i in row_matches]

    return mod_indices != sorted(mod_indices)


def compare_tables(orig_tables: List[Table], mod_tables: List[Table]) -> List[Tuple[Optional[Table], Optional[Table], Optional[TableDiff]]]:
    """Compare all tables between documents."""
    results = []

    # Match tables
    table_pairs = match_tables(orig_tables, mod_tables)

    for orig, mod in table_pairs:
        if orig is None:
            # Added table
            results.append((None, mod, None))
        elif mod is None:
            # Deleted table
            results.append((orig, None, None))
        else:
            # Compare matched tables
            diff = diff_tables(orig, mod)
            results.append((orig, mod, diff))

    return results


def diff_tables(orig: Table, mod: Table) -> TableDiff:
    """Generate detailed diff between two matched tables."""

    # Detect column reordering
    col_mapping = detect_column_reorder(orig, mod)

    # Find added/deleted columns
    orig_cols = set(range(orig.column_count))
    mod_cols = set(range(mod.column_count))
    mapped_orig = set(col_mapping.keys())
    mapped_mod = set(col_mapping.values())

    deleted_columns = list(orig_cols - mapped_orig)
    added_columns = list(mod_cols - mapped_mod)

    # Match rows
    row_matches = match_rows(orig, mod, col_mapping)

    # Detect re-sorting
    is_resorted = is_table_resorted(orig, mod, row_matches)

    # Generate row changes
    row_changes = []
    used_mod_rows = set()

    for orig_idx, orig_row in enumerate(orig.rows):
        if orig_idx in row_matches:
            mod_idx, confidence = row_matches[orig_idx]
            mod_row = mod.rows[mod_idx]
            used_mod_rows.add(mod_idx)

            # Compare cells
            cell_changes = []
            for orig_col, mod_col in col_mapping.items():
                if orig_col >= len(orig_row.cells):
                    continue

                orig_cell = orig_row.cells[orig_col]

                if mod_col < len(mod_row.cells):
                    mod_cell = mod_row.cells[mod_col]

                    if orig_cell.text.strip() == mod_cell.text.strip():
                        change_type = 'unchanged'
                    else:
                        change_type = 'modified'

                    cell_changes.append(CellChange(
                        change_type=change_type,
                        original_value=orig_cell.text,
                        modified_value=mod_cell.text,
                        row=orig_idx,
                        col=orig_col
                    ))
                else:
                    cell_changes.append(CellChange(
                        change_type='deleted',
                        original_value=orig_cell.text,
                        row=orig_idx,
                        col=orig_col
                    ))

            # Determine row change type
            if mod_idx != orig_idx:
                row_type = 'moved'
            elif any(c.change_type == 'modified' for c in cell_changes):
                row_type = 'modified'
            else:
                row_type = 'unchanged'

            row_changes.append(RowChange(
                change_type=row_type,
                original_index=orig_idx,
                modified_index=mod_idx,
                cells=cell_changes,
                moved_from=orig_idx if mod_idx != orig_idx else None
            ))
        else:
            # Deleted row
            cell_changes = [CellChange(
                change_type='deleted',
                original_value=c.text,
                row=orig_idx,
                col=c.col
            ) for c in orig_row.cells]

            row_changes.append(RowChange(
                change_type='deleted',
                original_index=orig_idx,
                cells=cell_changes
            ))

    # Added rows
    for mod_idx, mod_row in enumerate(mod.rows):
        if mod_idx not in used_mod_rows:
            cell_changes = [CellChange(
                change_type='added',
                modified_value=c.text,
                row=mod_idx,
                col=c.col
            ) for c in mod_row.cells]

            row_changes.append(RowChange(
                change_type='added',
                modified_index=mod_idx,
                cells=cell_changes
            ))

    return TableDiff(
        row_changes=row_changes,
        column_mapping=col_mapping,
        is_resorted=is_resorted,
        added_columns=added_columns,
        deleted_columns=deleted_columns
    )


# =============================================================================
# Output Generation
# =============================================================================

# Color scheme
DELETED_COLOR = RGBColor(255, 0, 0)    # Red
ADDED_COLOR = RGBColor(0, 0, 255)       # Blue
MOVED_COLOR = RGBColor(0, 128, 0)       # Green


def generate_redline_docx(para_diffs: List[DiffBlock],
                          table_results: List[Tuple],
                          output_path: str,
                          orig_file: str,
                          mod_file: str):
    """Generate Word document with redline markup."""
    if not HAS_DOCX:
        raise ImportError("python-docx not installed")

    doc = Document()

    # Add header
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run("DOCUMENT COMPARISON")
    run.bold = True
    run.font.size = Pt(14)

    # Add metadata
    meta = doc.add_paragraph()
    meta.add_run(f"Original: {os.path.basename(orig_file)}\n").italic = True
    meta.add_run(f"Modified: {os.path.basename(mod_file)}\n").italic = True
    meta.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n").italic = True

    doc.add_paragraph()  # Spacer

    # Add legend
    legend = doc.add_paragraph()
    legend.add_run("Legend: ").bold = True

    del_run = legend.add_run("Deleted text")
    del_run.font.strike = True
    del_run.font.color.rgb = DELETED_COLOR

    legend.add_run(" | ")

    add_run = legend.add_run("Added text")
    add_run.font.underline = True
    add_run.font.color.rgb = ADDED_COLOR

    legend.add_run(" | ")

    move_run = legend.add_run("Moved text")
    move_run.font.italic = True
    move_run.font.color.rgb = MOVED_COLOR

    doc.add_paragraph()  # Spacer

    # Add paragraph differences
    section_header = doc.add_paragraph()
    section_header.add_run("TEXT CHANGES").bold = True

    for diff in para_diffs:
        para = doc.add_paragraph()

        if diff.operation == 'equal':
            para.add_run(diff.original_text)

        elif diff.operation == 'delete':
            run = para.add_run(diff.original_text)
            run.font.strike = True
            run.font.color.rgb = DELETED_COLOR

        elif diff.operation == 'insert':
            run = para.add_run(diff.modified_text)
            run.font.underline = True
            run.font.color.rgb = ADDED_COLOR

        elif diff.operation == 'replace':
            # Show deleted text
            if diff.original_text:
                del_run = para.add_run(diff.original_text)
                del_run.font.strike = True
                del_run.font.color.rgb = DELETED_COLOR
                para.add_run(" ")

            # Show added text
            if diff.modified_text:
                add_run = para.add_run(diff.modified_text)
                add_run.font.underline = True
                add_run.font.color.rgb = ADDED_COLOR

    # Add table differences
    if table_results:
        doc.add_paragraph()  # Spacer
        section_header = doc.add_paragraph()
        section_header.add_run("TABLE CHANGES").bold = True

        for orig_table, mod_table, table_diff in table_results:
            # Table header
            table_para = doc.add_paragraph()

            if orig_table is None:
                table_para.add_run(f"[Table Added: {mod_table.id}]").font.color.rgb = ADDED_COLOR
            elif mod_table is None:
                table_para.add_run(f"[Table Deleted: {orig_table.id}]").font.color.rgb = DELETED_COLOR
            elif table_diff:
                table_para.add_run(f"Table: {orig_table.id}")

                if table_diff.is_resorted:
                    table_para.add_run(" [RESORTED]").font.color.rgb = MOVED_COLOR

                # Summarize changes
                added_rows = sum(1 for r in table_diff.row_changes if r.change_type == 'added')
                deleted_rows = sum(1 for r in table_diff.row_changes if r.change_type == 'deleted')
                modified_rows = sum(1 for r in table_diff.row_changes if r.change_type == 'modified')
                moved_rows = sum(1 for r in table_diff.row_changes if r.change_type == 'moved')

                summary = []
                if added_rows:
                    summary.append(f"+{added_rows} rows")
                if deleted_rows:
                    summary.append(f"-{deleted_rows} rows")
                if modified_rows:
                    summary.append(f"~{modified_rows} modified")
                if moved_rows:
                    summary.append(f"↕{moved_rows} moved")

                if summary:
                    table_para.add_run(f" ({', '.join(summary)})")

                # Show column changes
                if table_diff.added_columns:
                    col_para = doc.add_paragraph()
                    col_para.add_run(f"  Columns added: {table_diff.added_columns}").font.color.rgb = ADDED_COLOR

                if table_diff.deleted_columns:
                    col_para = doc.add_paragraph()
                    col_para.add_run(f"  Columns deleted: {table_diff.deleted_columns}").font.color.rgb = DELETED_COLOR

    # Save document
    doc.save(output_path)


# =============================================================================
# Main Comparison Function
# =============================================================================

def compare_documents(original_path: str, modified_path: str, output_path: str) -> Dict:
    """Compare two documents and generate redline output."""

    # Extract content
    emit("progress", percent=10, message=f"Extracting content from {os.path.basename(original_path)}...")
    orig_content = extract_content(original_path)

    emit("progress", percent=30, message=f"Extracting content from {os.path.basename(modified_path)}...")
    mod_content = extract_content(modified_path)

    # Compare paragraphs
    emit("progress", percent=50, message="Comparing paragraphs...")
    para_diffs = compare_paragraphs(orig_content.paragraphs, mod_content.paragraphs)

    # Compare tables
    emit("progress", percent=70, message="Comparing tables...")
    table_results = compare_tables(orig_content.tables, mod_content.tables)

    # Generate output
    emit("progress", percent=90, message="Generating redline document...")
    generate_redline_docx(para_diffs, table_results, output_path, original_path, modified_path)

    # Calculate statistics
    para_changes = sum(1 for d in para_diffs if d.operation != 'equal')
    table_changes = sum(1 for _, _, d in table_results if d is not None)

    return {
        'output_path': output_path,
        'paragraphs_compared': len(orig_content.paragraphs),
        'paragraph_changes': para_changes,
        'tables_compared': len(orig_content.tables),
        'table_changes': table_changes
    }


# =============================================================================
# Batch Processing
# =============================================================================

def process_single_pair(args: Tuple[str, str, str]) -> Dict:
    """Process a single document pair (for multiprocessing)."""
    original, modified, output = args
    try:
        result = compare_documents(original, modified, output)
        result['success'] = True
        return result
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'original': original,
            'modified': modified
        }


def process_batch(pairs: List[Dict], output_folder: str) -> List[Dict]:
    """Process multiple document pairs in parallel."""
    results = []
    total = len(pairs)

    # Prepare arguments
    args_list = []
    for i, pair in enumerate(pairs):
        original = pair['original']
        modified = pair['modified']

        # Generate output filename
        orig_name = os.path.splitext(os.path.basename(original))[0]
        mod_name = os.path.splitext(os.path.basename(modified))[0]
        output_name = f"Redline_{orig_name}_vs_{mod_name}.docx"
        output_path = os.path.join(output_folder, output_name)

        args_list.append((original, modified, output_path))

    # Process in parallel (use half of available CPUs)
    max_workers = max(1, mp.cpu_count() // 2)

    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(process_single_pair, args): i
                   for i, args in enumerate(args_list)}

        completed = 0
        for future in as_completed(futures):
            idx = futures[future]
            try:
                result = future.result()
                results.append(result)
            except Exception as e:
                results.append({
                    'success': False,
                    'error': str(e),
                    'index': idx
                })

            completed += 1
            emit("progress",
                 percent=int(completed / total * 100),
                 message=f"Completed {completed}/{total} comparisons")

    return results


# =============================================================================
# Main Entry Point
# =============================================================================

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

    try:
        if config.get('batch'):
            # Batch processing
            pairs = config.get('pairs', [])
            output_folder = config.get('output_folder', os.path.dirname(config_path))

            os.makedirs(output_folder, exist_ok=True)

            results = process_batch(pairs, output_folder)

            successful = sum(1 for r in results if r.get('success'))

            emit("result",
                 success=True,
                 mode="batch",
                 total=len(pairs),
                 successful=successful,
                 results=results)

        else:
            # Single comparison
            original = config.get('original')
            modified = config.get('modified')
            output = config.get('output')

            if not original or not modified:
                emit("error", message="Missing 'original' or 'modified' path")
                sys.exit(1)

            if not output:
                # Generate output path
                orig_name = os.path.splitext(os.path.basename(original))[0]
                mod_name = os.path.splitext(os.path.basename(modified))[0]
                output = os.path.join(
                    os.path.dirname(original),
                    f"Redline_{orig_name}_vs_{mod_name}.docx"
                )

            result = compare_documents(original, modified, output)

            emit("result",
                 success=True,
                 mode="single",
                 **result)

    except Exception as e:
        emit("error", message=f"Processing failed: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
