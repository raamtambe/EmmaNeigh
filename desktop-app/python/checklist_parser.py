#!/usr/bin/env python3
"""
EmmaNeigh - Transaction Checklist Parser
Parses Excel/CSV transaction checklists to extract document-signatory mappings.
"""

import pandas as pd
import re
import os
import sys
import json


# Keywords to identify document name column
DOC_COLUMN_KEYWORDS = [
    "document", "doc name", "agreement", "instrument", "deliverable", "item"
]

# Keywords to identify signatories column
SIGNATORY_COLUMN_KEYWORDS = [
    "signator", "signer", "parties", "execution", "who signs", "signed by"
]


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def find_column_by_keywords(columns, keywords):
    """Find a column name that matches any of the keywords."""
    columns_lower = {col: col.lower() for col in columns}

    for col, col_lower in columns_lower.items():
        for keyword in keywords:
            if keyword in col_lower:
                return col
    return None


def parse_signatories(signatories_str):
    """
    Parse a signatories string into a list of roles.
    Handles: "Borrower, Agent, Lenders" or "Borrower; Agent; Lenders"
    """
    if not signatories_str or pd.isna(signatories_str):
        return []

    # Convert to string if not already
    signatories_str = str(signatories_str)

    # Split by comma, semicolon, or "and"
    parts = re.split(r'[,;]|\band\b', signatories_str)

    # Clean up each part
    roles = []
    for part in parts:
        role = part.strip()
        if role and len(role) > 1:
            roles.append(role)

    return roles


def parse_checklist(filepath):
    """
    Parse a transaction checklist Excel/CSV file.

    Returns:
        dict: {
            "documents": {doc_name: [signatory_roles]},
            "all_roles": [unique_roles],
            "doc_column": str,
            "signatory_column": str
        }
    """
    # Determine file type and read
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.csv':
        df = pd.read_csv(filepath)
    elif ext in ['.xlsx', '.xls']:
        df = pd.read_excel(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    # Find document column
    doc_col = find_column_by_keywords(df.columns, DOC_COLUMN_KEYWORDS)
    if not doc_col:
        # Fall back to first column
        doc_col = df.columns[0]

    # Find signatories column
    sig_col = find_column_by_keywords(df.columns, SIGNATORY_COLUMN_KEYWORDS)
    if not sig_col:
        # Try to find a column that looks like it contains multiple parties
        for col in df.columns:
            if col != doc_col:
                # Check if values contain commas or semicolons
                sample = df[col].dropna().head(5)
                if any(',' in str(v) or ';' in str(v) for v in sample):
                    sig_col = col
                    break

    if not sig_col:
        raise ValueError("Could not identify signatories column. Please ensure your checklist has a column with headers like 'Signatories', 'Parties', or 'Signed By'.")

    # Build document -> signatories mapping
    documents = {}
    all_roles = set()

    for _, row in df.iterrows():
        doc_name = row[doc_col]
        if pd.isna(doc_name) or not str(doc_name).strip():
            continue

        doc_name = str(doc_name).strip()
        signatories = parse_signatories(row[sig_col])

        if signatories:
            documents[doc_name] = signatories
            all_roles.update(signatories)

    return {
        "documents": documents,
        "all_roles": sorted(list(all_roles)),
        "doc_column": doc_col,
        "signatory_column": sig_col,
        "total_documents": len(documents)
    }


def main():
    """CLI entry point for testing."""
    if len(sys.argv) < 2:
        emit("error", message="Usage: checklist_parser.py <checklist_file>")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.isfile(filepath):
        emit("error", message=f"File not found: {filepath}")
        sys.exit(1)

    try:
        emit("progress", percent=10, message="Parsing checklist...")
        result = parse_checklist(filepath)

        emit("progress", percent=100, message="Complete!")
        emit("result", success=True, **result)

    except Exception as e:
        emit("error", message=str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
