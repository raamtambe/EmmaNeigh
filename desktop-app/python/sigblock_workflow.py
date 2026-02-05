#!/usr/bin/env python3
"""
EmmaNeigh - Signature Block Workflow Orchestrator
Coordinates the full workflow: checklist → incumbency → signature blocks → documents
"""

import os
import sys
import json
import shutil

from checklist_parser import parse_checklist
from incumbency_parser import parse_incumbency
from sigblock_generator import generate_signature_page, insert_signature_page_docx, insert_signature_page_pdf


def emit(msg_type, **kwargs):
    """Output JSON message to stdout for the Electron app."""
    print(json.dumps({"type": msg_type, **kwargs}), flush=True)


def process_workflow(config):
    """
    Execute the full signature block workflow.

    Args:
        config: dict with:
            - checklist_path: Path to transaction checklist
            - incumbency_paths: List of paths to incumbency certificates
            - documents_folder: Path to unsigned documents
            - entity_mappings: Dict mapping entity_name -> {role, signer_index}
            - output_folder: Where to save processed documents

    Returns:
        dict: Results of the workflow
    """
    results = {
        "checklist": None,
        "incumbencies": [],
        "documents_processed": [],
        "errors": [],
        "output_folder": None
    }

    try:
        # Step 1: Parse checklist
        emit("progress", percent=5, message="Parsing transaction checklist...")
        checklist_data = parse_checklist(config['checklist_path'])
        results['checklist'] = checklist_data
        emit("progress", percent=15, message=f"Found {checklist_data['total_documents']} documents in checklist")

        # Step 2: Parse incumbency certificates
        emit("progress", percent=20, message="Parsing incumbency certificates...")
        incumbencies = {}
        for inc_path in config.get('incumbency_paths', []):
            try:
                inc_data = parse_incumbency(inc_path)
                entity_name = inc_data.get('entity_name') or os.path.basename(inc_path)
                incumbencies[entity_name] = inc_data
                results['incumbencies'].append(inc_data)
            except Exception as e:
                results['errors'].append({
                    "file": inc_path,
                    "error": f"Failed to parse incumbency: {str(e)}"
                })

        emit("progress", percent=35, message=f"Parsed {len(incumbencies)} incumbency certificates")

        # Step 3: Build entity-signer mapping from user configuration
        entity_signer_map = {}
        for entity_name, mapping in config.get('entity_mappings', {}).items():
            role = mapping.get('role')
            signer_index = mapping.get('signer_index', 0)

            if entity_name in incumbencies:
                inc_data = incumbencies[entity_name]
                signers = inc_data.get('signers', [])

                if signers and signer_index < len(signers):
                    signer = signers[signer_index]
                    entity_signer_map[role] = {
                        'entity_name': entity_name,
                        'signer_name': signer.get('name'),
                        'signer_title': signer.get('title')
                    }

        emit("progress", percent=45, message=f"Mapped {len(entity_signer_map)} entities to roles")

        # Step 4: Process documents
        emit("progress", percent=50, message="Processing documents...")

        output_folder = config.get('output_folder')
        if not output_folder:
            output_folder = os.path.join(
                os.path.dirname(config['documents_folder']),
                "sigblock_output"
            )

        os.makedirs(output_folder, exist_ok=True)
        results['output_folder'] = output_folder

        # Get document files
        documents_folder = config['documents_folder']
        doc_files = [f for f in os.listdir(documents_folder)
                     if f.lower().endswith(('.pdf', '.docx'))]

        if not doc_files:
            emit("error", message="No PDF or Word files found in documents folder.")
            sys.exit(1)

        total = len(doc_files)
        documents = checklist_data['documents']

        for idx, filename in enumerate(doc_files):
            percent = 50 + int((idx / total) * 45)
            emit("progress", percent=percent, message=f"Processing {filename}")

            filepath = os.path.join(documents_folder, filename)
            output_path = os.path.join(output_folder, filename)

            # Match document to checklist
            doc_name_base = os.path.splitext(filename)[0]
            signatories = None

            for checklist_doc, roles in documents.items():
                # Flexible matching
                check_lower = checklist_doc.lower()
                file_lower = doc_name_base.lower()

                if (file_lower in check_lower or
                    check_lower in file_lower or
                    file_lower.replace(' ', '') in check_lower.replace(' ', '')):
                    signatories = roles
                    break

            if not signatories:
                # No match - copy file without changes
                shutil.copy2(filepath, output_path)
                results['documents_processed'].append({
                    "file": filename,
                    "status": "copied",
                    "note": "No matching checklist entry - copied without signature page"
                })
                continue

            try:
                # Generate signature page
                sig_page = generate_signature_page(
                    doc_name=doc_name_base,
                    signatories=signatories,
                    entity_signer_map=entity_signer_map
                )

                # Insert into document
                if filename.lower().endswith('.docx'):
                    insert_signature_page_docx(filepath, sig_page, output_path)
                elif filename.lower().endswith('.pdf'):
                    insert_signature_page_pdf(filepath, sig_page, output_path)

                results['documents_processed'].append({
                    "file": filename,
                    "status": "processed",
                    "signatories": signatories,
                    "blocks_added": len(signatories)
                })

            except Exception as e:
                results['errors'].append({
                    "file": filename,
                    "error": str(e)
                })
                # Copy original file
                shutil.copy2(filepath, output_path)

        emit("progress", percent=100, message="Complete!")

        return results

    except Exception as e:
        emit("error", message=str(e))
        raise


def main():
    """
    CLI entry point.

    Usage: sigblock_workflow.py <config_json>

    The config JSON should contain:
    {
        "checklist_path": "/path/to/checklist.xlsx",
        "incumbency_paths": ["/path/to/inc1.pdf", "/path/to/inc2.pdf"],
        "documents_folder": "/path/to/documents",
        "entity_mappings": {
            "ABC Holdings, LLC": {"role": "Borrower", "signer_index": 0},
            "XYZ Bank, N.A.": {"role": "Administrative Agent", "signer_index": 0}
        },
        "output_folder": "/path/to/output"  // optional
    }
    """
    if len(sys.argv) < 2:
        emit("error", message="Usage: sigblock_workflow.py <config_json_path>")
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
        results = process_workflow(config)

        emit("result",
             success=True,
             outputPath=results['output_folder'],
             documentsProcessed=len(results['documents_processed']),
             errors=len(results['errors']),
             details=results)

    except Exception as e:
        emit("error", message=str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
