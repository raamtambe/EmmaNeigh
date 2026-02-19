#!/usr/bin/env python3
"""
EmmaNeigh - Unified Processor Dispatcher
v5.1.8: Single executable that dispatches to all Python processors.

This eliminates library duplication across 13 separate PyInstaller executables,
reducing total size from ~2.7GB to ~300MB by sharing fitz, docx, pandas, etc.

Usage: emna_processor <module_name> [args...]

Example: emna_processor signature_packets /path/to/config.json
"""

import sys
import os


def main():
    if len(sys.argv) < 2:
        print("Usage: emna_processor <module_name> [args...]", file=sys.stderr)
        print("Available modules:", file=sys.stderr)
        print("  signature_packets, execution_version, checklist_parser,", file=sys.stderr)
        print("  incumbency_parser, sigblock_workflow, document_collator,", file=sys.stderr)
        print("  email_csv_parser, time_tracker, checklist_updater,", file=sys.stderr)
        print("  punchlist_generator, email_nl_search, packet_shell_generator,", file=sys.stderr)
        print("  document_redline", file=sys.stderr)
        sys.exit(1)

    module_name = sys.argv[1]

    # Remove the module name from argv so the target module sees the right args
    # e.g., "emna_processor signature_packets /path/config.json"
    # becomes sys.argv = ["emna_processor", "/path/config.json"]
    sys.argv = [sys.argv[0]] + sys.argv[2:]

    # Dispatch to the correct module
    if module_name == 'signature_packets':
        import signature_packets
        signature_packets.main()
    elif module_name == 'execution_version':
        import execution_version
        execution_version.main()
    elif module_name == 'checklist_parser':
        import checklist_parser
        checklist_parser.main()
    elif module_name == 'incumbency_parser':
        import incumbency_parser
        incumbency_parser.main()
    elif module_name == 'sigblock_workflow':
        import sigblock_workflow
        sigblock_workflow.main()
    elif module_name == 'document_collator':
        import document_collator
        document_collator.main()
    elif module_name == 'email_csv_parser':
        import email_csv_parser
        email_csv_parser.main()
    elif module_name == 'time_tracker':
        import time_tracker
        time_tracker.main()
    elif module_name == 'checklist_updater':
        import checklist_updater
        checklist_updater.main()
    elif module_name == 'punchlist_generator':
        import punchlist_generator
        punchlist_generator.main()
    elif module_name == 'email_nl_search':
        import email_nl_search
        email_nl_search.main()
    elif module_name == 'packet_shell_generator':
        import packet_shell_generator
        packet_shell_generator.main()
    elif module_name == 'document_redline':
        import document_redline
        document_redline.main()
    else:
        print(f"Unknown module: {module_name}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
