#!/usr/bin/env python3
"""
EmmaNeigh - Signature Packet Automation
Main entry point for the Python processor.

This script handles:
1. Signature packet creation
2. Execution version creation (merging signed pages back into documents)

All progress and results are output as JSON lines to stdout for
communication with the Electron frontend.
"""

import json
import sys
import os

# Add the current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from processors.signature_packets import process_signature_packets
from processors.execution_version import create_execution_version


def emit_progress(stage: str, percent: int, message: str):
    """Send progress update to Electron frontend."""
    output = {
        "type": "progress",
        "stage": stage,
        "percent": percent,
        "message": message
    }
    print(json.dumps(output), flush=True)


def emit_error(error: str):
    """Send error to Electron frontend."""
    output = {
        "type": "error",
        "error": error
    }
    print(json.dumps(output), flush=True)


def emit_result(result: dict):
    """Send result to Electron frontend."""
    output = {
        "type": "result",
        **result
    }
    print(json.dumps(output), flush=True)


def main():
    if len(sys.argv) < 2:
        emit_error("No command provided. Use 'signature-packets' or 'execution-version'.")
        sys.exit(1)

    command = sys.argv[1]

    try:
        if command == "signature-packets":
            if len(sys.argv) < 3:
                emit_error("No input folder provided.")
                sys.exit(1)

            input_folder = sys.argv[2]
            result = process_signature_packets(input_folder, emit_progress)
            emit_result(result)

        elif command == "execution-version":
            if len(sys.argv) < 5:
                emit_error("Usage: execution-version <original_pdf> <signed_pdf> <insert_after_page>")
                sys.exit(1)

            original_pdf = sys.argv[2]
            signed_pdf = sys.argv[3]
            insert_after = int(sys.argv[4])

            result = create_execution_version(original_pdf, signed_pdf, insert_after, emit_progress)
            emit_result(result)

        else:
            emit_error(f"Unknown command: {command}")
            sys.exit(1)

    except Exception as e:
        emit_error(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
