import sys
import unittest
from pathlib import Path


TESTS_DIR = Path(__file__).resolve().parent
PYTHON_DIR = TESTS_DIR.parent
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

try:
    from incumbency_parser import extract_signers_from_table
except Exception:
    extract_signers_from_table = None


@unittest.skipUnless(extract_signers_from_table is not None, "incumbency parser dependencies are required")
class IncumbencyParserTests(unittest.TestCase):
    def test_extracts_headerless_name_title_rows(self):
        table_data = [
            ["Jane Doe", "Chief Financial Officer", ""],
            ["John Smith", "Vice President", ""],
        ]

        signers = extract_signers_from_table(table_data)

        self.assertEqual(
            [
                {"name": "Jane Doe", "title": "Chief Financial Officer"},
                {"name": "John Smith", "title": "Vice President"},
            ],
            signers,
        )

    def test_extracts_combined_name_and_title_cells(self):
        table_data = [
            ["Jane Doe - President"],
            ["John Smith - Secretary"],
        ]

        signers = extract_signers_from_table(table_data)

        self.assertEqual(
            [
                {"name": "Jane Doe", "title": "President"},
                {"name": "John Smith", "title": "Secretary"},
            ],
            signers,
        )


if __name__ == "__main__":
    unittest.main()
