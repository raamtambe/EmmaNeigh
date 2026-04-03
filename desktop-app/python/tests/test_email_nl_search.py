import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from email_nl_search import (
    prepare_email_context,
    score_attachment_title_match,
    split_attachment_names,
)


class EmailNaturalLanguageSearchTests(unittest.TestCase):
    def test_split_attachment_names_prefers_semicolons(self):
        attachments = split_attachment_names("Purchase Agreement, Draft.docx; Disclosure Schedule.pdf")
        self.assertEqual(
            attachments,
            ["Purchase Agreement, Draft.docx", "Disclosure Schedule.pdf"],
        )

    def test_attachment_matching_ignores_filename_separators(self):
        match = score_attachment_title_match(
            ["Acme_Purchase-Agreement_v3.docx"],
            "purchase agreement",
        )
        self.assertGreater(match["score"], 0)
        self.assertIn("Acme_Purchase-Agreement_v3.docx", match["matched_attachment_titles"])

    def test_prepare_email_context_prioritizes_attachment_name_hits(self):
        emails = [
            {
                "subject": "Random closing call",
                "from": "Paralegal",
                "to": "Me",
                "date_received": "2026-04-01T09:00:00",
                "body": "Please join the call.",
                "attachments": "",
                "has_attachments": False,
            },
            {
                "subject": "Docs attached",
                "from": "Seller Counsel",
                "to": "Me",
                "date_received": "2026-04-02T10:00:00",
                "body": "See attached.",
                "attachments": "Acme_Purchase_Agreement_v5.docx; Disclosure Schedule.pdf",
                "has_attachments": True,
            },
        ]

        context = prepare_email_context(emails, "purchase agreement", max_emails=10)

        self.assertGreaterEqual(len(context), 1)
        self.assertEqual(context[0]["index"], 1)
        self.assertIn("Acme_Purchase_Agreement_v5.docx", context[0]["attachment_titles"])
        self.assertIn("Acme_Purchase_Agreement_v5.docx", context[0]["matched_attachment_titles"])


if __name__ == "__main__":
    unittest.main()
