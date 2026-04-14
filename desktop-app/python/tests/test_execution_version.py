import sys
import tempfile
import unittest
from pathlib import Path

try:
    import fitz
except Exception:
    fitz = None


TESTS_DIR = Path(__file__).resolve().parent
PYTHON_DIR = TESTS_DIR.parent
if str(PYTHON_DIR) not in sys.path:
    sys.path.insert(0, str(PYTHON_DIR))

try:
    from execution_version import assign_signature_pages, build_page_features, clean_filename_stem, score_signature_page_match
except Exception:
    assign_signature_pages = None
    build_page_features = None
    clean_filename_stem = None
    score_signature_page_match = None


@unittest.skipUnless(fitz is not None and assign_signature_pages is not None, "PyMuPDF execution matcher dependencies are required")
class ExecutionVersionMatchingTests(unittest.TestCase):
    def create_pdf(self, path, page_texts):
        doc = fitz.open()
        for text in page_texts:
            page = doc.new_page()
            page.insert_textbox(
                fitz.Rect(50, 50, page.rect.width - 50, page.rect.height - 50),
                text,
                fontsize=12,
                fontname="helv",
            )
        doc.save(path)
        doc.close()

    def load_first_page_features(self, pdf_path):
        doc = fitz.open(pdf_path)
        try:
            return build_page_features(doc[0], Path(pdf_path).name, 0)
        finally:
            doc.close()

    def test_docusign_noise_still_matches_original_signature_page(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            original_path = tmp_path / "Credit Agreement.pdf"
            signed_path = tmp_path / "Signed Packet.pdf"

            original_text = "\n".join([
                "SIGNATURE PAGE TO CREDIT AGREEMENT",
                "ACME HOLDINGS LLC",
                "By: ____________________________",
                "Name: Jane Doe",
                "Title: Chief Financial Officer",
            ])
            signed_text = "\n".join([
                "DocuSign Envelope ID: 1234567890ABCDEF",
                "Completed by DocuSign",
                "SIGNATURE PAGE TO CREDIT AGREEMENT",
                "ACME HOLDINGS LLC",
                "By: /s/ Jane Doe",
                "Name: Jane Doe",
                "Title: Chief Financial Officer",
            ])

            self.create_pdf(original_path, [original_text])
            self.create_pdf(signed_path, [signed_text])

            original_feature = self.load_first_page_features(original_path)
            signed_feature = self.load_first_page_features(signed_path)
            original_doc = {
                "clean_name": clean_filename_stem(original_path.name),
                "detected_name": original_feature.get("doc_name"),
                "sig_pages": [original_feature],
            }

            score, details, plausible = score_signature_page_match(signed_feature, original_feature, original_doc)

            self.assertTrue(plausible)
            self.assertGreaterEqual(score, 0.58)
            self.assertGreaterEqual(details["signature_block_score"], 0.55)

    def test_assign_signature_pages_prefers_correct_document_when_multiple_candidates_exist(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            loan_path = tmp_path / "Loan Agreement.pdf"
            guaranty_path = tmp_path / "Guaranty Agreement.pdf"
            signed_path = tmp_path / "Signed Guaranty.pdf"

            loan_text = "\n".join([
                "SIGNATURE PAGE TO LOAN AGREEMENT",
                "ACME HOLDINGS LLC",
                "By: ____________________________",
                "Name: Jane Doe",
                "Title: Chief Financial Officer",
            ])
            guaranty_text = "\n".join([
                "SIGNATURE PAGE TO GUARANTY AGREEMENT",
                "ACME HOLDINGS LLC",
                "By: ____________________________",
                "Name: Jane Doe",
                "Title: Chief Financial Officer",
            ])
            signed_text = "\n".join([
                "DocuSign Envelope ID: 1234567890ABCDEF",
                "SIGNATURE PAGE TO GUARANTY AGREEMENT",
                "ACME HOLDINGS LLC",
                "By: /s/ Jane Doe",
                "Name: Jane Doe",
                "Title: Chief Financial Officer",
            ])

            self.create_pdf(loan_path, [loan_text])
            self.create_pdf(guaranty_path, [guaranty_text])
            self.create_pdf(signed_path, [signed_text])

            loan_feature = self.load_first_page_features(loan_path)
            guaranty_feature = self.load_first_page_features(guaranty_path)
            signed_feature = self.load_first_page_features(signed_path)
            signed_feature["filepath"] = str(signed_path)

            matches = assign_signature_pages(
                [signed_feature],
                {
                    loan_path.name: {
                        "clean_name": clean_filename_stem(loan_path.name),
                        "detected_name": loan_feature.get("doc_name"),
                        "sig_pages": [loan_feature],
                    },
                    guaranty_path.name: {
                        "clean_name": clean_filename_stem(guaranty_path.name),
                        "detected_name": guaranty_feature.get("doc_name"),
                        "sig_pages": [guaranty_feature],
                    },
                },
            )

            self.assertEqual(1, len(matches))
            self.assertEqual(guaranty_path.name, matches[0]["original_filename"])


if __name__ == "__main__":
    unittest.main()
