"""Tests for JSON-manifest batch result classification."""

from __future__ import annotations

import json
import os
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(ROOT, "samples")

from tradecheck.batch import run_batch


class TestBatchClassification(unittest.TestCase):
    def test_clean_sample_pairs_exit_successfully(self):
        report, exit_code, _ = run_batch(
            os.path.join(SAMPLES, "batch-manifest.example.json")
        )

        self.assertEqual(exit_code, 0)
        self.assertEqual(report["status"], "completed")
        self.assertEqual(report["counts"]["passed"], 2)
        self.assertEqual(report["counts"]["review_required"], 0)

    def test_difference_requires_review_and_nonzero_exit(self):
        pair = {
            "id": "short-shipment",
            "invoice": os.path.join(SAMPLES, "A", "qty_short_invoice.xlsx"),
            "packing": os.path.join(SAMPLES, "A", "qty_short_packing.xlsx"),
        }
        with tempfile.TemporaryDirectory() as directory:
            manifest_path = os.path.join(directory, "manifest.json")
            with open(manifest_path, "w", encoding="utf-8") as handle:
                json.dump({"template": "A", "pairs": [pair]}, handle)

            report, exit_code, _ = run_batch(manifest_path)

        self.assertEqual(exit_code, 1)
        self.assertEqual(report["status"], "review_required")
        self.assertEqual(report["counts"]["passed"], 0)
        self.assertEqual(report["counts"]["review_required"], 1)


if __name__ == "__main__":
    unittest.main()
