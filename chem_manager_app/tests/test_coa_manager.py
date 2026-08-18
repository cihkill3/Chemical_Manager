import os
import tempfile
import unittest
import time
import threading
from unittest.mock import patch

from core.coa_manager import COAManager, is_supported_vendor, lots_equal, normalize_expiration_date, normalize_lot, valid_cached_document
from core.excel_manager import reserve_missing_headers


class _Font:
    Color = None


class _Comment:
    def __init__(self):
        self.deleted = False

    def Delete(self):
        self.deleted = True


class _Cell:
    def __init__(self):
        self.Font = _Font()
        self.Comment = None
        self.note = ""

    def AddComment(self, note):
        self.note = note


class _Sheet:
    def __init__(self):
        self.cells = {}

    def Cells(self, row, column):
        return self.cells.setdefault((row, column), _Cell())


class COAManagerTests(unittest.TestCase):
    def test_ambiguous_numeric_expiration_is_not_guessed(self):
        self.assertEqual("", normalize_expiration_date("08/12/2027"))
        self.assertEqual("2027-08-13", normalize_expiration_date("13/08/2027"))
        self.assertEqual("2027-08-13", normalize_expiration_date("2027-08-13"))

    def test_missing_headers_are_appended_after_existing_user_columns(self):
        result = reserve_missing_headers({"Order No.": 1, "User Column": 24}, ["Lot No.", "COA Link"], 24)
        self.assertEqual(result["Lot No."], 25)
        self.assertEqual(result["COA Link"], 26)

    def test_identifiers_and_cache(self):
        self.assertEqual(normalize_lot(123.0), "123")
        self.assertTrue(lots_equal("0000494416", 494416.0))
        self.assertFalse(lots_equal("0000494416", "0000494417"))
        self.assertTrue(is_supported_vendor("Sigma-Aldrich"))
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "TCI_A123_LOT9_COA.pdf")
            with open(path, "wb") as stream:
                stream.write(b"%PDF" + b"x" * 1200)
            self.assertTrue(valid_cached_document(path, "A123", "LOT9"))

    def test_abcam_selects_coc_and_writes_orange_notes(self):
        result = {"documents": [
            {"document_type": "Datasheet", "path": "Abcam_AB1_L1_Datasheet.pdf", "source_url": "https://data"},
            {"document_type": "CoC", "path": "Abcam_AB1_L1_CoC.pdf", "source_url": "https://coc"},
        ]}
        payload = COAManager.payload(result)
        self.assertTrue(payload["is_coc"])
        sheet = _Sheet()
        COAManager.apply_metadata(sheet, 2, {"COA Link": 3, "COA Local Path": 4}, payload)
        for column in (3, 4):
            cell = sheet.Cells(2, column)
            self.assertIn("CoC", cell.note)
            self.assertIn("https://data", cell.note)
            self.assertIn("Datasheet.pdf", cell.note)

    def test_downloads_each_lot_once(self):
        manager = COAManager(".")
        requests = [{"vendor": "TCI", "catalog": "A", "lot": "L"}] * 2
        with patch("core.coa_manager.download_quality_documents", return_value={"documents": []}) as download:
            manager.download_many(object(), requests)
        self.assertEqual(download.call_count, 1)

    def test_coa_download_cancel_does_not_wait_for_vendor_timeout(self):
        stopped = {"value": False}
        manager = COAManager(".", stop_fn=lambda: stopped["value"])

        def slow_download(*_args, **_kwargs):
            time.sleep(2)
            return {"documents": []}

        threading.Timer(0.15, lambda: stopped.update(value=True)).start()
        started = time.monotonic()
        with patch("core.coa_manager.download_quality_documents", side_effect=slow_download), \
             self.assertRaisesRegex(RuntimeError, "중단"):
            manager.download_many(object(), [{"vendor": "TCI", "catalog": "A", "lot": "L"}])
        self.assertLess(time.monotonic() - started, 0.6)

    def test_same_reagent_with_different_lots_stays_separate(self):
        """Repeated orders of one reagent must never share another lot's COA."""
        manager = COAManager(".")
        requests = [
            {"order": "ORDER-101", "vendor": "TCI", "catalog": "A123", "lot": "LOT-A"},
            {"order": "ORDER-102", "vendor": "TCI", "catalog": "A123", "lot": "LOT-B"},
            # A duplicate source row for ORDER-101 must not trigger a third download.
            {"order": "ORDER-101", "vendor": "TCI", "catalog": "A123", "lot": "LOT-A"},
        ]

        def result_for_lot(_context, vendor, catalog, lot, _output_dir):
            return {
                "vendor": vendor,
                "catalog": catalog,
                "lot": lot,
                "status": "downloaded",
                "documents": [{
                    "document_type": "COA",
                    "path": f"TCI_{catalog}_{lot}_COA.pdf",
                    "source_url": f"https://example.test/{catalog}/{lot}",
                }],
            }

        with patch("core.coa_manager.download_quality_documents", side_effect=result_for_lot) as download:
            results = manager.download_many(object(), requests)

        self.assertEqual(download.call_count, 2)
        key_a = manager.key("TCI", "A123", "LOT-A")
        key_b = manager.key("TCI", "A123", "LOT-B")
        self.assertNotEqual(key_a, key_b)
        self.assertEqual(COAManager.payload(results[key_a])["path"], "TCI_A123_LOT-A_COA.pdf")
        self.assertEqual(COAManager.payload(results[key_b])["path"], "TCI_A123_LOT-B_COA.pdf")
        self.assertIn("LOT-A", COAManager.payload(results[key_a])["link"])
        self.assertIn("LOT-B", COAManager.payload(results[key_b])["link"])


if __name__ == "__main__":
    unittest.main()
