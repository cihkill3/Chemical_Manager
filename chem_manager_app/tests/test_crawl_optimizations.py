import os
import sys
import tempfile
import time
import threading
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.db_manager import DBManager
from scrapers.base_scraper import BaseScraper
from scrapers.registry import scraper_class


class DummyScraper(BaseScraper):
    def scrape(self, product_number):
        return {"Catalog No.": product_number}


class ReadyContext:
    def execute_script(self, script, *args):
        return True if args else "complete"

    def get_title(self):
        return "Product page"


class CrawlOptimizationTests(unittest.TestCase):
    def test_crawl_key_normalizes_alias_case_and_excel_number(self):
        self.assertEqual(
            DBManager.crawl_key("Sigma-Aldrich", "ABC123.0"),
            DBManager.crawl_key("ALDRICH", "abc123"),
        )

    def test_wait_for_page_returns_without_fixed_timeout(self):
        scraper = DummyScraper(browser_context=ReadyContext())
        started = time.monotonic()
        self.assertTrue(scraper.wait_for_page(timeout=5, selector="h1"))
        self.assertLess(time.monotonic() - started, 0.5)

    def test_blocking_network_call_cancels_without_waiting_for_timeout(self):
        stopped = {"value": False}
        scraper = DummyScraper(check_stop_fn=lambda: stopped["value"])

        def slow_call():
            time.sleep(2)
            return "late"

        threading.Timer(0.15, lambda: stopped.update(value=True)).start()
        started = time.monotonic()
        with self.assertRaisesRegex(RuntimeError, "중단"):
            scraper.run_cancellable(slow_call, "테스트 요청")
        self.assertLess(time.monotonic() - started, 0.6)

    def test_find_fresh_sds_matches_manufacturer_and_catalog(self):
        with tempfile.TemporaryDirectory() as root:
            sds_dir = os.path.join(root, "sds")
            os.makedirs(sds_dir)
            path = os.path.join(sds_dir, "Example (TCI, B1234).pdf")
            with open(path, "wb") as stream:
                stream.write(b"%PDF-1.4 test")
            scraper = DummyScraper(base_dir=root)
            self.assertEqual(scraper.find_fresh_sds("tci", "b1234"), os.path.abspath(path))

    def test_db_sds_path_is_preferred_even_with_nonstandard_filename(self):
        with tempfile.TemporaryDirectory() as root:
            path = os.path.join(root, "renamed-by-user.pdf")
            with open(path, "wb") as stream:
                stream.write(b"%PDF-1.4 test")
            scraper = DummyScraper(base_dir=root, existing_sds_path=path)
            self.assertEqual(scraper.find_fresh_sds("TCI", "B1234"), os.path.abspath(path))

    def test_registry_routes_aliases_consistently(self):
        self.assertEqual(scraper_class("Sigma-Aldrich").__name__, "AldrichScraper")
        self.assertEqual(scraper_class("Tokyo Kasei").__name__, "TciScraper")
        self.assertIsNone(scraper_class("Unsupported Vendor"))

    def test_corrupted_thermofisher_css_record_is_recrawled(self):
        record = {
            "Catalog No.": "A10436",
            "Product Name": "/* Hash: 27f2 Source: /chemicals/css/dist/header.min.css */ @charset UTF-8",
            "Detail_Link": "https://chemicals.thermofisher.kr/apac/product/A10436.14",
            "SDS_Link": "https://www.thermofisher.com/search/browse/results?term=A10436",
        }
        self.assertTrue(DBManager.needs_recrawl(record))

    def test_valid_record_does_not_force_recrawl(self):
        record = {
            "Catalog No.": "A10436",
            "Product Name": "Alexa Fluor™ 488 Hydrazide",
            "Detail_Link": "https://www.thermofisher.com/order/catalog/product/A10436",
            "SDS_Link": "https://documents.thermofisher.com/TFS-Assets/LSG/SDS/A10436_MTR-NALT_EN.pdf",
        }
        self.assertFalse(DBManager.needs_recrawl(record))

    def test_thermofisher_revision_datetime_is_stored_as_date_only(self):
        self.assertEqual(DBManager.normalize_revision_date("2026-08-13 6:00"), "2026-08-13")

    def test_excel_serial_revision_date_does_not_shift_timezone(self):
        self.assertEqual(DBManager.normalize_revision_date(46247.625), "2026-08-13")


if __name__ == "__main__":
    unittest.main()
