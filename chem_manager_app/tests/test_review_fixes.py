import pytest

from core.db_manager import DBManager
from scrapers.base_scraper import BaseScraper
from scrapers.coa_downloader import _document_filename


def test_db_record_preserves_supplied_sensitivity_when_hazard_text_has_none():
    record = DBManager.build_db_record(
        {
            "Product Name": "Reagent",
            "CAS No.": "123-45-6",
            "Storage Temp.": "room temperature",
            "Detailed Hazard Classification": "H315",
            "Sensitivity": "Moisture",
        },
        "sigma",
        "A100",
        "Fallback",
    )
    assert record["Manufacturer"] == "Aldrich"
    assert record["Sensitivity"] == "Moisture"
    assert record["Storage Temp."] == "RT"


def test_db_record_normalizes_search_failure_without_overwriting_fallback():
    record = DBManager.build_db_record(
        {"error": "Scraping error: timeout"}, "TCI", "B100", "Manual name"
    )
    assert record["Product Name"] == "Manual name"
    assert record["CAS No."] == "Search Failed"
    assert record["Detail_Link"] == "Product Not Found"


def test_invalid_pdf_bytes_are_rejected():
    assert not BaseScraper.validate_pdf_bytes(b"<html>not a pdf</html>")


def test_quality_document_filename_contains_product_and_document_type():
    assert _document_filename("시약/혼합물", "TCI", "A-123", "L/9", "COA") == (
        "시약_혼합물 (TCI, A-123, Lot_L_9) - COA.pdf"
    )
