import os
import unittest
from unittest.mock import Mock, patch

from scrapers.thermofisher import ThermofisherScraper


class ThermoFisherTests(unittest.TestCase):
    def test_rejects_css_asset_as_product_page(self):
        self.assertFalse(ThermofisherScraper._valid_product_page(
            "/* Hash: 27f2 Source: /chemicals/css/dist/header.min.css */ @charset UTF-8",
            "A10436 body.overflow-hidden{}",
            "A10436",
        ))

    def test_accepts_canonical_a10436_product_page(self):
        self.assertTrue(ThermofisherScraper._valid_product_page(
            "Alexa Fluor™ 488 Hydrazide",
            "Catalog number A10436 Quantity 1 mg",
            "A10436",
        ))

    def test_fresh_sds_keeps_direct_pdf_link_not_search_page(self):
        context = Mock()
        context.get_title.return_value = "Alexa Fluor 488 Hydrazide"
        context.get_text.return_value = "Alexa Fluor™ 488 Hydrazide"
        context.get_page_source.return_value = "Catalog number A10436 Quantity 1 mg"
        context.execute_script.return_value = "complete"
        scraper = ThermofisherScraper(browser_context=context, base_dir=os.getcwd())

        api_response = Mock(status_code=200)
        api_response.json.return_value = {"data": "https://documents.thermofisher.com/A10436_SDS.pdf"}
        with patch.object(scraper, "wait_for_page", return_value=True), \
             patch.object(scraper, "find_fresh_sds", return_value=r"C:\sds\A10436.pdf"), \
             patch("scrapers.thermofisher.requests.get", return_value=api_response):
            result = scraper.scrape("A10436")

        self.assertEqual(result["Product Name"], "Alexa Fluor™ 488 Hydrazide")
        self.assertEqual(result["Detail_Link"], "https://www.thermofisher.com/order/catalog/product/A10436")
        self.assertEqual(result["SDS_Link"], "https://documents.thermofisher.com/A10436_SDS.pdf")
        self.assertEqual(result["SDS_Local_Path"], r"C:\sds\A10436.pdf")


if __name__ == "__main__":
    unittest.main()
