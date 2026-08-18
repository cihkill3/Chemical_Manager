import os
import tempfile
import unittest
from unittest.mock import Mock, patch

from scrapers.abcam import AbcamScraper


class FakeContext:
    def execute_async_script(self, script, *args):
        if "proxy-gateway.abcam.com" in script:
            return {"documents": [{
                "url": "https://wercs-api-prod-bucket.example/AB92536_KGHS_EN.pdf?token=api",
                "languageCode": "en",
            }]}
        return {"data": list(b"%PDF-1.4\n" + b"test" * 50)}


class AbcamSdsTests(unittest.TestCase):
    def test_extracts_signed_wercs_sds_link(self):
        url = "https://wercs-api-prod-bucket.s3.eu-west-1.amazonaws.com/SDS/Public/AB92536/AB92536_KGHS_EN.pdf?X-Amz-Date=20260814T043248Z&amp;X-Amz-Signature=test"
        html = f'<a href="{url}">SDS</a>'
        self.assertEqual(
            url.replace("&amp;", "&"),
            AbcamScraper._sds_url_from_html(html, "https://www.abcam.com/ab92536"),
        )

    def test_downloads_abcam_sds_to_local_folder(self):
        with tempfile.TemporaryDirectory() as folder:
            scraper = AbcamScraper(browser_context=FakeContext(), base_dir=folder)
            result = {"Product Name": "Anti-MMP2 antibody", "SDS_Link": "", "SDS_Local_Path": ""}
            html = '<a href="https://wercs-api-prod-bucket.s3.eu-west-1.amazonaws.com/SDS/Public/AB92536/AB92536_KGHS_EN.pdf?token=1">SDS</a>'
            response = Mock(content=b"%PDF-1.4\n" + b"test" * 50)
            response.raise_for_status.return_value = None
            with patch("requests.get", return_value=response):
                scraper._download_sds(result, "ab92536", html, "https://www.abcam.com/ab92536")
            self.assertTrue(result["SDS_Link"].endswith("?token=1"))
            self.assertTrue(os.path.isfile(result["SDS_Local_Path"]))

    def test_uses_official_api_when_page_has_no_direct_link(self):
        with tempfile.TemporaryDirectory() as folder:
            scraper = AbcamScraper(browser_context=FakeContext(), base_dir=folder)
            result = {"Product Name": "Anti-MMP2 antibody", "SDS_Link": "", "SDS_Local_Path": ""}
            response = Mock(content=b"%PDF-1.4\n" + b"test" * 50)
            response.raise_for_status.return_value = None
            with patch("requests.get", return_value=response):
                scraper._download_sds(result, "ab92536", "<div class='sds-button'>SDS</div>", "https://www.abcam.com/ab92536")
            self.assertTrue(result["SDS_Link"].endswith("?token=api"))
            self.assertTrue(os.path.isfile(result["SDS_Local_Path"]))


if __name__ == "__main__":
    unittest.main()
