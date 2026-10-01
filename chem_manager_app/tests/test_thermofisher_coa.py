import json
from pathlib import Path
from unittest.mock import patch

from scrapers.coa_downloader import (
    _download_thermofisher,
    _thermo_asset_matches,
    _thermo_document_url,
    _thermo_pdf_catalog,
)


def test_thermo_asset_matches_root_catalog_for_package_skus():
    asset = {
        "lotNumber": "2668510",
        "documentType": "Certificate of Analysis",
        "sku": ["326891000", "326890010", "326890025"],
        "rootSku": ["32689"],
    }

    assert _thermo_asset_matches(asset, "32689", "2668510")
    assert not _thermo_asset_matches(asset, "32688", "2668510")


def test_thermo_ecertificate_uses_assets_host():
    assert _thermo_document_url("api/ecertificate/document-id") == (
        "https://assets.thermofisher.com/api/ecertificate/document-id"
    )
    assert _thermo_document_url("TFS-Assets/CCG/certificate/test.pdf") == (
        "https://documents.thermofisher.com/TFS-Assets/CCG/certificate/test.pdf"
    )


def test_thermo_package_sku_pdf_is_verified_by_root_catalog():
    asset = {"sku": ["326890010"], "rootSku": ["32689"]}

    assert _thermo_pdf_catalog(asset, "326890010") == "32689"
    assert _thermo_pdf_catalog(asset, "32689") == "32689"


def test_thermo_download_enables_partial_sku_search():
    payload = {
        "assetTypes": [
            {
                "documentTypes": [
                    {
                        "assets": [
                            {
                                "lotNumber": "2668510",
                                "documentType": "Certificate of Analysis",
                                "sku": ["326891000", "326890010", "326890025"],
                                "rootSku": ["32689"],
                                "path": "api/ecertificate/document-id",
                            }
                        ]
                    }
                ]
            }
        ]
    }

    class Context:
        def __init__(self):
            self.product_url = ""

        def get(self, url):
            self.product_url = url

    context = Context()
    pdf = b"%PDF-test"

    def browser_fetch(_context, url, **_kwargs):
        assert "partialSkuSearch=true" in url
        return {"content": json.dumps(payload).encode("utf-8")}

    with (
        patch("scrapers.coa_downloader._browser_fetch", side_effect=browser_fetch),
        patch("scrapers.coa_downloader._http_pdf", return_value=pdf) as http_pdf,
        patch("scrapers.coa_downloader._verify_pdf") as verify_pdf,
        patch("scrapers.coa_downloader._save_pdf", return_value="coa.pdf"),
    ):
        document = _download_thermofisher(
            context, "32689", "2668510", Path("."), "1,4-Dioxane"
        )

    assert context.product_url.endswith("/32689")
    assert document.source_url == (
        "https://assets.thermofisher.com/api/ecertificate/document-id"
    )
    http_pdf.assert_called_once_with(document.source_url, referer=context.product_url)
    verify_pdf.assert_called_once_with(pdf, "32689", "2668510", require_coa=True)
