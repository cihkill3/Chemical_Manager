"""Vendor quality-document downloader for a SeleniumBase browser context.

The vendor sites use four different delivery mechanisms:

* TCI returns a PDF Blob from a page-side XHR.
* Sigma-Aldrich returns either a signed PDF URL or an HTML COA page.
* Thermo Fisher exposes the certificate path through its same-origin JSON API.
* Abcam does not publish a lot COA for the tested antibody; its Datasheet and
  lot-specific CoC are saved as explicitly labelled fallback documents.

Never rename a Datasheet or CoC to COA.  The structured return value lets the
caller decide whether an Abcam fallback is acceptable for its quality system.
"""

from __future__ import annotations

import base64
import datetime as dt
import html
import json
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable
from urllib.parse import quote, urljoin

import requests


def _wait_ready(context, timeout: float = 3.0) -> bool:
    """Return as soon as navigation has a complete, non-empty document."""
    import time
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if context.execute_script(
                "return document.readyState === 'complete' && !!document.body"
            ):
                return True
        except Exception:
            pass
        time.sleep(0.1)
    return False


class QualityDocumentError(RuntimeError):
    """Raised when the requested official document cannot be verified."""


@dataclass(frozen=True)
class QualityDocument:
    document_type: str
    path: str
    source_url: str
    catalog: str
    lot: str
    verified: bool = True


@dataclass
class QualityDownloadResult:
    vendor: str
    catalog: str
    lot: str
    status: str
    documents: list[QualityDocument] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        data = asdict(self)
        data["documents"] = [asdict(item) for item in self.documents]
        return data


def _safe(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z._-]+", "_", value.strip()).strip("._")


def _compact(value: str) -> str:
    return re.sub(r"[^0-9a-z]+", "", value.casefold())


def _pdf_text(content: bytes) -> str:
    if len(content) < 1_000 or not content.startswith(b"%PDF"):
        raise QualityDocumentError("?묐떟???좏슚??PDF媛 ?꾨떃?덈떎.")
    try:
        import pymupdf

        doc = pymupdf.open(stream=content, filetype="pdf")
        if doc.page_count < 1:
            raise QualityDocumentError("PDF???섏씠吏媛 ?놁뒿?덈떎.")
        text = "\n".join(page.get_text() for page in doc)
        doc.close()
        return text
    except QualityDocumentError:
        raise
    except Exception as exc:
        raise QualityDocumentError(f"PDF ?뚯떛???ㅽ뙣?덉뒿?덈떎: {exc}") from exc


def _verify_pdf(
    content: bytes,
    catalog: str,
    lot: str | None,
    *,
    require_coa: bool,
) -> None:
    text = _pdf_text(content)
    compact = _compact(text)
    if _compact(catalog) not in compact:
        raise QualityDocumentError(f"PDF 蹂몃Ц?먯꽌 catalog {catalog!r}瑜??뺤씤?섏? 紐삵뻽?듬땲??")
    if lot and _compact(lot) not in compact:
        raise QualityDocumentError(f"PDF 蹂몃Ц?먯꽌 lot {lot!r}瑜??뺤씤?섏? 紐삵뻽?듬땲??")
    if require_coa and not any(
        marker in compact
        for marker in ("certificateofanalysis", "certificateofanalyses", "analysiscertificate")
    ):
        raise QualityDocumentError("PDF 蹂몃Ц?먯꽌 Certificate of Analysis ?쒓린瑜??뺤씤?섏? 紐삵뻽?듬땲??")


def _save_pdf(folder: Path, filename: str, content: bytes) -> str:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / filename
    target.write_bytes(content)
    return str(target.resolve())


def _http_pdf(url: str, *, referer: str = "") -> bytes:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127 Safari/537.36"
        ),
        "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
    }
    if referer:
        headers["Referer"] = referer
    response = requests.get(url, headers=headers, timeout=45, allow_redirects=True)
    response.raise_for_status()
    if not response.content.startswith(b"%PDF"):
        raise QualityDocumentError(
            f"PDF ???{response.headers.get('content-type', 'unknown')} ?묐떟??諛쏆븯?듬땲?? {url}"
        )
    return response.content


def _browser_fetch(context, url: str, *, accept: str = "application/pdf,*/*") -> dict:
    script = """
    const done = arguments[arguments.length - 1];
    fetch(__URL__, {credentials:'include', headers:{Accept:__ACCEPT__}})
      .then(async r => {
        const bytes = new Uint8Array(await r.arrayBuffer());
        let binary = '';
        for (let i=0; i<bytes.length; i+=0x8000)
          binary += String.fromCharCode.apply(null, bytes.subarray(i, i+0x8000));
        done({ok:r.ok, status:r.status, url:r.url,
              contentType:r.headers.get('content-type') || '', data:btoa(binary)});
      })
      .catch(e => done({ok:false, error:String(e)}));
    """.replace("__URL__", json.dumps(url)).replace("__ACCEPT__", json.dumps(accept))
    result = context.execute_async_script(script, timeout=45) or {}
    if not result.get("ok"):
        raise QualityDocumentError(result.get("error") or f"HTTP {result.get('status')}")
    result["content"] = base64.b64decode(result.pop("data"))
    return result


def _capture_window_open(context, selector: str, *, timeout_seconds: int = 15) -> str:
    script = """
    const done=arguments[arguments.length-1], selector=__SELECTOR__;
    let ended=false;
    const finish=x=>{if(!ended){ended=true;done(x)}};
    window.open=(url)=>{finish({ok:true,url:String(url)});return null};
    const node=document.querySelector(selector);
    if(!node)return finish({ok:false,error:'control not found: '+selector});
    node.click();
    setTimeout(()=>finish({ok:false,error:'window.open was not called'}), __TIMEOUT__);
    """.replace("__SELECTOR__", json.dumps(selector)).replace(
        "__TIMEOUT__", str(timeout_seconds * 1000)
    )
    result = context.execute_async_script(script, timeout=timeout_seconds + 5) or {}
    if not result.get("ok"):
        raise QualityDocumentError(result.get("error", "臾몄꽌 URL???살? 紐삵뻽?듬땲??"))
    return result["url"]


def _download_tci(context, catalog: str, lot: str, folder: Path) -> QualityDocument:
    product_url = f"https://www.tcichemicals.com/KR/ko/p/{quote(catalog.upper())}"
    context.get(product_url)
    script = """
    const done=arguments[arguments.length-1], lot=__LOT__;
    let ended=false;
    const finish=x=>{if(!ended){ended=true;done(x)}};
    const read=blob=>{
      const reader=new FileReader();
      reader.onloadend=()=>finish({ok:true,data:reader.result.split(',')[1]});
      reader.readAsDataURL(blob);
    };
    window.open=url=>{fetch(url,{credentials:'include'}).then(r=>r.blob()).then(read)
      .catch(e=>finish({ok:false,error:String(e)}));return null};
    const nativeSend=XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send=function(body){
      this.addEventListener('load',()=>{
        if(this.response instanceof Blob && this.response.size>1000)read(this.response);
      });
      return nativeSend.call(this,body);
    };
    const input=document.querySelector('#LotNumbere');
    const button=document.querySelector('button.js-cofa');
    if(!input || !button)return finish({ok:false,error:'TCI COA controls not found'});
    input.value=lot;
    input.dispatchEvent(new Event('input',{bubbles:true}));
    input.dispatchEvent(new Event('change',{bubbles:true}));
    button.click();
    setTimeout(()=>finish({ok:false,error:'TCI returned no COA for this lot'}),20000);
    """.replace("__LOT__", json.dumps(lot.upper()))
    result = context.execute_async_script(script, timeout=25) or {}
    if not result.get("ok"):
        raise QualityDocumentError(result.get("error", "TCI COA ?묐떟???놁뒿?덈떎."))
    content = base64.b64decode(result["data"])
    _verify_pdf(content, catalog, lot, require_coa=True)
    path = _save_pdf(folder, f"TCI_{_safe(catalog)}_{_safe(lot)}_COA.pdf", content)
    return QualityDocument("COA", path, product_url, catalog, lot)


def _find_aldrich_product(context, catalog: str) -> str:
    clean = catalog.split("-")[0]
    candidates = ("sigma", "aldrich", "sial", "supelco", "mm")
    for brand in candidates:
        url = f"https://www.sigmaaldrich.com/KR/ko/product/{brand}/{quote(clean.lower())}"
        context.get(url)
        _wait_ready(context, 3)
        title = (context.get_title() or "").casefold()
        if not any(marker in title for marker in ("404", "not found", "error")):
            page_text = context.execute_script("return document.body?.innerText || ''") or ""
            if _compact(clean) in _compact(page_text):
                return url
    raise QualityDocumentError(f"Sigma-Aldrich ?쒗뭹 ?섏씠吏瑜?李얠? 紐삵뻽?듬땲?? {catalog}")


def _aldrich_target(context, lot: str) -> str:
    marker_script = """
    (() => {
      const lot=__LOT__;
      const links=[...document.querySelectorAll('a.css-tg85n8-documentLink')];
      const hit=links.find(a=>a.textContent.trim().toLowerCase()===lot.toLowerCase());
      if(!hit)return false;
      hit.setAttribute('data-coa-target','1');return true;
    })()
    """.replace("__LOT__", json.dumps(lot))
    if not context.execute_script(marker_script):
        raise QualityDocumentError(f"Sigma-Aldrich ?섏씠吏??lot {lot}媛 ?놁뒿?덈떎.")
    return _capture_window_open(context, 'a[data-coa-target="1"]')


def _aldrich_search_target(context, catalog: str, lot: str) -> tuple[str, str]:
    """Search the official document form, including legacy zero-padded lots."""
    candidates = [lot]
    if lot.isdigit() and len(lot) < 10:
        candidates.append(lot.zfill(10))
    search_url = (
        "https://www.sigmaaldrich.com/KR/ko/documents-search?tab=coa"
        f"&productNumber={quote(catalog.upper())}"
    )
    for candidate in dict.fromkeys(candidates):
        context.get(search_url)
        context.wait_for_element_visible("#autocomplete-cofa_lot_number-input", timeout=15)
        context.type("#autocomplete-cofa_lot_number-input", candidate)
        try:
            target = _capture_window_open(context, "#COA-submit", timeout_seconds=10)
            return target, candidate
        except QualityDocumentError:
            # A failed form submission is not terminal: legacy numeric lots are
            # often indexed only in Sigma's ten-character zero-padded form.
            continue
    raise QualityDocumentError(
        f"Sigma-Aldrich 怨듭떇 COA 寃?됱뿉 寃곌낵媛 ?놁뒿?덈떎: {catalog} / {lot}"
    )


def _print_current_page_pdf(context, source_url: str, catalog: str, lot: str) -> bytes:
    body = context.execute_script("return document.body?.innerText || ''") or ""
    if _compact(catalog) not in _compact(body) or _compact(lot) not in _compact(body):
        raise QualityDocumentError("HTML COA ?섏씠吏?먯꽌 ?붿껌 catalog/lot瑜??뺤씤?섏? 紐삵뻽?듬땲??")
    if "/documents/coa/" not in source_url.casefold():
        raise QualityDocumentError("?대┛ HTML 臾몄꽌??怨듭떇 COA 寃쎈줈瑜??뺤씤?섏? 紐삵뻽?듬땲??")
    archived = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    context.execute_script(
        """(() => {
          const style=document.createElement('style');
          style.textContent=`@media print {
            header, footer, nav, [role="banner"], [role="contentinfo"] {
              display:none !important;
            }
          }`;
          document.head.appendChild(style);
        })()"""
    )
    header = (
        "<div style='font-size:8px;width:100%;padding:0 12mm;color:#555'>"
        f"Sigma-Aldrich Certificate of Analysis web archive ??Catalog {html.escape(catalog)}, "
        f"Lot {html.escape(lot)} ??{html.escape(archived)}</div>"
    )
    footer = (
        "<div style='font-size:7px;width:100%;padding:0 12mm;color:#666'>"
        f"Source: {html.escape(source_url)}"
        "<span style='float:right'>Page <span class='pageNumber'></span> / "
        "<span class='totalPages'></span></span></div>"
    )
    payload = context.driver.execute_cdp_cmd(
        "Page.printToPDF",
        {
            "printBackground": True,
            "displayHeaderFooter": True,
            "headerTemplate": header,
            "footerTemplate": footer,
            "marginTop": 0.7,
            "marginBottom": 0.7,
            "marginLeft": 0.35,
            "marginRight": 0.35,
            "preferCSSPageSize": False,
        },
    )
    return base64.b64decode(payload["data"])


def _download_aldrich(context, catalog: str, lot: str, folder: Path) -> QualityDocument:
    product_url = _find_aldrich_product(context, catalog)
    resolved_lot = lot
    try:
        target = _aldrich_target(context, lot)
    except QualityDocumentError:
        target, resolved_lot = _aldrich_search_target(context, catalog, lot)
    source_url = urljoin(product_url, target)
    if "/documents/coa/" in source_url.casefold():
        context.get(source_url)
        _wait_ready(context, 5)
        content = _print_current_page_pdf(context, source_url, catalog, lot)
    else:
        try:
            content = _browser_fetch(context, source_url)["content"]
        except QualityDocumentError:
            content = _http_pdf(source_url, referer=product_url)
    _verify_pdf(content, catalog, resolved_lot, require_coa=True)
    path = _save_pdf(folder, f"Aldrich_{_safe(catalog)}_{_safe(lot)}_COA.pdf", content)
    return QualityDocument("COA", path, source_url, catalog, lot)


def _flatten_thermo_assets(payload: dict) -> Iterable[dict]:
    for asset_type in payload.get("assetTypes", []):
        for document_type in asset_type.get("documentTypes", []):
            yield from document_type.get("assets", [])


def _download_thermofisher(context, catalog: str, lot: str, folder: Path) -> QualityDocument:
    catalog_upper = catalog.upper()
    product_url = f"https://www.thermofisher.com/order/catalog/product/{quote(catalog_upper)}"
    context.get(product_url)
    api = (
        "/api/store/Assets/Documents/Certificates/v2/search?"
        f"skus={quote(catalog_upper)}&country=kr&targetSite=TF&partialLotNumber=true"
        f"&erpType=Global_E1&lotNumbers={quote(lot)}"
    )
    response = _browser_fetch(context, api, accept="application/json")
    try:
        payload = json.loads(response["content"].decode("utf-8"))
    except Exception as exc:
        raise QualityDocumentError("Thermo Fisher ?몄쬆??API ?묐떟???댁꽍?섏? 紐삵뻽?듬땲??") from exc
    matches = [
        asset
        for asset in _flatten_thermo_assets(payload)
        if _compact(str(asset.get("lotNumber", ""))) == _compact(lot)
        and "certificate of analysis" in str(asset.get("documentType", "")).casefold()
        and any(_compact(sku) == _compact(catalog) for sku in asset.get("sku", []))
    ]
    if not matches:
        raise QualityDocumentError(f"Thermo Fisher COA瑜?李얠? 紐삵뻽?듬땲?? {catalog} / {lot}")
    asset = matches[0]
    source_url = "https://documents.thermofisher.com/" + quote(asset["path"], safe="/")
    try:
        content = _http_pdf(source_url, referer=product_url)
    except Exception:
        content = _browser_fetch(context, source_url)["content"]
    _verify_pdf(content, catalog, lot, require_coa=True)
    path = _save_pdf(folder, f"ThermoFisher_{_safe(catalog)}_{_safe(lot)}_COA.pdf", content)
    return QualityDocument("COA", path, source_url, catalog, lot)


def _download_abcam(context, catalog: str, lot: str, folder: Path) -> QualityDownloadResult:
    product_url = f"https://www.abcam.com/{catalog.lower()}"
    context.get(product_url)
    _wait_ready(context, 5)
    documents: list[QualityDocument] = []

    datasheet_url = urljoin(product_url, _capture_window_open(context, ".ds-button"))
    datasheet = _http_pdf(datasheet_url, referer=product_url)
    _verify_pdf(datasheet, catalog, None, require_coa=False)
    datasheet_path = _save_pdf(
        folder, f"Abcam_{_safe(catalog)}_{_safe(lot)}_Datasheet.pdf", datasheet
    )
    documents.append(QualityDocument("Datasheet", datasheet_path, datasheet_url, catalog, lot))

    context.click(".coc-button")
    context.wait_for_element_visible("input.lotNumber-input", timeout=10)
    context.type("input.lotNumber-input", lot)
    context.click(".coc-submit button")
    context.wait_for_element_visible(".coc-pdf-download-center a[href]", timeout=15)
    coc_url = context.get_attribute(".coc-pdf-download-center a[href]", "href")
    coc_url = urljoin(product_url, coc_url)
    coc = _http_pdf(coc_url, referer=product_url)
    _verify_pdf(coc, catalog, lot, require_coa=False)
    coc_text = _compact(_pdf_text(coc))
    if not any(
        marker in coc_text
        for marker in (
            "certificateofconformance",
            "certificateofconformity",
            "certificateofcompliance",
        )
    ):
        raise QualityDocumentError("Abcam 臾몄꽌?먯꽌 CoC ?쒓린瑜??뺤씤?섏? 紐삵뻽?듬땲??")
    coc_path = _save_pdf(folder, f"Abcam_{_safe(catalog)}_{_safe(lot)}_CoC.pdf", coc)
    documents.append(QualityDocument("CoC", coc_path, coc_url, catalog, lot))

    return QualityDownloadResult(
        vendor="Abcam",
        catalog=catalog,
        lot=lot,
        status="fallback",
        documents=documents,
        warnings=[
            "Abcam COA瑜?李얠? 紐삵빐 Datasheet? lot蹂?CoC瑜???ν뻽?듬땲?? "
            "CoC???쒗뿕 寃곌낵媛 湲곗옱??COA? ?숈씪??臾몄꽌媛 ?꾨떃?덈떎."
        ],
    )


def download_quality_documents(
    context,
    vendor: str,
    catalog: str,
    lot: str,
    output_dir: str | Path,
) -> dict:
    """Download and verify official quality documents for one product lot."""
    vendor_key = re.sub(r"[^a-z]", "", vendor.casefold())
    folder = Path(output_dir)
    if vendor_key == "tci":
        document = _download_tci(context, catalog, lot, folder)
        result = QualityDownloadResult("TCI", catalog, lot, "downloaded", [document])
    elif vendor_key in {"aldrich", "sigmaaldrich", "sigma"}:
        document = _download_aldrich(context, catalog, lot, folder)
        result = QualityDownloadResult("Aldrich", catalog, lot, "downloaded", [document])
    elif vendor_key in {"thermofisher", "thermo"}:
        document = _download_thermofisher(context, catalog, lot, folder)
        result = QualityDownloadResult("ThermoFisher", catalog, lot, "downloaded", [document])
    elif vendor_key == "abcam":
        result = _download_abcam(context, catalog, lot, folder)
    else:
        raise QualityDocumentError(f"吏?먰븯吏 ?딅뒗 ?쒖“?ъ엯?덈떎: {vendor}")
    return result.to_dict()


def download_coa(context, vendor: str, catalog: str, lot: str, output_dir: str | Path) -> str:
    """Compatibility helper returning the COA path; Abcam raises explicitly."""
    result = download_quality_documents(context, vendor, catalog, lot, output_dir)
    for document in result["documents"]:
        if document["document_type"] == "COA":
            return document["path"]
    raise QualityDocumentError(
        f"{vendor} {catalog}/{lot}?먮뒗 寃利앸맂 COA媛 ?놁뒿?덈떎. "
        "download_quality_documents()濡??泥?臾몄꽌 ?곹깭瑜??뺤씤?섏떗?쒖삤."
    )


