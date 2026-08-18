"""COA orchestration around the preserved portable vendor downloader."""

import datetime as dt
import os
import re
import shutil
import tempfile
import threading
from pathlib import Path

from scrapers.coa_downloader import QualityDocumentError, download_quality_documents
from scrapers.base_scraper import ScrapingCancelled
from utils.color_utils import hex_to_bgr

COA_HEADERS = ("Lot No.", "Expiration Date", "COA Link", "COA Local Path")
SUPPORTED_VENDORS = {"tci", "aldrich", "sigmaaldrich", "sigma", "thermofisher", "thermo", "abcam"}
ABCAM_ORANGE = "#F28C28"


def _compact(value):
    return re.sub(r"[^0-9a-z]+", "", str(value or "").casefold())


def normalize_lot(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def lot_comparison_key(value):
    """Compare numeric lots independently of Excel-dropped leading zeroes."""
    normalized = normalize_lot(value)
    if normalized.isdigit():
        return normalized.lstrip("0") or "0"
    return normalized.casefold()


def lots_equal(left, right):
    return bool(normalize_lot(left) and normalize_lot(right)) and (
        lot_comparison_key(left) == lot_comparison_key(right)
    )


def is_supported_vendor(vendor):
    return re.sub(r"[^a-z]", "", str(vendor or "").casefold()) in SUPPORTED_VENDORS


def valid_cached_document(path, catalog, lot):
    try:
        candidate = Path(str(path or ""))
        if not (candidate.is_file() and candidate.stat().st_size >= 1000):
            return False
        with candidate.open("rb") as stream:
            if stream.read(4) != b"%PDF":
                return False
        return _compact(catalog) in _compact(candidate.name) and _compact(lot) in _compact(candidate.name)
    except OSError:
        return False


def normalize_expiration_date(raw):
    raw = str(raw or "").replace("/", "-").replace(".", "-")
    try:
        parts = [int(part) for part in raw.split("-")]
    except (TypeError, ValueError):
        return ""
    formats = ["%Y-%m-%d"] if len(parts) == 3 and parts[0] > 31 else []
    if len(parts) == 3 and parts[0] <= 31:
        first, second, _year = parts
        if first > 12:
            formats = ["%d-%m-%Y"]
        elif second > 12:
            formats = ["%m-%d-%Y"]
    for fmt in formats:
        try:
            return dt.datetime.strptime(raw, fmt).date().isoformat()
        except ValueError:
            pass
    return ""


def extract_expiration_date(pdf_path, vendor=""):
    try:
        import pymupdf
        document = pymupdf.open(pdf_path)
        text = "\n".join(page.get_text() for page in document)
        document.close()
    except Exception:
        return ""
    patterns = (
        r"(?:expiration|expiry|exp\.?\s*date|use\s*before)\s*[:#-]?\s*(\d{4}[-/.]\d{1,2}[-/.]\d{1,2})",
        r"(?:expiration|expiry|exp\.?\s*date|use\s*before)\s*[:#-]?\s*(\d{1,2}[-/.]\d{1,2}[-/.]\d{4})",
    )
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            normalized = normalize_expiration_date(match.group(1))
            if normalized:
                return normalized
    return ""


class COAManager:
    def __init__(self, output_dir, log_fn=None, stop_fn=None):
        self.output_dir = os.path.abspath(output_dir)
        self.log = log_fn or (lambda _message: None)
        self.stop = stop_fn or (lambda: False)

    @staticmethod
    def key(vendor, catalog, lot):
        return (_compact(vendor), _compact(catalog), _compact(lot))

    def download_many(self, context, requests):
        results = {}
        unique_requests = []
        seen = set()
        for request in requests:
            key = self.key(request["vendor"], request["catalog"], request["lot"])
            if key not in seen:
                seen.add(key)
                unique_requests.append(request)
        for index, request in enumerate(unique_requests, 1):
            key = self.key(request["vendor"], request["catalog"], request["lot"])
            if self.stop():
                raise ScrapingCancelled("사용자 요청으로 COA 다운로드를 중단했습니다.")
            self.log(
                f"COA/CoC 다운로드 중 ({index}/{len(unique_requests)}): "
                f"{request['vendor']} {request['catalog']} / Lot {request['lot']}"
            )
            staging_dir = tempfile.mkdtemp(prefix="chemical_coa_")
            try:
                state = {}
                finished = threading.Event()

                def invoke():
                    try:
                        state["result"] = download_quality_documents(
                            context, request["vendor"], request["catalog"],
                            request["lot"], staging_dir
                        )
                    except BaseException as error:
                        state["error"] = error
                    finally:
                        finished.set()

                worker = threading.Thread(
                    target=invoke, name="ChemicalManager-COA", daemon=True
                )
                worker.start()
                while not finished.wait(0.1):
                    if self.stop():
                        driver = getattr(context, "driver", None)
                        if driver is not None:
                            try:
                                driver.quit()
                            except Exception:
                                pass
                        finished.wait(0.2)
                        if not finished.is_set():
                            threading.Thread(
                                target=lambda: (finished.wait(), shutil.rmtree(staging_dir, ignore_errors=True)),
                                name="ChemicalManager-COA-cleanup", daemon=True,
                            ).start()
                        raise ScrapingCancelled("사용자 요청으로 COA 다운로드를 중단했습니다.")
                if "error" in state:
                    raise state["error"]
                result = state.get("result") or {}
                os.makedirs(self.output_dir, exist_ok=True)
                for document in result.get("documents", []):
                    source = document.get("path", "")
                    if source and os.path.isfile(source):
                        destination = os.path.join(self.output_dir, os.path.basename(source))
                        os.replace(source, destination)
                        document["path"] = destination
                results[key] = result
            except ScrapingCancelled:
                raise
            except Exception as error:
                self.log(f"COA 다운로드 경고 [{request['vendor']} {request['catalog']} / {request['lot']}]: {error}")
                results[key] = {"status": "error", "error": str(error), "documents": []}
            finally:
                if finished.is_set():
                    shutil.rmtree(staging_dir, ignore_errors=True)
        return results

    @staticmethod
    def payload(result):
        documents = result.get("documents", []) if result else []
        coa = next((item for item in documents if item.get("document_type") == "COA"), None)
        coc = next((item for item in documents if item.get("document_type") == "CoC"), None)
        datasheet = next((item for item in documents if item.get("document_type") == "Datasheet"), None)
        selected = coa or coc
        if not selected:
            return None
        is_coc = selected is coc
        if is_coc and "coc" not in Path(selected.get("path", "")).name.casefold():
            raise QualityDocumentError("Abcam CoC 파일명에 CoC가 명시되지 않았습니다.")
        return {"expiration": "" if is_coc else extract_expiration_date(selected.get("path", ""), result.get("vendor", "")), "link": selected.get("source_url", ""), "path": selected.get("path", ""), "is_coc": is_coc, "datasheet_link": (datasheet or {}).get("source_url", ""), "datasheet_path": (datasheet or {}).get("path", "")}

    @staticmethod
    def apply_metadata(worksheet, row, column_map, payload):
        if not payload:
            return
        cells = (worksheet.Cells(row, column_map["COA Link"]), worksheet.Cells(row, column_map["COA Local Path"]))
        color = hex_to_bgr(ABCAM_ORANGE if payload["is_coc"] else "#000000")
        for cell in cells:
            cell.Font.Color = color
            try:
                if cell.Comment:
                    cell.Comment.Delete()
            except Exception:
                pass
        if payload["is_coc"]:
            note = "CoC (Certificate of Conformance) - COA 대체 문서\n" + f"Datasheet Link: {payload['datasheet_link']}\nDatasheet Local Path: {payload['datasheet_path']}"
            for cell in cells:
                try:
                    cell.AddComment(note)
                except Exception:
                    cell.NoteText(Text=note)
