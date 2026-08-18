import os
import glob
import time
import threading
from abc import ABC, abstractmethod


class ScrapingCancelled(RuntimeError):
    pass

class BaseScraper(ABC):
    coa_vendor = None
    _sds_index_cache = {}

    def __init__(self, browser_context=None, fast_mode=False, base_dir=None, check_stop_fn=None, existing_sds_path=None):
        self.context = browser_context
        self.fast_mode = fast_mode
        self.base_dir = base_dir if base_dir else os.getcwd()
        self.check_stop_fn = check_stop_fn
        self.existing_sds_path = existing_sds_path

    def is_stopped(self):
        return self.check_stop_fn() if self.check_stop_fn else False

    def sleep(self, duration):
        """Sleeps in 0.1s steps checking for stop requests instantly."""
        import time
        eff_time = self.get_sleep_time(duration)
        steps = int(eff_time * 10)
        for _ in range(max(1, steps)):
            if self.is_stopped():
                raise Exception("사용자에 의해 스크래핑이 중단되었습니다.")
            time.sleep(0.1)

    def run_cancellable(self, function, label="네트워크 요청"):
        """Run a blocking vendor call while polling cancellation every 0.1s."""
        state = {}
        finished = threading.Event()

        def invoke():
            try:
                state["result"] = function()
            except BaseException as error:
                state["error"] = error
            finally:
                finished.set()

        threading.Thread(target=invoke, name=f"ChemicalManager-{label}", daemon=True).start()
        while not finished.wait(0.1):
            if self.is_stopped():
                raise ScrapingCancelled(f"사용자 요청으로 {label}을 중단했습니다.")
        if "error" in state:
            raise state["error"]
        return state.get("result")

    def http_get(self, url, **kwargs):
        import requests
        return self.run_cancellable(
            lambda: requests.get(url, **kwargs), label="HTTP 다운로드"
        )

    def execute_async_script(self, script, *args, **kwargs):
        return self.run_cancellable(
            lambda: self.context.execute_async_script(script, *args, **kwargs),
            label="브라우저 네트워크 요청",
        )

    def wait_for_page(self, timeout=5.0, selector=None, reject_titles=None):
        """Wait until the current page is usable, returning early when ready."""
        deadline = time.monotonic() + self.get_sleep_time(timeout)
        reject_titles = tuple(x.lower() for x in (reject_titles or ()))
        while time.monotonic() < deadline:
            if self.is_stopped():
                raise Exception("Scraping was cancelled by the user.")
            try:
                ready = self.context.execute_script("return document.readyState") == "complete"
                title = (self.context.get_title() or "").strip().lower()
                title_ok = bool(title) and not any(x in title for x in reject_titles)
                selector_ok = True
                if selector:
                    selector_ok = bool(self.context.execute_script(
                        "return !!document.querySelector(arguments[0])", selector
                    ))
                if ready and title_ok and selector_ok:
                    return True
            except Exception:
                pass
            time.sleep(0.1)
        return False

    def find_fresh_sds(self, manufacturer, catalog_no, max_days=180):
        """Return a fresh local SDS for this manufacturer/catalog before networking."""
        from core.db_manager import DBManager
        if DBManager.is_sds_fresh(self.existing_sds_path, max_days):
            return os.path.abspath(self.existing_sds_path)
        sds_dir = os.path.join(self.base_dir, "sds")
        if not os.path.isdir(sds_dir):
            return None
        mfr = DBManager.clean_filename(manufacturer).casefold()
        cat = DBManager.clean_filename(catalog_no).casefold()
        suffix = f"({mfr}, {cat}).pdf"
        directory_mtime = os.stat(sds_dir).st_mtime_ns
        cached = self._sds_index_cache.get(sds_dir)
        if not cached or cached[0] != directory_mtime:
            index = {os.path.basename(path).casefold(): os.path.abspath(path)
                     for path in glob.glob(os.path.join(sds_dir, "*.pdf"))}
            self._sds_index_cache[sds_dir] = (directory_mtime, index)
        for filename, path in self._sds_index_cache[sds_dir][1].items():
            if filename.endswith(suffix) and DBManager.is_sds_fresh(path, max_days):
                return path
        return None

    def get_sleep_time(self, base_time):
        if self.fast_mode:
            return max(1.0, base_time / 3.0)
        return base_time

    def download_quality_documents(self, catalog_no, lot_no, output_dir=None):
        """Download a verified COA, or an explicitly labelled vendor fallback."""
        from scrapers.coa_downloader import download_quality_documents

        vendor = self.coa_vendor or self.__class__.__name__.replace("Scraper", "")
        target_dir = output_dir or os.path.join(self.base_dir, "coa")
        return download_quality_documents(
            self.context, vendor, catalog_no, lot_no, target_dir
        )

    @abstractmethod
    def scrape(self, product_number):
        """
        주어진 제품번호로 데이터를 스크래핑합니다.
        
        Returns:
            dict: {
                "Manufacturer": str,
                "Catalog No.": str,
                "Product Name": str,
                "CAS No.": str,
                "Storage Temp.": str,
                "Signal Word": str,
                "Key Hazards": str,
                "Detailed Hazard Classification": str,
                "Sensitivity": str,
                "Detail_Link": str,
                "SDS_Link": str,
                "SDS_Local_Path": str
            }
            또는 실패 시 None을 반환해야 합니다. (해당 값을 찾을 수 없는 경우 값은 "정보 없음")
        """
        pass
