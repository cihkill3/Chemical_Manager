import os
from abc import ABC, abstractmethod

class BaseScraper(ABC):
    def __init__(self, browser_context=None, fast_mode=False, base_dir=None):
        self.context = browser_context
        self.fast_mode = fast_mode
        self.base_dir = base_dir if base_dir else os.getcwd()

    def get_sleep_time(self, base_time):
        if self.fast_mode:
            return max(1.0, base_time / 3.0)
        return base_time

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
