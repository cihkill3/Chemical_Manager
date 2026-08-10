from abc import ABC, abstractmethod

class BaseScraper(ABC):
    def __init__(self, browser_context=None):
        self.context = browser_context

    @abstractmethod
    def scrape(self, product_number):
        """
        주어진 제품번호로 데이터를 스크래핑합니다.
        
        Returns:
            dict: {
                "제조사": str,
                "제품번호": str,
                "시약명": str,
                "CAS Number": str,
                "보관온도": str,
                "위험분류": str,
                "민감성": str,
                "SDS_Link": str,
                "SDS_Local_Path": str
            }
            또는 실패 시 None을 반환해야 합니다. (항목을 찾을 수 없는 경우 값은 "정보 없음")
        """
        pass
