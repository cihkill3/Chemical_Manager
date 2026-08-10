import os
import pandas as pd
from datetime import datetime
import re

class DBManager:
    def __init__(self, db_path="chemical_db.xlsx"):
        self.db_path = db_path
        self.columns = [
            "제조사", "제품번호", "시약명", "CAS Number", 
            "보관온도", "신호어", "주요위험", "상세 위험분류", "민감성", "상세정보_링크", "SDS_Link", 
            "SDS_Local_Path", "갱신일"
        ]
        self._init_db()

    def _init_db(self):
        """DB 파일이 없으면 초기 템플릿을 생성합니다."""
        if not os.path.exists(self.db_path):
            df = pd.DataFrame(columns=self.columns)
            df.to_excel(self.db_path, index=False)

    def load_db(self):
        """DB 파일을 pandas DataFrame으로 로드합니다."""
        return pd.read_excel(self.db_path)

    def get_product(self, manufacturer, product_number):
        """특정 제조사와 제품번호의 데이터를 DB에서 검색합니다."""
        df = self.load_db()
        mask = (
            (df["제조사"].str.lower().str.strip() == str(manufacturer).lower().strip()) & 
            (df["제품번호"].astype(str).str.strip() == str(product_number).strip())
        )
        result = df[mask]
        
        if not result.empty:
            return result.iloc[0].to_dict()
        return None

    def add_product(self, data_dict):
        """수집된 제품 데이터를 DB에 추가합니다."""
        df = self.load_db()
        
        data_dict["갱신일"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = pd.DataFrame([data_dict])
        
        mask = (
            (df["제조사"].str.lower().str.strip() == str(data_dict.get("제조사", "")).lower().strip()) & 
            (df["제품번호"].astype(str).str.strip() == str(data_dict.get("제품번호", "")).strip())
        )
        df = df[~mask]
        
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(self.db_path, index=False)

    @staticmethod
    def clean_filename(filename):
        if not isinstance(filename, str):
            return "unknown"
            
        cleaned = filename
        
        # 1.1. 무시할 특수문자 제거 (?, ", |)
        cleaned = re.sub(r'[?\"|]', '', cleaned)
        
        # 1.2. 시약명 맨 끝에 있는 특수문자 무시 (알파벳, 숫자, 닫는 괄호 등이 아니면 삭제)
        # ^a-zA-Z0-9) 가 아닌 문자로 끝나는 패턴을 공백으로 바꿈
        cleaned = re.sub(r'[^a-zA-Z0-9)]+$', '', cleaned)
        
        # 1.3. 윈도우에 파일명으로 쓸 수 없는 다른 특수문자는 언더바로 치환
        # 윈도우 예약 특수문자 중 남은 것: \ / : * < > 
        cleaned = re.sub(r'[\\/:*<>]+', '_', cleaned)
        
        cleaned = cleaned.replace('\n', ' ').replace('\r', '').strip()
        
        if not cleaned:
            return "unknown"
        return cleaned

    @staticmethod
    def normalize_temperature(temp_text):
        if not temp_text or pd.isna(temp_text):
            return "N/A"
            
        t = str(temp_text).lower()
        if any(kw in t for kw in ["deep freezer", "-80", "ultracold"]):
            return "deep freezer (-80도)"
        elif any(kw in t for kw in ["freezer", "냉동", "-20"]):
            return "freezer (-20도)"
        elif any(kw in t for kw in ["refrigerator", "냉장", "2-8", "2~8", "4도", "4°c"]):
            return "refrigerator (4 도)"
        elif any(kw in t for kw in ["rt", "room temp", "실온", "상온", "15-25", "ambient"]):
            return "RT"
            
        return temp_text.strip()
        
    @staticmethod
    def extract_sensitivity(text):
        if not text: return "N/A"
        t = str(text).lower()
        sensitivities = []
        
        if "moisture" in t or "hygroscopic" in t or "흡습" in t or "수분" in t:
            sensitivities.append("Moisture Sensitive")
        if "light" in t or "광" in t or "차광" in t:
            sensitivities.append("Light Sensitive")
        if "air" in t or "oxygen" in t or "공기" in t or "산소" in t:
            sensitivities.append("Air Sensitive")
        if "heat" in t or "열" in t:
            sensitivities.append("Heat Sensitive")
        if "argon" in t or "nitrogen" in t or "불활성" in t:
            sensitivities.append("Ar/N2 Required")
            
        if not sensitivities:
            return "N/A"
        return ", ".join(sensitivities)
