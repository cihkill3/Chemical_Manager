import os
import pandas as pd
from datetime import datetime
from PyPDF2 import PdfMerger
from db_manager import DBManager

class PDFManager:
    def __init__(self, sds_folder="sds"):
        self.sds_folder = sds_folder
        self.db_manager = DBManager()
        
        if not os.path.exists(self.sds_folder):
            os.makedirs(self.sds_folder)

    def merge_sds_after_date(self, target_date_str, output_filename=None):
        """
        주어진 날짜(YYYY-MM-DD) 이후에 다운로드(DB 갱신일 기준)된 SDS 파일들을 병합합니다.
        """
        try:
            target_date = datetime.strptime(target_date_str, "%Y-%m-%d")
        except ValueError:
            print(f"Error: 잘못된 날짜 형식입니다. 'YYYY-MM-DD' 형식으로 입력해주세요. (입력값: {target_date_str})")
            return False

        df = self.db_manager.load_db()
        if df.empty:
            print("DB가 비어있습니다.")
            return False

        merger = PdfMerger()
        files_merged = 0

        for index, row in df.iterrows():
            date_str = row.get("갱신일")
            file_path = row.get("SDS_Local_Path")
            
            if pd.isna(date_str) or pd.isna(file_path):
                continue
                
            try:
                # DB의 "YYYY-MM-DD HH:MM:SS" 형식 파싱
                updated_date = datetime.strptime(str(date_str), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue

            # target_date의 자정(00:00:00) 이후에 업데이트된 모든 항목 찾기
            if updated_date >= target_date:
                if os.path.exists(str(file_path)):
                    try:
                        merger.append(str(file_path))
                        files_merged += 1
                        print(f"병합 대상 추가: {file_path}")
                    except Exception as e:
                        print(f"파일 병합 실패 ({file_path}): {e}")
                else:
                    print(f"경고: DB에 기록된 파일이 실제 경로에 없습니다. ({file_path})")

        if files_merged > 0:
            if not output_filename:
                output_filename = f"sds_merged_{target_date.strftime('%Y%m%d')}.pdf"
            
            try:
                merger.write(output_filename)
                merger.close()
                print(f"성공: {files_merged}개의 PDF 파일이 '{output_filename}'로 병합되었습니다.")
                return True
            except Exception as e:
                print(f"최종 PDF 저장 실패: {e}")
                merger.close()
                return False
        else:
            print(f"정보: {target_date_str} 이후에 생성된 유효한 SDS PDF 파일이 없습니다.")
            merger.close()
            return False

if __name__ == "__main__":
    import pandas as pd # needed for isna check inside class when run independently
    # Simple test execution block
    manager = PDFManager()
    # manager.merge_sds_after_date("2023-01-01")
