"""
TCI SDS - 문서 섹션의 HTML 덤프
"""
import os
import time
from seleniumbase import Driver

def dump_tci_sds_html(product_number: str):
    d = Driver(uc=True, headless=True)
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    print(f"  [{product_number}] 페이지 로드 중...")
    d.get(url)
    time.sleep(3)
    
    # "문서" 또는 "Documents" 탭이 보이면 클릭 (필요 시)
    html = d.execute_script("""
        // 1. SDS 탭 클릭
        var links = document.querySelectorAll('a');
        for (var l of links) {
            if (l.textContent.trim().toLowerCase() === 'sds' || l.textContent.trim().toLowerCase() === 'documents') {
                l.click();
            }
        }
        
        // 2. 문서 섹션 전체 HTML 반환
        var sdsArea = document.getElementById('docomentsSectionPDP');
        if (sdsArea) {
            var parent = sdsArea.parentElement;
            return parent ? parent.outerHTML : sdsArea.outerHTML;
        }
        return "Not Found";
    """)
    
    print(f"--- HTML 덤프 시작 ---")
    if html and len(html) > 50:
        print(html[:2000])  # 앞부분 출력
    else:
        print(html)
        # 덤프 실패 시 전체 body 텍스트를 통해 sds 주변 확인
        text = d.execute_script("return document.body.innerText;")
        sds_idx = text.lower().find("sds")
        if sds_idx != -1:
            print("SDS 주변 텍스트:")
            print(text[max(0, sds_idx-100):min(len(text), sds_idx+500)])
    print(f"--- HTML 덤프 끝 ---")
    
    d.quit()

if __name__ == "__main__":
    dump_tci_sds_html("C0119")
