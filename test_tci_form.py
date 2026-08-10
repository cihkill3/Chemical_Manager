"""
TCI SDS - 문서 섹션 전체 텍스트 덤프 및 버튼 시뮬레이션
"""
import os
import time
from seleniumbase import Driver

def simulate_tci_sds_click(product_number: str):
    d = Driver(uc=True, headless=True)
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    print(f"  [{product_number}] 페이지 로드 중...")
    d.get(url)
    time.sleep(3)
    
    # 폼 요소들 확인
    form_elements = d.execute_script("""
        var form = {};
        form.brand = document.getElementById('brandSelector') ? document.getElementById('brandSelector').value : null;
        form.prod = document.getElementById('sdsProductCode') ? document.getElementById('sdsProductCode').value : null;
        form.lang = document.getElementById('langSelector') ? document.getElementById('langSelector').value : null;
        
        // sds 버튼 찾기
        var buttons = [];
        document.querySelectorAll('button, a, input[type="button"], input[type="submit"]').forEach(function(el) {
            var txt = el.textContent.trim() || el.value || '';
            if (txt.toLowerCase().includes('sds') || el.id.toLowerCase().includes('sds') || (el.className && typeof el.className === 'string' && el.className.toLowerCase().includes('sds'))) {
                buttons.push({id: el.id, class: el.className, text: txt});
            }
        });
        form.buttons = buttons;
        return form;
    """)
    print("  폼 요소들:")
    print(f"    brand: {form_elements.get('brand')}")
    print(f"    prod: {form_elements.get('prod')}")
    print(f"    lang: {form_elements.get('lang')}")
    print(f"    buttons: {form_elements.get('buttons')}")
    
    d.quit()

if __name__ == "__main__":
    simulate_tci_sds_click("C0119")
