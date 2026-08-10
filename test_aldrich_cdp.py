"""
Aldrich SDS 다운로드 - Selenium CDP Network.requestWillBeSent 이벤트 사용
실제 PDF가 로드될 때의 URL을 포착하기 위한 진짜 네트워크 감시
"""
import json
import time
import os
import requests

from seleniumbase import Driver

def capture_aldrich_pdf_url(product_number: str, brand: str = "aldrich") -> str:
    """CDP 이벤트로 실제 PDF URL 포착"""
    
    driver = Driver(uc=True, headless=False)
    
    # CDP Network 활성화
    driver.execute_cdp_cmd("Network.enable", {})
    
    # 네트워크 요청을 가로채기 위한 리스너 설정
    # Selenium 4.x에서는 CDP listener를 별도로 등록해야 함
    
    captured_urls = []
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    print(f"  로딩: {url}")
    driver.get(url)
    time.sleep(12)
    
    # 방법 1: JavaScript로 실시간 fetch 가로채기
    # 페이지가 로드된 후 XHR/fetch로 SDS를 가져오는 경우
    all_requests = driver.execute_script("""
        // Service Worker가 있는지 확인
        var sw = navigator.serviceWorker ? 'exists' : 'none';
        
        // 현재 페이지의 모든 script 태그
        var scripts = Array.from(document.querySelectorAll('script[src]')).map(s => s.src);
        
        // 네트워크 요청을 보내는 코드 찾기
        return {
            sw: sw,
            scriptUrls: scripts.slice(0, 5)
        };
    """)
    print(f"  페이지 정보: {all_requests}")
    
    # 방법 2: 페이지의 document.body HTML 중 SDS URL 관련 내용 출력
    body_html = driver.execute_script("return document.body.innerHTML.slice(0, 5000);")
    
    import re
    # PDF, SDS, download 관련 URL 패턴 찾기
    urls = re.findall(r'https?://[^\s"\'<>]+(?:pdf|sds|download|document)[^\s"\'<>]*', body_html, re.IGNORECASE)
    print(f"  HTML 내 PDF/SDS URL: {urls[:5]}")
    
    # 방법 3: 실제 다운로드 버튼이 있는지 확인 (스크린샷에서 보임)
    # Chrome PDF 뷰어의 다운로드 버튼은 Shadow DOM 안에 있음
    shadow_download = driver.execute_script("""
        // Chrome PDF 뷰어의 다운로드 버튼 탐색
        var allElements = document.querySelectorAll('*');
        var pdfViewer = null;
        allElements.forEach(function(el) {
            if (el.shadowRoot) {
                var downloadBtn = el.shadowRoot.querySelector('#download');
                if (downloadBtn) pdfViewer = {found: true, href: downloadBtn.href || ''};
            }
        });
        return pdfViewer;
    """)
    print(f"  Shadow DOM 다운로드 버튼: {shadow_download}")
    
    # 방법 4: 현재 페이지 자체가 PDF인 경우 (직접 URL이 PDF)
    # 다른 URL 탐색: /en/sds/aldrich/803200 (언어 변경)
    for lang_url in [
        f"https://www.sigmaaldrich.com/US/en/sds/{brand}/{product_number}",
        f"https://www.sigmaaldrich.com/KR/en/sds/{brand}/{product_number}",
    ]:
        driver.get(lang_url)
        time.sleep(5)
        current = driver.current_url
        ct_html = driver.execute_script("return document.contentType || document.mimeType;")
        print(f"  {lang_url}")
        print(f"    -> 현재 URL: {current}, Content-Type: {ct_html}")
    
    driver.quit()
    return ""


if __name__ == "__main__":
    result = capture_aldrich_pdf_url("803200", "aldrich")
    print(f"결과: {result or '실패'}")
