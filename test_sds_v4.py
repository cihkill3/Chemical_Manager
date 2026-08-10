"""
SDS PDF 다운로드 - 최종 구현
Aldrich: 브라우저 내 PDF 뷰어의 다운로드 버튼 클릭
TF: SDS 전용 문서 URL 탐색
TCI: SDS 섹션 탐색 후 PDF URL 추출
"""
import os
import re
import time
import requests
import PyPDF2
import io
from html import unescape

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36',
    'Referer': 'https://chemicals.thermofisher.kr/',
    'Accept': 'application/pdf,*/*',
}

def read_pdf_text(content: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:3]:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"[오류: {e}]"


# ================================================================
# Sigma-Aldrich: 브라우저 내장 PDF 뷰어 다운로드 버튼 클릭
# ================================================================
def download_aldrich_sds(product_number: str, brand: str, save_dir: str) -> str:
    """
    Aldrich SDS 페이지 PDF 뷰어에서 다운로드 버튼 클릭
    스크린샷에서 보이는 상단 다운로드 버튼(⬇️ 아이콘) 클릭
    """
    import glob
    from seleniumbase import Driver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    
    abs_save_dir = os.path.abspath(save_dir)
    existing_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf")))
    
    d = Driver(uc=True, headless=False)
    
    # 다운로드 경로 설정
    d.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": abs_save_dir
    })
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    print(f"  [{product_number}] 로딩: {url}")
    d.get(url)
    time.sleep(8)
    
    # 방법 1: 브라우저 내장 PDF 뷰어 다운로드 버튼
    # 스크린샷에서 보이는 뷰어는 Chrome 내장 PDF 뷰어
    # 키보드 단축키 Ctrl+S 로 저장 시도
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    
    # 페이지가 iframe 없이 직접 PDF를 표시하는 경우 (Chrome PDF 뷰어)
    # 이 경우 URL이 pdf파일 직접 URL이어야 함
    current_url = d.current_url
    print(f"  현재 URL: {current_url}")
    
    # 방법 2: SDS PDF를 iframe으로 표시하는 경우 src 가져오기
    iframes = d.execute_script("""
        return Array.from(document.querySelectorAll('iframe, embed')).map(el => ({
            tag: el.tagName,
            src: el.src || el.getAttribute('src') || el.data || '',
            id: el.id, class: el.className
        }));
    """)
    print(f"  iframe/embed: {iframes}")
    
    # 방법 3: 실제 PDF URL이 페이지에 숨겨진 경우 찾기
    hidden_pdf = d.execute_script("""
        // window.__NEXT_DATA__ 또는 redux store 확인
        var nextData = window.__NEXT_DATA__;
        if (nextData) {
            return JSON.stringify(nextData).slice(0, 500);
        }
        // 전역 변수 확인
        for (var key of Object.keys(window)) {
            try {
                var val = window[key];
                if (val && typeof val === 'object' && JSON.stringify(val).includes('.pdf')) {
                    return key + ': ' + JSON.stringify(val).slice(0, 200);
                }
            } catch(e) {}
        }
        return null;
    """)
    print(f"  Hidden PDF data: {str(hidden_pdf)[:200] if hidden_pdf else 'None'}")
    
    # 방법 4: 직접 다운로드 URL 구성 시도
    # sigmaaldrich API endpoint 탐색
    # /rest/documents/sds/{brand}/{product_number}
    api_endpoints = [
        f"https://www.sigmaaldrich.com/api/sds/download/{product_number}",
        f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}?format=pdf",
        f"https://www.sigmaaldrich.com/KR/ko/product/{brand}/{product_number}?lang=ko&region=KR&downloadSDS=true",
    ]
    
    for api_url in api_endpoints:
        try:
            # Selenium 세션의 쿠키로 requests 요청
            cookies = {c['name']: c['value'] for c in d.get_cookies()}
            r = requests.get(api_url, headers={
                **HEADERS,
                'Referer': url,
                'Cookie': '; '.join([f"{k}={v}" for k, v in cookies.items()])
            }, timeout=15, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            print(f"  API [{api_url[:60]}]: {r.status_code}, {ct[:30]}")
            if r.status_code == 200 and 'pdf' in ct.lower():
                text = read_pdf_text(r.content)
                print(f"  PDF 성공! 내용: {text[:200]}")
                fname = os.path.join(abs_save_dir, f"SDS_Aldrich_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(r.content)
                d.quit()
                return fname
        except Exception as e:
            print(f"  오류: {e}")
    
    # 방법 5: 스크린샷에서 보이는 PDF는 서버 렌더링 PDF (전체 페이지가 PDF viewer)
    # JavaScript로 PDF blob URL 가져오기
    pdf_blob = d.execute_script("""
        // fetch API로 현재 페이지 내용을 PDF로 가져오기 시도
        return document.querySelector('embed') ? document.querySelector('embed').src : null;
    """)
    print(f"  PDF blob/embed URL: {pdf_blob}")
    
    # 방법 6: Ctrl+S 키로 저장 시도  
    try:
        body = d.find_element('css selector', 'body')
        body.send_keys(Keys.CONTROL, 's')
        time.sleep(3)
        print("  Ctrl+S 실행됨")
    except Exception as e:
        print(f"  Ctrl+S 오류: {e}")
    
    time.sleep(5)
    new_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf"))) - existing_pdfs
    if new_pdfs:
        print(f"  새 PDF: {new_pdfs}")
    
    # 방법 7: 헤드리스 없이 실행된 Chrome에서 PDF 뷰어 다운로드 버튼 위치 파악
    # 스크린샷 저장
    screenshot_path = os.path.join(abs_save_dir, f"aldrich_{product_number}_v2.png")
    d.save_screenshot(screenshot_path)
    print(f"  스크린샷 저장: {screenshot_path}")
    
    d.quit()
    return ""


# ================================================================
# TCI - SDS 버튼 탐색 및 직접 다운로드
# ================================================================
def download_tci_sds(product_number: str, cas_number: str, save_dir: str) -> str:
    """
    TCI SDS 다운로드
    #docomentsSectionPDP 탭의 SDS 버튼에서 실제 PDF URL 추출
    """
    import glob
    from seleniumbase import Driver
    
    abs_save_dir = os.path.abspath(save_dir)
    existing_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf")))
    
    d = Driver(uc=True, headless=True)
    d.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": abs_save_dir
    })
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}#docomentsSectionPDP"
    print(f"  [{product_number}] 로딩: {url}")
    d.get(url)
    time.sleep(6)
    
    # SDS 탭 클릭 시도
    try:
        sds_tab = d.find_element('css selector', '[href="#docomentsSectionPDP"], [data-tab="sds"], a[href*="docoment"]')
        d.execute_script("arguments[0].click();", sds_tab)
        time.sleep(3)
        print(f"  SDS 탭 클릭됨")
    except Exception as e:
        print(f"  SDS 탭 탐색 오류: {e}")
    
    # 문서 섹션 로드
    doc_section_html = d.execute_script("""
        var el = document.getElementById('docomentsSectionPDP') || 
                 document.querySelector('[id*="document"]') ||
                 document.querySelector('[class*="document"]');
        return el ? el.outerHTML.slice(0, 2000) : 'Not found';
    """)
    print(f"  문서 섹션 HTML: {doc_section_html[:500]}")
    
    # SDS 관련 링크에서 실제 PDF URL 추출
    sds_pdf_urls = d.execute_script("""
        var urls = [];
        // 모든 링크 확인
        document.querySelectorAll('a').forEach(function(a) {
            var href = a.href || '';
            var text = a.textContent.toLowerCase();
            // SDS 문서 직접 URL
            if ((text.includes('sds') || text.includes('safety data')) && 
                (href.includes('sds') || href.includes('.pdf') || href.includes('download'))) {
                urls.push({text: text.trim().slice(0,40), href: href});
            }
        });
        return urls;
    """)
    print(f"  SDS PDF 링크: {sds_pdf_urls}")
    
    # 알려진 TCI SDS URL 패턴 시도
    # https://www.tcichemicals.com/assets/sds/{CAS}-KR.pdf 등
    tci_sds_patterns = [
        f"https://www.tcichemicals.com/assets/sds/{cas_number.replace('-', '')}.pdf",
        f"https://www.tcichemicals.com/sds/{product_number}.pdf",
        f"https://www.tcichemicals.com/sds-download/{product_number}",
        f"https://www.tcichemicals.com/api/sds/{product_number}?lang=en&region=KR",
    ]
    
    cookies = {c['name']: c['value'] for c in d.get_cookies()}
    
    for api_url in tci_sds_patterns:
        try:
            r = requests.get(api_url, headers={
                **HEADERS,
                'Referer': url,
                'Cookie': '; '.join([f"{k}={v}" for k, v in cookies.items()])
            }, timeout=15, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            print(f"  TCI API [{api_url[:60]}]: {r.status_code}, {ct[:30]}")
            if r.status_code == 200 and 'pdf' in ct.lower():
                text = read_pdf_text(r.content)
                print(f"  PDF 성공! 내용: {text[:200]}")
                fname = os.path.join(abs_save_dir, f"SDS_TCI_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(r.content)
                d.quit()
                return fname
        except Exception as e:
            print(f"  TCI API 오류: {e}")
    
    d.quit()
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("Sigma-Aldrich SDS 다운로드 버튼 클릭")
    print("=" * 60)
    result = download_aldrich_sds("803200", "aldrich", save_dir)
    print(f"\n결과: {result or '실패'}")
    
    print("\n" + "=" * 60)
    print("TCI SDS 다운로드 탐색")
    print("=" * 60)
    result = download_tci_sds("C0119", "530-62-1", save_dir)
    print(f"\n결과: {result or '실패'}")
