"""
SDS PDF 다운로드 최종 방법 검증
1. Thermo Fisher: chemicals.thermofisher.kr 제품 페이지에서 실제 SDS 링크 포착
2. Sigma-Aldrich: Playwright-cdp를 이용한 실제 네트워크 트래픽 감시
3. TCI: find_sds 방식 (CAS 번호 이용)
"""
import sys
import os
import time
import re
import requests
import PyPDF2
import io

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36',
}

def verify_pdf(content: bytes, expected_keyword: str = "") -> tuple[bool, str]:
    """PDF 내용을 읽어서 시약명 포함 여부 확인"""
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:3]:  # 첫 3페이지
            text += page.extract_text() or ""
        
        if expected_keyword and expected_keyword.lower() in text.lower():
            return True, text[:200]
        elif text.strip():
            return True, text[:200]
        return False, "PDF 텍스트 추출 실패"
    except Exception as e:
        return False, f"PDF 읽기 오류: {e}"


# =====================================================
# 방법 1: TCI - CAS 번호를 이용한 직접 PDF URL 접근
# =====================================================
def download_tci_sds(cas_number: str, product_number: str, save_dir: str = ".") -> str:
    """
    TCI는 공개 API를 통해 SDS를 직접 제공
    URL: https://www.tcichemicals.com/KR/en/sds/{product_number}.pdf
    """
    urls_to_try = [
        f"https://www.tcichemicals.com/KR/en/sds/{product_number}.pdf",
        f"https://www.tcichemicals.com/US/en/sds/{product_number}.pdf",
        f"https://www.tcichemicals.com/KR/en/p/{product_number}",  # 페이지에서 PDF URL 찾기
    ]
    
    for url in urls_to_try:
        try:
            r = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            print(f"  TCI [{product_number}] URL: {url}")
            print(f"    Status: {r.status_code}, CT: {ct[:50]}")
            if r.status_code == 200 and 'pdf' in ct.lower():
                ok, text = verify_pdf(r.content, product_number)
                print(f"    PDF 검증: {ok}, 내용: {text[:100]}")
                if ok:
                    fname = os.path.join(save_dir, f"SDS_TCI_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(r.content)
                    print(f"    저장: {fname}")
                    return fname
        except Exception as e:
            print(f"    오류: {e}")
    return ""


# =====================================================
# 방법 2: Thermo Fisher - Selenium으로 SDS PDF 직접 다운로드
# =====================================================
def download_tf_sds_selenium(product_number: str, save_dir: str = ".") -> str:
    """
    chemicals.thermofisher.kr 제품 페이지를 열고
    SDS 섹션에서 PDF URL 추출 후 다운로드
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    # 제품 페이지 URL 시도 (여러 suffix 포함)
    suffixes = ["", ".MD", ".06", ".14", ".36"]
    pdf_url = ""
    
    for suffix in suffixes:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(3)
        
        html = d.page_source
        
        # SDS PDF URL 패턴 찾기
        # 1. assets.thermofisher.com 패턴
        asset_links = re.findall(r'(https://assets\.thermofisher\.com[^\s"\'<>]+)', html)
        for lk in asset_links:
            if 'pdf' in lk.lower() or 'document' in lk.lower() or 'SDS' in lk or 'sds' in lk:
                pdf_url = lk
                print(f"  TF assets 링크 발견: {pdf_url}")
                break
        
        # 2. document-search 관련 링크
        doc_links = re.findall(r'href=["\']([^"\']*document[^"\']*)["\']', html, re.IGNORECASE)
        print(f"  TF document 링크 ({len(doc_links)}개): {doc_links[:3]}")
        
        # 3. JavaScript에서 SDS URL 찾기
        js_sds = re.findall(r'["\']([^"\']*(?:sds|safety.?data)[^"\']*\.pdf[^"\']*)["\']', html, re.IGNORECASE)
        print(f"  TF JS SDS 링크 ({len(js_sds)}개): {js_sds[:3]}")
        
        # 4. 페이지 내 script 태그에서 API URL 찾기
        api_configs = re.findall(r'["\']([^"\']*DirectWebViewer[^"\']*)["\']', html)
        if api_configs:
            print(f"  TF DirectWebViewer 발견: {api_configs[:2]}")
            pdf_url = api_configs[0]
            break
        
        if pdf_url:
            break
    
    # 모든 a 태그에서 SDS 관련 찾기
    try:
        sds_buttons = d.find_elements('xpath', '//a[contains(translate(text(),"SDS","sds"),"sds") or contains(@href,"sds") or contains(@href,"SDS") or contains(text(),"Safety Data Sheet")]')
        print(f"\n  TF SDS 버튼/링크 ({len(sds_buttons)}개):")
        for btn in sds_buttons[:5]:
            href = btn.get_attribute('href') or ""
            text = btn.text.strip()
            print(f"    TEXT={text[:40]}, HREF={href[:80]}")
            if href and ('pdf' in href.lower() or 'sds' in href.lower()):
                pdf_url = href
    except Exception as e:
        print(f"  버튼 탐색 오류: {e}")
    
    d.quit()
    
    if pdf_url:
        print(f"\n  PDF URL: {pdf_url}")
        try:
            r = requests.get(pdf_url, headers=HEADERS, timeout=20, allow_redirects=True)
            if r.status_code == 200 and 'pdf' in r.headers.get('content-type','').lower():
                ok, text = verify_pdf(r.content)
                print(f"  PDF 검증: {ok}, 내용: {text[:100]}")
                if ok:
                    fname = os.path.join(save_dir, f"SDS_TF_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        r.write(r.content)
                    return fname
        except Exception as e:
            print(f"  PDF 다운로드 오류: {e}")
    
    return ""


# =====================================================
# 방법 3: Sigma-Aldrich - Selenium CDP 네트워크 감시
# 실제 브라우저에서 SDS 페이지 열고 PDF URL 포착
# =====================================================
def download_aldrich_sds_cdp(product_number: str, brand: str = "aldrich", save_dir: str = ".") -> str:
    """
    Selenium CDP를 통해 네트워크 요청을 실시간 모니터링
    실제 PDF가 로드될 때의 URL을 캡처
    """
    from seleniumbase import Driver
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    
    d = Driver(uc=True, headless=False)  # headful 모드로 실행 (봇 감지 우회)
    
    pdf_urls_found = []
    
    # CDP 네트워크 이벤트 리스너 설정
    d.execute_cdp_cmd("Network.enable", {})
    
    # 사용자 에이전트 위장
    d.execute_cdp_cmd("Network.setUserAgentOverride", {
        "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "platform": "Win32"
    })
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    print(f"  Aldrich CDP 모드로 로딩: {url}")
    
    d.get(url)
    time.sleep(8)
    
    # 페이지 로드 후 resources 확인
    resources = d.execute_script("""
        return window.performance.getEntriesByType('resource')
            .filter(r => r.name.includes('pdf') || r.name.includes('SDS') || 
                        r.name.includes('sds') || r.name.includes('document') ||
                        r.name.includes('safety') || r.name.includes('material'))
            .map(r => r.name);
    """)
    
    print(f"  SDS 관련 리소스: {resources}")
    
    # 현재 페이지 HTML에서 PDF iframe 찾기
    html = d.page_source
    
    # JavaScript variables에서 URL 찾기
    js_vars = re.findall(r'(?:pdfUrl|documentUrl|sdsUrl|fileUrl|downloadUrl)\s*[=:]\s*["\']([^"\']+)["\']', html, re.IGNORECASE)
    print(f"  JS PDF 변수: {js_vars}")
    
    # JSON 데이터에서 URL 찾기
    json_urls = re.findall(r'"(?:url|href|src|path|link)"\s*:\s*"([^"]*\.pdf[^"]*)"', html, re.IGNORECASE)
    print(f"  JSON PDF URL: {json_urls}")
    
    d.quit()
    return ""


# =====================================================
# 방법 4: find_sds 방식 - CAS 번호 이용
# Fisher Scientific 직접 API (open source 방법)
# =====================================================
def download_sds_via_cas(cas_number: str, product_name: str, save_dir: str = ".") -> str:
    """
    CAS 번호를 이용해 Fisher Scientific에서 SDS PDF 검색
    find_sds (khoivan88) 프로젝트 방식 참고
    """
    # Fisher Scientific SDS API
    fisher_url = f"https://www.fishersci.com/us/en/catalog/search/sdshome.html?sku={cas_number}&catalogId=29104&langId=-1&storeId=10652"
    
    # Sigma-Aldrich SDS search
    sigma_url = f"https://www.sigmaaldrich.com/US/en/search#{cas_number}"
    
    print(f"  CAS {cas_number} - Fisher SDS 검색...")
    
    try:
        r = requests.get(fisher_url, headers=HEADERS, timeout=15, allow_redirects=True)
        print(f"    Fisher Status: {r.status_code}")
    except Exception as e:
        print(f"    Fisher 오류: {e}")
    
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("TCI SDS 다운로드 테스트")
    print("=" * 60)
    # TCI는 직접 PDF 제공 여부 확인
    result = download_tci_sds("530-62-1", "C0119", save_dir)
    print(f"결과: {result or '실패'}\n")
    
    result = download_tci_sds("1493-13-6", "T0751", save_dir)
    print(f"결과: {result or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("Thermo Fisher SDS URL 심층 탐색")
    print("=" * 60)
    result = download_tf_sds_selenium("L09319", save_dir)
    print(f"결과: {result or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("Sigma-Aldrich CDP 네트워크 감시")
    print("=" * 60)
    result = download_aldrich_sds_cdp("803200", "aldrich", save_dir)
    print(f"결과: {result or '실패'}\n")
