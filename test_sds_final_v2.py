"""
핵심 발견: Aldrich /KR/en/ URL = 직접 PDF!
이제 세 제조사 모두 다운로드 구현
"""
import os
import re
import time
import requests
import PyPDF2
import io
from html import unescape

HEADERS_BROWSER = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}

def read_pdf_text(content: bytes, pages: int = 3) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:pages]:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"[PDF 파싱 오류: {e}]"


# ================================================================
# 1. Sigma-Aldrich SDS PDF 직접 다운로드
# ================================================================
def download_aldrich_sds(product_number: str, brand: str = "aldrich", save_dir: str = ".") -> str:
    """
    Aldrich SDS PDF 다운로드
    핵심: /KR/en/ URL이 직접 PDF를 반환함
    """
    # 영어 URL이 직접 PDF 반환
    url = f"https://www.sigmaaldrich.com/KR/en/sds/{brand}/{product_number}"
    
    s = requests.Session()
    s.headers.update(HEADERS_BROWSER)
    
    try:
        r = s.get(url, timeout=30, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        print(f"  [{product_number}] URL: {url}")
        print(f"    Status: {r.status_code}, Content-Type: {ct}, Size: {len(r.content)}")
        
        if r.status_code == 200 and 'pdf' in ct.lower():
            text = read_pdf_text(r.content)
            print(f"    PDF 내용 (첫 200자): {text[:200]}")
            fname = os.path.join(save_dir, f"SDS_Aldrich_{brand}_{product_number}.pdf")
            with open(fname, 'wb') as f:
                f.write(r.content)
            print(f"    저장: {fname}")
            return fname
        else:
            # Selenium으로 시도
            print(f"    requests 실패, Selenium 시도...")
            return download_aldrich_sds_selenium(product_number, brand, save_dir)
    except Exception as e:
        print(f"    오류: {e}")
        return download_aldrich_sds_selenium(product_number, brand, save_dir)


def download_aldrich_sds_selenium(product_number: str, brand: str, save_dir: str) -> str:
    """Selenium으로 Aldrich SDS 다운로드 (requests 실패시 fallback)"""
    from seleniumbase import Driver
    import glob
    
    abs_save_dir = os.path.abspath(save_dir)
    existing_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf")))
    
    d = Driver(uc=True, headless=True)
    d.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": abs_save_dir
    })
    
    url = f"https://www.sigmaaldrich.com/KR/en/sds/{brand}/{product_number}"
    d.get(url)
    time.sleep(8)
    
    ct = d.execute_script("return document.contentType;")
    current_url = d.current_url
    print(f"    Selenium -> URL: {current_url}, CT: {ct}")
    
    # PDF content를 blob으로 가져오기
    if ct == 'application/pdf':
        try:
            pdf_bytes = d.execute_script("""
                var xhr = new XMLHttpRequest();
                xhr.open('GET', document.URL, false);
                xhr.responseType = 'arraybuffer';
                xhr.send();
                var arr = new Uint8Array(xhr.response);
                var result = [];
                for (var i = 0; i < arr.length; i++) result.push(arr[i]);
                return result;
            """)
            if pdf_bytes:
                content = bytes(pdf_bytes)
                text = read_pdf_text(content)
                print(f"    PDF 획득! 내용: {text[:200]}")
                fname = os.path.join(abs_save_dir, f"SDS_Aldrich_{brand}_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(content)
                d.quit()
                return fname
        except Exception as e:
            print(f"    PDF blob 오류: {e}")
    
    d.quit()
    return ""


# ================================================================
# 2. Thermo Fisher SDS PDF 다운로드 (제품 스펙 URL 패턴 사용)
# ================================================================
def download_tf_sds(product_number: str, save_dir: str = ".") -> str:
    """
    Thermo Fisher SDS PDF 다운로드
    제품 페이지에서 retrievePdf URL 추출
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    pdf_url = ""
    for suffix in ["", ".MD", ".06", ".14", ".36"]:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(4)
        
        html = unescape(d.page_source)
        
        matches = re.findall(
            r'https://assets\.thermofisher\.com/chem-specs-pdf/retrievePdf\?[^\s"\'<>]+',
            html
        )
        if matches:
            pdf_url = matches[0]
            print(f"  [{product_number}] URL 발견 (suffix={suffix!r}): {pdf_url}")
            break
    
    d.quit()
    
    if not pdf_url:
        print(f"  [{product_number}] URL 미발견")
        return ""
    
    try:
        r = requests.get(pdf_url, headers=HEADERS_BROWSER, timeout=20, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        print(f"  Status: {r.status_code}, CT: {ct}")
        
        if r.status_code == 200 and 'pdf' in ct.lower():
            text = read_pdf_text(r.content)
            print(f"  PDF 내용 (첫 200자): {text[:200]}")
            fname = os.path.join(save_dir, f"SDS_TF_{product_number}.pdf")
            with open(fname, 'wb') as f:
                f.write(r.content)
            print(f"  저장: {fname}")
            return fname
    except Exception as e:
        print(f"  오류: {e}")
    
    return ""


# ================================================================
# 3. TCI SDS PDF 다운로드
# ================================================================
def download_tci_sds(product_number: str, save_dir: str = ".") -> str:
    """
    TCI SDS PDF 다운로드
    제품 페이지 문서 섹션에서 SDS 링크 추출
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    d.get(url)
    time.sleep(6)
    
    pdf_url = ""
    
    # SDS 다운로드 버튼/링크 찾기
    sds_links = d.execute_script("""
        var results = [];
        document.querySelectorAll('a[href]').forEach(function(a) {
            var href = a.href || '';
            var text = (a.textContent || '').toLowerCase().trim();
            // SDS 관련 링크
            if (text === 'sds' || text.includes('safety data') || 
                href.toLowerCase().includes('/sds/') || 
                (href.includes('.pdf') && (text.includes('sds') || text.includes('safety')))) {
                results.push({text: text.slice(0,40), href: href.slice(0,150)});
            }
        });
        return results;
    """)
    print(f"  [{product_number}] SDS 링크: {sds_links}")
    
    # SDS 버튼 클릭 시도
    for link in sds_links:
        href = link.get('href', '')
        if href and href != url and '#' not in href:
            pdf_url = href
            break
    
    # 만약 링크가 없으면 버튼 클릭
    if not pdf_url:
        try:
            # SDS 버튼 클릭 (모달이나 다운로드 유발)
            btn = d.find_element('xpath', '//a[normalize-space(text())="SDS" or contains(@class,"sds")]')
            original_url = d.current_url
            d.execute_script("arguments[0].click();", btn)
            time.sleep(3)
            
            new_url = d.current_url
            if new_url != original_url:
                pdf_url = new_url
                print(f"  클릭 후 새 URL: {new_url}")
        except Exception as e:
            print(f"  SDS 버튼 클릭 오류: {e}")
    
    d.quit()
    
    if pdf_url and pdf_url.endswith('.pdf'):
        try:
            r = requests.get(pdf_url, headers=HEADERS_BROWSER, timeout=20, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            if r.status_code == 200 and 'pdf' in ct.lower():
                text = read_pdf_text(r.content)
                print(f"  TCI PDF 내용 (첫 200자): {text[:200]}")
                fname = os.path.join(save_dir, f"SDS_TCI_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(r.content)
                print(f"  저장: {fname}")
                return fname
        except Exception as e:
            print(f"  TCI PDF 오류: {e}")
    
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("1. Sigma-Aldrich SDS 직접 다운로드 테스트")
    print("=" * 60)
    for pn, br in [("803200", "aldrich"), ("909270", "aldrich")]:
        r = download_aldrich_sds(pn, br, save_dir)
        print(f"결과: {r or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("2. Thermo Fisher SDS 다운로드 테스트")
    print("=" * 60)
    for pn in ["L09319", "L16400"]:
        r = download_tf_sds(pn, save_dir)
        print(f"결과: {r or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("3. TCI SDS 다운로드 테스트")
    print("=" * 60)
    for pn in ["C0119", "T0751"]:
        r = download_tci_sds(pn, save_dir)
        print(f"결과: {r or '실패'}\n")
