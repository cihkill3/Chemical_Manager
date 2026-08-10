"""
Thermo Fisher PDF URL 발견! 실제 다운로드 테스트
URL 패턴: https://assets.thermofisher.com/chem-specs-pdf/retrievePdf?rootSku={sku}&sku={sku}.MD&erp_type=KR_E1_LCD&countryCode=kr
"""
import os
import re
import time
import requests
import PyPDF2
import io

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
    except:
        return ""

def get_tf_pdf_url(product_number: str, suffix: str = "MD") -> str:
    """
    Thermo Fisher SDS PDF URL 생성
    """
    # HTML entity decode: &amp; -> &
    return (
        f"https://assets.thermofisher.com/chem-specs-pdf/retrievePdf"
        f"?rootSku={product_number}&sku={product_number}.{suffix}&erp_type=KR_E1_LCD&countryCode=kr"
    )

def download_tf_sds(product_number: str, save_dir: str = ".") -> str:
    """
    Thermo Fisher SDS PDF 다운로드
    제품 페이지에서 실제 PDF URL 추출 후 다운로드
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    pdf_url = ""
    for suffix in ["", ".MD", ".06"]:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(3)
        
        html = d.page_source
        
        # assets.thermofisher.com/chem-specs-pdf/retrievePdf 패턴 찾기
        matches = re.findall(r'https://assets\.thermofisher\.com/chem-specs-pdf/retrievePdf[^"\'<\s&]+', html)
        if not matches:
            # HTML entity encoded 버전도 찾기
            matches_encoded = re.findall(r'https://assets\.thermofisher\.com/chem-specs-pdf/retrievePdf[^"\'<\s]+', html)
            if matches_encoded:
                matches = [m.replace('&amp;', '&') for m in matches_encoded]
        
        if matches:
            pdf_url = matches[0]
            print(f"  발견된 PDF URL: {pdf_url}")
            break
    
    d.quit()
    
    if not pdf_url:
        # Fallback: 알려진 패턴으로 URL 직접 생성
        pdf_url = get_tf_pdf_url(product_number, "MD")
        print(f"  Fallback URL: {pdf_url}")
    
    # PDF 다운로드 시도
    try:
        r = requests.get(pdf_url, headers=HEADERS, timeout=20, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        print(f"  상태: {r.status_code}, Content-Type: {ct}")
        
        if r.status_code == 200 and 'pdf' in ct.lower():
            text = read_pdf_text(r.content)
            print(f"  PDF 내용 (첫 300자): {text[:300]}")
            fname = os.path.join(save_dir, f"SDS_TF_{product_number}.pdf")
            with open(fname, 'wb') as f:
                f.write(r.content)
            print(f"  저장: {fname}")
            return fname
        else:
            print(f"  PDF 아님, 응답 크기: {len(r.content)}, 첫 200바이트: {r.content[:200]}")
    except Exception as e:
        print(f"  오류: {e}")
    
    return ""

def download_aldrich_sds(product_number: str, brand: str = "aldrich", save_dir: str = ".") -> str:
    """
    Sigma-Aldrich SDS PDF 다운로드
    제품 페이지에서 실제 PDF URL 찾기
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    d.get(url)
    time.sleep(10)
    
    html = d.page_source
    
    # 1. data-testid="sds-download" 등의 다운로드 버튼 확인
    pdf_links = re.findall(r'href=["\'](https?://[^"\']*\.pdf[^"\']*)["\']', html, re.IGNORECASE)
    print(f"  Aldrich PDF 링크: {pdf_links[:3]}")
    
    # 2. JSON 응답에서 URL 찾기
    json_patterns = re.findall(r'"(?:documentUrl|pdfUrl|url|downloadUrl)"\s*:\s*"(https?://[^"]+)"', html, re.IGNORECASE)
    print(f"  JSON URL 패턴: {json_patterns[:3]}")
    
    # 3. JavaScript 변수에서 URL 찾기
    js_patterns = re.findall(r'(?:pdfUrl|documentUrl|sdsUrl)\s*=\s*["\']?(https?://[^"\';\s]+)', html, re.IGNORECASE)
    print(f"  JS 변수 URL: {js_patterns[:3]}")
    
    # 4. script 태그에서 fetch URL 찾기
    fetch_patterns = re.findall(r'fetch\(["\']([^"\']+)["\']', html, re.IGNORECASE)
    sds_fetches = [u for u in fetch_patterns if 'sds' in u.lower() or 'document' in u.lower()]
    print(f"  Fetch SDS URL: {sds_fetches[:3]}")
    
    d.quit()
    return ""

if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("Thermo Fisher SDS PDF 다운로드 테스트")
    print("=" * 60)
    
    print("\n[L09319 - Fluorescein isothiocyanate]")
    result = download_tf_sds("L09319", save_dir)
    print(f"결과: {result or '실패'}\n")
    
    print("\n[L16400 - Acryloyloxy propyltrimethoxysilane]")
    result = download_tf_sds("L16400", save_dir)
    print(f"결과: {result or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("Sigma-Aldrich SDS PDF 탐색")
    print("=" * 60)
    
    print("\n[803200 - DTSSP]")
    result = download_aldrich_sds("803200", "aldrich", save_dir)
    print(f"결과: {result or '실패'}\n")
