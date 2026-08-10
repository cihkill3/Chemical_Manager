"""
Thermo Fisher PDF URL 파라미터 조합을 정확히 추출하여 다운로드 시도
"""
import os
import re
import time
import requests
import urllib.parse
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
    except Exception as e:
        return f"[오류: {e}]"

def get_tf_sds_url_from_page(product_number: str) -> str:
    """제품 페이지에서 실제 SDS PDF URL 추출 (HTML entity 올바르게 디코딩)"""
    from seleniumbase import Driver
    from html import unescape
    
    d = Driver(uc=True, headless=True)
    
    pdf_url = ""
    for suffix in ["", ".MD", ".06", ".14"]:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(4)
        
        html = d.page_source
        
        # HTML entity decode 후 URL 찾기
        html_decoded = unescape(html)
        
        # assets.thermofisher.com/chem-specs-pdf/retrievePdf 패턴 찾기
        matches = re.findall(
            r'https://assets\.thermofisher\.com/chem-specs-pdf/retrievePdf\?[^\s"\'<>]+',
            html_decoded
        )
        
        if matches:
            pdf_url = matches[0]
            print(f"  발견된 URL (suffix={suffix!r}): {pdf_url}")
            break
        
        # JavaScript 실행으로 URL 가져오기
        try:
            js_result = d.execute_script("""
                var links = [];
                document.querySelectorAll('a[href]').forEach(function(a) {
                    if (a.href && a.href.includes('chem-specs-pdf')) {
                        links.push(a.href);
                    }
                });
                return links;
            """)
            if js_result:
                print(f"  JS에서 발견된 URL ({suffix!r}): {js_result}")
                pdf_url = js_result[0]
                break
        except Exception as e:
            pass
    
    d.quit()
    return pdf_url


def try_tf_sds_download(product_number: str, save_dir: str) -> str:
    """Thermo Fisher SDS PDF 다운로드"""
    print(f"\n[{product_number}]")
    
    pdf_url = get_tf_sds_url_from_page(product_number)
    
    if not pdf_url:
        # 발견된 URL 패턴으로 수동 구성
        # 원본: https://assets.thermofisher.com/chem-specs-pdf/retrievePdf?rootSku=L09319&sku=L09319.MD&erp_type=KR_E1_LCD&countryCode=kr
        pdf_url = (
            f"https://assets.thermofisher.com/chem-specs-pdf/retrievePdf"
            f"?rootSku={product_number}&sku={product_number}.MD&erp_type=KR_E1_LCD&countryCode=kr"
        )
        print(f"  Fallback URL: {pdf_url}")
    
    # URL 검사
    print(f"  최종 URL: {pdf_url}")
    
    try:
        r = requests.get(pdf_url, headers=HEADERS, timeout=20, allow_redirects=True)
        ct = r.headers.get('content-type', '')
        print(f"  Status: {r.status_code}, Content-Type: {ct}")
        
        if r.status_code == 200 and 'pdf' in ct.lower():
            text = read_pdf_text(r.content)
            print(f"  PDF 성공! 내용 (첫 300자): {text[:300]}")
            fname = os.path.join(save_dir, f"SDS_TF_{product_number}.pdf")
            with open(fname, 'wb') as f:
                f.write(r.content)
            print(f"  저장: {fname}")
            return fname
        else:
            print(f"  실패. 응답: {r.content[:200]}")
    except Exception as e:
        print(f"  오류: {e}")
    
    return ""


def try_aldrich_sds_via_selenium(product_number: str, brand: str, save_dir: str) -> str:
    """
    Selenium으로 Aldrich SDS 페이지를 열고
    CDP를 통해 실제 PDF 바이너리를 다운로드
    """
    from seleniumbase import Driver
    from html import unescape
    
    d = Driver(uc=True, headless=False)  # headful
    
    # SDS 페이지 로드
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    print(f"  Loading: {url}")
    d.get(url)
    time.sleep(10)  # JavaScript 렌더링 대기
    
    html = d.page_source
    html_decoded = unescape(html)
    
    # 모든 가능한 PDF URL 탐색
    # 1. 직접 링크
    pdf_hrefs = re.findall(r'href=["\']([^"\']*\.pdf[^"\']*)["\']', html_decoded, re.IGNORECASE)
    print(f"  PDF hrefs: {pdf_hrefs[:5]}")
    
    # 2. script/json 내부
    data_urls = re.findall(r'"(?:url|href|path|link|document[Uu]rl)"\s*:\s*"([^"]*\.pdf[^"]*)"', html_decoded, re.IGNORECASE)
    print(f"  Data URLs: {data_urls[:5]}")
    
    # 3. 현재 URL 확인 (redirect된 경우)
    current_url = d.current_url
    print(f"  현재 URL: {current_url}")
    
    # 4. Chrome downloads API 대신 직접 링크 다운로드 가능 여부
    # CDP를 통해 페이지 내부의 a[download] 요소 찾기
    download_links = d.execute_script("""
        var links = [];
        document.querySelectorAll('a[download], a[href*=".pdf"], button[onclick*="download"]').forEach(function(el) {
            links.push({
                tag: el.tagName,
                href: el.href || '',
                text: el.textContent.trim().substring(0, 50),
                onclick: el.getAttribute('onclick') || ''
            });
        });
        return links;
    """)
    print(f"  다운로드 링크 ({len(download_links)}개):")
    for lk in download_links[:5]:
        print(f"    {lk}")
    
    # 5. 페이지의 PDF 뷰어 내용 확인
    iframes = d.execute_script("""
        return Array.from(document.querySelectorAll('iframe, embed, object')).map(el => ({
            tag: el.tagName,
            src: el.src || el.data || '',
            type: el.type || ''
        }));
    """)
    print(f"  iframes/embeds ({len(iframes)}개):")
    for iframe in iframes:
        print(f"    {iframe}")
    
    # 6. JavaScript 변수에서 PDF URL 찾기
    js_vars = d.execute_script("""
        var results = [];
        // 전역 변수에서 찾기
        for (var key in window) {
            try {
                var val = window[key];
                if (typeof val === 'string' && (val.includes('.pdf') || val.includes('sds') || val.includes('SDS'))) {
                    results.push(key + ': ' + val.substring(0, 100));
                }
            } catch(e) {}
        }
        return results.slice(0, 10);
    """)
    print(f"  JS 변수 PDF: {js_vars[:5]}")
    
    d.quit()
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("Thermo Fisher SDS 정확한 URL 테스트")
    print("=" * 60)
    
    try_tf_sds_download("L09319", save_dir)
    try_tf_sds_download("L16400", save_dir)
    
    print("\n" + "=" * 60)
    print("Sigma-Aldrich 심층 탐색")
    print("=" * 60)
    
    try_aldrich_sds_via_selenium("803200", "aldrich", save_dir)
