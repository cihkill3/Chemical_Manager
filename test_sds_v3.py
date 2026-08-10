"""
SDS PDF 다운로드 - 최종 구현
- Thermo Fisher: SDS 전용 URL 찾기 (Product Spec 말고 SDS)
- Sigma-Aldrich: Selenium 다운로드 버튼 클릭 방식
- TCI: SDS 다운로드 재시도
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
# Thermo Fisher - 제품 페이지에서 SDS 링크 구체적 탐색
# ================================================================
def find_tf_sds_all_docs(product_number: str) -> dict:
    """
    TF 제품 페이지에서 모든 문서 타입과 URL을 탐색
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    urls = {}
    for suffix in ["", ".MD", ".06", ".14"]:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(4)
        
        html = d.page_source
        html_dec = unescape(html)
        
        # 모든 chem-specs-pdf URL 수집 (document type 포함)
        all_spec_urls = re.findall(
            r'https://assets\.thermofisher\.com/chem-specs-pdf/[^\s"\'<>]+',
            html_dec
        )
        
        if all_spec_urls:
            print(f"  [{product_number}{suffix}] 문서 URL들:")
            for u in all_spec_urls[:10]:
                print(f"    {u}")
                urls[u] = True
            break
        
        # 'SDS' 또는 'Safety Data' 포함 링크 찾기
        sds_links = d.execute_script("""
            var links = [];
            document.querySelectorAll('a').forEach(function(a) {
                var text = a.textContent.toLowerCase();
                var href = a.href || '';
                if (text.includes('sds') || text.includes('safety data') || 
                    href.includes('sds') || href.includes('safety')) {
                    links.push({text: a.textContent.trim().slice(0, 50), href: href.slice(0, 100)});
                }
            });
            return links;
        """)
        if sds_links:
            print(f"  [{product_number}{suffix}] SDS 링크: {sds_links[:5]}")
        
        # 'document' 텍스트가 있는 섹션 확인
        doc_sections = d.execute_script("""
            var titles = [];
            document.querySelectorAll('h2, h3, h4, button, .tab, [class*="document"]').forEach(function(el) {
                var t = el.textContent.trim();
                if (t.length < 100 && (t.toLowerCase().includes('document') || t.toLowerCase().includes('certificate') || t.toLowerCase().includes('sds'))) {
                    titles.push(t);
                }
            });
            return [...new Set(titles)].slice(0, 10);
        """)
        if doc_sections:
            print(f"  [{product_number}{suffix}] 문서 섹션: {doc_sections}")
    
    d.quit()
    return urls


def download_tf_sds_specific(product_number: str, save_dir: str) -> str:
    """
    TF SDS PDF 다운로드 (Spec과 SDS 구분)
    """
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    all_doc_urls = {}
    for suffix in ["", ".MD", ".06", ".14"]:
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
        d.get(url)
        time.sleep(4)
        
        html = d.page_source
        html_dec = unescape(html)
        
        # 모든 chem-specs-pdf URL
        matches = re.findall(r'https://assets\.thermofisher\.com/chem-specs-pdf/[^\s"\'<>&]+', html_dec)
        if matches:
            for m in matches:
                all_doc_urls[m] = True
            
            # SDS 관련 파라미터를 가진 URL 우선 선택
            # 보통 SDS는 documentType=SDS 또는 type=STIS 등
            sds_url = None
            for u in matches:
                ul = u.lower()
                if 'sds' in ul or 'stis' in ul or 'safety' in ul or 'msds' in ul:
                    sds_url = u
                    break
            
            if not sds_url and matches:
                # 첫 번째 URL로 시도하되, Product Spec이면 다른 URL 탐색
                sds_url = matches[0]
            
            if sds_url:
                print(f"  [{product_number}] 문서 URL: {sds_url}")
                break
    
    # JavaScript로 SDS 버튼 찾기
    try:
        js_result = d.execute_script("""
            var urls = [];
            document.querySelectorAll('a[href*="chem-specs-pdf"]').forEach(function(a) {
                urls.push({text: a.textContent.trim().slice(0, 60), href: a.href});
            });
            return urls;
        """)
        print(f"  JS에서 발견한 PDF 링크: {js_result[:5]}")
        for item in js_result:
            href = item.get('href', '')
            text = item.get('text', '').lower()
            if 'sds' in text or 'safety' in text or 'sds' in href.lower():
                all_doc_urls[href] = True
    except Exception as e:
        print(f"  JS 오류: {e}")
    
    d.quit()
    
    # 모든 URL 시도 - SDS 우선
    print(f"\n  총 {len(all_doc_urls)}개 문서 URL 발견:")
    for url in list(all_doc_urls.keys()):
        print(f"    {url}")
        try:
            r = requests.get(url, headers=HEADERS, timeout=20, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            if r.status_code == 200 and 'pdf' in ct.lower():
                text = read_pdf_text(r.content)
                print(f"    -> 다운로드 성공! 내용: {text[:200]}")
                # SDS인지 확인
                text_low = text.lower()
                if any(kw in text_low for kw in ['safety data', 'hazard', 'ghs', 'h302', 'h314']):
                    fname = os.path.join(save_dir, f"SDS_TF_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(r.content)
                    print(f"    -> SDS 확인! 저장: {fname}")
                    return fname
                elif any(kw in text_low for kw in ['specification', 'purity', 'cas number']):
                    fname = os.path.join(save_dir, f"SPEC_TF_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(r.content)
                    print(f"    -> 제품 사양서 저장: {fname}")
                    # SDS 계속 탐색
        except Exception as e:
            print(f"    오류: {e}")
    
    return ""


# ================================================================
# Sigma-Aldrich - Selenium 다운로드 버튼 클릭
# ================================================================
def download_aldrich_sds_click(product_number: str, brand: str, save_dir: str) -> str:
    """
    Selenium으로 Aldrich SDS 페이지를 열고 다운로드 버튼 클릭
    """
    from seleniumbase import Driver
    import glob
    
    abs_save_dir = os.path.abspath(save_dir)
    
    # 기존 PDF 파일 목록 (새 파일 감지용)
    existing_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf")))
    
    # Chrome 옵션으로 다운로드 디렉토리 설정
    d = Driver(uc=True, headless=False)
    
    # CDP로 다운로드 경로 설정
    d.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": abs_save_dir
    })
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/{brand}/{product_number}"
    print(f"  [{product_number}] 로딩: {url}")
    d.get(url)
    time.sleep(8)
    
    # 다운로드 버튼 탐색 및 클릭
    download_button_found = False
    
    # 1. 다운로드 버튼 텍스트로 찾기
    try:
        buttons = d.find_elements('css selector', 'button, a[href], [role="button"]')
        for btn in buttons:
            text = btn.text.strip().lower()
            if any(kw in text for kw in ['download', '다운로드', 'pdf', 'sds']):
                print(f"  버튼 발견: '{btn.text.strip()}'")
                try:
                    d.execute_script("arguments[0].click();", btn)
                    time.sleep(5)
                    download_button_found = True
                    print(f"  클릭 완료!")
                    break
                except Exception as e:
                    print(f"  클릭 실패: {e}")
    except Exception as e:
        print(f"  버튼 탐색 오류: {e}")
    
    # 2. 페이지 스크린샷 저장
    screenshot_path = os.path.join(abs_save_dir, f"aldrich_{product_number}_page.png")
    d.save_screenshot(screenshot_path)
    print(f"  스크린샷 저장: {screenshot_path}")
    
    # 새 PDF 파일 감지
    time.sleep(5)
    new_pdfs = set(glob.glob(os.path.join(abs_save_dir, "*.pdf"))) - existing_pdfs
    if new_pdfs:
        for pdf in new_pdfs:
            print(f"  새 PDF 다운로드됨: {pdf}")
            text = read_pdf_text(open(pdf, 'rb').read())
            print(f"  내용: {text[:200]}")
        d.quit()
        return list(new_pdfs)[0]
    
    # 페이지 현재 URL 및 소스에서 PDF 링크 재탐색
    print(f"  현재 URL: {d.current_url}")
    
    # PDF URL 직접 실행
    try:
        pdf_links = d.execute_script("""
            return Array.from(document.querySelectorAll('*'))
                .map(el => el.href || el.src || el.getAttribute('data-url') || el.getAttribute('data-href') || '')
                .filter(url => url && url.includes('.pdf'))
                .slice(0, 5);
        """)
        print(f"  발견된 PDF 링크: {pdf_links}")
    except Exception as e:
        print(f"  JS 오류: {e}")
    
    d.quit()
    return ""


# ================================================================
# TCI - 다른 접근법 (Selenium으로 직접)
# ================================================================
def download_tci_sds_selenium(product_number: str, save_dir: str) -> str:
    """
    TCI SDS 다운로드 - Selenium으로 직접 접속 후 PDF 버튼 클릭
    """
    from seleniumbase import Driver
    import glob
    
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
    
    # SDS 링크 찾기
    sds_links = d.execute_script("""
        var links = [];
        document.querySelectorAll('a').forEach(function(a) {
            var text = a.textContent.toLowerCase();
            var href = a.href || '';
            if (text.includes('sds') || text.includes('safety') || 
                href.includes('sds') || href.includes('.pdf')) {
                links.push({text: text.trim().slice(0, 40), href: href.slice(0, 100)});
            }
        });
        return links.slice(0, 10);
    """)
    print(f"  TCI SDS 링크: {sds_links}")
    
    # PDF 직접 URL 찾기
    pdf_urls = d.execute_script("""
        var urls = [];
        document.querySelectorAll('[href*=".pdf"], [src*=".pdf"], [data-url*=".pdf"]').forEach(function(el) {
            urls.push(el.href || el.src || el.getAttribute('data-url'));
        });
        return urls.slice(0, 5);
    """)
    print(f"  TCI PDF URLs: {pdf_urls}")
    
    # downloads 섹션 찾기
    doc_section = d.execute_script("""
        var section = document.querySelector('#docomentsSectionPDP, [id*="document"], [class*="document"]');
        return section ? section.innerHTML.slice(0, 500) : 'Not found';
    """)
    print(f"  TCI 문서 섹션: {doc_section[:200] if doc_section else 'None'}")
    
    d.quit()
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("Thermo Fisher - 모든 문서 URL 탐색 (SDS 구분)")
    print("=" * 60)
    
    print("\n[L09319]")
    find_tf_sds_all_docs("L09319")
    
    print("\n[L16400]")  
    find_tf_sds_all_docs("L16400")
    
    print("\n" + "=" * 60)
    print("Sigma-Aldrich - 다운로드 버튼 클릭")
    print("=" * 60)
    
    result = download_aldrich_sds_click("803200", "aldrich", save_dir)
    print(f"결과: {result or '실패'}")
    
    print("\n" + "=" * 60)
    print("TCI - Selenium 직접 탐색")
    print("=" * 60)
    
    result = download_tci_sds_selenium("C0119", save_dir)
    print(f"결과: {result or '실패'}")
