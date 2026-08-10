"""
TCI SDS - documentSearch.js API를 통한 SDS PDF URL 추출
Selenium으로 TCI 페이지를 열고 documentSearch API 호출 분석
"""
import os
import time
import requests
import PyPDF2
import io

def read_pdf_text(content: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:3]:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"[오류: {e}]"


def find_tci_sds_url(product_number: str) -> str:
    """TCI 제품 페이지에서 실제 SDS PDF URL 추출"""
    from seleniumbase import Driver
    
    d = Driver(uc=True, headless=True)
    
    # fetch 인터셉터 설치
    intercept_code = """
    window._allRequests = [];
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
        var url = typeof input === 'string' ? input : (input.url || '');
        window._allRequests.push({type: 'fetch', url: url.slice(0, 200)});
        return origFetch.apply(this, arguments);
    };
    
    // XHR 인터셉터
    const origOpen = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function(method, url) {
        this._reqUrl = url;
        window._allRequests.push({type: 'xhr', url: String(url).slice(0, 200), method: method});
        return origOpen.apply(this, arguments);
    };
    """
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    d.get(url)
    
    # 인터셉터 설치
    d.execute_script(intercept_code)
    time.sleep(5)
    
    # SDS 탭 클릭 (JavaScript로)
    d.execute_script("""
        var links = document.querySelectorAll('a');
        for (var l of links) {
            if (l.textContent.trim().toLowerCase() === 'sds') {
                l.click();
                break;
            }
        }
    """)
    time.sleep(4)
    
    # 포착된 요청 확인
    requests_log = d.execute_script("return window._allRequests || [];")
    print(f"  [{product_number}] 모든 요청 ({len(requests_log)}개):")
    for req in requests_log:
        url_str = req.get('url', '')
        if any(kw in url_str.lower() for kw in ['sds', 'pdf', 'document', 'search', 'download']):
            print(f"    {req}")
    
    # Performance API로 document/sds 관련 요청
    perf = d.execute_script("""
        return window.performance.getEntriesByType('resource')
            .filter(r => r.name.includes('document') || r.name.includes('sds') || r.name.includes('pdf') || r.name.includes('search'))
            .map(r => ({url: r.name.slice(0, 200), type: r.initiatorType, size: r.encodedBodySize}));
    """)
    print(f"  Performance 리소스: {perf}")
    
    # documentSearch API 엔드포인트 찾기 (JS 소스에서)
    scripts = d.execute_script("""
        var inline_scripts = [];
        document.querySelectorAll('script:not([src])').forEach(function(s) {
            var text = s.textContent;
            if (text.includes('documentSearch') || text.includes('sds') || text.includes('SDS') || text.includes('pdf')) {
                inline_scripts.push(text.slice(0, 500));
            }
        });
        return inline_scripts.slice(0, 3);
    """)
    print(f"  인라인 스크립트: {scripts}")
    
    d.quit()
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    
    print("TCI SDS URL 탐색")
    print("=" * 60)
    find_tci_sds_url("C0119")
    print()
    find_tci_sds_url("T0751")
