"""
Aldrich: Selenium이 PDF를 로드한 상태에서 CDP로 binary 내용 가져오기
TCI: 실제 SDS PDF URL 네트워크에서 포착
"""
import os
import time
import re
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
        return f"[PDF 오류: {e}]"


def download_aldrich_sds_async_xhr(product_number: str, brand: str, save_dir: str) -> str:
    """Selenium + 비동기 XHR로 Aldrich PDF 다운로드"""
    from seleniumbase import Driver
    import glob
    
    abs_save_dir = os.path.abspath(save_dir)
    d = Driver(uc=True, headless=True)
    
    # 먼저 메인 사이트 방문해 쿠키 획득
    d.get("https://www.sigmaaldrich.com/KR/ko")
    time.sleep(3)
    
    # PDF URL
    pdf_url = f"https://www.sigmaaldrich.com/KR/en/sds/{brand}/{product_number}"
    
    # 비동기 fetch로 PDF 바이너리 가져오기
    print(f"  [{product_number}] 비동기 XHR로 PDF 획득 시도")
    try:
        result = d.execute_script(f"""
            return new Promise((resolve, reject) => {{
                fetch('{pdf_url}', {{
                    headers: {{
                        'Accept': 'application/pdf,*/*',
                        'Referer': 'https://www.sigmaaldrich.com/KR/ko'
                    }},
                    credentials: 'include'
                }})
                .then(r => {{
                    console.log('Status:', r.status, 'CT:', r.headers.get('content-type'));
                    if (!r.ok) return resolve({{error: 'HTTP ' + r.status}});
                    return r.arrayBuffer();
                }})
                .then(buf => {{
                    var arr = Array.from(new Uint8Array(buf));
                    resolve({{data: arr, length: arr.length}});
                }})
                .catch(e => resolve({{error: e.toString()}}));
            }});
        """)
        
        print(f"  결과: {str(result)[:200] if result else 'None'}")
        
        if result and 'data' in result:
            content = bytes(result['data'])
            print(f"  PDF 크기: {len(content)}")
            text = read_pdf_text(content)
            print(f"  PDF 내용 (첫 200자): {text[:200]}")
            fname = os.path.join(abs_save_dir, f"SDS_Aldrich_{brand}_{product_number}.pdf")
            with open(fname, 'wb') as f:
                f.write(content)
            print(f"  저장: {fname}")
            d.quit()
            return fname
        elif result and 'error' in result:
            print(f"  fetch 오류: {result['error']}")
    except Exception as e:
        print(f"  비동기 XHR 오류: {e}")
    
    d.quit()
    return ""


def download_tci_sds_network_capture(product_number: str, save_dir: str) -> str:
    """TCI SDS - Selenium 네트워크 감시로 실제 PDF URL 포착"""
    from seleniumbase import Driver
    import glob
    
    abs_save_dir = os.path.abspath(save_dir)
    d = Driver(uc=True, headless=True)
    
    # 페이지 로드 전에 fetch 인터셉터 설치
    d.execute_cdp_cmd("Network.enable", {})
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    
    # fetch/XHR 인터셉터 설치
    intercept_code = """
    window._capturedPdfUrls = [];
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
        var url = typeof input === 'string' ? input : (input.url || '');
        if (url.includes('.pdf') || url.includes('sds') || url.includes('document') || url.includes('download')) {
            window._capturedPdfUrls.push({type: 'fetch', url: url, method: (init && init.method) || 'GET'});
        }
        return origFetch.apply(this, arguments);
    };
    """
    
    d.get(url)
    d.execute_script(intercept_code)
    time.sleep(5)
    
    # SDS 탭으로 이동하기 위해 앵커 로드
    d.get(url + "#docomentsSectionPDP")
    time.sleep(4)
    
    # 포착된 URL 확인
    captured = d.execute_script("return window._capturedPdfUrls || [];")
    print(f"  [{product_number}] 포착된 URL: {captured}")
    
    # SDS 링크를 직접 클릭
    try:
        # 문서 섹션 찾기
        doc_section = d.execute_script("""
            var el = document.getElementById('docomentsSectionPDP');
            if (!el) return null;
            var html = el.nextElementSibling ? el.nextElementSibling.outerHTML : el.parentElement.outerHTML;
            return html.slice(0, 3000);
        """)
        print(f"  문서 섹션 HTML: {doc_section[:500] if doc_section else 'None'}")
        
        # SDS 버튼 탐색
        sds_btns = d.execute_script("""
            var btns = [];
            // 모든 링크 탐색
            document.querySelectorAll('a, button, [role="button"]').forEach(function(el) {
                var text = el.textContent.trim().toLowerCase();
                if (text === 'sds' || text === 'sds download' || text.includes('safety data sheet')) {
                    var href = el.href || el.getAttribute('data-url') || el.getAttribute('onclick') || '';
                    btns.push({tag: el.tagName, text: text.slice(0,30), href: href.slice(0,100), id: el.id});
                }
            });
            return btns;
        """)
        print(f"  SDS 버튼들: {sds_btns}")
        
        # 버튼 클릭 시도
        if sds_btns:
            for btn_info in sds_btns:
                if btn_info.get('id'):
                    try:
                        d.execute_script(f"document.getElementById('{btn_info['id']}').click();")
                        time.sleep(3)
                    except Exception as e:
                        print(f"  ID 클릭 오류: {e}")
                
            # 클릭 후 포착된 URL 다시 확인
            captured2 = d.execute_script("return window._capturedPdfUrls || [];")
            print(f"  클릭 후 포착된 URL: {captured2}")
    
    except Exception as e:
        print(f"  탐색 오류: {e}")
    
    # 현재 페이지에서 PDF URL 직접 탐색 (Performance API)
    pdf_resources = d.execute_script("""
        return window.performance.getEntriesByType('resource')
            .filter(r => r.name.includes('.pdf') || r.name.includes('sds') || r.name.includes('document'))
            .map(r => r.name);
    """)
    print(f"  Performance API PDF: {pdf_resources}")
    
    d.quit()
    return ""


if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("Aldrich - 비동기 XHR PDF 다운로드")
    print("=" * 60)
    r = download_aldrich_sds_async_xhr("803200", "aldrich", save_dir)
    print(f"결과: {r or '실패'}\n")
    
    print("\n" + "=" * 60)
    print("TCI - 네트워크 감시 SDS URL 포착")
    print("=" * 60)
    r = download_tci_sds_network_capture("C0119", save_dir)
    print(f"결과: {r or '실패'}\n")
