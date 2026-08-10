"""
TCI SDS - 상세 페이지에서 SDS 버튼 클릭 후 네트워크 요청 캡처 (XHR + Fetch)
"""
import os
import time
from seleniumbase import Driver

def download_tci_sds_click(product_number: str):
    d = Driver(uc=True, headless=True)
    
    # XHR, Fetch 가로채기 코드
    intercept_code = """
    window._capturedRequests = [];
    
    // Fetch 인터셉터
    const origFetch = window.fetch;
    window.fetch = function(input, init) {
        var url = typeof input === 'string' ? input : (input.url || '');
        var method = (init && init.method) ? init.method : 'GET';
        var body = (init && init.body) ? init.body.toString() : '';
        window._capturedRequests.push({type: 'fetch', method: method, url: url, body: body.substring(0, 100)});
        
        return origFetch.apply(this, arguments).then(response => {
            var clone = response.clone();
            clone.blob().then(blob => {
                if (blob.type === 'application/pdf') {
                    window._capturedPdfBlob = blob;
                    window._capturedPdfUrl = url;
                }
            }).catch(e => {});
            return response;
        });
    };
    
    // XHR 인터셉터
    const origOpen = XMLHttpRequest.prototype.open;
    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url) {
        this._reqUrl = url;
        this._reqMethod = method;
        window._capturedRequests.push({type: 'xhr', method: method, url: String(url)});
        return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function(body) {
        this.addEventListener('load', function() {
            if (this.responseType === 'blob' && this.response && this.response.type === 'application/pdf') {
                window._capturedPdfBlob = this.response;
                window._capturedPdfUrl = this._reqUrl;
            }
        });
        return origSend.apply(this, arguments);
    };
    """
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    print(f"  [{product_number}] 페이지 로드 중...")
    d.get(url)
    time.sleep(2)
    
    # 인터셉터 설치
    d.execute_script(intercept_code)
    time.sleep(3)
    
    # SDS 버튼 찾아서 클릭
    print(f"  [{product_number}] SDS 버튼 클릭 시도...")
    clicked = d.execute_script("""
        var links = document.querySelectorAll('a, button');
        for (var l of links) {
            var text = l.textContent.trim().toLowerCase();
            if (text === 'sds' || text === 'safety data sheet') {
                l.click();
                return true;
            }
        }
        return false;
    """)
    
    print(f"  버튼 클릭 성공 여부: {clicked}")
    
    if clicked:
        # 파일이 다운로드되거나 요청이 완료될 때까지 대기
        for i in range(10):
            time.sleep(1)
            pdf_url = d.execute_script("return window._capturedPdfUrl;")
            if pdf_url:
                print(f"  PDF Blob 포착됨! URL: {pdf_url}")
                # Blob 데이터를 Base64로 가져오기
                b64 = d.execute_script("""
                    return new Promise((resolve) => {
                        var reader = new FileReader();
                        reader.readAsDataURL(window._capturedPdfBlob);
                        reader.onloadend = function() {
                            resolve(reader.result);
                        }
                    });
                """)
                if b64:
                    print(f"  Base64 데이터 크기: {len(b64)}")
                break
        
        # 포착된 모든 네트워크 요청 출력 (PDF 포착 실패 시 분석용)
        reqs = d.execute_script("return window._capturedRequests;")
        print(f"  포착된 요청 수: {len(reqs)}")
        for r in reqs:
            if 'pdf' in r['url'].lower() or 'sds' in r['url'].lower() or 'document' in r['url'].lower() or r['method'].upper() == 'POST':
                print(f"    - {r['method']} {r['url']}")
                if 'body' in r:
                    print(f"      Body: {r['body']}")
    
    d.quit()

if __name__ == "__main__":
    download_tci_sds_click("C0119")
