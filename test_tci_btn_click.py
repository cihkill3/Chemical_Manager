"""
TCI SDS - sdsSearchButton 클릭 시뮬레이션 및 다운로드 인터셉트
"""
import os
import time
import PyPDF2
import io
from seleniumbase import Driver

def read_pdf_text(content: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:3]:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"[오류: {e}]"

def download_tci_sds_btn_click(product_number: str, save_dir: str):
    abs_save_dir = os.path.abspath(save_dir)
    d = Driver(uc=True, headless=True)
    
    # XMLHttpRequest 몽키 패치로 Blob 가로채기
    intercept_code = """
    window._pdfBlob = null;
    const origOpen = XMLHttpRequest.prototype.open;
    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url) {
        this._reqUrl = url;
        return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function(body) {
        this.addEventListener('load', function() {
            if (this._reqUrl && this._reqUrl.includes('SDSSearchDoc')) {
                if (this.responseType === 'blob' && this.response && this.response.type === 'application/pdf') {
                    window._pdfBlob = this.response;
                }
            }
        });
        return origSend.apply(this, arguments);
    };
    """
    
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    print(f"  [{product_number}] 페이지 로드 중...")
    d.get(url)
    
    d.execute_script(intercept_code)
    time.sleep(3)
    
    print(f"  [{product_number}] 검색 버튼 클릭 시도...")
    
    # 영문(EN) 언어로 변경 후 sdsSearchButton 클릭
    clicked = d.execute_script("""
        var btn = document.getElementById('sdsSearchButton');
        var lang = document.getElementById('langSelector');
        if (lang) lang.value = 'EN';
        
        if (btn) {
            btn.click();
            return true;
        }
        return false;
    """)
    
    if not clicked:
        print("  버튼을 찾을 수 없습니다.")
        d.quit()
        return ""
        
    print("  버튼 클릭 성공! 응답 대기 중...")
    
    for i in range(15):
        time.sleep(1)
        # 윈도우 객체에 PDF Blob이 담겼는지 확인
        has_blob = d.execute_script("return window._pdfBlob !== null;")
        if has_blob:
            print("  PDF Blob 획득 성공!")
            # Blob 데이터를 Base64로 가져오기
            b64_data = d.execute_script("""
                return new Promise((resolve) => {
                    var reader = new FileReader();
                    reader.readAsDataURL(window._pdfBlob);
                    reader.onloadend = function() {
                        resolve(reader.result); // "data:application/pdf;base64,JVBERi..."
                    }
                });
            """)
            if b64_data and "," in b64_data:
                import base64
                b64 = b64_data.split(",")[1]
                content = base64.b64decode(b64)
                
                print(f"  PDF 크기: {len(content)} bytes")
                text = read_pdf_text(content)
                print(f"  PDF 내용 (첫 200자): {text[:200]}")
                
                fname = os.path.join(abs_save_dir, f"SDS_TCI_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(content)
                print(f"  저장 완료: {fname}")
                d.quit()
                return fname
    
    print("  타임아웃: PDF Blob을 가져오지 못했습니다.")
    d.quit()
    return ""

if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("TCI SDS - Search 버튼 시뮬레이션 방식 테스트")
    print("=" * 60)
    
    download_tci_sds_btn_click("C0119", save_dir)
    print("-" * 30)
    download_tci_sds_btn_click("T0751", save_dir)
