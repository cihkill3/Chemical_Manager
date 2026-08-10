"""
TCI SDS - API를 직접 호출하여 다운로드 시도 (수정)
ACC.config.encodedContextPath 동적 획득
"""
import os
import time
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


def download_tci_sds_api(product_number: str, save_dir: str) -> str:
    from seleniumbase import Driver
    
    abs_save_dir = os.path.abspath(save_dir)
    d = Driver(uc=True, headless=True)
    
    # 1. 메인 페이지 로드하여 쿠키 및 컨텍스트 획득
    print(f"  [{product_number}] 메인 페이지 로드 중...")
    d.get(f"https://www.tcichemicals.com/KR/en/p/{product_number}")
    time.sleep(4)
    
    # 2. ACC.config.encodedContextPath 획득 및 XHR POST 요청 실행
    print(f"  [{product_number}] API 호출 중...")
    try:
        result = d.execute_script(f"""
            return new Promise((resolve) => {{
                var contextPath = (typeof ACC !== 'undefined' && ACC.config && ACC.config.encodedContextPath) ? ACC.config.encodedContextPath : '/KR/en';
                var url = contextPath + "/documentSearch/productSDSSearchDoc";
                console.log("Using URL:", url);
                
                var formData = new URLSearchParams();
                formData.append('brandCode', 'TCI');
                formData.append('productCode', '{product_number}');
                formData.append('langSelector', 'KO'); // SDS 한국어 시도, 없으면 EN
                formData.append('selectedCountry', 'KR');
                
                fetch(url, {{
                    method: 'POST',
                    headers: {{
                        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                        'X-Requested-With': 'XMLHttpRequest',
                        'Accept': '*/*'
                    }},
                    body: formData.toString()
                }})
                .then(r => {{
                    console.log('Status:', r.status);
                    var disp = r.headers.get('Content-Disposition') || '';
                    if (!r.ok) return resolve({{error: 'HTTP ' + r.status, url: url}});
                    
                    return r.arrayBuffer().then(buf => ({{
                        status: r.status,
                        disposition: disp,
                        url: url,
                        data: Array.from(new Uint8Array(buf))
                    }}));
                }})
                .then(res => resolve(res))
                .catch(e => resolve({{error: String(e), url: url}}));
            }});
        """)
        
        if not result:
            print("  결과가 없습니다.")
        elif 'error' in result:
            print(f"  오류 발생: {result['error']} (URL: {result.get('url')})")
        else:
            print(f"  URL: {result.get('url')}")
            print(f"  Status: {result.get('status')}")
            print(f"  Content-Disposition: {result.get('disposition')}")
            
            if 'data' in result and result['data']:
                content = bytes(result['data'])
                print(f"  크기: {len(content)} bytes")
                
                # HTML 에러 페이지인지 확인 (PDF가 아닌 경우)
                if len(content) < 1000 and b'<html' in content[:500].lower():
                    print("  서버에서 HTML을 반환했습니다 (PDF 아님).")
                else:
                    text = read_pdf_text(content)
                    print(f"  PDF 내용 (첫 200자): {text[:200]}")
                    
                    fname = os.path.join(abs_save_dir, f"SDS_TCI_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(content)
                    print(f"  저장 완료: {fname}")
                    d.quit()
                    return fname
            else:
                print("  데이터가 없습니다.")
    except Exception as e:
        print(f"  Exception: {e}")
        
    d.quit()
    return ""

if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    save_dir = "sds_downloads"
    
    print("=" * 60)
    print("TCI SDS - API (POST) 방식 테스트 v2")
    print("=" * 60)
    
    download_tci_sds_api("C0119", save_dir)
    print("-" * 30)
    download_tci_sds_api("T0751", save_dir)
