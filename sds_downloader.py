import os
import io
import time
import requests
import PyPDF2
from seleniumbase import Driver
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def read_pdf_text(content: bytes) -> str:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page in reader.pages[:3]:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"PDF 읽기 오류: {e}"


def download_thermofisher_sds(product_number: str, save_dir: str) -> str:
    """Thermo Fisher SDS 다운로드 (API 기반 URL 획득)"""
    abs_save_dir = os.path.abspath(save_dir)
    os.makedirs(abs_save_dir, exist_ok=True)
    
    # 여러 Sku 조합을 API에 전달
    child_skus = f"{product_number}.MF,{product_number}.03,{product_number}.MD,{product_number}.06,{product_number}.14,{product_number}"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }

    # 한국어 시도, 없으면 영어 시도
    for lang in ['ko', 'en']:
        api_url = f"https://chemicals.thermofisher.kr/apac/api/document/search/sds?childSkus={child_skus}&language={lang}"
        try:
            logger.info(f"[Thermo Fisher] {product_number} API 요청 ({lang}): {api_url}")
            api_resp = requests.get(api_url, headers=headers, timeout=10)
            if api_resp.status_code == 200:
                data = api_resp.json()
                pdf_url = data.get("data")
                if pdf_url and isinstance(pdf_url, str) and pdf_url.startswith("http"):
                    logger.info(f"[Thermo Fisher] {product_number} SDS URL 획득: {pdf_url}")
                    
                    pdf_resp = requests.get(pdf_url, headers=headers, timeout=30, allow_redirects=True)
                    if pdf_resp.status_code == 200 and 'application/pdf' in pdf_resp.headers.get('content-type', '').lower():
                        content = pdf_resp.content
                        text = read_pdf_text(content)
                        if len(text) > 50:
                            logger.info(f"[Thermo Fisher] {product_number} SDS PDF 확인 완료 (크기: {len(content)} bytes)")
                            fname = os.path.join(abs_save_dir, f"SDS_ThermoFisher_{product_number}.pdf")
                            with open(fname, 'wb') as f:
                                f.write(content)
                            return fname
                        else:
                            logger.warning(f"[Thermo Fisher] {product_number} 다운로드 실패: 유효하지 않은 PDF 텍스트")
        except Exception as e:
            logger.error(f"[Thermo Fisher] {product_number} 언어 {lang} 오류: {e}")
            
    # Fallback for global products (e.g., PeproTech)
    base_sku = product_number
    if "-" in product_number:
        parts = product_number.split("-")
        if len(parts) >= 2:
            base_sku = f"{parts[0]}-{parts[1]}"
            
    fallback_urls = [
        f"https://documents.thermofisher.com/TFS-Assets/LSG/SDS/{base_sku}_MTR-NALT_EN.pdf",
        f"https://documents.thermofisher.com/TFS-Assets/LSG/SDS/{product_number}_MTR-NALT_EN.pdf"
    ]
    
    for f_url in fallback_urls:
        try:
            logger.info(f"[Thermo Fisher] {product_number} 글로벌 폴백(Fallback) 시도: {f_url}")
            pdf_resp = requests.get(f_url, headers=headers, timeout=10, allow_redirects=True)
            if pdf_resp.status_code == 200 and 'application/pdf' in pdf_resp.headers.get('content-type', '').lower():
                content = pdf_resp.content
                text = read_pdf_text(content)
                if len(text) > 50:
                    logger.info(f"[Thermo Fisher] {product_number} SDS PDF 폴백 획득 성공 (크기: {len(content)} bytes)")
                    fname = os.path.join(abs_save_dir, f"SDS_ThermoFisher_{product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(content)
                    return fname
        except Exception as e:
            logger.error(f"[Thermo Fisher] 폴백 에러: {e}")
            
    return ""


def download_aldrich_sds(product_number: str, brand: str, save_dir: str) -> str:
    """Sigma-Aldrich SDS 다운로드 (Selenium + 비동기 XHR Blob 캡처)"""
    abs_save_dir = os.path.abspath(save_dir)
    os.makedirs(abs_save_dir, exist_ok=True)
    
    # 용량이 함께 있는 경우 (예: 909270-10g) '-' 뒤의 부분을 제거
    clean_product_number = product_number.split('-')[0]
    
    d = Driver(uc=True, headless=True)
    try:
        # 1. 초기 쿠키 세팅을 위해 메인 페이지 방문
        logger.info(f"[Aldrich] {product_number} 쿠키 획득용 메인 페이지 접속")
        d.get("https://www.sigmaaldrich.com/KR/ko")
        time.sleep(3)
        
        # 2. 비동기 XHR로 PDF Blob 가져오기
        pdf_url = f"https://www.sigmaaldrich.com/KR/en/sds/{brand}/{clean_product_number}"
        logger.info(f"[Aldrich] {product_number} PDF Blob 비동기 요청 (Clean Number: {clean_product_number})")
        
        result = d.execute_script(f"""
            return new Promise((resolve) => {{
                fetch('{pdf_url}', {{
                    headers: {{ 'Accept': 'application/pdf,*/*', 'Referer': 'https://www.sigmaaldrich.com/KR/ko' }},
                    credentials: 'include'
                }})
                .then(r => {{
                    if (!r.ok) return resolve({{error: 'HTTP ' + r.status}});
                    return r.arrayBuffer();
                }})
                .then(buf => {{
                    var arr = Array.from(new Uint8Array(buf));
                    resolve({{data: arr}});
                }})
                .catch(e => resolve({{error: e.toString()}}));
            }});
        """)
        
        if result and 'data' in result:
            content = bytes(result['data'])
            if len(content) > 1000:
                text = read_pdf_text(content)
                logger.info(f"[Aldrich] {product_number} PDF 확인 완료 (크기: {len(content)} bytes)")
                fname = os.path.join(abs_save_dir, f"SDS_Aldrich_{product_number}.pdf")
                with open(fname, 'wb') as f:
                    f.write(content)
                return fname
            else:
                logger.warning(f"[Aldrich] {product_number} 다운로드 실패: 용량이 너무 작음 (HTML 에러 페이지 의심)")
        elif result and 'error' in result:
            logger.error(f"[Aldrich] {product_number} 다운로드 실패: {result['error']}")
            
    except Exception as e:
        logger.error(f"[Aldrich] 오류: {e}")
    finally:
        d.quit()
        
    return ""


def download_tci_sds(product_number: str, save_dir: str) -> str:
    """TCI SDS 다운로드 (브라우저 내 jQuery $.post + FileReader Base64 캡처)"""
    abs_save_dir = os.path.abspath(save_dir)
    os.makedirs(abs_save_dir, exist_ok=True)
    
    clean_product_number = product_number.strip().upper()
    d = Driver(uc=True, headless=True)
    try:
        url = f"https://www.tcichemicals.com/KR/en/p/{clean_product_number}"
        logger.info(f"[TCI] {clean_product_number} 접속 중...")
        d.get(url)
        time.sleep(3)
        
        script = """
        var prodNum = arguments[0];
        return new Promise((resolve) => {
            var url = (typeof ACC !== 'undefined' && ACC.config && ACC.config.encodedContextPath) 
                      ? ACC.config.encodedContextPath + "/documentSearch/productSDSSearchDoc" 
                      : "/KR/en/documentSearch/productSDSSearchDoc";
            var brandCode = $("#brandSelector").val() ? $("#brandSelector").val().toUpperCase() : "TCI";
            var productCode = $("#sdsProductCode").val() ? $("#sdsProductCode").val().toUpperCase() : prodNum;
            var langSelector = $("#langSelector").val() || "EN";
            var selectedCountry = $("#selectedCountry").val() || "KR";
            
            $.post({
                url: url,
                data: {
                    brandCode: brandCode,
                    productCode: productCode,
                    langSelector: langSelector,
                    selectedCountry: selectedCountry
                },
                xhrFields: {responseType: 'blob'}
            }).done(function (response, status, xhr) {
                var reader = new FileReader();
                reader.readAsDataURL(response);
                reader.onloadend = function() {
                    resolve({
                        status: 'success',
                        dataUrl: reader.result,
                        disposition: xhr.getResponseHeader('Content-Disposition')
                    });
                };
            }).fail(function (jqXHR, textStatus, errorThrown) {
                resolve({
                    status: 'fail',
                    httpStatus: jqXHR.status,
                    textStatus: textStatus,
                    errorThrown: errorThrown
                });
            });
        });
        """
        
        res = d.execute_script(script, clean_product_number)
        if res and res.get("status") == "success":
            data_url = res.get("dataUrl", "")
            if "," in data_url:
                import base64
                b64 = data_url.split(",")[1]
                content = base64.b64decode(b64)
                if len(content) > 1000:
                    text = read_pdf_text(content)
                    logger.info(f"[TCI] {clean_product_number} PDF 확인 완료 (크기: {len(content)} bytes)")
                    fname = os.path.join(abs_save_dir, f"SDS_TCI_{clean_product_number}.pdf")
                    with open(fname, 'wb') as f:
                        f.write(content)
                    return fname
                else:
                    logger.warning(f"[TCI] {clean_product_number} 다운로드 실패: 용량이 너무 작음")
        elif res and res.get("status") == "fail":
            logger.error(f"[TCI] {clean_product_number} 다운로드 실패: HTTP {res.get('httpStatus')}")
            
    except Exception as e:
        logger.error(f"[TCI] 오류: {e}")
    finally:
        d.quit()
        
    return ""

if __name__ == "__main__":
    os.makedirs("sds_downloads", exist_ok=True)
    print("--- Thermo Fisher Test ---")
    download_thermofisher_sds("L09319", "sds_downloads")
    
    print("\n--- Sigma-Aldrich Test ---")
    download_aldrich_sds("803200", "aldrich", "sds_downloads")
