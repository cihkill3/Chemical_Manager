"""
SDS PDF 다운로드 방법 검증 스크립트
"""
import requests
import PyPDF2
import io
import re

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36'
}

def read_pdf_text(content: bytes) -> str:
    reader = PyPDF2.PdfReader(io.BytesIO(content))
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

# =====================================================
# 1. Thermo Fisher: DirectWebViewer API 방식
# =====================================================
def test_tf_sds(product_number: str, sku_prefix="ALFAAL"):
    """
    Thermo Fisher SDS PDF API 다운로드 테스트
    URL 패턴: https://assets.thermofisher.com/DirectWebViewer/private/results.aspx
    """
    sku = f"{sku_prefix}{product_number}"
    
    for lang_code, lang_label in [("EN", "영어"), ("KO", "한국어")]:
        for subformat in ["CGV4", "KOSD"]:
            url = (
                f"https://assets.thermofisher.com/DirectWebViewer/private/results.aspx"
                f"?page=NewSearch&LANGUAGE=d__{lang_code}&SUBFORMAT=d__{subformat}"
                f"&SKU={sku}&PLANT=d__ALF"
            )
            try:
                r = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)
                print(f"  [{lang_label}/{subformat}] Status: {r.status_code}, Content-Type: {r.headers.get('content-type','')[:40]}")
                if r.status_code == 200 and 'pdf' in r.headers.get('content-type','').lower():
                    print(f"  ✅ PDF 다운로드 성공! 크기: {len(r.content)} bytes")
                    text = read_pdf_text(r.content)
                    print(f"  PDF 첫 200자: {text[:200]}")
                    return r.content
            except Exception as e:
                print(f"  ❌ 오류: {e}")
    return None


# =====================================================
# 2. Sigma-Aldrich: API 방식 탐색
# =====================================================
def test_aldrich_sds(product_number: str, brand="aldrich"):
    """
    Sigma-Aldrich SDS API 탐색
    Merck는 SDS PDF를 제공하는 API를 갖고 있음
    """
    # 방법 1: 직접 PDF URL 패턴
    # sigmaaldrich.com/sds API
    urls_to_test = [
        f"https://www.sigmaaldrich.com/deepweb/assets/sigmaaldrich/product/documents/sds/{product_number}_SDS_EN_EU.pdf",
        f"https://www.sigmaaldrich.com/deepweb/assets/sigmaaldrich/product/documents/sds/{product_number}_SDS_EN_US.pdf",
        f"https://www.sigmaaldrich.com/deepweb/assets/sigmaaldrich/product/documents/sds/{product_number}_SDS_KO_KR.pdf",
        # Merck SDS API 패턴
        f"https://www.sigmaaldrich.com/etc.clientlibs/sigma/clientlibs/clientlib-site/resources/assets/pdf/{product_number}_sds.pdf",
    ]
    
    for url in urls_to_test:
        try:
            r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            print(f"  URL: {url[:70]}")
            print(f"    Status: {r.status_code}, Content-Type: {r.headers.get('content-type','')[:40]}")
            if r.status_code == 200 and 'pdf' in r.headers.get('content-type','').lower():
                print(f"    ✅ PDF 다운로드 성공!")
                text = read_pdf_text(r.content)
                print(f"    PDF 첫 200자: {text[:200]}")
                return r.content
        except Exception as e:
            print(f"    ❌ 오류: {e}")
    
    # 방법 2: Selenium으로 페이지 로드 후 PDF iframe src 추출
    print("  → API 직접 접근 실패, Selenium으로 내부 PDF URL 탐색...")
    return None


# =====================================================
# 3. Selenium을 이용한 네트워크 탭 PDF URL 포착
# =====================================================
def test_selenium_network_intercept(product_number: str, manufacturer: str):
    """
    Selenium CDP를 통해 네트워크 요청을 가로채서 PDF URL 탐색
    """
    from seleniumbase import Driver
    from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
    
    pdf_urls = []
    
    d = Driver(uc=True, headless=True)
    
    # CDP Network 이벤트 활성화
    d.execute_cdp_cmd("Network.enable", {})
    
    # 응답 이벤트 리스너 (CDP)
    if manufacturer == "aldrich":
        target_url = f"https://www.sigmaaldrich.com/KR/ko/sds/aldrich/{product_number}"
    elif manufacturer == "thermofisher":
        target_url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}"
    
    d.get(target_url)
    import time; time.sleep(8)
    
    # 페이지 내의 모든 링크/리소스 URL 점검
    all_links = d.execute_script("""
        let resources = window.performance.getEntriesByType('resource');
        return resources.map(r => r.name);
    """)
    
    print(f"\n  [{manufacturer}] 페이지 로드된 리소스 ({len(all_links)}개):")
    for url in all_links:
        if 'pdf' in url.lower() or 'sds' in url.lower() or 'document' in url.lower():
            print(f"    🔗 {url[:100]}")
            pdf_urls.append(url)
    
    d.quit()
    return pdf_urls


# ============================================================
# 메인 실행
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("Thermo Fisher SDS 다운로드 테스트")
    print("=" * 60)
    
    print("\n[L09319 - Fluorescein isothiocyanate]")
    pdf = test_tf_sds("L09319")
    if pdf:
        with open("test_sds_L09319.pdf", "wb") as f:
            f.write(pdf)
        print("  → test_sds_L09319.pdf 저장 완료")
    
    print("\n[L16400 - Acryloyloxy propyltrimethoxysilane]")
    pdf = test_tf_sds("L16400")
    if pdf:
        with open("test_sds_L16400.pdf", "wb") as f:
            f.write(pdf)
        print("  → test_sds_L16400.pdf 저장 완료")
    
    print("\n" + "=" * 60)
    print("Sigma-Aldrich SDS 다운로드 테스트")
    print("=" * 60)
    
    print("\n[803200 - DTSSP]")
    pdf = test_aldrich_sds("803200")

    print("\n[909270-10G - Acrylic anhydride]")
    pdf = test_aldrich_sds("909270-10G")
    
    print("\n" + "=" * 60)
    print("Selenium 네트워크 인터셉트 테스트 (Sigma-Aldrich)")
    print("=" * 60)
    
    print("\n[803200 - DTSSP (Aldrich)]")
    urls = test_selenium_network_intercept("803200", "aldrich")
    if not urls:
        print("  PDF/SDS 관련 리소스를 찾지 못했습니다.")
