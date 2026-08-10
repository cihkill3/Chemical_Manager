"""
SDS PDF 다운로드 - 심층 탐색
1. Thermo Fisher: 올바른 SKU 패턴 찾기
2. Aldrich: Selenium + CDP 네트워크 인터셉트로 PDF URL 포착
"""
import sys
import os
import time
import re
import requests

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

# =====================================================
# 1. Thermo Fisher: Selenium으로 실제 SKU/API URL 찾기
# =====================================================
def find_tf_sds_url(product_number: str):
    """
    chemicals.thermofisher.kr 제품 페이지를 열고,
    SDS 다운로드 버튼 클릭 시 발생하는 요청 URL을 CDP로 캡처
    """
    from seleniumbase import Driver
    
    found_pdf_urls = []
    
    d = Driver(uc=True, headless=True)
    
    # CDP 네트워크 모니터링 활성화
    d.execute_cdp_cmd("Network.enable", {})
    
    url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}"
    d.get(url)
    time.sleep(4)
    
    html = d.page_source
    
    # SDS 관련 버튼/링크 텍스트 찾기
    sds_texts = re.findall(r'(?:href|src|url)[=\s]*["\']([^"\']*(?:sds|SDS|pdf|PDF|document|DirectWebViewer)[^"\']*)["\']', html)
    print(f"  SDS 관련 URL 패턴 ({len(sds_texts)}개):")
    for u in sds_texts[:10]:
        print(f"    {u}")
    
    # data-* 속성에서 SKU 또는 문서 ID 찾기
    sku_matches = re.findall(r'(?:sku|SKU|product-number|productNumber)["\s:=]+([A-Za-z0-9\-]+)', html)
    print(f"\n  SKU/제품코드 패턴: {list(set(sku_matches))[:5]}")
    
    # assets.thermofisher.com 관련 링크 모두 찾기
    asset_links = re.findall(r'(https://assets\.thermofisher\.com[^"\'<\s]+)', html)
    print(f"\n  assets.thermofisher.com 링크:")
    for lk in asset_links[:5]:
        print(f"    {lk}")
    
    d.quit()
    return sds_texts, sku_matches


# =====================================================
# 2. Aldrich: Selenium + CDP 네트워크 감시
# =====================================================
def capture_aldrich_network(product_number: str):
    """
    Selenium CDP를 통해 sigmaaldrich.com SDS 페이지에서
    발생하는 모든 네트워크 요청을 감시하여 PDF URL 포착
    """
    from seleniumbase import Driver
    import json
    
    network_requests = []
    
    d = Driver(uc=True, headless=True)
    
    # JS 로깅으로 XMLHttpRequest / fetch 가로채기
    d.execute_cdp_cmd("Network.enable", {})
    
    url = f"https://www.sigmaaldrich.com/KR/ko/sds/aldrich/{product_number}"
    print(f"  Loading: {url}")
    d.get(url)
    time.sleep(8)
    
    # performance API로 모든 네트워크 요청 보기
    all_resources = d.execute_script("""
        return window.performance.getEntriesByType('resource').map(r => ({
            name: r.name,
            type: r.initiatorType,
            duration: r.duration
        }));
    """)
    
    print(f"\n  전체 리소스 요청 ({len(all_resources)}개)")
    pdf_related = []
    for res in all_resources:
        name = res.get('name', '')
        if any(kw in name.lower() for kw in ['pdf', 'sds', 'document', 'safety', 'mds', 'material']):
            print(f"    [관련] {name[:100]}")
            pdf_related.append(name)
        elif 'merck' in name.lower() or 'sigma' in name.lower():
            print(f"    [merck] {name[:100]}")
    
    print(f"\n  현재 페이지 URL: {d.current_url}")
    
    # iframe src 찾기
    iframes = d.execute_script("""
        return Array.from(document.querySelectorAll('iframe')).map(f => ({
            src: f.src,
            id: f.id,
            class: f.className
        }));
    """)
    print(f"\n  iframe ({len(iframes)}개):")
    for iframe in iframes:
        print(f"    src={iframe.get('src','')[:80]}, id={iframe.get('id','')}, class={iframe.get('class','')}")
    
    # embed/object 태그
    embeds = d.execute_script("""
        return Array.from(document.querySelectorAll('embed, object')).map(e => ({
            src: e.src || e.data,
            type: e.type
        }));
    """)
    print(f"\n  embed/object ({len(embeds)}개):")
    for embed in embeds:
        print(f"    src={embed.get('src','')[:80]}, type={embed.get('type','')}")
    
    # 페이지 소스에서 PDF URL 패턴 찾기
    html = d.page_source
    pdf_patterns = re.findall(r'["\']([^"\']*\.pdf[^"\']*)["\']', html)
    print(f"\n  HTML 내 .pdf 패턴 ({len(pdf_patterns)}개):")
    for p in pdf_patterns[:10]:
        print(f"    {p[:100]}")
    
    # JavaScript 내 API URL 패턴 찾기
    api_patterns = re.findall(r'["\']([^"\']*(?:api|getDocument|sdsData|downloadSDS|fetchSDS)[^"\']*)["\']', html, re.IGNORECASE)
    print(f"\n  API 관련 패턴 ({len(api_patterns)}개):")
    for p in api_patterns[:10]:
        print(f"    {p[:100]}")
    
    d.quit()
    return pdf_related


# =====================================================
# 3. requests로 Aldrich API 엔드포인트 탐색
# =====================================================
def explore_aldrich_api(product_number: str):
    """
    Merck/Sigma-Aldrich의 공개 또는 반공개 SDS API 탐색
    """
    import requests
    
    # 알려진 Merck API 엔드포인트 시도
    endpoints = [
        # Merck eSDSDirect API
        f"https://www.merck.com/eSDSDirect/esdsSearch.do?serialNumber=1&term=803200",
        # Sigma REST API  
        f"https://www.sigmaaldrich.com/store/controller/Detail;jsessionid=?tab=msds&catalogNumber={product_number}&brand=ALDRICH&symbol=KR&language=ko&region=KR",
        # 공개 SDS 다운로드 API (Merck 계열)
        f"https://www.emdmillipore.com/msds.do?id={product_number}",
        # 직접 CDN 패턴
        f"https://www.sigmaaldrich.com/deepweb/assets/sigmaaldrich/product/documents/{product_number[:-3]}000/{product_number}_SDS_EN_EU.pdf",
    ]
    
    session = requests.Session()
    session.headers.update(HEADERS)
    
    for url in endpoints:
        try:
            r = session.get(url, timeout=10, allow_redirects=True)
            ct = r.headers.get('content-type', '')
            print(f"  URL: {url[:70]}")
            print(f"    Status: {r.status_code}, CT: {ct[:40]}, Size: {len(r.content)}")
            if r.status_code == 200 and 'pdf' in ct.lower():
                print("    >>> PDF 발견!")
                return r.content
        except Exception as e:
            print(f"    오류: {str(e)[:60]}")


if __name__ == "__main__":
    print("=" * 60)
    print("1. Thermo Fisher SDS URL 탐색")
    print("=" * 60)
    find_tf_sds_url("L09319")
    
    print("\n" + "=" * 60)
    print("2. Aldrich 네트워크 감시")
    print("=" * 60)
    capture_aldrich_network("803200")
    
    print("\n" + "=" * 60)
    print("3. Aldrich API 탐색")
    print("=" * 60)
    explore_aldrich_api("803200")
