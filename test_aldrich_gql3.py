"""Aldrich GraphQL - getDocLib와 실제 SDS 페이지에서 사용되는 쿼리 추출"""
import requests
import json
import re
import time

product_number = '803200'
brand = 'aldrich'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36',
    'Referer': 'https://www.sigmaaldrich.com/KR/ko/sds/aldrich/803200',
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Origin': 'https://www.sigmaaldrich.com',
}

# getDocLib 탐색 (다양한 인수 조합)
print("=== getDocLib 탐색 ===")
doclib_queries = [
    '{ getDocLib(productNumber: "803200") { url } }',
    '{ getDocLib(sku: "803200") { url downloadUrl } }',
    '{ getDocLib(productNumber: "803200", brand: "aldrich") { url type } }',
    '{ getDocLib(id: "aldrich/803200") { url } }',
]
for q in doclib_queries:
    r = requests.post('https://www.sigmaaldrich.com/api/sds/download',
                      headers=headers, json={'query': q}, timeout=15)
    print(f"Query: {q[:60]}")
    print(f"  {r.status_code}: {r.text[:200]}\n")

# getProductDetail 탐색 (다른 인수 조합)
print("=== getProductDetail 탐색 ===")
detail_queries = [
    '{ getProductDetail(id: "803200") { name } }',
    '{ getProductDetail(catalogNumber: "803200") { name sds { url } } }',
    '{ getProductDetail(sku: "803200") { name } }',
]
for q in detail_queries:
    r = requests.post('https://www.sigmaaldrich.com/api/sds/download',
                      headers=headers, json={'query': q}, timeout=15)
    print(f"Query: {q[:60]}")
    print(f"  {r.status_code}: {r.text[:200]}\n")

# getVectorDocument 탐색 (다른 인수)
print("=== getVectorDocument 탐색 ===")
vector_queries = [
    '{ getVectorDocument(id: "803200") { url } }',
    '{ getVectorDocument(sku: "803200") { url } }',
    '{ getVectorDocument(catalogNumber: "803200") { url } }',
    '{ getVectorDocument(vectorId: "803200") { url sequence } }',
]
for q in vector_queries:
    r = requests.post('https://www.sigmaaldrich.com/api/sds/download',
                      headers=headers, json={'query': q}, timeout=15)
    print(f"Query: {q[:60]}")
    print(f"  {r.status_code}: {r.text[:200]}\n")

# 실제 SDS 페이지를 브라우저로 열고 Chrome DevTools로 실제 요청 포착
print("=== Selenium Chrome DevTools 실시간 XHR 포착 ===")
from seleniumbase import Driver

d = Driver(uc=True, headless=False)

# XHR 인터셉터 설치
d.execute_cdp_cmd("Network.enable", {})
d.execute_cdp_cmd("Page.enable", {})

# localStorage에 intercept 설정
intercept_script = """
    window._interceptedRequests = [];
    const originalFetch = window.fetch;
    window.fetch = function(...args) {
        var url = typeof args[0] === 'string' ? args[0] : args[0].url;
        if (url && (url.includes('api') || url.includes('sds') || url.includes('pdf'))) {
            var bodyClone = args[1] ? JSON.parse(JSON.stringify(args[1])) : {};
            window._interceptedRequests.push({url: url, method: bodyClone.method || 'GET', body: bodyClone.body});
        }
        return originalFetch.apply(this, args);
    };
    
    const originalXHR = XMLHttpRequest.prototype.open;
    XMLHttpRequest.prototype.open = function(method, url) {
        if (url && (url.includes('api') || url.includes('sds') || url.includes('pdf'))) {
            this._url = url;
            this._method = method;
        }
        return originalXHR.apply(this, arguments);
    };
    const originalSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.send = function(data) {
        if (this._url) {
            window._interceptedRequests.push({url: this._url, method: this._method, body: data});
        }
        return originalSend.apply(this, arguments);
    };
"""

url = "https://www.sigmaaldrich.com/KR/ko/sds/aldrich/803200"
d.execute_script("window._interceptedRequests = [];")
d.execute_script(intercept_script.replace('\n', ' '))
d.get(url)
time.sleep(12)

# 포착된 요청 확인
intercepted = d.execute_script("return window._interceptedRequests || [];")
print(f"포착된 API 요청 ({len(intercepted)}개):")
for req in intercepted[:20]:
    print(f"  {req}")

# Performance API로 리소스 목록
resources = d.execute_script("""
    return window.performance.getEntriesByType('resource')
        .map(r => ({name: r.name.slice(0, 100), type: r.initiatorType}))
        .filter(r => r.name.includes('api') || r.name.includes('sds') || r.name.includes('pdf') || r.name.includes('graphql'))
        .slice(0, 20);
""")
print(f"\nPerformance API 리소스:")
for r in resources:
    print(f"  {r}")

d.quit()
