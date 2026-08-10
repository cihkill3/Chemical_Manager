"""Aldrich GraphQL API - 정확한 필드명으로 탐색"""
import requests
import json

product_number = '803200'
brand = 'aldrich'

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/126.0 Safari/537.36',
    'Referer': 'https://www.sigmaaldrich.com/KR/ko/sds/aldrich/803200',
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Origin': 'https://www.sigmaaldrich.com',
}

# 1. getVectorDocument 시도
queries = [
    {
        'name': 'getVectorDocument',
        'query': '{ getVectorDocument(productNumber: "803200", brand: "aldrich") { url fileName } }'
    },
    {
        'name': 'getVectorDocument with language',
        'query': '{ getVectorDocument(productNumber: "803200", brand: "aldrich", language: "KO", country: "KR") { url fileName } }'
    },
    # getDocument 관련 이름들
    {
        'name': 'getDocument',
        'query': '{ getDocument(productNumber: "803200", brand: "aldrich", type: "SDS") { url } }'
    },
    {
        'name': 'getPdf',
        'query': '{ getPdf(productNumber: "803200", brand: "aldrich") { url } }'
    },
    # product 기반
    {
        'name': 'product',
        'query': '{ product(productNumber: "803200", brand: "aldrich") { sdsUrl documents { type url } } }'
    },
    {
        'name': 'getProductDetail',
        'query': '{ getProductDetail(productNumber: "803200", brand: "aldrich") { sds { url } } }'
    },
]

print("GraphQL 필드 탐색:")
for q in queries:
    try:
        r = requests.post(
            'https://www.sigmaaldrich.com/api/sds/download',
            headers=headers,
            json={'query': q['query']},
            timeout=15,
            allow_redirects=True
        )
        msg = r.text[:300]
        print(f"\n[{q['name']}] Status: {r.status_code}")
        print(f"  {msg}")
    except Exception as e:
        print(f"\n[{q['name']}] 오류: {e}")

# 2. 실제 SDS 페이지를 로드해서 네트워크에서 GraphQL 요청 포착
print("\n\nSelenium으로 실제 GraphQL 요청 포착:")
from seleniumbase import Driver
import time

d = Driver(uc=True, headless=True)
d.execute_cdp_cmd("Network.enable", {})

requests_log = []

url = "https://www.sigmaaldrich.com/KR/ko/sds/aldrich/803200"
d.get(url)
time.sleep(10)

# 실행된 모든 네트워크 요청 확인
performance_entries = d.execute_script("""
    return window.performance.getEntriesByType('resource')
        .filter(r => r.name.includes('api') || r.name.includes('graphql') || r.name.includes('sds'))
        .map(r => r.name)
        .slice(0, 20);
""")
print(f"API 관련 요청들: {performance_entries}")

# XHR/Fetch 로그 확인
xhr_log = d.execute_script("""
    return window._xhrLog || [];
""")
print(f"XHR 로그: {xhr_log}")

# localStorage와 sessionStorage에서 찾기
storage_data = d.execute_script("""
    var result = {};
    for (var i = 0; i < localStorage.length; i++) {
        var key = localStorage.key(i);
        if (key.toLowerCase().includes('sds') || key.toLowerCase().includes('pdf') || key.toLowerCase().includes('document')) {
            result[key] = localStorage.getItem(key).slice(0, 200);
        }
    }
    return result;
""")
print(f"LocalStorage PDF: {storage_data}")

# 페이지 소스에서 GraphQL 관련 정보 찾기
html = d.page_source
import re
gql_queries = re.findall(r'query\s+(\w+)\s*\(', html)
gql_mutations = re.findall(r'mutation\s+(\w+)\s*\(', html)
print(f"페이지 내 GraphQL queries: {set(gql_queries)}")
print(f"페이지 내 GraphQL mutations: {set(gql_mutations)}")

# API URL 패턴 찾기
api_urls = re.findall(r'["\'](https?://[^"\']*api[^"\']*)["\']', html)
sds_api_urls = [u for u in api_urls if 'sds' in u.lower() or 'document' in u.lower()]
print(f"API URL 패턴들: {list(set(sds_api_urls))[:10]}")

d.quit()
