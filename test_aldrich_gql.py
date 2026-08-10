"""Aldrich GraphQL API 탐색"""
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

# GraphQL 쿼리 시도 1: 직접 필드명 추측
queries = [
    {
        'query': '{ getSds(productNumber: "803200", brand: "aldrich") { url } }'
    },
    {
        'query': '''
        query {
            getSdsDocument(productNumber: "803200", brand: "aldrich", language: "KO", country: "KR") {
                url
                fileName
            }
        }
        '''
    },
    {
        'query': '''
        query {
            getSdsForProduct(productNumber: "803200", brand: "aldrich") {
                downloadUrl
            }
        }
        '''
    },
]

for q in queries:
    try:
        r = requests.post(
            'https://www.sigmaaldrich.com/api/sds/download',
            headers=headers,
            json=q,
            timeout=15,
            allow_redirects=True
        )
        print(f'Status: {r.status_code}')
        print(f'Body: {r.text[:500]}')
        print()
    except Exception as e:
        print(f'오류: {e}')

# GraphQL introspection 시도 (스키마 확인)
introspection_query = {
    'query': '{ __schema { queryType { name fields { name } } } }'
}
try:
    r = requests.post(
        'https://www.sigmaaldrich.com/api/sds/download',
        headers=headers,
        json=introspection_query,
        timeout=15
    )
    print('Introspection 결과:')
    print(r.text[:1000])
except Exception as e:
    print(f'Introspection 오류: {e}')
