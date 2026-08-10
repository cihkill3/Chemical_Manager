import requests
import json
import urllib.parse
from bs4 import BeautifulSoup
import time

def test_tf_sds(product_number):
    child_skus = f"{product_number}.MF,{product_number}.03,{product_number}.MD,{product_number}"
    url = f"https://chemicals.thermofisher.kr/apac/api/document/search/sds?childSkus={child_skus}&language=ko"
    
    print(f"Requesting API: {url}")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json"
    }
    
    response = requests.get(url, headers=headers)
    print(f"API Status: {response.status_code}")
    
    try:
        data = response.json()
        print("API Response JSON:", json.dumps(data, indent=2))
        
        # Extract direct URL if available
        pdf_url = data.get("data", "")
        
        print("Found PDF URL:", pdf_url)
                
        if not pdf_url:
            for item in data:
                if isinstance(item, dict) and item.get('url'):
                    pdf_url = item['url']
                    break
        
        # fallback to look for any url
        if not pdf_url:
            print("Could not find exact URL field, stringifying to search...")
            s = json.dumps(data)
            if 'http' in s:
                print("URL might be inside:", s)
        
        print("Found PDF URL:", pdf_url)
        
        if pdf_url:
            print(f"Fetching PDF from {pdf_url}")
            pdf_resp = requests.get(pdf_url, headers=headers, allow_redirects=True)
            print(f"PDF Fetch Status: {pdf_resp.status_code}")
            print(f"Content-Type: {pdf_resp.headers.get('Content-Type')}")
            print(f"Content-Disposition: {pdf_resp.headers.get('Content-Disposition')}")
            
            content = pdf_resp.content
            print(f"Received bytes: {len(content)}")
            
            if 'html' in pdf_resp.headers.get('Content-Type', '').lower():
                print("HTML content received. Might be a viewer page. First 500 chars:")
                print(content.decode('utf-8', errors='ignore')[:500])
            else:
                with open(f"TF_test_{product_number}.pdf", "wb") as f:
                    f.write(content)
                print(f"Saved to TF_test_{product_number}.pdf")
            
    except Exception as e:
        print("Error parsing JSON or fetching:", e)

if __name__ == "__main__":
    test_tf_sds("L09319")
