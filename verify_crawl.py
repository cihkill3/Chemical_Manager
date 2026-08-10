import os
import requests
import time
from playwright.sync_api import sync_playwright
from playwright_stealth import stealth

def clean_filename(filename):
    import re
    if not isinstance(filename, str): return "unknown"
    
    # 1.1 무시할 특수문자 제거
    cleaned = re.sub(r'[?\"|]', '', filename)
    
    # 1.2 맨 끝 특수문자 제거 (알파벳, 숫자, 괄호가 아닌 문자로 끝나는 경우)
    # 정규식으로 끝에 있는 특수문자들 제거
    cleaned = re.sub(r'[^a-zA-Z0-9)]+$', '', cleaned)
    
    # 1.3 나머지 윈도우 불가 문자를 언더바로 치환
    cleaned = re.sub(r'[\\/:*<>_]+', '_', cleaned)
    return cleaned.strip()

def test_crawling():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 봇 탐지 우회를 위해 User-Agent 설정
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        products = [
            ("Thermo Fisher", "326871000", "https://chemicals.thermofisher.kr/apac/product/326871000"),
            ("Sigma-Aldrich", "909270", "https://www.sigmaaldrich.com/KR/en/product/aldrich/909270"), # 영어 추출 위해 /en/ 사용
            ("TCI", "B0527", "https://www.tcichemicals.com/KR/en/p/B0527") # 영어 추출 위해 /en/ 사용
        ]
        
        for vendor, prod, url in products:
            print(f"--- Testing {vendor} ({prod}) ---")
            page = context.new_page()
            stealth(page)
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=20000)
                time.sleep(3) # 추가 렌더링 대기
                print(f"[{vendor}] Page Title: {page.title()}")
                
                # HTML 덤프(디버깅 용)
                html = page.content()
                if "Not Found" in html or "404" in page.title():
                    print("Page not found")
                else:
                    print("Page loaded successfully.")
                    
                    #간단한 시약명 추출 시도
                    name = "Unknown"
                    try:
                        if vendor == "TCI":
                            name = page.locator("h1.product-name").first.inner_text().strip()
                        else:
                            name = page.locator("h1").first.inner_text().strip()
                    except:
                        pass
                    print(f"[{vendor}] Extracted Name: {name}")
                    print(f"[{vendor}] Cleaned filename: {clean_filename(name)}")
                    
            except Exception as e:
                print(f"[{vendor}] Error: {e}")
            finally:
                page.close()
        
        browser.close()

if __name__ == "__main__":
    test_crawling()
