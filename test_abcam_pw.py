import asyncio
from playwright.async_api import async_playwright

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        
        urls = []
        page.on("request", lambda request: urls.append(request.url))
        
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        await page.goto(url, wait_until="networkidle")
        
        print("Clicking SDS button...")
        try:
            # Bypass overlays by evaluating JS directly
            await page.evaluate("document.querySelector('div.sds-button').click();")
        except Exception as e:
            print("Click error:", e)
            
        await page.wait_for_timeout(5000)
        
        for u in set(urls):
            if 'wercs' in u.lower() or 'sds' in u.lower() or 'pdf' in u.lower() or 'document' in u.lower():
                print("FOUND URL:", u)
                
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
