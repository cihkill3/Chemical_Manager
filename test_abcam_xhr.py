from seleniumbase import Driver
import time
import json

def test_xhr_intercept():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        driver.get(url)
        time.sleep(5)
        
        # Inject script to intercept XHR and fetch
        inject_js = """
        window.interceptedUrls = [];
        
        // Intercept fetch
        const originalFetch = window.fetch;
        window.fetch = async function() {
            window.interceptedUrls.push(arguments[0]);
            return originalFetch.apply(this, arguments);
        };
        
        // Intercept XHR
        const XHR = XMLHttpRequest.prototype;
        const open = XHR.open;
        XHR.open = function(method, url) {
            window.interceptedUrls.push(url);
            return open.apply(this, arguments);
        };
        """
        driver.execute_script(inject_js)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        time.sleep(5)
        
        urls = driver.execute_script("return window.interceptedUrls;")
        
        print("Intercepted URLs:")
        for u in urls:
            if isinstance(u, str):
                print(u)
            else:
                print(u.get('url', str(u)))
                
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_xhr_intercept()
