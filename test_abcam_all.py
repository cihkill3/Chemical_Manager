from seleniumbase import Driver
import time

def test_abcam():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        driver.get(url)
        time.sleep(5)
        
        inject_js = """
        window.interceptedUrls = [];
        
        // Hook fetch
        const originalFetch = window.fetch;
        window.fetch = async function() {
            window.interceptedUrls.push({type: 'fetch', url: arguments[0]});
            return originalFetch.apply(this, arguments);
        };
        
        // Hook XHR
        const XHR = XMLHttpRequest.prototype;
        const open = XHR.open;
        XHR.open = function(method, url) {
            window.interceptedUrls.push({type: 'xhr', url: url});
            return open.apply(this, arguments);
        };
        
        // Hook window.open
        const originalOpen = window.open;
        window.open = function(url, target, features) {
            window.interceptedUrls.push({type: 'window.open', url: url});
            return null; // prevent popup
        };
        
        // Hook form submission
        const originalSubmit = HTMLFormElement.prototype.submit;
        HTMLFormElement.prototype.submit = function() {
            window.interceptedUrls.push({type: 'form.submit', url: this.action});
            originalSubmit.apply(this, arguments);
        };
        """
        driver.execute_script(inject_js)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        time.sleep(5)
        
        urls = driver.execute_script("return window.interceptedUrls;")
        print("Intercepted Events:")
        for u in urls:
            print(u)
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_abcam()
