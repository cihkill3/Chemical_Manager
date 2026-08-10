from seleniumbase import Driver
import time
import json

def test_fetch_intercept():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        driver.get(url)
        time.sleep(5)
        
        # Inject script to intercept fetch requests
        inject_js = """
        window.interceptedUrls = [];
        const originalFetch = window.fetch;
        window.fetch = async function() {
            window.interceptedUrls.push(arguments[0]);
            const response = await originalFetch.apply(this, arguments);
            const clone = response.clone();
            clone.text().then(body => {
                if(body.includes('wercs-api') || body.includes('SDS')) {
                    window.foundSdsResponse = body;
                }
            }).catch(e => {});
            return response;
        };
        """
        driver.execute_script(inject_js)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        time.sleep(5)
        
        urls = driver.execute_script("return window.interceptedUrls;")
        sds_resp = driver.execute_script("return window.foundSdsResponse;")
        
        print("Intercepted Fetch URLs:")
        for u in urls:
            if isinstance(u, str):
                print(u)
            else:
                print(u.get('url', str(u)))
                
        print("\nSDS Response Body:")
        print(sds_resp)
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_fetch_intercept()
