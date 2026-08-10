from seleniumbase import Driver
import time

def test_abcam():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        # Override window.open and link clicks
        inject_js = """
        window.interceptedUrls = [];
        
        const originalOpen = window.open;
        window.open = function(url, target, features) {
            window.interceptedUrls.push(url);
            return null;
        };
        
        document.addEventListener('click', function(e) {
            let target = e.target.closest('a');
            if (target && target.href) {
                if(target.href.includes('wercs') || target.href.includes('sds')) {
                    window.interceptedUrls.push(target.href);
                    e.preventDefault();
                }
            }
        }, true);
        """
        driver.execute_script(inject_js)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        time.sleep(5)
        
        urls = driver.execute_script("return window.interceptedUrls;")
        print("Intercepted URLs:", urls)
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_abcam()
