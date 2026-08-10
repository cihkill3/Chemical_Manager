from seleniumbase import Driver
import time
import json

def get_resources():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        # Click SDS button just to trigger any fetch
        try:
            driver.execute_script("document.querySelector('div.sds-button').click();")
            time.sleep(2)
        except:
            pass
        
        # Dump performance entries
        log = driver.execute_script("return window.performance.getEntriesByType('resource').map(e => e.name);")
        with open('abcam_resources.json', 'w') as f:
            json.dump(log, f, indent=2)
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    get_resources()
