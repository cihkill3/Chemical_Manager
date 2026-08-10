from seleniumbase import Driver
import time
import os
import glob

def test_download_sds():
    driver = None
    try:
        # Default download dir might be the user's Downloads or seleniumbase default
        print("Starting Driver with uc=True and external_pdf=True...")
        driver = Driver(uc=True, headless=True, external_pdf=True)
        
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Clicking SDS button via JS...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        
        print("Waiting for something to happen...")
        time.sleep(10)
        
        print("Checking window handles...")
        print("Handles:", driver.window_handles)
        for handle in driver.window_handles:
            driver.switch_to.window(handle)
            print("URL:", driver.current_url)
            
    except Exception as e:
        print("Error:", e)
    finally:
        if driver:
            driver.quit()

if __name__ == '__main__':
    test_download_sds()
