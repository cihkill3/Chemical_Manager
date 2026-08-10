from seleniumbase import Driver
import time
import os
import glob
from selenium.webdriver.chrome.options import Options

def test_download_sds():
    dl_dir = os.path.abspath('abcam_dl_test3')
    os.makedirs(dl_dir, exist_ok=True)
    
    # Empty dir
    for f in glob.glob(os.path.join(dl_dir, '*')):
        os.remove(f)
        
    driver = None
    try:
        # Create standard options
        options = Options()
        options.add_experimental_option('prefs', {
            'download.default_directory': dl_dir,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'plugins.always_open_pdf_externally': True
        })
        
        print("Starting Driver with uc=True and options...")
        driver = Driver(uc=True, headless=True, options=options)
        
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Clicking SDS button via JS...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        
        print("Waiting for download...")
        for _ in range(20):
            time.sleep(1)
            files = glob.glob(os.path.join(dl_dir, '*.*'))
            if files:
                print("Downloaded files:", files)
                break
        else:
            print("Download did not complete. Checking window handles...")
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
