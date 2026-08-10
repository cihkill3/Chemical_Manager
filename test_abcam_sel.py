from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import json
import glob
import os

def test_download_and_logs():
    dl_dir = os.path.abspath('abcam_dl_test2')
    os.makedirs(dl_dir, exist_ok=True)
    for f in glob.glob(os.path.join(dl_dir, '*')):
        os.remove(f)

    options = Options()
    options.add_argument('--headless')
    options.add_experimental_option('prefs', {
        'download.default_directory': dl_dir,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'plugins.always_open_pdf_externally': True
    })
    options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
    
    # We will use webdriver.Chrome directly with seleniumbase's downloaded driver or system driver
    # Assuming system driver is available or we use selenium-manager
    driver = webdriver.Chrome(options=options)
    
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Clicking SDS button via JS...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        
        print("Waiting for download or logs...")
        time.sleep(10)
        
        files = glob.glob(os.path.join(dl_dir, '*'))
        print("Downloaded files:", files)
        
        print("Checking performance logs...")
        logs = driver.get_log('performance')
        for log in logs:
            try:
                msg = json.loads(log['message'])['message']
                if msg['method'] == 'Network.requestWillBeSent':
                    req_url = msg['params']['request']['url']
                    if 'wercs' in req_url or 'sds' in req_url.lower():
                        print("FOUND SDS API:", req_url)
                if msg['method'] == 'Network.responseReceived':
                    res_url = msg['params']['response']['url']
                    if 'wercs' in res_url or 'sds' in res_url.lower():
                        print("FOUND SDS RESPONSE:", res_url)
            except Exception:
                pass
                
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_download_and_logs()
