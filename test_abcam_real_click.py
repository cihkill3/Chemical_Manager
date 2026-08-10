from seleniumbase import Driver
import time
import os
import glob

def test_download_sds():
    dl_dir = os.path.abspath('abcam_dl_final')
    os.makedirs(dl_dir, exist_ok=True)
    
    # Empty dir
    for f in glob.glob(os.path.join(dl_dir, '*')):
        os.remove(f)
        
    driver = None
    try:
        driver = Driver(uc=True, headless=True)
        
        # Enable CDP for download behavior
        try:
            driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": dl_dir
            })
        except:
            pass
            
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Hiding cookie banner...")
        driver.execute_script("document.querySelectorAll('.onetrust-pc-dark-filter, #onetrust-banner-sdk').forEach(el => el.remove());")
        time.sleep(1)
        
        print("Real clicking SDS button...")
        driver.click("div.sds-button")
        
        print("Waiting for download to finish...")
        file_path = None
        for _ in range(30):
            time.sleep(1)
            # Check for non-crdownload files
            files = [f for f in glob.glob(os.path.join(dl_dir, '*.*')) if not f.endswith('.crdownload')]
            if files:
                file_path = files[0]
                print("Download finished:", file_path)
                break
        else:
            print("Download did not finish in time. Files in dir:")
            print(glob.glob(os.path.join(dl_dir, '*.*')))
            
    except Exception as e:
        print("Error:", e)
    finally:
        if driver:
            driver.quit()

if __name__ == '__main__':
    test_download_sds()
