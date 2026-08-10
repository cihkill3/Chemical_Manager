from seleniumbase import Driver
import time
import os
import glob

def test_download_sds():
    dl_dir = os.path.abspath('abcam_dl_gui')
    os.makedirs(dl_dir, exist_ok=True)
    
    # Empty dir
    for f in glob.glob(os.path.join(dl_dir, '*')):
        os.remove(f)
        
    driver = None
    try:
        print("Starting Driver with uc=True and headless=False...")
        driver = Driver(uc=True, headless=False)
        
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
        time.sleep(10) # Wait longer for Cloudflare in GUI
        
        print("Hiding cookie banner...")
        driver.execute_script("document.querySelectorAll('.onetrust-pc-dark-filter, #onetrust-banner-sdk').forEach(el => el.remove());")
        time.sleep(2)
        
        print("Real clicking SDS button...")
        driver.click("div.sds-button")
        
        print("Waiting for download to finish...")
        for _ in range(20):
            time.sleep(1)
            files = [f for f in glob.glob(os.path.join(dl_dir, '*.*')) if not f.endswith('.crdownload')]
            if files:
                print("Download finished:", files[0])
                break
        else:
            print("Download did not finish. Checking default download folder...")
            def_dir = os.path.join(os.getcwd(), 'downloaded_files')
            print(glob.glob(os.path.join(def_dir, '*.*')))
            
    except Exception as e:
        print("Error:", e)
    finally:
        if driver:
            driver.quit()

if __name__ == '__main__':
    test_download_sds()
