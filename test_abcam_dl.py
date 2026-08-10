from seleniumbase import Driver
import time
import os
import glob

def test_download_sds():
    dl_dir = os.path.abspath('abcam_dl_test')
    os.makedirs(dl_dir, exist_ok=True)
    
    # Empty dir
    for f in glob.glob(os.path.join(dl_dir, '*')):
        os.remove(f)
        
    options = [
        f'--download.default_directory={dl_dir}',
    ]
    
    driver = Driver(uc=True, headless=True)
    
    # In newer selenium versions, we can set download dir via execute_cdp_cmd
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": dl_dir
        })
    except Exception as e:
        print("CDP cmd failed:", e)
        
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        
        print("Waiting for download...")
        # Wait up to 20 seconds for a pdf file to appear
        for _ in range(20):
            time.sleep(1)
            files = glob.glob(os.path.join(dl_dir, '*.pdf'))
            if files:
                print("Downloaded file:", files[0])
                break
        else:
            print("Download did not complete.")
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_download_sds()
