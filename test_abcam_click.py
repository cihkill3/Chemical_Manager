from seleniumbase import Driver
import time
import os

def test_click_sds():
    options = [
        "--disable-popup-blocking",
    ]
    
    driver = Driver(uc=True, headless=True)
    
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        # Click the SDS button via Javascript to avoid overlay
        print("Clicking SDS button via JS...")
        # using JS click
        driver.execute_script("document.querySelector('div.sds-button').click();")
        
        time.sleep(5) # wait for action
        
        # Check window handles (did it open a new tab?)
        handles = driver.window_handles
        if len(handles) > 1:
            driver.switch_to.window(handles[-1])
            print("New window URL:", driver.current_url)
        else:
            print("Did not open new window. Current URL:", driver.current_url)
            
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    test_click_sds()
