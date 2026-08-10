from seleniumbase import Driver
import time

def scrape_abcam_sds():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(10)  # Wait longer for JS
        
        # scroll down a bit
        driver.execute_script("window.scrollTo(0, 500);")
        time.sleep(2)
        
        html = driver.page_source
        with open('abcam_source.html', 'w', encoding='utf-8') as f:
            f.write(html)
        
        # take screenshot
        driver.save_screenshot('abcam_screenshot.png')
        
        print("Saved source and screenshot.")
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    scrape_abcam_sds()
