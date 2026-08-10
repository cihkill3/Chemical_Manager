from seleniumbase import Driver
import time
import json

def get_network_logs():
    options = [
        "--disable-popup-blocking",
    ]
    # To get performance logs, we need to pass capability
    from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
    caps = DesiredCapabilities.CHROME.copy()
    caps['goog:loggingPrefs'] = {'performance': 'ALL'}
    
    driver = Driver(uc=True, headless=True, desired_capabilities=caps)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)
        
        print("Clicking SDS button...")
        driver.execute_script("document.querySelector('div.sds-button').click();")
        time.sleep(5)
        
        # Fetch logs
        logs = driver.get_log('performance')
        urls = []
        for log in logs:
            try:
                msg = json.loads(log['message'])['message']
                if msg['method'] == 'Network.requestWillBeSent':
                    req_url = msg['params']['request']['url']
                    urls.append(req_url)
                    if 'wercs' in req_url or 'sds' in req_url.lower():
                        print("FOUND SDS API:", req_url)
                if msg['method'] == 'Network.responseReceived':
                    res_url = msg['params']['response']['url']
                    urls.append(res_url)
                    if 'wercs' in res_url or 'sds' in res_url.lower():
                        print("FOUND SDS RESPONSE:", res_url)
            except Exception:
                pass
                
        with open('abcam_urls.txt', 'w', encoding='utf-8') as f:
            for u in set(urls):
                f.write(u + '\n')
                
        print("Done capturing logs.")
    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    get_network_logs()
