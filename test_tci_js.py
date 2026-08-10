"""TCI documentSearch.js 내용 분석 - Selenium 사용"""
import time
from seleniumbase import Driver

d = Driver(uc=True, headless=True)
d.get("https://www.tcichemicals.com/KR/en/p/C0119")
time.sleep(5)

# documentSearch.js 내용 가져오기
js_content = d.execute_script("""
    return new Promise((resolve) => {
        fetch('https://www.tcichemicals.com/_ui/responsive/common/js/documentSearch.js?92a8199f54943294ae2c9912068c432d', {
            credentials: 'include'
        })
        .then(r => r.text())
        .then(text => resolve(text))
        .catch(e => resolve('오류: ' + e));
    });
""")

print("documentSearch.js 내용:")
print(js_content if js_content else "내용 없음")

d.quit()
