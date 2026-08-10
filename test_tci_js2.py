"""TCI documentSearch.js - 파일로 저장해서 확인"""
import time
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from seleniumbase import Driver

d = Driver(uc=True, headless=True)
d.get("https://www.tcichemicals.com/KR/en/p/C0119")
time.sleep(5)

js_content = d.execute_script("""
    return new Promise((resolve) => {
        fetch('https://www.tcichemicals.com/_ui/responsive/common/js/documentSearch.js?92a8199f54943294ae2c9912068c432d', {
            credentials: 'include'
        })
        .then(r => r.text())
        .then(text => resolve(text))
        .catch(e => resolve('error: ' + e));
    });
""")

d.quit()

if js_content:
    with open('tci_document_search.js', 'w', encoding='utf-8') as f:
        f.write(js_content)
    print(f"저장 완료: {len(js_content)} bytes")
    print("첫 200자:")
    print(js_content[:200])
else:
    print("내용 없음")
