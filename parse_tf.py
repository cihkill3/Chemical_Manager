import re
import urllib.parse
with open('tf_sds_search_global.html', 'r', encoding='utf-8') as f:
    text = f.read()

urls = re.findall(r'href=[\"\'](.*?\.pdf.*?)[\"\']', text, flags=re.IGNORECASE)
for u in list(set(urls)):
    print(u)
