import json
from bs4 import BeautifulSoup
with open('tf_protein_page.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')
links = []
for a in soup.find_all('a', href=True):
    href = a['href']
    if 'pdf' in href.lower() or 'sds' in href.lower() or 'documents' in href.lower():
        links.append({'text': a.text.strip(), 'href': href})
with open('tf_all_links.json', 'w', encoding='utf-8') as f:
    json.dump(links, f, indent=2)
