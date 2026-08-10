from bs4 import BeautifulSoup
with open('tf_protein_page.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')
for a in soup.find_all('a'):
    if 'sds' in a.text.lower() or 'sds' in str(a.get('href')).lower():
        print(a.attrs, a.text.strip())
