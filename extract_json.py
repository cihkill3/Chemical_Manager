import re
import json
with open('tf_protein_page.html', 'r', encoding='utf-8') as f:
    html = f.read()
    
# Extract any JSON-like strings
scripts = re.findall(r'<script.*?>\s*({.*?})\s*</script>', html, flags=re.DOTALL)
for i, s in enumerate(scripts):
    if '300-13' in s or 'sds' in s.lower() or 'document' in s.lower():
        with open(f'script_{i}.json', 'w', encoding='utf-8') as out:
            out.write(s)
            
scripts2 = re.findall(r'window\.[a-zA-Z0-9_]+ *?= *?({.*?});', html, flags=re.DOTALL)
for i, s in enumerate(scripts2):
    if '300-13' in s or 'sds' in s.lower() or 'document' in s.lower():
        with open(f'window_{i}.json', 'w', encoding='utf-8') as out:
            out.write(s)
