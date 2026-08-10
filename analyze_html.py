import os
from bs4 import BeautifulSoup

def analyze(name, html_file, out_file):
    out_file.write(f"\n======== {name} ========\n")
    with open(html_file, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f.read(), 'html.parser')
        
    out_file.write(f"Title: {soup.title.string if soup.title else 'No Title'}\n")
    
    keywords = ["cas", "storage", "temp", "hazard", "sensitive", "h2", "h3", "h4", "h319", "h225", "h314", "danger", "warning"]
    
    # Extract structural elements
    for el in soup.find_all(['tr', 'li', 'div', 'p', 'span']):
        text = el.get_text(separator=' ', strip=True).lower()
        if len(text) < 200 and any(kw in text for kw in keywords):
            t = text.replace('\n', ' ')
            if "sds" not in t and "privacy" not in t and "cookie" not in t:
                out_file.write(f"Found [{el.name}]: {t[:100]}\n")

if __name__ == "__main__":
    with open("analysis.txt", "w", encoding="utf-8") as out:
        analyze("Thermo", "html_dumps/thermo.html", out)
        analyze("TCI", "html_dumps/tci.html", out)
        analyze("Aldrich", "html_dumps/aldrich.html", out)
