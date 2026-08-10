import codecs, glob
for f in glob.glob('scrapers/*.py'):
    text = codecs.open(f, 'r', 'utf-8').read()
    text = text.replace('"N/A"', '"정보 없음"')
    codecs.open(f, 'w', 'utf-8').write(text)
