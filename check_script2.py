import pandas as pd
df = pd.read_excel('test_quick_output.xlsx')
with open('result2.txt', 'w', encoding='utf-8') as f:
    for idx, row in df.iterrows():
        f.write(f"[{row.get('제조사')} - {row.get('제품번호')}]\n")
        f.write(f"보관온도: {row.get('보관온도')}\n")
        f.write("---\n")
