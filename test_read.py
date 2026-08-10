import pandas as pd
df = pd.read_excel('result.xlsx')
for idx, row in df.iterrows():
    print(f"{row['제품번호']} -> {row['상세정보_링크']}")
