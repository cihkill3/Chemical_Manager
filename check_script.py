import pandas as pd
df = pd.read_excel("test_quick_output.xlsx")
with open("result.txt", "w", encoding="utf-8") as f:
    for idx, row in df.iterrows():
        f.write(f"[{row.get('제조사')} - {row.get('제품번호')}]\n")
        f.write(f"신호어: {row.get('신호어')}\n")
        f.write(f"주요위험: {row.get('주요위험')}\n")
        f.write(f"상세위험: {row.get('상세 위험분류')}\n")
        f.write("---\n")
