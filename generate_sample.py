import pandas as pd

data = {
    "제조사": ["thermo fisher", "aldrich", "TCI", "UnknownBrand"],
    "제품번호": ["326871000", "909270", "B0527", "12345"]
}

df = pd.DataFrame(data)
df.to_excel("sample_order.xlsx", index=False)
print("sample_order.xlsx generated.")
