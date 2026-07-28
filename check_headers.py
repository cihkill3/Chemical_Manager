import openpyxl

file_path = r"c:\Users\Jeonghun Lee\.gemini\antigravity\scratch\chemical manager\시약관리.xlsm"
wb = openpyxl.load_workbook(file_path, data_only=True, keep_vba=True)
ws = wb["시약리스트"]

print("Headers in 시약리스트:")
for c in range(1, 20):
    val = ws.cell(row=1, column=c).value
    print(f"Col {c}: '{val}'")
