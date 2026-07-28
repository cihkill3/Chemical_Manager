import win32com.client
import traceback
try:
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Add()
    ws = wb.Sheets(1)
    ws.Name = 'ChemicalList'
    ws.Range('G2').Value = 'Test'
    
    # Test FormatConditions
    try:
        rng = ws.Range('G2:G10')
        print('Testing FormatConditions with Operator=None...')
        fc = rng.FormatConditions.Add(2, None, '=AND(TRIM(G2)="", TRIM(D2)<>"")')
        print('Success!')
    except Exception as e:
        print('FormatConditions failed:', e)
        
    try:
        print('Testing FormatConditions with empty string Operator...')
        fc = rng.FormatConditions.Add(2, "", '=AND(TRIM(G2)="", TRIM(D2)<>"")')
        print('Success!')
    except Exception as e:
        print('FormatConditions Operator="" failed:', e)
        
    try:
        print('Testing FormatConditions with Type.Missing...')
        import pythoncom
        fc = rng.FormatConditions.Add(2, pythoncom.Missing, '=AND(TRIM(G2)="", TRIM(D2)<>"")')
        print('Success!')
    except Exception as e:
        print('FormatConditions pythoncom.Missing failed:', e)
        
    # Test Validation
    wb.Sheets.Add().Name = 'index'
    idx_ws = wb.Sheets('index')
    idx_ws.Range('A2:A10').Value = 'Room1'
    idx_ws.Range('B2:B10').Value = 'Temp1'
    idx_ws.Range('C2:C10').Value = 'Cab1'
    idx_ws.Range('D2:D10').Formula = '=A2&"_"&B2'
    
    try:
        print('Testing Validation 1...')
        ws.Range('I2:I10').Validation.Add(3, 1, 1, "='index'!$A$2:$A$1000")
        print('Success!')
    except Exception as e:
        print('Validation 1 failed:', e)
        
    # Remove dummy values
    ws.Range('I2:I10').ClearContents()
    ws.Range('J2:J10').ClearContents()
    try:
        print('Testing Validation formula with IFERROR...')
        f = "=OFFSET('index'!$C$1, IFERROR(MATCH(I2&\"_\"&J2, 'index'!$D:$D, 0)-1, 1), 0, MAX(1, COUNTIF('index'!$D:$D, I2&\"_\"&J2)), 1)"
        ws.Range('K2:K10').Validation.Add(3, 1, 1, f)
        print('Validation with IFERROR SUCCESS!')
    except Exception as e:
        print('Validation formula failed:', e)

    wb.Close(SaveChanges=False)
    excel.Quit()
except Exception as e:
    traceback.print_exc()
