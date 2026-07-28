import win32com.client
import os
import traceback

try:
    excel = win32com.client.DispatchEx('Excel.Application')
    wb = excel.Workbooks.Add()
    ws = wb.Sheets(1)
    ws.Range('A1').Value = 'Test'
    pdf_fwd = os.path.abspath('test_fwd.pdf').replace('\\', '/')
    pdf_back = os.path.abspath('test_back.pdf')
    
    print('Forward:', pdf_fwd)
    try:
        wb.ExportAsFixedFormat(0, pdf_fwd)
        print('Forward success')
    except Exception as e:
        print('Forward fail', e)
        
    print('Backward:', pdf_back)
    try:
        wb.ExportAsFixedFormat(0, pdf_back)
        print('Backward success')
    except Exception as e:
        print('Backward fail', e)
        
    wb.Close(False)
    excel.Quit()
except Exception as e:
    traceback.print_exc()
