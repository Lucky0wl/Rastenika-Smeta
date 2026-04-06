import os
import win32com.client as win32
import pythoncom

def test_template_export():
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        tmpl_path = os.path.abspath('Шаблон.xlsx')
        pdf_path = os.path.abspath('Template_Sanity_Check.pdf')
        
        wb = excel.Workbooks.Open(tmpl_path)
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)
        excel.Quit()
        print(f"Exported sanity check to {pdf_path}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    test_template_export()
