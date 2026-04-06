"""
Convert both Шаблон.xlsx and latest estimate to images for visual comparison.
Uses xlwings + win32com to open in Excel and take screenshot.
"""
import os
import subprocess
import shutil

# Open the template in Excel and save as PDF for visual review
file1 = r'C:\Users\Server\Desktop\Rastenika Smeta\Шаблон.xlsx'

# Find the latest generated estimate
temp_dir = r'C:\Users\Server\Desktop\Rastenika Smeta\temp_pdfs'
files = [f for f in os.listdir(temp_dir) if f.endswith('.xlsx')]
if files:
    latest = max(files, key=lambda f: os.path.getmtime(os.path.join(temp_dir, f)))
    file2 = os.path.join(temp_dir, latest)
    print(f"Latest estimate: {file2}")
else:
    print("No estimates found")
    exit()

# Use win32com to convert both to PNG via Excel
import win32com.client as win32
import pythoncom

pythoncom.CoInitialize()
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

out1 = r'C:\Users\Server\.gemini\antigravity\brain\d4e6b417-c281-4ec1-902a-f3df9f745fd2\template_view.pdf'
out2 = r'C:\Users\Server\.gemini\antigravity\brain\d4e6b417-c281-4ec1-902a-f3df9f745fd2\estimate_view.pdf'

try:
    print("Converting template...")
    wb1 = excel.Workbooks.Open(file1)
    wb1.ExportAsFixedFormat(0, out1)
    wb1.Close(False)
    print(f"Saved: {out1}")

    print("Converting latest estimate...")
    wb2 = excel.Workbooks.Open(file2)
    wb2.ExportAsFixedFormat(0, out2)
    wb2.Close(False)
    print(f"Saved: {out2}")
finally:
    excel.Quit()
    pythoncom.CoUninitialize()

print("Done! Check the PDFs to compare visually.")
