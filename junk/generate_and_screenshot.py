import sys
import os
import glob
import time
import win32com.client as win32
import win32gui
import pythoncom
from PIL import ImageGrab

# --- Генерируем отчёт с более реалистичными данными ---
os.chdir(r'C:\Users\Server\Desktop\Rastenika Smeta')
sys.path.insert(0, r'C:\Users\Server\Desktop\Rastenika Smeta')

test_data = {
    "items": [
        {"name": "Туя западная Aureospicata", "parameters": "180-200 rb", "quantity": 3, "price": 16875, "planting": 5906.25, "total": 68343.75},
        {"name": "Ель обыкновенная Inversa", "parameters": "100-120 C35", "quantity": 2, "price": 12000, "planting": 4200, "total": 32400},
        {"name": "Берёза повислая Юнги (штамб)", "parameters": "Ств. 50-70 C25", "quantity": 1, "price": 25000, "planting": 8750, "total": 33750},
        {"name": "Сосна горная Мугус", "parameters": "40-50 C5", "quantity": 5, "price": 3500, "planting": 1225, "total": 23625},
        {"name": "Можжевельник казацкий Тамарисцифолиа", "parameters": "20-30 C3", "quantity": 10, "price": 1800, "planting": 630, "total": 24300},
    ],
    "material_total": 0,
    "tax_rate": 0
}

import app as flask_app
filepath = flask_app.create_app_xlsx(test_data)
print(f"Generated: {filepath}")

# Открыть и сделать скриншот всего листа
pythoncom.CoInitialize()
xl = win32.DispatchEx('Excel.Application')
xl.Visible = True
xl.DisplayAlerts = False

wb = xl.Workbooks.Open(os.path.abspath(filepath))
ws = wb.Sheets(1)
ws.Activate()

# Ctrl+Home прокрутка к началу
xl.ActiveWindow.Zoom = 70
xl.ActiveWindow.ScrollRow = 1

time.sleep(2)

hwnd = xl.Hwnd
rect = win32gui.GetWindowRect(hwnd)
img = ImageGrab.grab(rect)
out1 = r'C:\Users\Server\.gemini\antigravity\brain\d4e6b417-c281-4ec1-902a-f3df9f745fd2\report_top.png'
img.save(out1)
print(f"Top screenshot: {out1}")

# Прокрутка вниз
xl.ActiveWindow.ScrollRow = 10
time.sleep(1)
img2 = ImageGrab.grab(rect)
out2 = r'C:\Users\Server\.gemini\antigravity\brain\d4e6b417-c281-4ec1-902a-f3df9f745fd2\report_bottom.png'
img2.save(out2)
print(f"Bottom screenshot: {out2}")

wb.Close(False)
xl.Quit()
pythoncom.CoUninitialize()
print("Done")
