import requests
import os
import glob

def run_test():
    payload = {
        'items': [
            {'name': 'Туя западная Смарагд (очень длинное название для проверки переноса строк)', 'parameters': '100-120см', 'quantity': 5, 'price': 1000, 'planting': 350, 'total': 6750},
            {'name': 'Сосна горная Пумилио', 'parameters': 'С5', 'quantity': 2, 'price': 2500, 'planting': 875, 'total': 6750}
        ],
        'material_total': 10000,
        'labor_total': 3500,
        'tax_rate': 0,
        'tax_amount': 0,
        'grand_total': 13500
    }
    
    # Clean old temp files 
    for f in glob.glob('temp_pdfs/estimate_*.pdf'):
        try: os.remove(f)
        except: pass

    resp = requests.post('http://127.0.0.1:5000/generate-pdf', json=payload)
    if resp.status_code == 200:
        print("PDF Generation Successful (200 OK)")
        
        # Output is returned directly in the stream, but the temp file is also left on disk
        new_files = glob.glob('temp_pdfs/estimate_*.pdf')
        if not new_files:
            print("Could not find generated PDF temp file.")
            return
            
        pdf_path = new_files[0]
        size = os.path.getsize(pdf_path)
        print(f"Generated PDF found at {pdf_path}. Size: {size} bytes.")
        
    else:
        print(f"Failed with status: {resp.status_code}")
        print(resp.text)

if __name__ == '__main__':
    run_test()
