import os
import time
import uuid
import win32com.client as win32
import pythoncom

class PDFGenerator:
    def __init__(self):
        self.BASE_DIR = os.path.abspath(os.path.dirname(__file__))
        self.temp_dir = os.path.join(self.BASE_DIR, 'temp_pdfs')
        os.makedirs(self.temp_dir, exist_ok=True)
        self._cleanup_old_files()

    def _cleanup_old_files(self, max_age_seconds=3600):
        """Удаляет PDF-файлы старше max_age_seconds (по умолчанию 1 час)."""
        now = time.time()
        try:
            for fname in os.listdir(self.temp_dir):
                if not fname.endswith('.pdf'):
                    continue
                fpath = os.path.join(self.temp_dir, fname)
                try:
                    if now - os.path.getmtime(fpath) > max_age_seconds:
                        os.remove(fpath)
                except OSError:
                    pass
        except OSError:
            pass

    def create_pdf_from_excel(self, excel_path):
        """
        Takes a perfectly formatted XLSX file and uses Excel's native
        ExportAsFixedFormat to guarantee a 1-to-1 visual PDF match.
        """
        # Required for multithreaded environments like Flask
        pythoncom.CoInitialize()
        
        filename = f"estimate_{uuid.uuid4()}.pdf"
        pdf_path = os.path.join(self.temp_dir, filename)
        
        # Open Excel headlessly using DispatchEx to ensure a fresh COM instance
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            # Must use absolute paths for win32com
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            
            # 0 представляет xlTypePDF
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            wb.Close(False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()
            
        return pdf_path
