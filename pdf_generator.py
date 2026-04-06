import os
import time
import uuid
import subprocess
import platform

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
        Конвертирует XLSX в PDF. 
        На Windows использует Excel (если доступен), на Linux — LibreOffice.
        """
        filename = f"estimate_{uuid.uuid4()}.pdf"
        pdf_path = os.path.join(self.temp_dir, filename)
        
        abs_excel_path = os.path.abspath(excel_path)
        abs_pdf_path = os.path.abspath(pdf_path)

        if platform.system() == 'Windows':
            try:
                import win32com.client as win32
                import pythoncom
                pythoncom.CoInitialize()
                excel = win32.DispatchEx('Excel.Application')
                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    wb = excel.Workbooks.Open(abs_excel_path)
                    wb.ExportAsFixedFormat(0, abs_pdf_path)
                    wb.Close(False)
                finally:
                    excel.Quit()
                    pythoncom.CoUninitialize()
                return pdf_path
            except Exception as e:
                print(f"Excel COM failed, trying fallback: {e}")

        # Fallback/Linux: LibreOffice
        try:
            # На Render/Railway/Ubuntu libreoffice обычно доступен как 'soffice' или 'libreoffice'
            commands = ['libreoffice', 'soffice']
            success = False
            for cmd in commands:
                try:
                    subprocess.run([
                        cmd, '--headless', '--convert-to', 'pdf', 
                        '--outdir', self.temp_dir, abs_excel_path
                    ], check=True, timeout=30)
                    # LibreOffice сохраняет файл с тем же именем, что и оригинал, но с .pdf
                    orig_name = os.path.splitext(os.path.basename(excel_path))[0]
                    generated_pdf = os.path.join(self.temp_dir, f"{orig_name}.pdf")
                    if os.path.exists(generated_pdf):
                        os.rename(generated_pdf, pdf_path)
                        success = True
                        break
                except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
                    continue
            
            if not success:
                raise RuntimeError("Could not find LibreOffice for PDF conversion")
                
        except Exception as e:
            print(f"PDF conversion failed: {e}")
            raise e

        return pdf_path
