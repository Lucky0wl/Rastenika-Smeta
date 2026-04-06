import os
import time
import uuid
from playwright.sync_api import sync_playwright

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

    def create_pdf_from_html(self, html_content):
        """
        Использует Playwright (Chromium) для рендеринга HTML в PDF.
        Это гарантирует идеальное качество и работу на Linux-хостинге.
        """
        filename = f"estimate_{uuid.uuid4()}.pdf"
        pdf_path = os.path.join(self.temp_dir, filename)
        
        with sync_playwright() as p:
            # Запуск браузера в безголовом режиме.
            # Аргумент --no-sandbox часто нужен в Docker/Linux
            browser = p.chromium.launch(headless=True, args=["--no-sandbox"])
            context = browser.new_context()
            page = context.new_page()
            
            # Устанавливаем контент страницы
            page.set_content(html_content)
            
            # Ждем загрузки всех ресурсов (шрифты, стили)
            page.wait_for_load_state("networkidle")
            
            # Генерируем PDF
            page.pdf(
                path=pdf_path,
                format="A4",
                print_background=True,
                margin={"top": "0px", "right": "0px", "bottom": "0px", "left": "0px"}
            )
            
            browser.close()
            
        return pdf_path
