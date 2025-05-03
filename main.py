import tkinter as tk
import logging
import sys
import os

from word_processor_1 import WordProcessor
from gui import AutoOfficeGUI
from update import AutoOfficeUpdater, get_application_path

# Thiết lập đường dẫn ứng dụng
app_path = get_application_path()

# Thiết lập logging
logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(app_path, "autooffice.log"), encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def main():
    """Hàm chính khởi động ứng dụng."""
    try:
        # Khởi tạo các thành phần
        logger.info("Khởi động ứng dụng Auto Office")
        
        # Tạo cửa sổ chính
        root = tk.Tk()
        
        # Thiết lập icon nếu có
        try:
            logo_path = os.path.join(app_path, "Logo.png")
            if os.path.exists(logo_path):
                logo_icon = tk.PhotoImage(file=logo_path)
                root.iconphoto(True, logo_icon)
                logger.info("Đã tải logo ứng dụng")
            else:
                logger.warning(f"Không tìm thấy file logo tại {logo_path}")
        except Exception as e:
            logger.warning(f"Không thể thiết lập icon: {e}")
        
        # Khởi tạo các module
        word_processor = WordProcessor()
        updater = AutoOfficeUpdater()
        
        # Khởi tạo giao diện
        app = AutoOfficeGUI(root, word_processor, updater)
        
        # Chạy ứng dụng
        root.mainloop()
        
    except Exception as e:
        logger.error(f"Lỗi không xác định: {e}")
        import traceback
        logger.error(traceback.format_exc())

if __name__ == "__main__":
    main()
