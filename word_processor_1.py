from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.shared import Pt
from docx2python import docx2python
import os
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WordProcessor:
    def __init__(self):
        self.document = None
        self.file_path = None
        self.sections_info = []
        
    def open_document(self, file_path):
        """Mở tệp Word và đọc dữ liệu."""
        try:
            self.file_path = file_path
            self.document = Document(file_path)
            logger.info(f"Đã mở tệp: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi mở tệp: {e}")
            return False
    
    def analyze_document(self):
        """Phân tích tài liệu để tìm các ngắt phần."""
        if not self.document:
            logger.error("Chưa mở tệp nào.")
            return False
            
        self.sections_info = []
        
        # Lấy thông tin về các section trong tài liệu
        for i, section in enumerate(self.document.sections):
            section_type = section.start_type
            section_info = {
                "index": i,
                "type": section_type,
                "type_name": self._get_section_type_name(section_type),
                "page_break": section_type == WD_SECTION_START.NEW_PAGE,
                "needs_conversion": section_type == WD_SECTION_START.NEW_PAGE
            }
            self.sections_info.append(section_info)
            
        logger.info(f"Đã phân tích tệp: Tìm thấy {len(self.sections_info)} phần.")
        return self.sections_info
    
    def _get_section_type_name(self, section_type):
        """Trả về tên kiểu ngắt phần."""
        section_types = {
            WD_SECTION_START.CONTINUOUS: "Continuous",
            WD_SECTION_START.NEW_COLUMN: "New Column",
            WD_SECTION_START.NEW_PAGE: "Next Page",
            WD_SECTION_START.EVEN_PAGE: "Even Page",
            WD_SECTION_START.ODD_PAGE: "Odd Page"
        }
        return section_types.get(section_type, "Unknown")
    
    def fix_empty_pages(self):
        """Sửa các trang trắng bằng cách chuyển ngắt phần sang Continuous."""
        if not self.document:
            logger.error("Chưa mở tệp nào.")
            return False
            
        changes_made = 0
        
        for i, section in enumerate(self.document.sections):
            if section.start_type == WD_SECTION_START.NEW_PAGE:
                section.start_type = WD_SECTION_START.CONTINUOUS
                changes_made += 1
                logger.info(f"Đã chuyển phần {i} từ 'Next Page' sang 'Continuous'")
        
        logger.info(f"Đã thực hiện {changes_made} thay đổi.")
        return changes_made
    
    def save_document(self, output_path=None):
        """Lưu tài liệu đã chỉnh sửa."""
        if not self.document:
            logger.error("Chưa mở tệp nào.")
            return False
            
        if not output_path:
            # Tạo tên tệp mới nếu không được chỉ định
            file_name, file_ext = os.path.splitext(self.file_path)
            output_path = f"{file_name}_fixed{file_ext}"
            
        try:
            self.document.save(output_path)
            logger.info(f"Đã lưu tệp vào: {output_path}")
            return output_path
        except Exception as e:
            logger.error(f"Lỗi khi lưu tệp: {e}")
            return False
    
    def get_document_info(self):
        """Lấy thông tin cơ bản về tài liệu."""
        if not self.document:
            return None
            
        info = {
            "sections": len(self.document.sections),
            "paragraphs": len(self.document.paragraphs),
            "tables": len(self.document.tables)
        }
        return info
