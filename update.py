import json
import os
import requests
import logging
import zipfile
import io
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def get_application_path():
    """Lấy đường dẫn đến thư mục chứa ứng dụng, hoạt động cả với file exe và mã nguồn."""
    if getattr(sys, 'frozen', False):
        # Nếu đang chạy từ file exe (đã đóng gói)
        application_path = os.path.dirname(sys.executable)
    else:
        # Nếu đang chạy từ mã nguồn
        application_path = os.path.dirname(os.path.abspath(__file__))
        
    logger.info(f"Đường dẫn ứng dụng: {application_path}")
    return application_path

class AutoOfficeUpdater:
    def __init__(self, repo_url="https://github.com/truong-29/AutoOffice"):
        self.repo_url = repo_url
        self.repo_owner = "truong-29"
        self.repo_name = "AutoOffice"
        self.api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/contents/version.json"
        self.app_path = get_application_path()
        self.current_version = self._get_current_version()
        
    def _get_current_version(self):
        """Lấy phiên bản hiện tại từ file version.json."""
        try:
            version_path = os.path.join(self.app_path, "version.json")
            
            if not os.path.exists(version_path):
                # Tạo file version.json nếu không tồn tại
                version_data = {
                    "version": "1.0.0",
                    "release_date": "2023-07-01",
                    "changes": ["Phiên bản đầu tiên"]
                }
                
                with open(version_path, 'w', encoding='utf-8') as f:
                    json.dump(version_data, f, ensure_ascii=False, indent=4)
                    
                logger.info("Đã tạo file version.json với phiên bản 1.0.0")
                return "1.0.0"
            
            with open(version_path, 'r', encoding='utf-8') as f:
                version_data = json.load(f)
                
            logger.info(f"Phiên bản hiện tại: {version_data.get('version', '1.0.0')}")
            return version_data.get("version", "1.0.0")
            
        except Exception as e:
            logger.error(f"Lỗi khi đọc phiên bản hiện tại: {e}")
            return "1.0.0"
    
    def check_for_updates(self):
        """Kiểm tra cập nhật từ GitHub repository."""
        try:
            logger.info("Đang kiểm tra cập nhật...")
            
            response = requests.get(self.api_url)
            
            if response.status_code != 200:
                logger.error(f"Lỗi khi kiểm tra cập nhật: HTTP {response.status_code}")
                return False, None
            
            # Lấy nội dung file version.json từ GitHub
            content_data = response.json()
            if "content" not in content_data:
                logger.error("Không tìm thấy trường 'content' trong phản hồi API")
                return False, None
                
            import base64
            content = base64.b64decode(content_data["content"]).decode("utf-8")
            remote_version_data = json.loads(content)
            
            remote_version = remote_version_data.get("version", "1.0.0")
            
            # So sánh phiên bản
            has_update = self._compare_versions(self.current_version, remote_version)
            
            if has_update:
                logger.info(f"Có phiên bản mới: {remote_version}")
                return True, remote_version
            else:
                logger.info("Không có cập nhật mới")
                return False, remote_version
                
        except Exception as e:
            logger.error(f"Lỗi khi kiểm tra cập nhật: {e}")
            return False, None
    
    def _compare_versions(self, current, remote):
        """So sánh phiên bản hiện tại với phiên bản mới từ server."""
        try:
            current_parts = list(map(int, current.split('.')))
            remote_parts = list(map(int, remote.split('.')))
            
            # Thêm các phần thiếu nếu cần
            while len(current_parts) < 3:
                current_parts.append(0)
            while len(remote_parts) < 3:
                remote_parts.append(0)
            
            # So sánh từng phần
            for i in range(3):
                if remote_parts[i] > current_parts[i]:
                    return True
                if remote_parts[i] < current_parts[i]:
                    return False
            
            # Nếu giống nhau
            return False
            
        except Exception as e:
            logger.error(f"Lỗi khi so sánh phiên bản: {e}")
            return False
    
    def get_update_url(self):
        """Trả về URL để người dùng tải bản cập nhật mới."""
        return self.repo_url + "/releases"
        
    def download_update(self):
        """Tải xuống bản cập nhật mới từ GitHub."""
        try:
            logger.info("Đang tải xuống bản cập nhật mới...")
            
            # URL để tải xuống repository dưới dạng zip
            download_url = f"https://github.com/{self.repo_owner}/{self.repo_name}/archive/refs/heads/main.zip"
            
            response = requests.get(download_url, stream=True)
            
            if response.status_code != 200:
                logger.error(f"Lỗi khi tải xuống bản cập nhật: HTTP {response.status_code}")
                return False
                
            # Tạo bộ nhớ đệm để lưu trữ dữ liệu zip
            buffer = io.BytesIO()
            
            # Tải xuống và lưu vào bộ nhớ đệm
            for chunk in response.iter_content(chunk_size=1024):
                if chunk:
                    buffer.write(chunk)
                    
            buffer.seek(0)
            
            # Tạo thư mục tạm thời để giải nén
            temp_dir = os.path.join(self.app_path, "temp_update")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            # Giải nén tệp zip
            with zipfile.ZipFile(buffer) as zip_ref:
                zip_ref.extractall(temp_dir)
                
            # Tìm thư mục giải nén
            extracted_dir = None
            for item in os.listdir(temp_dir):
                item_path = os.path.join(temp_dir, item)
                if os.path.isdir(item_path):
                    extracted_dir = item_path
                    break
                    
            if not extracted_dir:
                logger.error("Không tìm thấy thư mục sau khi giải nén")
                return False
                
            logger.info(f"Đã giải nén file vào: {extracted_dir}")
            return extracted_dir
            
        except Exception as e:
            logger.error(f"Lỗi khi tải xuống bản cập nhật: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
            
    def start_update_launcher(self, extracted_dir, new_version):
        """Khởi động update_launcher.py để tiến hành cập nhật."""
        try:
            logger.info("Đang khởi động update launcher...")
            
            # Đường dẫn đến update_launcher.py
            update_launcher_path = os.path.join(self.app_path, "update_launcher.py")
            
            # Kiểm tra xem file update_launcher.py có tồn tại không
            if not os.path.exists(update_launcher_path):
                # Copy update_launcher.py từ thư mục giải nén nếu có
                src_launcher_path = os.path.join(extracted_dir, "update_launcher.py")
                if os.path.exists(src_launcher_path):
                    shutil.copy2(src_launcher_path, update_launcher_path)
                else:
                    logger.error("Không tìm thấy update_launcher.py")
                    return False
            
            # Khởi động update_launcher.py với tham số là đường dẫn thư mục giải nén và phiên bản mới
            python_exe = sys.executable
            
            # Kiểm tra nếu đang chạy từ file exe
            if getattr(sys, 'frozen', False):
                # Nếu đã được đóng gói thành exe, tìm file update_launcher.exe
                update_launcher_exe = os.path.join(self.app_path, "update_launcher.exe")
                if os.path.exists(update_launcher_exe):
                    subprocess.Popen([update_launcher_exe, extracted_dir, new_version])
                else:
                    # Nếu không có file exe, sử dụng python
                    subprocess.Popen(["python", update_launcher_path, extracted_dir, new_version], cwd=self.app_path)
            else:
                # Nếu đang chạy từ mã nguồn
                subprocess.Popen([python_exe, update_launcher_path, extracted_dir, new_version], cwd=self.app_path)
                
            logger.info("Đã khởi động update launcher thành công")
            
            # Thoát ứng dụng hiện tại để update_launcher tiếp quản
            os._exit(0)
            
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi khởi động update launcher: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
            
    def update_with_confirmation(self, parent_window=None):
        """Kiểm tra, xác nhận và thực hiện cập nhật."""
        has_update, new_version = self.check_for_updates()
        
        if not has_update:
            return False
            
        # Hiển thị hộp thoại xác nhận
        if parent_window:
            result = messagebox.askyesno(
                "Cập nhật mới", 
                f"Có phiên bản mới: {new_version}\nBạn có muốn cập nhật ngay bây giờ không?",
                parent=parent_window
            )
        else:
            # Tạo cửa sổ tạm thời nếu không có cửa sổ chính
            temp_window = tk.Tk()
            temp_window.withdraw()  # Ẩn cửa sổ
            result = messagebox.askyesno(
                "Cập nhật mới", 
                f"Có phiên bản mới: {new_version}\nBạn có muốn cập nhật ngay bây giờ không?"
            )
            temp_window.destroy()
            
        if not result:
            logger.info("Người dùng đã từ chối cập nhật")
            return False
            
        # Tiến hành cập nhật
        logger.info("Người dùng đã chọn cập nhật")
        
        # Hiển thị thông báo đang cập nhật
        if parent_window:
            messagebox.showinfo(
                "Đang cập nhật", 
                "Ứng dụng đang tải xuống bản cập nhật. Vui lòng đợi...",
                parent=parent_window
            )
        
        # Tải xuống bản cập nhật
        extracted_dir = self.download_update()
        if not extracted_dir:
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể tải xuống bản cập nhật.", parent=parent_window)
            return False
            
        # Khởi động update_launcher để tiến hành cập nhật
        if parent_window:
            messagebox.showinfo(
                "Cập nhật", 
                "Ứng dụng sẽ đóng và tiến hành cập nhật. Vui lòng không tắt máy tính.",
                parent=parent_window
            )
            
        # Khởi động update_launcher và thoát ứng dụng hiện tại
        self.start_update_launcher(extracted_dir, new_version)
        return True
