import json
import os
import logging
import sys
import tkinter as tk
from tkinter import messagebox
import requests
import zipfile
import shutil
import subprocess

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AutoOfficeUpdater:
    def __init__(self, repo_url="https://github.com/truong-29/AutoOffice"):
        self.repo_url = repo_url
        self.repo_owner = "truong-29"
        self.repo_name = "AutoOffice"
        self.current_version = self._get_current_version()
        self.app_path = os.path.dirname(os.path.abspath(__file__))
        self.temp_dir = os.path.join(self.app_path, "temp_update")
        
    def _get_current_version(self):
        """Lấy phiên bản hiện tại từ file version.json."""
        try:
            version_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "version.json")
            
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
        """Kiểm tra phiên bản mới nhất từ GitHub."""
        try:
            logger.info("Đang kiểm tra cập nhật...")
            
            # Lấy thông tin phiên bản mới nhất từ GitHub
            api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
            response = requests.get(api_url, timeout=10)
            
            if response.status_code != 200:
                logger.error(f"Không thể lấy thông tin phiên bản: {response.status_code}")
                return False, None
                
            release_data = response.json()
            latest_version = release_data.get("tag_name", "").replace("v", "")
            
            logger.info(f"Phiên bản mới nhất: {latest_version}")
            logger.info(f"Phiên bản hiện tại: {self.current_version}")
            
            # So sánh phiên bản
            if self._compare_versions(latest_version, self.current_version) > 0:
                return True, latest_version
            else:
                return False, latest_version
                
        except Exception as e:
            logger.error(f"Lỗi khi kiểm tra cập nhật: {e}")
            return False, None
    
    def _compare_versions(self, version1, version2):
        """So sánh hai phiên bản. Trả về 1 nếu version1 > version2, 0 nếu bằng, -1 nếu nhỏ hơn."""
        v1_parts = [int(x) for x in version1.split(".")]
        v2_parts = [int(x) for x in version2.split(".")]
        
        # Đảm bảo cả hai danh sách có cùng độ dài
        while len(v1_parts) < 3:
            v1_parts.append(0)
        while len(v2_parts) < 3:
            v2_parts.append(0)
            
        for i in range(3):
            if v1_parts[i] > v2_parts[i]:
                return 1
            elif v1_parts[i] < v2_parts[i]:
                return -1
                
        return 0
    
    def get_update_url(self):
        """Trả về URL để người dùng tải bản cập nhật mới."""
        return f"{self.repo_url}/releases/latest"
    
    def download_update(self, version):
        """Tải xuống bản cập nhật mới từ GitHub."""
        try:
            # Tạo thư mục tạm thời nếu chưa tồn tại
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                
            os.makedirs(self.temp_dir, exist_ok=True)
            
            # Lấy thông tin bản phát hành mới nhất
            api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
            response = requests.get(api_url, timeout=10)
            
            if response.status_code != 200:
                logger.error(f"Không thể lấy thông tin bản phát hành: {response.status_code}")
                return False
                
            release_data = response.json()
            
            # Tìm asset ZIP để tải xuống
            zip_asset = None
            for asset in release_data.get("assets", []):
                if asset.get("name", "").endswith(".zip"):
                    zip_asset = asset
                    break
                    
            if not zip_asset:
                logger.error("Không tìm thấy file ZIP trong bản phát hành")
                return False
                
            # Tải xuống file ZIP
            download_url = zip_asset.get("browser_download_url")
            zip_path = os.path.join(self.temp_dir, "update.zip")
            
            logger.info(f"Đang tải xuống bản cập nhật từ: {download_url}")
            
            response = requests.get(download_url, stream=True, timeout=60)
            
            if response.status_code != 200:
                logger.error(f"Không thể tải xuống bản cập nhật: {response.status_code}")
                return False
                
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    
            logger.info(f"Đã tải xuống bản cập nhật vào: {zip_path}")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi tải xuống bản cập nhật: {e}")
            return False
    
    def apply_update(self):
        """Giải nén và áp dụng bản cập nhật."""
        try:
            # Đường dẫn đến file ZIP
            zip_path = os.path.join(self.temp_dir, "update.zip")
            extract_path = os.path.join(self.temp_dir, "extracted")
            
            if not os.path.exists(zip_path):
                logger.error("Không tìm thấy file ZIP")
                return False
                
            # Tạo thư mục giải nén
            if os.path.exists(extract_path):
                shutil.rmtree(extract_path)
                
            os.makedirs(extract_path, exist_ok=True)
            
            # Giải nén file ZIP
            logger.info(f"Đang giải nén file {zip_path}...")
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_path)
                
            logger.info(f"Đã giải nén vào: {extract_path}")
            
            # Tìm thư mục gốc trong bản giải nén
            root_dir = extract_path
            subdirs = [d for d in os.listdir(extract_path) if os.path.isdir(os.path.join(extract_path, d))]
            
            if len(subdirs) == 1:
                root_dir = os.path.join(extract_path, subdirs[0])
                
            # Sao chép các file từ bản cập nhật vào thư mục ứng dụng
            logger.info("Đang sao chép các file cập nhật...")
            
            for item in os.listdir(root_dir):
                source = os.path.join(root_dir, item)
                destination = os.path.join(self.app_path, item)
                
                # Bỏ qua các file không cần thiết
                if item == ".gitattributes":
                    continue
                    
                # Sao chép file/thư mục
                if os.path.isdir(source):
                    if os.path.exists(destination):
                        shutil.rmtree(destination)
                    shutil.copytree(source, destination)
                else:
                    shutil.copy2(source, destination)
                    
            logger.info("Đã áp dụng cập nhật thành công")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi áp dụng bản cập nhật: {e}")
            return False
        finally:
            # Dọn dẹp
            try:
                if os.path.exists(self.temp_dir):
                    shutil.rmtree(self.temp_dir)
            except Exception:
                pass
    
    def restart_application(self):
        """Khởi động lại ứng dụng sau khi cập nhật."""
        try:
            logger.info("Đang khởi động lại ứng dụng...")
            
            # Tạo lệnh để khởi động lại ứng dụng
            is_frozen = getattr(sys, 'frozen', False)
            
            if is_frozen:
                # Nếu là file exe (đã đóng gói)
                app_path = sys.executable
                subprocess.Popen([app_path], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
            else:
                # Nếu là script Python
                python = sys.executable
                main_script = os.path.join(self.app_path, "main.py")
                subprocess.Popen([python, main_script], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
            
            # Thoát tiến trình hiện tại
            os._exit(0)
            
        except Exception as e:
            logger.error(f"Lỗi khi khởi động lại ứng dụng: {e}")
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
                "Ứng dụng sẽ tự động khởi động lại sau khi cập nhật hoàn tất.",
                parent=parent_window
            )
        
        # Tải xuống bản cập nhật
        if not self.download_update(new_version):
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể tải xuống bản cập nhật.", parent=parent_window)
            return False
            
        # Áp dụng bản cập nhật
        if not self.apply_update():
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể áp dụng bản cập nhật.", parent=parent_window)
            return False
            
        # Khởi động lại ứng dụng
        self.restart_application()
        return True
            