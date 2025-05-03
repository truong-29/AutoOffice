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

class AutoOfficeUpdater:
    def __init__(self, repo_url="https://github.com/truong-29/AutoOffice"):
        self.repo_url = repo_url
        self.repo_owner = "truong-29"
        self.repo_name = "AutoOffice"
        self.api_url = f"https://api.github.com/repos/{self.repo_owner}/{self.repo_name}/contents/version.json"
        self.current_version = self._get_current_version()
        self.app_path = os.path.dirname(os.path.abspath(__file__))
        self.is_frozen = getattr(sys, 'frozen', False)
        self.exe_path = sys.executable if self.is_frozen else None
        
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
        """Kiểm tra cập nhật từ GitHub repository."""
        try:
            logger.info("Đang kiểm tra cập nhật...")
            
            # Thiết lập timeout cho requests
            timeout_seconds = 10
            
            # Sử dụng session để tái sử dụng kết nối
            session = requests.Session()
            session.headers.update({'User-Agent': 'AutoOffice-UpdateChecker'})
            
            try:
                response = session.get(self.api_url, timeout=timeout_seconds)
            except requests.exceptions.Timeout:
                logger.error(f"Timeout khi kết nối đến GitHub API sau {timeout_seconds} giây")
                return False, None, []
            except requests.exceptions.ConnectionError:
                logger.error("Lỗi kết nối đến GitHub API. Vui lòng kiểm tra kết nối internet.")
                return False, None, []
            
            if response.status_code != 200:
                logger.error(f"Lỗi khi kiểm tra cập nhật: HTTP {response.status_code}")
                logger.error(f"Nội dung phản hồi: {response.text[:200]}")
                return False, None, []
            
            # Lấy nội dung file version.json từ GitHub
            content_data = response.json()
            if "content" not in content_data:
                logger.error("Không tìm thấy trường 'content' trong phản hồi API")
                logger.error(f"Dữ liệu nhận được: {str(content_data)[:200]}")
                return False, None, []
                
            import base64
            try:
                content = base64.b64decode(content_data["content"]).decode("utf-8")
                remote_version_data = json.loads(content)
            except Exception as e:
                logger.error(f"Lỗi khi giải mã nội dung version.json: {e}")
                return False, None, []
            
            remote_version = remote_version_data.get("version", "1.0.0")
            logger.info(f"Phiên bản hiện tại: {self.current_version}, Phiên bản mới: {remote_version}")
            
            # So sánh phiên bản
            has_update = self._compare_versions(self.current_version, remote_version)
            
            if has_update:
                # Hiển thị thêm thông tin về bản cập nhật
                changes = remote_version_data.get("changes", [])
                changes_str = "\n".join([f"- {change}" for change in changes])
                logger.info(f"Có phiên bản mới: {remote_version}")
                logger.info(f"Các thay đổi:\n{changes_str}")
                return True, remote_version, changes
            else:
                logger.info("Không có cập nhật mới")
                return False, remote_version, []
                
        except Exception as e:
            logger.error(f"Lỗi khi kiểm tra cập nhật: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False, None, []
    
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
            
    def apply_update(self, extracted_dir):
        """Áp dụng bản cập nhật bằng cách thay thế tất cả các file."""
        try:
            logger.info("Đang áp dụng bản cập nhật...")
            
            # Danh sách file bỏ qua
            ignore_files = [".gitattributes", ".git"]
            
            # Copy tất cả file từ thư mục giải nén vào thư mục ứng dụng
            for item in os.listdir(extracted_dir):
                if item in ignore_files:
                    continue
                    
                src_path = os.path.join(extracted_dir, item)
                dst_path = os.path.join(self.app_path, item)
                
                # Nếu là thư mục
                if os.path.isdir(src_path):
                    if os.path.exists(dst_path):
                        shutil.rmtree(dst_path)
                    shutil.copytree(src_path, dst_path)
                # Nếu là file
                else:
                    if os.path.exists(dst_path):
                        os.remove(dst_path)
                    shutil.copy2(src_path, dst_path)
                    
                logger.info(f"Đã cập nhật: {item}")
                
            # Xóa thư mục tạm
            temp_dir = os.path.dirname(extracted_dir)
            
            # Nếu đang chạy từ file exe, cần build lại exe
            if self.is_frozen:
                self.build_new_exe(temp_dir)
            else:
                # Nếu không, xóa thư mục tạm
                shutil.rmtree(temp_dir)
            
            logger.info("Đã áp dụng bản cập nhật thành công")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi áp dụng bản cập nhật: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
            
    def build_new_exe(self, temp_dir):
        """Đóng gói ứng dụng thành file exe mới và thay thế file exe cũ."""
        try:
            logger.info("Đang đóng gói ứng dụng thành file exe mới...")
            
            # Kiểm tra xem PyInstaller đã được cài đặt chưa
            try:
                import PyInstaller
            except ImportError:
                # Nếu chưa cài đặt, cài đặt PyInstaller
                logger.info("Đang cài đặt PyInstaller...")
                pip_command = [sys.executable, "-m", "pip", "install", "pyinstaller"]
                subprocess.call(pip_command)
            
            # Tạo file batch để build exe
            build_script = os.path.join(temp_dir, "build_exe.bat")
            exe_output_dir = os.path.join(temp_dir, "dist")
            main_script = os.path.join(self.app_path, "main.py")
            
            # Chuẩn bị lệnh PyInstaller
            pyinstaller_cmd = f'"{sys.executable}" -m PyInstaller --onefile --windowed --icon="{os.path.join(self.app_path, "Logo.png")}" --name="AutoOffice" "{main_script}"'
            
            # Ghi file batch
            with open(build_script, 'w') as f:
                f.write(f"cd /d {self.app_path}\n")
                f.write(pyinstaller_cmd)
            
            # Chạy file batch
            logger.info("Đang chạy lệnh build exe...")
            build_result = subprocess.call(build_script, shell=True)
            
            if build_result != 0:
                logger.error(f"Lỗi khi build exe: Mã lỗi {build_result}")
                return False
                
            # Đường dẫn của file exe mới
            new_exe_path = os.path.join(self.app_path, "dist", "AutoOffice.exe")
            
            # Kiểm tra xem file exe mới đã được tạo chưa
            if os.path.exists(new_exe_path):
                logger.info(f"Đã tạo thành công file exe mới: {new_exe_path}")
                
                # Tạo file batch để khởi động lại ứng dụng sau khi cập nhật
                restart_script = os.path.join(temp_dir, "restart_app.bat")
                with open(restart_script, 'w') as f:
                    f.write("@echo off\n")
                    f.write("echo Khoi dong lai ung dung sau khi cap nhat...\n")
                    f.write(f'timeout /t 2 /nobreak > nul\n')
                    f.write(f'start "" "{self.exe_path}"\n')
                    f.write("exit\n")
                
                # Tạo file batch để thay thế file exe cũ
                replace_script = os.path.join(temp_dir, "replace_exe.bat")
                
                with open(replace_script, 'w') as f:
                    f.write("@echo off\n")
                    f.write("echo Dang thay the file exe cu...\n")
                    # Đợi lâu hơn để đảm bảo ứng dụng cũ đã đóng hoàn toàn
                    f.write("timeout /t 5 /nobreak > nul\n")
                    # Thử copy nhiều lần trong trường hợp file bị khóa
                    f.write(":retry_copy\n")
                    f.write(f'if exist "{self.exe_path}" (\n')
                    f.write(f'  copy /Y "{new_exe_path}" "{self.exe_path}" 2>nul\n')
                    f.write("  if errorlevel 1 (\n")
                    f.write("    echo Dang thu lai...\n")
                    f.write("    timeout /t 2 /nobreak > nul\n")
                    f.write("    goto retry_copy\n")
                    f.write("  )\n")
                    f.write(")\n")
                    # Dọn dẹp và khởi động lại
                    f.write("echo Da thay the file exe thanh cong!\n")
                    f.write(f'start "" "{restart_script}"\n')
                    f.write(f'rmdir /S /Q "{temp_dir}" 2>nul\n')
                    f.write(f'rmdir /S /Q "{os.path.join(self.app_path, "build")}" 2>nul\n')
                    f.write(f'rmdir /S /Q "{os.path.join(self.app_path, "dist")}" 2>nul\n')
                    f.write("exit\n")
                
                # Tạo file VBS để yêu cầu quyền Admin khi chạy batch script
                elevate_vbs = os.path.join(temp_dir, "elevate.vbs")
                with open(elevate_vbs, 'w') as f:
                    f.write('Set UAC = CreateObject("Shell.Application")\n')
                    f.write(f'UAC.ShellExecute "cmd.exe", "/c """{replace_script}"""", "", "runas", 1\n')
                
                # Chạy file VBS để thực thi script với quyền admin
                logger.info("Đang chuẩn bị thay thế file exe với quyền admin...")
                subprocess.Popen(["wscript.exe", elevate_vbs], 
                                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
                
                logger.info("Đã tạo lệnh thay thế file exe, ứng dụng sẽ khởi động lại sau khi hoàn tất")
                return True
            else:
                logger.error(f"Không thể tạo file exe mới. Đường dẫn {new_exe_path} không tồn tại.")
                return False
                
        except Exception as e:
            logger.error(f"Lỗi khi đóng gói ứng dụng thành file exe: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def restart_application(self):
        """Khởi động lại ứng dụng sau khi cập nhật."""
        try:
            logger.info("Đang khởi động lại ứng dụng...")
            
            # Nếu đang chạy từ exe
            if self.is_frozen:
                # Tạo một file batch để đợi tiến trình hiện tại kết thúc và hiển thị thông báo cho người dùng
                temp_dir = os.path.join(self.app_path, "temp_update")
                os.makedirs(temp_dir, exist_ok=True)
                
                message_script = os.path.join(temp_dir, "update_message.vbs")
                with open(message_script, 'w') as f:
                    f.write('WScript.Sleep 1000\n')  # Đợi 1 giây
                    f.write('MsgBox "Ứng dụng đang được cập nhật. Vui lòng đợi trong giây lát...", 64, "Đang cập nhật"\n')
                
                # Chạy script thông báo trong tiến trình riêng biệt
                subprocess.Popen(["wscript.exe", message_script],
                               creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
                
                logger.info("Ứng dụng sẽ được khởi động lại sau khi thay thế file exe")
                # Thoát ứng dụng để file exe không bị khóa
                os._exit(0)
            else:
                # Nếu đang chạy từ mã nguồn Python
                python = sys.executable
                main_script = os.path.join(self.app_path, "main.py")
                
                # Khởi động một tiến trình mới
                logger.info(f"Khởi động lại ứng dụng từ {main_script}")
                subprocess.Popen([python, main_script], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
                
                # Thoát tiến trình hiện tại
                os._exit(0)
            
        except Exception as e:
            logger.error(f"Lỗi khi khởi động lại ứng dụng: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
            
    def update_with_confirmation(self, parent_window=None):
        """Kiểm tra, xác nhận và thực hiện cập nhật."""
        has_update, new_version, changes = self.check_for_updates()
        
        if not has_update:
            return False
            
        # Tạo chuỗi thông tin thay đổi
        changes_text = ""
        if changes:
            changes_text = "\n".join([f"- {change}" for change in changes])
        
        # Hiển thị hộp thoại xác nhận
        if parent_window:
            result = messagebox.askyesno(
                "Cập nhật mới", 
                f"Có phiên bản mới: {new_version}\nBạn có muốn cập nhật ngay bây giờ không?\n\nCác thay đổi:\n{changes_text}",
                parent=parent_window
            )
        else:
            # Tạo cửa sổ tạm thời nếu không có cửa sổ chính
            temp_window = tk.Tk()
            temp_window.withdraw()  # Ẩn cửa sổ
            result = messagebox.askyesno(
                "Cập nhật mới", 
                f"Có phiên bản mới: {new_version}\nBạn có muốn cập nhật ngay bây giờ không?\n\nCác thay đổi:\n{changes_text}"
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
        extracted_dir = self.download_update()
        if not extracted_dir:
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể tải xuống bản cập nhật.", parent=parent_window)
            return False
            
        # Áp dụng bản cập nhật
        if self.apply_update(extracted_dir):
            # Khởi động lại ứng dụng
            self.restart_application()
            return True
        else:
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể áp dụng bản cập nhật.", parent=parent_window)
            return False
