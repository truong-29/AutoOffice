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
from tkinter import messagebox, ttk
import time
import psutil

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
        self.original_exe_name = os.path.basename(self.exe_path) if self.is_frozen else None
        
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
            shutil.rmtree(temp_dir)
            
            logger.info("Đã áp dụng bản cập nhật thành công")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi áp dụng bản cập nhật: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def build_exe(self):
        """Build file exe từ mã nguồn đã cập nhật."""
        try:
            logger.info("Đang chuẩn bị build file exe...")
            
            # Cài đặt pyinstaller nếu chưa có
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
                logger.info("Đã cài đặt PyInstaller")
            except Exception as e:
                logger.error(f"Không thể cài đặt PyInstaller: {e}")
                return False
            
            # Tạo thư mục build nếu chưa có
            build_dir = os.path.join(self.app_path, "build")
            if not os.path.exists(build_dir):
                os.makedirs(build_dir)
            
            # Tạo thư mục dist nếu chưa có
            dist_dir = os.path.join(self.app_path, "dist")
            if not os.path.exists(dist_dir):
                os.makedirs(dist_dir)
            
            # Xây dựng lệnh để build exe
            main_script = os.path.join(self.app_path, "main.py")
            icon_path = os.path.join(self.app_path, "Logo.png")
            
            # Tên exe mới
            exe_name = self.original_exe_name if self.original_exe_name else "AutoOffice.exe"
            
            # Xóa file exe cũ trong thư mục dist nếu có
            old_dist_exe = os.path.join(dist_dir, exe_name)
            if os.path.exists(old_dist_exe):
                try:
                    os.remove(old_dist_exe)
                    logger.info(f"Đã xóa file exe cũ trong thư mục dist: {old_dist_exe}")
                except Exception as e:
                    logger.warning(f"Không thể xóa file exe cũ trong dist: {e}")
            
            # Thực hiện lệnh build
            build_cmd = [
                sys.executable,
                "-m",
                "PyInstaller",
                "--noconfirm",
                "--onefile",
                "--windowed",
                f"--name={os.path.splitext(exe_name)[0]}",
                f"--add-data={icon_path};.",
                main_script
            ]
            
            logger.info(f"Đang chạy lệnh build: {' '.join(build_cmd)}")
            result = subprocess.run(build_cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                logger.error(f"Build exe thất bại: {result.stderr}")
                return False
            
            # Kiểm tra xem file exe mới có tồn tại không
            new_exe = os.path.join(dist_dir, exe_name)
            if not os.path.exists(new_exe):
                logger.error(f"Không tìm thấy file exe mới sau khi build: {new_exe}")
                return False
                
            logger.info(f"Build exe thành công: {new_exe}")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi build exe: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def get_pid_by_name(self, process_name):
        """Lấy PID của process theo tên."""
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] == process_name:
                    return proc.info['pid']
            return None
        except Exception as e:
            logger.error(f"Lỗi khi tìm PID: {e}")
            return None
    
    def replace_and_start_exe(self):
        """Thay thế và chạy file exe mới."""
        try:
            if not self.is_frozen:
                logger.info("Không phải chạy từ exe, bỏ qua thay thế exe")
                return True
                
            logger.info("Đang chuẩn bị thay thế file exe...")
            
            # Đường dẫn đến file exe cũ
            old_exe = self.exe_path
            exe_name = os.path.basename(old_exe)
            exe_dir = os.path.dirname(old_exe)
            
            # Đường dẫn đến file exe mới
            new_exe = os.path.join(self.app_path, "dist", exe_name)
            if not os.path.exists(new_exe):
                logger.error(f"Không tìm thấy file exe mới: {new_exe}")
                return False
            
            # Tạo batch file để thay thế exe với đường dẫn đầy đủ
            batch_path = os.path.join(self.app_path, "update_exe.bat")
            
            # Lấy PID của process hiện tại
            current_pid = os.getpid()
            
            with open(batch_path, 'w') as f:
                f.write('@echo off\n')
                f.write('echo Updating AutoOffice...\n')
                
                # Đợi process hiện tại kết thúc
                f.write(f'echo Waiting for process {current_pid} to end...\n')
                f.write(f'set /a wait_count=0\n')
                f.write(':wait_loop\n')
                f.write(f'tasklist /fi "PID eq {current_pid}" | find "{current_pid}" > nul\n')
                f.write('if not errorlevel 1 (\n')
                f.write('  timeout /t 1 /nobreak > nul\n')
                f.write('  set /a wait_count+=1\n')
                f.write('  if %wait_count% gtr 30 (\n')
                f.write('    echo Timeout waiting for process to end\n')
                f.write('    goto :force_close\n')
                f.write('  )\n')
                f.write('  goto wait_loop\n')
                f.write(')\n')
                
                # Nếu quá thời gian chờ, thử kết thúc bằng taskkill
                f.write(':force_close\n')
                f.write(f'taskkill /F /PID {current_pid} /T > nul 2>&1\n')
                f.write('timeout /t 2 /nobreak > nul\n')
                
                # Thay thế file exe
                f.write('echo Replacing executable...\n')
                f.write(f'copy /y "{new_exe}" "{old_exe}"\n')
                f.write('if errorlevel 1 (\n')
                f.write('  echo Failed to copy new executable\n')
                f.write('  pause\n')
                f.write('  exit /b 1\n')
                f.write(')\n')
                
                # Chạy exe mới
                f.write('echo Starting new version...\n')
                f.write(f'start "" "{old_exe}"\n')
                
                # Xóa batch file
                f.write('timeout /t 2 /nobreak > nul\n')
                f.write('del "%~f0"\n')
            
            # Tạo shortcut để chạy batch file với quyền admin
            vbs_path = os.path.join(self.app_path, "run_as_admin.vbs")
            with open(vbs_path, 'w') as f:
                f.write('Set UAC = CreateObject("Shell.Application")\n')
                f.write(f'UAC.ShellExecute "{batch_path}", "", "", "runas", 1\n')
            
            # Chạy VBS file để mở batch với quyền admin
            subprocess.Popen(['wscript.exe', vbs_path], 
                            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
            
            logger.info("Đã chuẩn bị thay thế và khởi động lại exe")
            # Chờ 2 giây để VBS kịp chạy batch
            time.sleep(2)
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi thay thế exe: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
            
    def restart_application(self):
        """Khởi động lại ứng dụng sau khi cập nhật."""
        try:
            logger.info("Đang khởi động lại ứng dụng...")
            
            # Nếu đang chạy từ file exe
            if self.is_frozen:
                # Thực hiện thay thế và chạy exe mới
                if self.replace_and_start_exe():
                    # Thoát tiến trình hiện tại
                    os._exit(0)
                else:
                    logger.error("Không thể thay thế và chạy exe mới")
                    return False
            else:
                # Nếu chạy từ Python, khởi động lại bình thường
                python = sys.executable
                main_script = os.path.join(self.app_path, "main.py")
                
                # Khởi động một tiến trình mới
                subprocess.Popen([python, main_script], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
                
                # Thoát tiến trình hiện tại
                os._exit(0)
            
        except Exception as e:
            logger.error(f"Lỗi khi khởi động lại ứng dụng: {e}")
            return False
            
    def create_update_batch(self, extracted_dir=None):
        """Tạo script để build và thay thế exe sau khi ứng dụng đóng."""
        try:
            logger.info("Đang tạo script PowerShell để hoàn tất cập nhật...")
            
            # Đường dẫn đến file exe hiện tại
            if not self.is_frozen:
                logger.info("Không phải chạy từ exe, bỏ qua tạo script update")
                return False
                
            # Lấy thông tin cần thiết
            current_pid = os.getpid()
            old_exe = self.exe_path
            exe_name = os.path.basename(old_exe)
            exe_dir = os.path.dirname(old_exe)
            app_dir = self.app_path
            
            # Đường dẫn tuyệt đối cho PowerShell script
            ps_script_path = os.path.join(self.app_path, "update_autooffice.ps1")
            log_path = os.path.join(self.app_path, "update_log.txt")
            
            with open(ps_script_path, 'w', encoding='utf-8') as f:
                # Thêm header và thiết lập basic
                f.write('# AutoOffice Update Script\n')
                f.write('$ErrorActionPreference = "Stop"\n')
                f.write('$logFile = "' + log_path.replace('\\', '\\\\') + '"\n\n')
                
                # Hàm ghi log
                f.write('function Write-Log($message) {\n')
                f.write('    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"\n')
                f.write('    "$timestamp - $message" | Out-File -FilePath $logFile -Append -Encoding utf8\n')
                f.write('    Write-Host $message\n')
                f.write('}\n\n')
                
                # Bắt đầu cập nhật
                f.write('Clear-Host\n')
                f.write('Write-Host "===== AUTO OFFICE UPDATE =====" -ForegroundColor Green\n')
                f.write('Write-Log "Bắt đầu quy trình cập nhật AutoOffice"\n\n')
                
                # Chuyển đến thư mục ứng dụng
                f.write('# Chuyển đến thư mục ứng dụng\n')
                f.write('Set-Location -Path "' + app_dir.replace('\\', '\\\\') + '"\n')
                f.write('Write-Log "Chuyển đến thư mục: $(Get-Location)"\n\n')
                
                # Đợi process hiện tại kết thúc
                f.write('# Đợi process hiện tại kết thúc\n')
                f.write('Write-Log "Đang đợi process có PID ' + str(current_pid) + ' kết thúc"\n')
                f.write('$waitCount = 0\n')
                f.write('Start-Sleep -Seconds 3\n')
                f.write('$process = Get-Process -Id ' + str(current_pid) + ' -ErrorAction SilentlyContinue\n')
                f.write('while ($process -ne $null -and $waitCount -lt 15) {\n')
                f.write('    Write-Host "Process vẫn đang chạy, đợi 1 giây..."\n')
                f.write('    Start-Sleep -Seconds 1\n')
                f.write('    $waitCount++\n')
                f.write('    $process = Get-Process -Id ' + str(current_pid) + ' -ErrorAction SilentlyContinue\n')
                f.write('}\n\n')
                
                # Kết thúc process nếu vẫn còn chạy
                f.write('if ($process -ne $null) {\n')
                f.write('    Write-Log "Process vẫn chạy sau khi đợi, thử kết thúc cưỡng bức"\n')
                f.write('    try {\n')
                f.write('        Stop-Process -Id ' + str(current_pid) + ' -Force -ErrorAction SilentlyContinue\n')
                f.write('        Write-Log "Đã kết thúc process thành công"\n')
                f.write('    } catch {\n')
                f.write('        Write-Log "Không thể kết thúc process: $_"\n')
                f.write('    }\n')
                f.write('    Start-Sleep -Seconds 2\n')
                f.write('} else {\n')
                f.write('    Write-Log "Process đã kết thúc"\n')
                f.write('}\n\n')
                
                # Kiểm tra Python và PyInstaller
                f.write('# Kiểm tra Python và PyInstaller\n')
                f.write('Write-Log "Kiểm tra Python và PyInstaller"\n')
                f.write('try {\n')
                f.write('    $pythonPath = (Get-Command python -ErrorAction Stop).Path\n')
                f.write('    Write-Log "Tìm thấy Python: $pythonPath"\n')
                f.write('} catch {\n')
                f.write('    Write-Log "Lỗi: Python không được cài đặt hoặc không có trong PATH"\n')
                f.write('    Write-Host "Lỗi: Python không được cài đặt. Vui lòng cài đặt Python và thử lại." -ForegroundColor Red\n')
                f.write('    Read-Host "Nhấn Enter để thoát"\n')
                f.write('    exit 1\n')
                f.write('}\n\n')
                
                # Cài đặt PyInstaller
                f.write('Write-Log "Đang cài đặt PyInstaller"\n')
                f.write('try {\n')
                f.write('    & python -m pip install pyinstaller -q\n')
                f.write('    if ($LASTEXITCODE -ne 0) { throw "Lỗi cài đặt PyInstaller" }\n')
                f.write('    Write-Log "Đã cài đặt PyInstaller thành công"\n')
                f.write('} catch {\n')
                f.write('    Write-Log "Lỗi cài đặt PyInstaller: $_"\n')
                f.write('    Write-Host "Lỗi cài đặt PyInstaller. Vui lòng thử lại." -ForegroundColor Red\n')
                f.write('    Read-Host "Nhấn Enter để thoát"\n')
                f.write('    exit 1\n')
                f.write('}\n\n')
                
                # Chuẩn bị build exe
                f.write('# Chuẩn bị build exe\n')
                f.write('Write-Log "Chuẩn bị build exe mới"\n')
                f.write('if (-not (Test-Path "dist")) { New-Item -ItemType Directory -Path "dist" | Out-Null }\n')
                f.write('if (Test-Path "dist\\\\' + exe_name + '") {\n')
                f.write('    Remove-Item "dist\\\\' + exe_name + '" -Force\n')
                f.write('    Write-Log "Đã xóa file exe cũ trong thư mục dist"\n')
                f.write('}\n\n')
                
                # Build exe mới
                icon_path = os.path.join(app_dir, "Logo.png").replace('\\', '\\\\')
                main_script = os.path.join(app_dir, "main.py").replace('\\', '\\\\')
                
                f.write('# Build exe mới\n')
                f.write('Write-Log "Đang build exe mới"\n')
                f.write('Write-Host "Quá trình này có thể mất vài phút..." -ForegroundColor Yellow\n')
                f.write('try {\n')
                f.write('    $buildOutput = & python -m PyInstaller --noconfirm --onefile --windowed ')
                f.write(f'--name="{os.path.splitext(exe_name)[0]}" ')
                f.write(f'--add-data="{icon_path};." "{main_script}" 2>&1\n')
                f.write('    $buildOutput | Out-File -FilePath "build_output.log" -Encoding utf8\n')
                f.write('    if ($LASTEXITCODE -ne 0) { throw "Lỗi build exe" }\n')
                f.write('    Write-Log "Build exe thành công"\n')
                f.write('} catch {\n')
                f.write('    Write-Log "Lỗi khi build exe: $_"\n')
                f.write('    Write-Log "Chi tiết lỗi build có thể xem trong file build_output.log"\n')
                f.write('    Write-Host "Lỗi khi build exe. Xem chi tiết trong build_output.log" -ForegroundColor Red\n')
                f.write('    Read-Host "Nhấn Enter để thoát"\n')
                f.write('    exit 1\n')
                f.write('}\n\n')
                
                # Kiểm tra file exe mới
                f.write('# Kiểm tra file exe mới\n')
                f.write('if (-not (Test-Path "dist\\\\' + exe_name + '")) {\n')
                f.write('    Write-Log "Không tìm thấy file exe mới sau khi build"\n')
                f.write('    Write-Host "Lỗi: Không tìm thấy file exe mới sau khi build" -ForegroundColor Red\n')
                f.write('    Read-Host "Nhấn Enter để thoát"\n')
                f.write('    exit 1\n')
                f.write(')\n\n')
                
                # Đóng các tiến trình liên quan
                f.write('# Kiểm tra và kết thúc các tiến trình liên quan\n')
                f.write('Write-Log "Kiểm tra và đóng các tiến trình liên quan"\n')
                f.write('$exeProcesses = Get-Process -Name "' + os.path.splitext(exe_name)[0] + '" -ErrorAction SilentlyContinue\n')
                f.write('if ($exeProcesses) {\n')
                f.write('    Write-Log "Đang đóng các tiến trình liên quan đến exe"\n')
                f.write('    $exeProcesses | ForEach-Object { \n')
                f.write('        try {\n')
                f.write('            $_.Kill()\n')
                f.write('            Write-Log "Đã kết thúc process $($_.Id)"\n')
                f.write('        } catch {\n')
                f.write('            Write-Log "Không thể kết thúc process $($_.Id): $_"\n')
                f.write('        }\n')
                f.write('    }\n')
                f.write('    Start-Sleep -Seconds 2\n')
                f.write('}\n\n')
                
                # Thay thế file exe cũ
                old_exe_escaped = old_exe.replace('\\', '\\\\')
                f.write('# Thay thế file exe cũ\n')
                f.write('Write-Log "Đang thay thế file exe cũ"\n')
                f.write('try {\n')
                f.write('    Copy-Item -Path "dist\\\\' + exe_name + '" -Destination "' + old_exe_escaped + '" -Force\n')
                f.write('    Write-Log "Đã thay thế file exe thành công"\n')
                f.write('} catch {\n')
                f.write('    Write-Log "Lỗi khi thay thế file exe: $_"\n')
                f.write('    Write-Host "Lỗi: Không thể thay thế file exe cũ. Có thể cần quyền admin." -ForegroundColor Red\n')
                f.write('    Write-Host "Đang thử lại với quyền admin..." -ForegroundColor Yellow\n')
                f.write('    Start-Process powershell -ArgumentList "-Command Copy-Item -Path ""dist\\\\' + exe_name + '"" -Destination ""' + old_exe_escaped + '"" -Force" -Verb RunAs\n')
                f.write('    Start-Sleep -Seconds 3\n')
                f.write('}\n\n')
                
                # Hoàn thành và chạy exe mới
                f.write('# Hoàn thành và chạy exe mới\n')
                f.write('Write-Log "Cập nhật hoàn tất"\n')
                f.write('Write-Host "Cập nhật hoàn tất thành công!" -ForegroundColor Green\n')
                f.write('Write-Host "Đang khởi động lại ứng dụng..." -ForegroundColor Cyan\n')
                f.write('Start-Sleep -Seconds 1\n')
                f.write('try {\n')
                f.write('    Start-Process -FilePath "' + old_exe_escaped + '"\n')
                f.write('    Write-Log "Đã khởi động lại ứng dụng"\n')
                f.write('} catch {\n')
                f.write('    Write-Log "Lỗi khi khởi động lại ứng dụng: $_"\n')
                f.write('}\n')
                f.write('Start-Sleep -Seconds 2\n')
                f.write('Remove-Item -Path "build_output.log" -ErrorAction SilentlyContinue\n')
                f.write('# End of script\n')
            
            # Tạo BAT file đơn giản để chạy PowerShell script với quyền admin (không dùng vbs)
            admin_bat_path = os.path.join(self.app_path, "run_update.bat")
            with open(admin_bat_path, 'w') as f:
                f.write('@echo off\n')
                f.write('cls\n')
                f.write('echo Dang chuan bi cap nhat AutoOffice...\n')
                f.write(f'echo Duong dan thu muc ung dung: {app_dir}\n')
                ps_path = os.path.join(app_dir, "update_autooffice.ps1").replace('\\', '\\\\')
                f.write(f'powershell -ExecutionPolicy Bypass -File "{ps_path}"\n')
                f.write('if errorlevel 1 (\n')
                f.write('    echo Loi khi chay PowerShell script. Vui long xem log de biet chi tiet.\n')
                f.write('    pause\n')
                f.write(')\n')
            
            logger.info(f"Đã tạo PowerShell script cập nhật: {ps_script_path}")
            logger.info(f"Đã tạo BAT file chạy script: {admin_bat_path}")
            return True
            
        except Exception as e:
            logger.error(f"Lỗi khi tạo script cập nhật: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def update_with_confirmation(self, parent_window=None):
        """Kiểm tra, xác nhận và thực hiện cập nhật."""
        has_update, new_version = self.check_for_updates()
        
        if not has_update:
            if parent_window:
                messagebox.showinfo("Cập nhật", "Bạn đang sử dụng phiên bản mới nhất.", parent=parent_window)
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
        update_window = None
        update_label = None
        progress_var = None
        
        if parent_window:
            # Tạo cửa sổ thông báo tiến trình
            update_window = tk.Toplevel(parent_window)
            update_window.title("Đang cập nhật")
            update_window.geometry("300x100")
            update_window.resizable(False, False)
            update_window.transient(parent_window)
            
            # Tạo các widget trong cửa sổ
            update_label = tk.Label(update_window, text="Đang tải bản cập nhật...")
            update_label.pack(pady=10)
            
            progress_var = tk.IntVar()
            progress_bar = ttk.Progressbar(update_window, variable=progress_var, maximum=100)
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            
            # Cập nhật giao diện
            update_window.update()
        
        # Tải xuống bản cập nhật
        extracted_dir = self.download_update()
        if not extracted_dir:
            if update_window:
                update_window.destroy()
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể tải xuống bản cập nhật.", parent=parent_window)
            return False
            
        # Cập nhật tiến trình
        if update_label:
            update_label.config(text="Đang áp dụng bản cập nhật...")
            progress_var.set(40)
            update_window.update()
        
        # Áp dụng bản cập nhật
        if not self.apply_update(extracted_dir):
            if update_window:
                update_window.destroy()
            if parent_window:
                messagebox.showerror("Lỗi", "Không thể áp dụng bản cập nhật.", parent=parent_window)
            return False
        
        # Cập nhật tiến trình
        if update_label:
            progress_var.set(60)
            update_window.update()
        
        # Nếu đang chạy từ file exe, tạo PowerShell script để hoàn tất cập nhật
        if self.is_frozen:
            if update_label:
                update_label.config(text="Đang chuẩn bị cập nhật exe...")
                progress_var.set(80)
                update_window.update()
            
            # Tạo batch file trực tiếp build và thay thế exe
            update_batch_path = self._create_simple_update_batch()
            if not update_batch_path:
                if update_window:
                    update_window.destroy()
                if parent_window:
                    messagebox.showerror("Lỗi", "Không thể tạo quy trình cập nhật exe.", parent=parent_window)
                return False
            
            # Cập nhật tiến trình
            if update_label:
                update_label.config(text="Đang chuẩn bị khởi động lại...")
                progress_var.set(90)
                update_window.update()
            
            # Hiển thị thông báo thành công
            if update_window:
                update_window.destroy()
                
            if parent_window:
                restart_result = messagebox.askyesno(
                    "Cập nhật thành công", 
                    "Quá trình tải và cập nhật mã nguồn đã hoàn tất.\n\n"
                    "Ứng dụng cần đóng để hoàn tất cập nhật và build file exe mới.\n"
                    "Quá trình này có thể mất vài phút.\n\n"
                    "Bạn có muốn tiếp tục không?",
                    parent=parent_window
                )
                
                if not restart_result:
                    messagebox.showinfo(
                        "Cập nhật hoãn lại",
                        "Quá trình cập nhật đã bị hoãn lại.\n"
                        "Mã nguồn đã được cập nhật, nhưng file exe chưa được cập nhật.\n"
                        "Bạn có thể tiếp tục sử dụng phiên bản hiện tại.",
                        parent=parent_window
                    )
                    return False
            
            try:
                # Chạy batch file với quyền admin bằng cách gọi runas
                update_cmd_path = update_batch_path.replace("/", "\\")
                logger.info(f"Đang chạy batch file cập nhật: {update_cmd_path}")
                
                # Tạo file VBS tạm thời để chạy với quyền admin
                temp_vbs_path = os.path.join(self.app_path, "run_as_admin.vbs")
                with open(temp_vbs_path, 'w') as f:
                    # Script VBS đơn giản hơn, trực tiếp gọi batch file
                    f.write('Set UAC = CreateObject("Shell.Application")\n')
                    f.write(f'UAC.ShellExecute "{update_cmd_path}", "", "", "runas", 1\n')
                
                # Chạy file VBS
                os.system(f'wscript.exe "{temp_vbs_path}"')
                
                # Hiển thị thông báo hướng dẫn
                if parent_window:
                    messagebox.showinfo(
                        "Đang cập nhật",
                        "Quá trình cập nhật đã bắt đầu.\n"
                        "Vui lòng đợi một lát để hoàn tất quá trình cập nhật.\n"
                        "Ứng dụng sẽ tự động đóng sau khi xác nhận.",
                        parent=parent_window
                    )
                
                # Đóng ứng dụng
                logger.info("Đóng ứng dụng để hoàn tất cập nhật")
                time.sleep(2)  # Đợi 2 giây để đảm bảo các thông báo được hiển thị
                os._exit(0)
                
            except Exception as e:
                logger.error(f"Lỗi khi chạy batch file cập nhật: {e}")
                if parent_window:
                    messagebox.showerror(
                        "Lỗi cập nhật", 
                        f"Không thể chạy batch file cập nhật: {e}\n"
                        "Vui lòng thử lại hoặc cập nhật thủ công.",
                        parent=parent_window
                    )
                return False
        else:
            # Nếu chạy từ Python, chỉ cần khởi động lại
            if update_window:
                update_window.destroy()
                
            if parent_window:
                messagebox.showinfo(
                    "Cập nhật thành công", 
                    "Ứng dụng sẽ tự động khởi động lại để hoàn tất cập nhật.",
                    parent=parent_window
                )
            
            # Khởi động lại ứng dụng
            self.restart_application()
        
        return True
        
    def _create_simple_update_batch(self):
        """Tạo batch file đơn giản để build và thay thế exe."""
        try:
            if not self.is_frozen:
                return False
                
            # Lấy thông tin cần thiết
            exe_path = self.exe_path
            exe_name = os.path.basename(exe_path)
            app_dir = self.app_path
                
            # Tạo batch file đơn giản để build exe và thay thế
            batch_path = os.path.join(app_dir, "update_exe.bat")
            
            with open(batch_path, 'w') as f:
                f.write('@echo off\n')
                f.write('title Auto Office Update\n')
                f.write('color 0A\n')
                f.write('cls\n')
                f.write('echo ===== AUTO OFFICE UPDATE =====\n')
                f.write('echo.\n')
                
                # Ghi thông tin chẩn đoán
                f.write('echo Thong tin cap nhat:\n')
                f.write(f'echo - Thu muc ung dung: {app_dir}\n')
                f.write(f'echo - File exe goc: {exe_path}\n')
                f.write(f'echo - Ten file exe: {exe_name}\n')
                f.write('echo.\n')
                
                f.write('echo Dang cho ung dung dong...\n')
                f.write(f'timeout /t 5 /nobreak\n')
                
                # Kiểm tra và đóng tiến trình nếu còn chạy
                f.write(f'tasklist /fi "imagename eq {exe_name}" | find "{exe_name}" > nul\n')
                f.write('if not errorlevel 1 (\n')
                f.write(f'  echo {exe_name} van dang chay, se ket thuc tien trinh...\n')
                f.write(f'  taskkill /f /im {exe_name} > nul\n')
                f.write('  timeout /t 2 /nobreak > nul\n')
                f.write(')\n\n')
                
                # Cài đặt PyInstaller
                f.write('echo Dang cai dat PyInstaller...\n')
                f.write('pip install pyinstaller > nul\n')
                f.write('if errorlevel 1 (\n')
                f.write('  echo Loi cai dat PyInstaller\n')
                f.write('  pause\n')
                f.write('  exit /b 1\n')
                f.write(')\n\n')
                
                # Build exe mới
                f.write('echo Dang build file exe moi...\n')
                f.write(f'cd /d "{app_dir}"\n')
                icon_path = os.path.join(app_dir, "Logo.png")
                main_path = os.path.join(app_dir, "main.py")
                f.write(f'pyinstaller --noconfirm --onefile --windowed --name="{os.path.splitext(exe_name)[0]}" --add-data="{icon_path};." "{main_path}"\n')
                f.write('if errorlevel 1 (\n')
                f.write('  echo Loi build file exe\n')
                f.write('  pause\n')
                f.write('  exit /b 1\n')
                f.write(')\n\n')
                
                # Kiểm tra file exe mới
                f.write(f'if not exist "dist\\{exe_name}" (\n')
                f.write('  echo Khong tim thay file exe moi\n')
                f.write('  pause\n')
                f.write('  exit /b 1\n')
                f.write(')\n\n')
                
                # Tạo bản sao lưu của file exe cũ
                f.write('echo Tao ban sao luu cua file exe cu...\n')
                f.write(f'if exist "{exe_path}.bak" del /f /q "{exe_path}.bak"\n')
                f.write(f'copy /y "{exe_path}" "{exe_path}.bak"\n')
                f.write('echo Da sao luu file exe cu.\n\n')
                
                # Thay thế file exe cũ - sử dụng nhiều phương pháp
                f.write('echo Dang thay the file exe cu...\n')
                
                # Phương pháp 1: Thử copy trực tiếp
                f.write('echo Phuong phap 1: Copy truc tiep\n')
                f.write(f'copy /y "dist\\{exe_name}" "{exe_path}"\n')
                f.write('if not errorlevel 1 (\n')
                f.write('  echo Da thay the file exe thanh cong.\n')
                f.write('  goto success\n')
                f.write(')\n\n')
                
                # Phương pháp 2: Thử đổi tên file cũ và copy file mới
                f.write('echo Phuong phap 2: Doi ten va copy\n')
                f.write(f'ren "{exe_path}" "temp_{exe_name}"\n')
                f.write(f'copy /y "dist\\{exe_name}" "{exe_path}"\n')
                f.write('if not errorlevel 1 (\n')
                f.write(f'  del /f /q "{os.path.dirname(exe_path)}\\temp_{exe_name}"\n')
                f.write('  echo Da thay the file exe thanh cong.\n')
                f.write('  goto success\n')
                f.write(')\n')
                f.write(f'ren "{os.path.dirname(exe_path)}\\temp_{exe_name}" "{exe_name}"\n\n')
                
                # Phương pháp 3: Sử dụng robocopy
                f.write('echo Phuong phap 3: Su dung robocopy\n')
                f.write(f'robocopy "dist" "{os.path.dirname(exe_path)}" {exe_name} /NFL /NDL /NJH /NJS /NC /NS /MT:4\n')
                f.write('if %ERRORLEVEL% LEQ 7 (\n')
                f.write('  echo Da thay the file exe thanh cong (robocopy).\n')
                f.write('  goto success\n')
                f.write(')\n\n')
                
                # Báo lỗi nếu không thể thay thế
                f.write('echo Loi: Khong the thay the file exe.\n')
                f.write('echo Co the can chay voi quyen quan tri.\n')
                f.write('echo Thu tao shortcut va chay voi quyen admin...\n')
                
                # Tạo script thay thế với quyền admin
                f.write('echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\\admin_copy.vbs"\n')
                f.write(f'echo UAC.ShellExecute "cmd.exe", "/c copy /y ""dist\\{exe_name}"" ""{exe_path}"" & echo Thanh cong!", "", "runas", 1 >> "%temp%\\admin_copy.vbs"\n')
                f.write('"%temp%\\admin_copy.vbs"\n')
                f.write('timeout /t 5 /nobreak > nul\n')
                
                f.write(f'if not exist "{exe_path}" (\n')
                f.write('  echo Loi: Khong the thay the file exe.\n')
                f.write('  echo Phuc hoi tu ban sao luu...\n')
                f.write(f'  copy /y "{exe_path}.bak" "{exe_path}"\n')
                f.write('  pause\n')
                f.write('  exit /b 1\n')
                f.write(')\n\n')
                
                # Xác nhận thành công
                f.write(':success\n')
                f.write('cls\n')
                f.write('echo ===== CAP NHAT THANH CONG =====\n')
                f.write('echo.\n')
                f.write('echo Da thay the file exe thanh cong.\n')
                f.write(f'echo File exe moi: {exe_path}\n')
                f.write('echo.\n')
                f.write('echo Dang khoi dong lai ung dung...\n')
                f.write('timeout /t 3 > nul\n')
                f.write(f'start "" "{exe_path}"\n')
                f.write('echo Hoan tat!\n')
                f.write('timeout /t 2 > nul\n')
                f.write('del "%~f0"\n')
                f.write('exit\n')
            
            logger.info(f"Đã tạo batch file cập nhật: {batch_path}")
            return batch_path
            
        except Exception as e:
            logger.error(f"Lỗi khi tạo batch file cập nhật: {e}")
            return False
