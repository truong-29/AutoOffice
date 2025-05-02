import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import os
import time
import re
import sys
import requests
import subprocess
import shutil
import datetime
import tempfile
import ctypes
import traceback
import atexit

class WordCleanerApp:
    def __init__(self, root, word_processor):
        """
        Khởi tạo giao diện người dùng
        
        Args:
            root (tk.Tk): Cửa sổ gốc Tkinter
            word_processor (WordProcessor): Đối tượng xử lý file Word
        """
        self.root = root
        self.root.title("Ứng dụng xóa trang trắng trong Word")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Đảm bảo các thư viện cần thiết được load
        self._ensure_libraries_loaded()
        
        self.word_processor = word_processor
        self.selected_file = None
        self.analysis_results = []
        
        # Định nghĩa phiên bản hiện tại
        self.current_version = "1.0.0"
        self.github_repo = "truong-29/AutoOffice"
        
        # Tạo file version.json nếu chạy từ thư mục có exe
        if getattr(sys, 'frozen', False):
            try:
                self._create_version_json()
            except Exception as e:
                self.word_processor.log(f"Không thể tạo version.json: {e}")
        
        # Cài đặt xử lý khi đóng chương trình
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Kiểm tra xem có module pywin32 không
        self.has_pywin32 = self._check_pywin32()
        
        self.create_widgets()
        
        # Hiển thị trạng thái
        self.show_status("Sẵn sàng")
        
        # Hiển thị phiên bản ở thanh tiêu đề
        self.update_title_with_version()
        
        # Kiểm tra cập nhật sau khi khởi động
        threading.Timer(2.0, self.check_for_updates).start()
        
    def _ensure_libraries_loaded(self):
        """Đảm bảo các thư viện cần thiết được load thành công"""
        try:
            # Thử import một số thư viện quan trọng để đảm bảo chúng đã được load
            import win32com.client
            import pythoncom
            
            # Khởi tạo COM cho thread hiện tại
            pythoncom.CoInitialize()
            
            # Đăng ký hàm dọn dẹp khi đóng ứng dụng
            atexit.register(pythoncom.CoUninitialize)
            
            # Đăng ký COM component cho Word nếu cần
            self._register_word_com()
            
            print("Đã tải các thư viện cần thiết thành công")
        except Exception as e:
            print(f"Lỗi khi tải thư viện: {e}")
            # Hiển thị thông báo lỗi nếu cần
            self.show_error_dialog(f"Không thể tải thư viện cần thiết: {e}\n\nỨng dụng có thể không hoạt động đúng!")

    def _check_pywin32(self):
        """Kiểm tra xem pywin32 đã được cài đặt đúng cách chưa"""
        try:
            import win32com
            import win32api
            import pythoncom
            return True
        except ImportError:
            return False
        except Exception:
            return False

    def _register_word_com(self):
        """Đăng ký COM component cho Word nếu cần"""
        try:
            # Chỉ thực hiện trên Windows
            if sys.platform.startswith('win'):
                # Kiểm tra xem Word COM đã được đăng ký chưa
                try:
                    import win32com.client
                    word_app = win32com.client.Dispatch("Word.Application")
                    word_app.Quit()
                    del word_app
                    # Nếu không có lỗi, Word COM đã được đăng ký
                except Exception:
                    # Thử đăng ký lại Word COM
                    import os
                    import subprocess
                    
                    # Tìm makepy.py trong thư mục win32com
                    import win32com
                    win32com_dir = os.path.dirname(win32com.__file__)
                    makepy_path = os.path.join(win32com_dir, 'client', 'makepy.py')
                    
                    if os.path.exists(makepy_path):
                        # Chạy makepy.py để đăng ký Word
                        python_exe = sys.executable
                        cmd = [python_exe, makepy_path, '-i', 'Microsoft Word 16.0 Object Library']
                        subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        print("Đã đăng ký Word COM")
        except Exception as e:
            print(f"Lỗi khi đăng ký Word COM: {e}")

    def _on_closing(self):
        """Xử lý sự kiện khi đóng ứng dụng"""
        try:
            # Dọn dẹp tài nguyên
            if hasattr(self, 'word_processor') and self.word_processor:
                self.word_processor.cleanup()
                
            # Đóng COM nếu đã khởi tạo
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
                
            # Đóng ứng dụng
            self.root.destroy()
        except Exception as e:
            print(f"Lỗi khi đóng ứng dụng: {e}")
            # Vẫn phải đóng ứng dụng
            self.root.destroy()

    def create_widgets(self):
        """Tạo các thành phần giao diện"""
        # Frame chính
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook (Tab)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab xử lý
        process_tab = ttk.Frame(self.notebook)
        self.notebook.add(process_tab, text="Xử lý tệp")
        
        # Tab xem log
        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(log_tab, text="Nhật ký")
        
        # === Tạo nội dung cho Tab xử lý ===
        # Frame chọn tệp
        file_frame = ttk.LabelFrame(process_tab, text="Chọn tệp Word", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        self.file_path_var = tk.StringVar()
        ttk.Label(file_frame, textvariable=self.file_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Duyệt...", command=self.browse_file).pack(side=tk.RIGHT)
        
        # Frame phân tích
        analyze_frame = ttk.LabelFrame(process_tab, text="Phân tích tài liệu", padding="10")
        analyze_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        analyze_button_frame = ttk.Frame(analyze_frame)
        analyze_button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(analyze_button_frame, text="Phân tích tài liệu", command=self.analyze_document).pack(side=tk.LEFT, padx=5)
        
        # Thêm filter cho trang trắng
        self.show_only_blank_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(analyze_button_frame, text="Chỉ hiển thị các trang trắng", 
                       variable=self.show_only_blank_var, 
                       command=self.filter_results).pack(side=tk.LEFT, padx=20)
        
        # Thêm nút kiểm tra cập nhật ở góc phải
        ttk.Button(analyze_button_frame, text="Kiểm tra cập nhật", 
                  command=self.check_for_updates).pack(side=tk.RIGHT, padx=5)
        
        # Treeview hiển thị kết quả phân tích
        columns = ("page_num", "section", "is_blank", "reason", "status")
        self.results_tree = ttk.Treeview(analyze_frame, columns=columns, show="headings")
        
        # Đặt tiêu đề cột
        self.results_tree.heading("page_num", text="Số trang")
        self.results_tree.heading("section", text="Section")
        self.results_tree.heading("is_blank", text="Trạng thái")
        self.results_tree.heading("reason", text="Chi tiết")
        self.results_tree.heading("status", text="Tình trạng xử lý")
        
        # Đặt chiều rộng cột
        self.results_tree.column("page_num", width=70, anchor=tk.CENTER)
        self.results_tree.column("section", width=70, anchor=tk.CENTER)
        self.results_tree.column("is_blank", width=100, anchor=tk.CENTER)
        self.results_tree.column("reason", width=200)
        self.results_tree.column("status", width=100, anchor=tk.CENTER)
        
        # Định nghĩa các tags cho Treeview
        self.results_tree.tag_configure("processed", background="#E8F5E9")  # Màu xanh nhạt cho đã xử lý
        self.results_tree.tag_configure("failed", background="#FFEBEE")  # Màu đỏ nhạt cho lỗi
        self.results_tree.tag_configure("pending", background="#FFF8E1")  # Màu vàng nhạt cho chờ xử lý
        
        # Thanh cuộn
        scrollbar = ttk.Scrollbar(analyze_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscroll=scrollbar.set)
        
        # Sự kiện khi nhấp đúp vào một dòng
        self.results_tree.bind("<Double-1>", self.on_item_double_click)
        
        # Hiển thị TreeView với scrollbar
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame xử lý
        process_frame = ttk.LabelFrame(process_tab, text="Xử lý tài liệu", padding="10")
        process_frame.pack(fill=tk.X, pady=5)
        
        processing_method_frame = ttk.Frame(process_frame)
        processing_method_frame.pack(fill=tk.X, pady=5)
        
        # Ẩn lựa chọn phương pháp xử lý vì sẽ thử tất cả các phương pháp
        self.processing_method = tk.StringVar(value="auto")
        
        # Thêm một nút duy nhất trên giao diện
        process_button = ttk.Button(
            process_frame, 
            text="Xử lý triệt để trang trắng", 
            command=self.advanced_process_with_tracking,
            style="Primary.TButton"
        )
        process_button.pack(fill=tk.X, pady=10, padx=20, ipady=5)
        
        # Tạo style cho nút xử lý
        style = ttk.Style()
        style.configure("Primary.TButton", 
                       font=("Arial", 11, "bold"),
                       padding=10)
        style.map("Primary.TButton",
                 background=[("active", "#45a049"), ("!active", "#4CAF50")],
                 foreground=[("active", "white"), ("!active", "white")])
        
        # === Tạo nội dung cho Tab xem log ===
        log_control_frame = ttk.Frame(log_tab)
        log_control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(log_control_frame, text="Làm mới", command=self.refresh_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_control_frame, text="Xóa log", command=self.clear_log_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(log_control_frame, text="Mở thư mục log", command=self.open_log_folder).pack(side=tk.LEFT, padx=5)
        
        # Text hiển thị log
        log_frame = ttk.Frame(log_tab)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Thanh trạng thái (chung cho cả hai tab)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_var = tk.StringVar(value="Sẵn sàng")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=200, mode='indeterminate')
        self.progress.pack(side=tk.RIGHT)
        
        # Đọc log ban đầu
        self.refresh_log()
    
    def check_for_updates(self):
        """Kiểm tra xem có phiên bản mới trên GitHub không"""
        self.word_processor.log("Checking for updates from GitHub...")
        self.status_var.set("Checking for updates...")
        
        def perform_check():
            try:
                # Sử dụng timeout để tránh treo ứng dụng
                import main
                
                # Sử dụng URL raw để lấy version từ file JSON thay vì parse file python
                url = f"https://raw.githubusercontent.com/{self.github_repo}/main/version.json"
                self.word_processor.log(f"Checking for updates from: {url}")
                response = requests.get(url, timeout=10)
                
                if response.status_code != 200:
                    # Thử phương án dự phòng - trích xuất từ gui.py
                    url = f"https://raw.githubusercontent.com/{self.github_repo}/main/gui.py"
                    self.word_processor.log(f"Could not get version.json, trying gui.py: {url}")
                    response = requests.get(url, timeout=10)
                    
                    if response.status_code != 200:
                        self.word_processor.log(f"Could not connect to GitHub. Status code: {response.status_code}", error=True)
                        self.root.after(0, lambda: self.status_var.set(f"Ready - v{self.current_version}"))
                        return
                    
                    # Lưu content vào file tạm để đảm bảo encoding
                    content = response.text
                    temp_file = os.path.join(tempfile.gettempdir(), f"autooffice_temp_gui_{int(time.time())}.py")
                    with open(temp_file, 'w', encoding='utf-8') as f:
                        f.write(content)
                    
                    # Đọc file tạm
                    with open(temp_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Xóa file tạm
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                    
                    # Tìm version trong nội dung file
                    version_match = re.search(r'self\.current_version\s*=\s*["\']([^"\']+)["\']', content)
                    
                    if not version_match:
                        self.word_processor.log("Could not determine version from gui.py on GitHub", error=True)
                        self.root.after(0, lambda: self.status_var.set(f"Ready - v{self.current_version}"))
                        return
                    
                    github_version = version_match.group(1)
                else:
                    # Parse JSON response
                    try:
                        version_data = response.json()
                        github_version = version_data.get("version", "0.0.0")
                    except:
                        self.word_processor.log("Error parsing JSON from version.json", error=True)
                        # Fallback to string parsing
                        content = response.text
                        version_match = re.search(r'"version"\s*:\s*"([^"]+)"', content)
                        if version_match:
                            github_version = version_match.group(1)
                        else:
                            self.word_processor.log("Could not determine version from version.json", error=True)
                            self.root.after(0, lambda: self.status_var.set(f"Ready - v{self.current_version}"))
                            return
                
                self.word_processor.log(f"Current version: {self.current_version}, GitHub version: {github_version}")
                
                # So sánh phiên bản
                if github_version != self.current_version:
                    # Hiển thị thông báo cập nhật
                    self.root.after(0, lambda: self.prompt_update(github_version))
                else:
                    self.word_processor.log("Using the latest version")
                    self.root.after(0, lambda: self.status_var.set(f"Ready - v{self.current_version} (Latest version)"))
            
            except Exception as e:
                self.word_processor.log(f"Error checking for updates: {e}", error=True)
                self.root.after(0, lambda: self.status_var.set(f"Ready - v{self.current_version}"))
        
        # Tạo thread mới để tránh đóng băng giao diện
        threading.Thread(target=perform_check, daemon=True).start()
    
    def prompt_update(self, new_version):
        """Hỏi người dùng có muốn cập nhật lên phiên bản mới không"""
        response = messagebox.askyesno(
            "New Version Available",
            f"A new version v{new_version} is available (Current version: v{self.current_version}).\n\n"
            "Do you want to update?\n\n"
            "Note: The update process will close the current application and create a new exe file."
        )
        
        if response:
            self.update_application(new_version)
        else:
            self.status_var.set(f"Ready - v{self.current_version} (New version available: v{new_version})")
    
    def update_application(self, new_version):
        """Cập nhật ứng dụng lên phiên bản mới"""
        self.status_var.set("Updating...")
        self.progress.start()
        self.word_processor.log(f"Starting update to version {new_version}")
        
        def perform_update():
            try:
                # Kiểm tra xem ứng dụng đang chạy từ file exe hay không
                is_exe = getattr(sys, 'frozen', False)
                if not is_exe:
                    self.word_processor.log("Cannot update because application is not running from exe")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Cannot Auto-Update",
                        "Auto-update only works when running from exe.\n"
                        "Please download the new version from the official website."
                    ))
                    self.status_var.set("Cannot auto-update")
                    self.progress.stop()
                    return
                
                # Lấy đường dẫn đến file exe hiện tại
                exe_path = sys.executable
                self.word_processor.log(f"Current exe path: {exe_path}")
                
                # Lấy đường dẫn đến thư mục chứa exe
                exe_dir = os.path.dirname(exe_path)
                self.word_processor.log(f"Exe directory: {exe_dir}")
                
                # Tạo thư mục tạm thời để lưu trữ các file tạm
                temp_dir = os.path.join(tempfile.gettempdir(), f"AutoOffice_Update_{int(time.time())}")
                os.makedirs(temp_dir, exist_ok=True)
                self.word_processor.log(f"Created temp directory: {temp_dir}")
                
                # Tạo marker cập nhật để đánh dấu quá trình cập nhật đang diễn ra
                marker_path = os.path.join(exe_dir, "updating.marker")
                with open(marker_path, 'w', encoding='ascii') as f:
                    f.write(f"Update in progress: {time.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Tạo file update_log.txt trong thư mục tạm - KHÔNG DÙNG TIẾNG VIỆT
                update_log_path = os.path.join(temp_dir, "update_log.txt")
                with open(update_log_path, 'w', encoding='ascii', errors='replace') as f:
                    f.write(f"Starting update to version {new_version} at {datetime.datetime.now()}\n")
                
                # Tạo bản sao lưu của các file py
                py_files = ['main.py', 'gui.py', 'word_processor.py']
                for py_file in py_files:
                    file_path = os.path.join(exe_dir, py_file)
                    if os.path.exists(file_path):
                        backup_path = os.path.join(temp_dir, f"{py_file}.bak")
                        try:
                            shutil.copy2(file_path, backup_path)
                            self.word_processor.log(f"Created backup of {py_file} at {backup_path}")
                        except Exception as e:
                            self.word_processor.log(f"Could not create backup of {py_file}: {e}")
                
                # Tạo version.json trong thư mục update để dễ dàng kiểm tra phiên bản
                version_json_path = os.path.join(temp_dir, "version.json")
                with open(version_json_path, 'w', encoding='ascii') as f:
                    f.write(f'{{"version": "{new_version}"}}')
                    
                # Thêm file này vào danh sách cần tải
                files_to_download = ['main.py', 'gui.py', 'word_processor.py', 'version.json']
                downloaded_files = []
                
                # Thêm file version.json vào đã tải
                downloaded_files.append('version.json')
                
                # Tải các file mã nguồn còn lại từ GitHub
                for file in files_to_download:
                    # Bỏ qua version.json vì đã tạo ở trên
                    if file == 'version.json':
                        continue
                        
                    url = f"https://raw.githubusercontent.com/{self.github_repo}/main/{file}"
                    self.word_processor.log(f"Downloading file {file} from {url}")
                    
                    # Đường dẫn tạm cho file tải về
                    temp_file_path = os.path.join(temp_dir, file)
                    
                    # Thử tải file với số lần thử lại
                    max_retries = 3
                    retry_count = 0
                    success = False
                    
                    while retry_count < max_retries and not success:
                        try:
                            response = requests.get(url, timeout=10)
                            if response.status_code == 200:
                                with open(temp_file_path, 'wb') as f:
                                    f.write(response.content)
                                self.word_processor.log(f"Downloaded file {file}")
                                # Kiểm tra file đã tải
                                if os.path.exists(temp_file_path) and os.path.getsize(temp_file_path) > 0:
                                    downloaded_files.append(file)
                                    success = True
                                else:
                                    raise Exception(f"File {file} is empty or does not exist")
                            else:
                                raise Exception(f"HTTP error: {response.status_code}")
                        except Exception as e:
                            retry_count += 1
                            self.word_processor.log(f"Error downloading {file} (attempt {retry_count}): {e}")
                            if retry_count >= max_retries:
                                self.word_processor.log(f"Could not download {file} after {max_retries} attempts")
                            else:
                                time.sleep(2)  # Đợi 2 giây trước khi thử lại
                
                # Kiểm tra xem tất cả các file có được tải thành công không
                if len(downloaded_files) != len(files_to_download):
                    missing_files = [f for f in files_to_download if f not in downloaded_files]
                    error_msg = f"Could not download files: {', '.join(missing_files)}"
                    self.word_processor.log(error_msg)
                    raise Exception(error_msg)
                
                # Cập nhật file version.txt - không dùng tiếng Việt có dấu
                version_path = os.path.join(exe_dir, "version.txt")
                with open(version_path, 'w', encoding='ascii') as f:
                    f.write(new_version)
                self.word_processor.log(f"Updated version.txt with version {new_version}")
                
                # Thông báo người dùng về việc cập nhật
                result = messagebox.askyesno(
                    "Update Software",
                    f"New version {new_version} is ready to install.\n"
                    "The application will close to perform the update.\n\n"
                    "Do you want to update now?"
                )
                
                if result:
                    # Hẹn giờ đóng ứng dụng và cập nhật
                    self.word_processor.log("User confirmed update, preparing to close application...")
                    
                    # ----NEW: Sửa đổi updater script để đảm bảo sao chép file thành công----
                    # Tạo Python updater script
                    updater_script_path = os.path.join(temp_dir, "updater.py")
                    
                    # Nội dung script cập nhật - KHÔNG DÙNG TIẾNG VIỆT
                    updater_script = f"""# Updater script for AutoOffice
import os
import sys
import time
import shutil
import subprocess
import traceback

def log(message):
    with open("{update_log_path.replace('\\', '\\\\')}", "a", encoding="ascii", errors="replace") as f:
        f.write(f"{{message}}\\n")

log(f"Updater script started at {{time.strftime('%Y-%m-%d %H:%M:%S')}}")

# Wait for application to exit
log("Waiting 5 seconds for application to close...")
time.sleep(5)

# Paths
exe_dir = "{exe_dir.replace('\\', '\\\\')}"
temp_dir = "{temp_dir.replace('\\', '\\\\')}"
exe_path = "{exe_path.replace('\\', '\\\\')}"
marker_path = os.path.join(exe_dir, "updating.marker")

files_to_copy = {files_to_download}
successful = True

try:
    # Kiểm tra nếu folder đích có quyền ghi
    test_file = os.path.join(exe_dir, "test_write_permission.tmp")
    try:
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        log(f"Destination folder has write permission: {{exe_dir}}")
    except Exception as e:
        log(f"ERROR: No write permission to destination folder: {{e}}")
        log(f"Trying to run updater with admin privileges...")
        
        # Thử chạy với quyền admin
        if sys.platform.startswith('win'):
            try:
                import ctypes
                if not ctypes.windll.shell32.IsUserAnAdmin():
                    log("Re-launching updater as Administrator")
                    # Re-run python script with admin rights
                    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
                    sys.exit(0)
            except Exception as e:
                log(f"Failed to elevate privileges: {{e}}")
    
    # Copy files - ENHANCED ERROR HANDLING
    for file in files_to_copy:
        src_path = os.path.join(temp_dir, file)
        dst_path = os.path.join(exe_dir, file)
        log(f"Copying {{file}} to {{dst_path}}...")
        try:
            # Check if source file exists and has content
            if not os.path.exists(src_path):
                log(f"ERROR: Source file does not exist: {{src_path}}")
                successful = False
                continue
                
            if os.path.getsize(src_path) == 0:
                log(f"ERROR: Source file is empty: {{src_path}}")
                successful = False
                continue
                
            # Try to make destination writable if it exists
            if os.path.exists(dst_path):
                try:
                    os.chmod(dst_path, 0o666)  # Make file writable
                    log(f"Made destination file writable: {{dst_path}}")
                except Exception as e:
                    log(f"Warning: Could not change file permissions: {{e}}")
            
            # First try direct copy
            try:
                shutil.copy2(src_path, dst_path)
                log(f"Successfully copied {{file}}")
            except Exception as e:
                log(f"Direct copy failed: {{e}}")
                log("Trying to create intermediate file...")
                
                # Try with intermediate file
                temp_dst = dst_path + ".new"
                try:
                    shutil.copy2(src_path, temp_dst)
                    if os.path.exists(dst_path):
                        os.remove(dst_path)
                    os.rename(temp_dst, dst_path)
                    log(f"Successfully copied {{file}} using intermediate file")
                except Exception as e2:
                    log(f"ERROR: All copy methods failed for {{file}}: {{e2}}")
                    successful = False
            
            # Verify file was copied
            if not os.path.exists(dst_path) or os.path.getsize(dst_path) == 0:
                log(f"ERROR: Verification failed, destination file missing or empty: {{dst_path}}")
                successful = False
            else:
                log(f"Verified file was copied successfully: {{dst_path}}")
                
        except Exception as e:
            log(f"ERROR: Unexpected error copying {{file}}: {{e}}")
            log(traceback.format_exc())
            successful = False
    
    # Remove update marker
    if os.path.exists(marker_path):
        try:
            os.remove(marker_path)
            log("Removed update marker")
        except Exception as e:
            log(f"Error removing update marker: {{e}}")
    
    if successful:
        log("Update completed successfully")
        # Start application with updated flag
        try:
            subprocess.Popen([exe_path, "--updated", "--from-update"])
            log("Application restarted")
        except Exception as e:
            log(f"Error restarting application: {{e}}")
    else:
        log("Update failed, restoring from backup...")
        # Restore from backup
        for file in ['main.py', 'gui.py', 'word_processor.py']:
            backup_path = os.path.join(temp_dir, f"{{file}}.bak")
            if os.path.exists(backup_path):
                dst_path = os.path.join(exe_dir, file)
                try:
                    shutil.copy2(backup_path, dst_path)
                    log(f"Restored {{file}} from backup")
                except Exception as e:
                    log(f"Error restoring {{file}}: {{e}}")
        
        # Start application with restore flag
        try:
            subprocess.Popen([exe_path, "--from-failed-update"])
            log("Application restarted after restore")
        except Exception as e:
            log(f"Error restarting application after restore: {{e}}")
except Exception as e:
    log(f"Unexpected error: {{e}}")
    log(traceback.format_exc())
    # Start application with error flag
    try:
        subprocess.Popen([exe_path, "--from-failed-update"])
        log("Application restarted after error")
    except:
        log("Could not restart application")

# Cleanup temp directory after delay
time.sleep(5)
try:
    shutil.rmtree(temp_dir)
    print("Removed temp directory")
except Exception as e:
    print(f"Could not remove temp directory: {{e}}")
"""

                    # Lưu script cập nhật
                    with open(updater_script_path, 'w', encoding='utf-8') as f:
                        f.write(updater_script)
                    self.word_processor.log(f"Created updater script at {updater_script_path}")
                    
                    # Hàm đóng ứng dụng và chạy updater script
                    def close_and_update():
                        try:
                            self.word_processor.log("Preparing to run updater script...")
                            
                            # Import thư viện cần thiết
                            import subprocess
                            
                            # Sử dụng hàm từ main để lấy đường dẫn ngắn nếu có thể
                            try:
                                from main import get_short_path
                                python_exe = get_short_path(sys.executable)
                                updater_script = get_short_path(updater_script_path)
                            except ImportError:
                                python_exe = sys.executable
                                updater_script = updater_script_path
                            
                            # Tạo command để chạy script Python
                            cmd = [
                                python_exe,  # Sử dụng Python hiện tại
                                updater_script  # Chạy script cập nhật
                            ]
                            
                            self.word_processor.log(f"Running updater with command: {cmd}")
                            
                            # Khởi chạy script cập nhật với cửa sổ ẩn
                            process = subprocess.Popen(
                                cmd,
                                shell=False,
                                creationflags=subprocess.CREATE_NO_WINDOW
                            )
                            
                            self.word_processor.log(f"Started updater process with PID: {process.pid}")
                            self.word_processor.log("Update process started, closing application...")
                            
                            # Đợi 1 giây để đảm bảo script đã được khởi chạy
                            time.sleep(1)
                            
                            # Đóng ứng dụng
                            self.root.quit()
                            
                        except Exception as e:
                            self.word_processor.log(f"Error running updater script: {e}")
                            messagebox.showerror("Update Error", f"An error occurred during update: {e}")
                            self.status_var.set("Update failed")
                            self.progress.stop()
                            
                            # Xóa marker cập nhật nếu có lỗi
                            if os.path.exists(marker_path):
                                try:
                                    os.remove(marker_path)
                                except:
                                    pass
                    
                    # Hẹn giờ đóng ứng dụng và cập nhật
                    self.root.after(1000, close_and_update)
                else:
                    self.word_processor.log("User cancelled update")
                    self.status_var.set("Update cancelled")
                    self.progress.stop()
                    
                    # Xóa marker cập nhật nếu người dùng hủy
                    if os.path.exists(marker_path):
                        try:
                            os.remove(marker_path)
                        except:
                            pass
                
            except Exception as e:
                error_msg = f"Error preparing update: {e}"
                self.word_processor.log(error_msg)
                self.word_processor.log(f"Error details: {traceback.format_exc()}")
                messagebox.showerror("Update Error", error_msg)
                self.status_var.set("Update failed")
                self.progress.stop()
                
                # Xóa marker cập nhật nếu có lỗi
                marker_path = os.path.join(os.path.dirname(sys.executable), "updating.marker")
                if os.path.exists(marker_path):
                    try:
                        os.remove(marker_path)
                    except:
                        pass
        
        # Thực hiện quá trình cập nhật trong một luồng riêng
        self.root.after(0, perform_update)
    
    def browse_file(self):
        """Mở hộp thoại chọn tệp Word"""
        file_path = filedialog.askopenfilename(
            title="Chọn tệp Word",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.selected_file = file_path
            self.file_path_var.set(file_path)
            # Xóa kết quả phân tích cũ
            self.clear_results(False)
    
    def clear_results(self, reset_analysis=True):
        """Xóa kết quả phân tích cũ
        
        Args:
            reset_analysis (bool): Nếu True, sẽ xóa cả dữ liệu phân tích, mặc định là True
        """
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if reset_analysis:
            self.analysis_results = []
            self.status_var.set("Chưa phân tích")
    
    def filter_results(self):
        """Lọc kết quả hiển thị dựa vào checkbox"""
        # Chỉ cập nhật lại kết quả hiển thị dựa trên dữ liệu hiện có
        self.update_analysis_results()
    
    def analyze_document(self):
        """Phân tích tài liệu tìm kiếm trang trắng"""
        if not self.selected_file:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp Word trước!")
            return
        
        # Hiển thị thông tin trạng thái và bắt đầu thanh tiến trình
        self.status_var.set("Đang phân tích...")
        self.progress.start()
        
        # Xóa kết quả cũ trước khi phân tích mới
        self.clear_results(True)
        
        # Log thông tin khi bắt đầu phân tích
        self.word_processor.log(f"Bắt đầu phân tích tài liệu: {self.selected_file}")
        
        # Sử dụng thread để tránh đóng băng giao diện
        def perform_analysis():
            try:
                # Mở tài liệu Word
                result = self.word_processor.open_document(self.selected_file)
                if not result:
                    self.root.after(0, lambda: messagebox.showerror("Lỗi", "Không thể mở tệp Word!"))
                    self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi khi mở tệp"))
                    self.root.after(0, self.refresh_log)
                    return
                
                # Phân tích section breaks
                self.word_processor.log("Đang phân tích cấu trúc section...")
                analysis_results = self.word_processor.analyze_document()
                
                if not analysis_results:
                    self.word_processor.log("Không tìm thấy kết quả phân tích section hoặc có lỗi", error=True)
                    self.root.after(0, lambda: messagebox.showwarning("Cảnh báo", "Không tìm thấy section nào trong tài liệu hoặc có lỗi khi phân tích!"))
                    self.root.after(0, lambda: self.status_var.set("Không tìm thấy section"))
                    self.root.after(0, self.refresh_log)
                    return
                    
                self.word_processor.log(f"Đã phân tích được {len(analysis_results)} section, tiếp tục phát hiện trang trắng...")
                
                # Phát hiện thêm thông tin về trang
                updated_sections, all_pages_info = self.word_processor.detect_blank_pages(analysis_results)
                
                # Lưu thông tin trang để xử lý sau này
                self.analysis_results = updated_sections
                self.all_pages_info = all_pages_info
                
                # Tính số trang nếu all_pages_info tồn tại, nếu không thì để là 0
                pages_count = 0
                blank_pages_count = 0
                
                if all_pages_info is not None:
                    pages_count = len(all_pages_info)
                    # Đếm số trang trắng
                    blank_pages_count = sum(1 for page in all_pages_info if page.get("is_blank", False))
                
                self.word_processor.log(f"Đã nhận kết quả phân tích với {len(self.analysis_results)} section và {pages_count} trang (trong đó có {blank_pages_count} trang trắng)")
                
                # Cập nhật UI từ main thread
                self.root.after(0, self.update_analysis_results)
                # Cập nhật log
                self.root.after(0, self.refresh_log)
                
            except Exception as e:
                self.word_processor.log(f"Lỗi khi phân tích tài liệu: {e}", error=True)
                self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi khi phân tích: {e}"))
                self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi"))
                self.root.after(0, self.refresh_log)
            finally:
                self.root.after(0, self.progress.stop)
        
        threading.Thread(target=perform_analysis, daemon=True).start()
    
    def update_analysis_results(self):
        """Cập nhật kết quả phân tích vào TreeView"""
        # Xóa dữ liệu cũ
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
            
        # Lọc kết quả theo checkbox
        show_only_blank = self.show_only_blank_var.get()
        
        # Cập nhật TreeView với dữ liệu mới
        if hasattr(self, 'all_pages_info') and self.all_pages_info:
            for page in self.all_pages_info:
                page_num = page.get("page_number", "")
                section = page.get("section_number", "")
                is_blank = page.get("is_blank", False)
                reason = page.get("reason", "")
                
                # Lấy trạng thái xử lý từ thuộc tính mới
                processing_status = page.get("processing_status", "Chưa xử lý")
                
                # Nếu đang lọc, chỉ hiển thị trang trắng
                if show_only_blank and not is_blank:
                    continue
                    
                # Chuẩn bị giá trị hiển thị
                status_text = "Trang trắng" if is_blank else "Trang bình thường"
                
                # Xác định tag cho dòng dựa vào trạng thái xử lý
                row_tags = ()
                if processing_status == "Đã xử lý":
                    row_tags = ("processed",)
                elif processing_status == "Lỗi":
                    row_tags = ("failed",)
                elif processing_status == "Đang chờ xử lý":
                    row_tags = ("pending",)
                    
                # Chèn vào TreeView
                self.results_tree.insert(
                    "", "end", values=(page_num, section, status_text, reason, processing_status),
                    tags=row_tags
                )
        
        # Cập nhật thanh trạng thái
        if hasattr(self, 'all_pages_info') and self.all_pages_info:
            total_pages = len(self.all_pages_info)
            if self.show_only_blank_var.get():
                blank_count = sum(1 for page in self.all_pages_info if page.get("is_blank", False))
                self.status_var.set(f"Hiển thị {blank_count} trang trắng")
            else:
                self.status_var.set(f"Đã hiển thị {total_pages} trang")
        else:
            self.status_var.set(f"Sẵn sàng - v{self.current_version}")
    
    def on_item_double_click(self, event):
        """Xử lý sự kiện nhấp đúp vào một mục trong TreeView"""
        # Lấy ID của mục được chọn
        try:
            if not self.results_tree.selection():
                return
                
            item_id = self.results_tree.selection()[0]
            
            # Lấy thông tin chi tiết và hiển thị hộp thoại
            values = self.results_tree.item(item_id, "values")
            if not values or len(values) < 3:
                messagebox.showinfo("Thông báo", "Không có thông tin chi tiết cho mục này.")
                return
                
            try:
                page_num = int(values[0])
                section_num = int(values[1])
            except ValueError:
                self.word_processor.log(f"Lỗi khi chuyển đổi số trang: {values[0]} hoặc section: {values[1]}", error=True)
                messagebox.showinfo("Thông báo", "Số trang hoặc section không hợp lệ")
                return
            
            # Tìm thông tin trang trong all_pages_info
            page_info = None
            for page in self.all_pages_info:
                if page.get("page_number") == page_num:
                    page_info = page
                    break
            
            if not page_info:
                messagebox.showinfo("Thông báo", f"Không tìm thấy thông tin chi tiết cho trang {page_num}.")
                return
                
            # Tìm section tương ứng
            section_info = None
            for section in self.analysis_results:
                if section.get("index") == section_num:
                    section_info = section
                    break
                
            # Tạo thông điệp hiển thị
            message = f"Thông tin chi tiết về Trang {page_num}:\n\n"
            message += f"- Thuộc Section: {section_num}\n"
            message += f"- Trạng thái: {'Trang trắng' if page_info.get('is_blank', False) else 'Trang có nội dung'}\n"
            
            # Thêm thông tin về section break
            message += f"- Chứa Section Break: {'Có' if page_info.get('contains_section_break', False) else 'Không'}\n"
            
            # Thêm thông tin về độ dài văn bản
            if "text_length" in page_info:
                message += f"- Độ dài văn bản: {page_info.get('text_length')} ký tự\n"
            
            # Thêm lý do nếu là trang trắng
            if page_info.get("is_blank", False):
                reasons = []
                if page_info.get("has_only_section_break", False):
                    reasons.append("chỉ chứa section break")
                if page_info.get("has_only_whitespace", False):
                    reasons.append("chỉ chứa khoảng trắng")
                if page_info.get("has_page_break", False):
                    reasons.append("chứa page break")
                if page_info.get("has_blank_pattern", False):
                    reasons.append("chứa văn bản biểu thị trang trắng")
                    
                if reasons:
                    message += f"- Lý do trang trắng: {', '.join(reasons)}\n"
            
            # Thêm thông tin về document_summary nếu có từ section
            if section_info and section_info.get("document_summary"):
                doc_summary = section_info.get("document_summary")
                message += f"\nThông tin tài liệu:\n"
                message += f"- Tổng số trang: {doc_summary.get('total_pages', 'N/A')}\n"
                message += f"- Tổng số trang trắng: {doc_summary.get('total_blank_pages', 'N/A')}\n"
                message += f"- Số section có trang trắng: {doc_summary.get('sections_with_blank_pages', 'N/A')}\n"
            
            # Hiển thị hộp thoại thông tin chi tiết
            messagebox.showinfo(f"Chi tiết Trang {page_num}", message)
            
        except Exception as e:
            error_msg = f"Lỗi khi hiển thị chi tiết trang: {e}"
            self.word_processor.log(error_msg, error=True)
            messagebox.showerror("Lỗi", error_msg)
    
    def advanced_process_with_tracking(self, selected_only=False):
        """
        Xử lý toàn diện tất cả trang trắng theo một quy trình liên tục với theo dõi trạng thái
        
        Quy trình này sẽ thử tất cả các phương pháp theo thứ tự, đánh dấu các trang đã xử lý,
        và tiếp tục cho đến khi tất cả trang trắng được xử lý hoặc đã thử tất cả phương pháp.
        """
        if not self.selected_file or not hasattr(self, 'all_pages_info') or self.all_pages_info is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp và phân tích trước!")
            return
            
        # Khởi tạo danh sách các trang cần xử lý
        pages_to_process = []
        
        # Lấy các trang được chọn trong bảng nếu có
        selected_items = self.results_tree.selection()
        
        if selected_items:
            # Xác nhận từ người dùng
            confirm = messagebox.askyesno(
                "Xác nhận", 
                "Bạn đã chọn một số trang cụ thể. Bạn có muốn chỉ xử lý các trang đã chọn không?\n"
                "(Chọn 'Không' để xử lý tất cả trang trắng)"
            )
            selected_only = confirm
        
        if selected_only and selected_items:
            # Xác định các trang đã chọn
            for item in selected_items:
                values = self.results_tree.item(item, "values")
                try:
                    page_num = int(values[0])
                    for page in self.all_pages_info:
                        if page.get("page_number") == page_num and page.get("is_blank", False):
                            if page.get("processing_status", "Chưa xử lý") != "Đã xử lý":
                                pages_to_process.append(page_num)
                            break
                except (ValueError, IndexError):
                    continue
        else:
            # Lấy tất cả trang trắng chưa xử lý
            for page in self.all_pages_info:
                if page.get("is_blank", False) and page.get("processing_status", "Chưa xử lý") != "Đã xử lý":
                    pages_to_process.append(page.get("page_number"))
                    
        if not pages_to_process:
            messagebox.showinfo("Thông báo", "Không có trang trắng nào cần xử lý!")
            return
            
        # Xác nhận từ người dùng
        confirm = messagebox.askyesno(
            "Xác nhận xử lý tự động", 
            f"Phát hiện {len(pages_to_process)} trang trắng cần xử lý.\n"
            f"Phần mềm sẽ thử tất cả các phương pháp để xử lý triệt để các trang trắng.\n"
            f"Quá trình này có thể mất nhiều thời gian. Tiếp tục?"
        )
        
        if not confirm:
            return
        
        # Cập nhật tất cả trang trong giao diện thành "Đang chờ xử lý"
        for page in self.all_pages_info:
            if page.get("page_number") in pages_to_process:
                page["processing_status"] = "Đang chờ xử lý"
                
        # Cập nhật giao diện
        self.update_analysis_results()
        
        # Bắt đầu xử lý
        self.status_var.set("Đang bắt đầu quy trình xử lý tự động...")
        self.progress.start()
        
        # Định nghĩa callbacks
        def on_page_processed(page_number):
            # Tìm trang trong all_pages_info và cập nhật trạng thái
            for page in self.all_pages_info:
                if page.get("page_number") == page_number:
                    page["processing_status"] = "Đã xử lý"
                    break
            # Cập nhật giao diện
            self.root.after(0, self.update_analysis_results)
            
        def on_page_failed(page_number):
            # Tìm trang trong all_pages_info và cập nhật trạng thái
            for page in self.all_pages_info:
                if page.get("page_number") == page_number:
                    page["processing_status"] = "Lỗi"
                    break
            # Cập nhật giao diện
            self.root.after(0, self.update_analysis_results)
            
        def on_progress(percent, message):
            # Cập nhật thanh tiến trình và thông báo
            self.root.after(0, lambda: self.status_var.set(message))
            
        def on_complete(result):
            # Xử lý kết quả
            pass  # Không cần làm gì vì xử lý kết quả ở thread chính
        
        # Tạo dictionary callbacks
        callbacks = {
            "on_page_processed": on_page_processed,
            "on_page_failed": on_page_failed,
            "on_progress": on_progress,
            "on_complete": on_complete
        }
        
        # Sử dụng thread để tránh đóng băng giao diện
        def perform_comprehensive_processing():
            try:
                # Tạo đường dẫn file kết quả
                file_dir = os.path.dirname(self.selected_file)
                file_name = os.path.basename(self.selected_file)
                base_name, ext = os.path.splitext(file_name)
                output_path = os.path.join(file_dir, f"{base_name}_processed_advanced{ext}")
                
                # Gọi phương thức advanced_process_with_tracking mới
                result = self.word_processor.advanced_process_with_tracking(
                    self.selected_file,
                    marked_pages=pages_to_process,
                    output_path=output_path,
                    callbacks=callbacks
                )
                
                # Xử lý kết quả
                if result["success"]:
                    # Hiển thị thông báo thành công
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Hoàn tất", 
                        f"Đã xử lý thành công tất cả {len(result['processed_pages'])} trang trắng!\n"
                        f"Số lần thử: {result['attempts']}\n"
                        f"Tệp đã được lưu tại:\n{result['path']}"
                    ))
                    self.root.after(0, lambda: self.status_var.set(f"Đã xử lý thành công {len(result['processed_pages'])} trang trắng"))
                else:
                    # Hiển thị thông báo xử lý một phần
                    if result["processed_pages"]:
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Hoàn tất một phần", 
                            f"Đã xử lý {len(result['processed_pages'])}/{len(result['processed_pages']) + len(result['remaining_pages'])} trang trắng.\n"
                            f"Còn {len(result['remaining_pages'])} trang không thể xử lý được bằng tất cả các phương pháp.\n"
                            f"Số lần thử: {result['attempts']}\n"
                            f"Tệp đã được lưu tại:\n{result['path']}"
                        ))
                        self.root.after(0, lambda: self.status_var.set(f"Đã xử lý một phần: {len(result['processed_pages'])}/{len(pages_to_process)} trang"))
                    else:
                        self.root.after(0, lambda: messagebox.showerror(
                            "Lỗi", 
                            "Không thể xử lý bất kỳ trang trắng nào. Vui lòng kiểm tra log để biết thêm chi tiết."
                        ))
                        self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi khi xử lý"))
                        self.root.after(0, lambda: self.notebook.select(1))  # Chuyển sang tab log
                
            except Exception as e:
                # Cập nhật trạng thái lỗi cho các trang
                for page in self.all_pages_info:
                    if page.get("page_number") in pages_to_process and page.get("processing_status") == "Đang chờ xử lý":
                        page["processing_status"] = "Lỗi"
                        
                self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi khi xử lý tự động: {e}"))
                self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi"))
                self.root.after(0, lambda: self.notebook.select(1))
            finally:
                # Cập nhật giao diện
                self.root.after(0, self.update_analysis_results)
                # Cập nhật log
                self.root.after(0, self.refresh_log)
                self.root.after(0, self.progress.stop)
        
        threading.Thread(target=perform_comprehensive_processing, daemon=True).start()
    
    def remove_blank_pages(self):
        """Xóa tất cả các trang trắng trực tiếp"""
        if not self.selected_file or not hasattr(self, 'all_pages_info') or self.all_pages_info is None:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp và phân tích trước!")
            return
        
        # Tìm các trang trắng
        blank_pages = []
        for page in self.all_pages_info:
            if page.get("is_blank", False):
                blank_pages.append(page["page_number"])
        
        if not blank_pages:
            messagebox.showinfo("Thông báo", "Không tìm thấy trang trắng nào trong tài liệu!")
            return
        
        # Xác nhận từ người dùng
        confirm = messagebox.askyesno(
            "Xác nhận", 
            f"Phát hiện {len(blank_pages)} trang trắng.\nBạn có muốn xóa trực tiếp tất cả các trang trắng không?"
        )
        
        if not confirm:
            return
            
        self.status_var.set("Đang xóa trang trắng...")
        self.progress.start()
        
        # Sử dụng thread để tránh đóng băng giao diện
        def perform_processing():
            try:
                # Gọi phương thức remove_blank_pages mới
                result = self.word_processor.remove_blank_pages(self.selected_file, blank_pages)
                
                if result:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Hoàn tất", 
                        f"Đã xóa thành công {len(blank_pages)} trang trắng!\nTệp đã được lưu tại:\n{result}"
                    ))
                    self.root.after(0, lambda: self.status_var.set("Đã xóa trang trắng thành công"))
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Lỗi", 
                        "Không thể xóa trang trắng. Vui lòng kiểm tra log để biết thêm chi tiết."
                    ))
                    self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi"))
                    self.root.after(0, lambda: self.notebook.select(1))  # Chuyển sang tab log
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Lỗi", f"Lỗi khi xóa trang trắng: {e}"))
                self.root.after(0, lambda: self.status_var.set("Đã xảy ra lỗi"))
                self.root.after(0, lambda: self.notebook.select(1))
            finally:
                # Cập nhật log
                self.root.after(0, self.refresh_log)
                self.root.after(0, self.progress.stop)
        
        threading.Thread(target=perform_processing, daemon=True).start()

    def show_processing_methods_info(self, event):
        """Hiển thị thông tin về các phương pháp xử lý"""
        info = """Các phương pháp xử lý:
        
1. combined: Kết hợp cả hai phương pháp (mặc định)
   - Đầu tiên sửa section break
   - Sau đó xóa trực tiếp trang trắng còn lại
   - Hiệu quả nhất với hầu hết tài liệu

2. section_fix: Chỉ sửa section break
   - Chuyển section break thành Continuous
   - Không xóa trực tiếp trang
   - Hiệu quả với trang trắng do section break

3. page_remove: Chỉ xóa trang trắng
   - Xóa trực tiếp các trang trắng
   - Không thay đổi section break
   - Hiệu quả khi sửa section không khắc phục được"""
           
        toplevel = tk.Toplevel(self.root)
        toplevel.title("Thông tin phương pháp xử lý")
        toplevel.geometry("500x300")
        
        # Đặt cửa sổ mới ở trung tâm cửa sổ cha
        x = self.root.winfo_x() + (self.root.winfo_width() // 2 - 250)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2 - 150)
        toplevel.geometry(f"+{x}+{y}")
        
        # Frame chính
        main_frame = ttk.Frame(toplevel, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        ttk.Label(main_frame, text="Thông tin phương pháp xử lý trang trắng",
                 font=("Arial", 10, "bold")).pack(pady=(0, 10))
        
        # Nội dung thông tin
        info_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD)
        info_text.pack(fill=tk.BOTH, expand=True)
        info_text.insert(tk.END, info)
        info_text.configure(state="disabled")  # Chỉ đọc
        
        # Nút đóng
        ttk.Button(main_frame, text="Đóng", command=toplevel.destroy).pack(pady=(10, 0))
    
    def refresh_log(self):
        """Cập nhật nội dung log"""
        try:
            # Lấy đường dẫn đến file log từ word processor
            log_file = self.word_processor.log_file
            
            # Xóa nội dung hiện tại
            self.log_text.delete(1.0, tk.END)
            
            # Kiểm tra xem file log có tồn tại không
            if os.path.exists(log_file):
                with open(log_file, "r", encoding="utf-8") as f:
                    content = f.read()
                    self.log_text.insert(tk.END, content)
                    # Cuộn xuống cuối
                    self.log_text.see(tk.END)
            else:
                self.log_text.insert(tk.END, "Chưa có file log.")
        except Exception as e:
            self.log_text.insert(tk.END, f"Lỗi khi đọc log: {e}")
    
    def clear_log_file(self):
        """Xóa nội dung file log"""
        try:
            log_file = self.word_processor.log_file
            
            # Xác nhận xóa
            if messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn xóa toàn bộ log?"):
                # Tạo file mới (xóa nội dung cũ)
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write("")
                
                # Ghi một dòng khởi tạo mới
                self.word_processor.log("Đã xóa log và khởi tạo lại")
                
                # Cập nhật hiển thị
                self.refresh_log()
                messagebox.showinfo("Thông báo", "Đã xóa log thành công")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xóa file log: {e}")
    
    def open_log_folder(self):
        """Mở thư mục chứa file log"""
        try:
            log_file = self.word_processor.log_file
            log_dir = os.path.dirname(os.path.abspath(log_file))
            
            # Mở thư mục chứa file log
            os.startfile(log_dir)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở thư mục log: {e}")
    
    def update_title_with_version(self):
        """Cập nhật tiêu đề của cửa sổ để hiển thị phiên bản"""
        self.root.title(f"Ứng dụng xóa trang trắng trong Word - v{self.current_version}")

    def show_status(self, status):
        """Hiển thị trạng thái của ứng dụng"""
        self.status_var.set(status)

    def show_error_dialog(self, message):
        """Hiển thị hộp thoại thông báo lỗi"""
        messagebox.showerror("Lỗi", message)

    def _create_version_json(self):
        """Tạo file version.json trong thư mục hiện tại"""
        try:
            exe_path = sys.executable
            exe_dir = os.path.dirname(exe_path)
            
            version_json_path = os.path.join(exe_dir, "version.json")
            
            # Nếu file đã tồn tại, không cần tạo lại
            if os.path.exists(version_json_path):
                return
                
            # Tạo file version.json đơn giản
            version_data = {
                "version": self.current_version,
                "app_name": "AutoOffice",
                "build_date": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # Lưu file với mã hóa ASCII
            with open(version_json_path, 'w', encoding='ascii') as f:
                import json
                json.dump(version_data, f, ensure_ascii=True)
                
            self.word_processor.log(f"Đã tạo file version.json: {version_json_path}")
        except Exception as e:
            self.word_processor.log(f"Lỗi khi tạo version.json: {e}", error=True) 