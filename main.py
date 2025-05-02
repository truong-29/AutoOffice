import tkinter as tk
from tkinter import messagebox
import os
import sys
import subprocess
import datetime
import time
import traceback
import ctypes
import socket
import tempfile
import winreg
import importlib.util
import atexit
import platform

# Hàm kiểm tra và nhập module động từ file
def import_module_from_file(module_name, file_path):
    """
    Nhập module từ file .py thay vì từ thư viện hệ thống
    
    Args:
        module_name (str): Tên của module
        file_path (str): Đường dẫn đến file .py
        
    Returns:
        module: Module đã nhập hoặc None nếu lỗi
    """
    try:
        # Kiểm tra xem file tồn tại không
        if not os.path.exists(file_path):
            print(f"Không tìm thấy file {file_path}")
            return None
            
        # Tạo đặc tả từ file
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec is None:
            print(f"Không thể tạo spec từ {file_path}")
            return None
            
        # Tạo module từ đặc tả
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        print(f"Đã nhập module {module_name} từ {file_path}")
        return module
    except Exception as e:
        print(f"Lỗi khi nhập module {module_name} từ {file_path}: {e}")
        return None

# Nhập từ file nếu đang chạy từ exe
def import_modules():
    """
    Nhập các module cần thiết từ file .py trong thư mục ứng dụng
    
    Returns:
        tuple: (WordProcessor, WordCleanerApp) hoặc (None, None) nếu lỗi
    """
    # Kiểm tra xem đang chạy từ exe hay không
    is_frozen = getattr(sys, 'frozen', False)
    
    if is_frozen:
        # Lấy thư mục chứa exe
        app_dir = os.path.dirname(sys.executable)
        
        # Nhập WordProcessor từ file
        word_processor_path = os.path.join(app_dir, 'word_processor.py')
        word_processor_module = import_module_from_file('word_processor', word_processor_path)
        
        # Nhập GUI từ file
        gui_path = os.path.join(app_dir, 'gui.py')
        gui_module = import_module_from_file('gui', gui_path)
        
        if word_processor_module and gui_module:
            return (word_processor_module.WordProcessor, gui_module.WordCleanerApp)
        else:
            print("Không thể nhập các module cần thiết từ file")
            return (None, None)
    else:
        # Nếu đang chạy từ mã nguồn, nhập trực tiếp
        try:
            from word_processor import WordProcessor
            from gui import WordCleanerApp
            return (WordProcessor, WordCleanerApp)
        except Exception as e:
            print(f"Lỗi khi nhập module: {e}")
            return (None, None)

# Thiết lập môi trường để tải DLL
def setup_dll_environment():
    """Thiết lập môi trường để tránh lỗi tải DLL"""
    try:
        # Đặt thư mục hiện tại làm thư mục làm việc
        if getattr(sys, 'frozen', False):
            # Nếu đang chạy từ file exe đã đóng gói
            application_path = os.path.dirname(sys.executable)
            os.chdir(application_path)
            
            # Thêm đường dẫn vào PATH
            if application_path not in os.environ['PATH'].split(os.pathsep):
                os.environ['PATH'] = application_path + os.pathsep + os.environ['PATH']
        
        # Đổi tên process cho dễ quản lý
        try:
            if platform.system() == 'Windows':
                ctypes.windll.kernel32.SetConsoleTitleW("AutoOffice")
        except:
            pass
            
        print(f"Cài đặt môi trường: {os.environ['PATH']}")
        print(f"Thư mục hiện tại: {os.getcwd()}")
    except Exception as e:
        print(f"Lỗi khi thiết lập môi trường: {e}")
        print(traceback.format_exc())

# Thực hiện thiết lập môi trường
setup_dll_environment()

# Biến toàn cục để giữ socket
singleton_socket = None

# Cơ chế ngăn chạy nhiều bản sao đồng thời
def prevent_multiple_instances():
    """
    Ngăn chặn nhiều phiên bản của ứng dụng chạy cùng lúc bằng cách tạo socket
    
    Returns:
        bool: True nếu đây là phiên bản duy nhất, False nếu đã có phiên bản khác đang chạy
    """
    global singleton_socket
    
    # Cố gắng bắt một port cố định (có thể thay đổi nếu cần)
    singleton_port = 47852
    
    try:
        # Đóng socket cũ nếu có
        if singleton_socket:
            try:
                singleton_socket.close()
            except:
                pass
            singleton_socket = None
        
        # Thiết lập timeout cho socket để tránh treo
        singleton_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        singleton_socket.settimeout(2)  # 2 giây timeout
        
        # Gắn socket vào localhost với port đã chọn
        singleton_socket.bind(('localhost', singleton_port))
        
        # Nếu bind thành công, đây là phiên bản duy nhất
        return True
    except socket.error:
        # Nếu không bind được, đã có phiên bản khác đang chạy
        try:
            # Thử gửi tin nhắn đến phiên bản đang chạy
            client_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            client_socket.settimeout(1)
            client_socket.sendto(b"CHECK_RUNNING", ('localhost', singleton_port))
        except:
            pass
        finally:
            try:
                client_socket.close()
            except:
                pass
                
        return False
    except Exception as e:
        # Xử lý lỗi khác
        print(f"Lỗi khi kiểm tra phiên bản duy nhất: {e}")
        return True  # Cho phép chạy nếu có lỗi

# Thử đóng các phiên bản khác đang chạy
def try_terminate_other_instances():
    """Thử đóng các phiên bản khác của AutoOffice đang chạy"""
    try:
        # Lấy tên của process hiện tại
        current_exe = os.path.basename(sys.executable)
        current_pid = os.getpid()
        
        print(f"Đang thử đóng các phiên bản khác. Process hiện tại: {current_exe}, PID: {current_pid}")
        
        # Sử dụng tasklist để lấy danh sách các process có tên giống nhau
        output = subprocess.check_output(['tasklist', '/FI', f'IMAGENAME eq {current_exe}', '/NH'], 
                                         shell=True).decode('utf-8')
        
        # Phân tích kết quả để lấy PIDs
        lines = output.strip().split('\n')
        killed_count = 0
        
        for line in lines:
            if not line.strip():
                continue
                
            parts = line.split()
            if len(parts) >= 2:
                try:
                    pid = int(parts[1])
                    # Không kill process hiện tại
                    if pid != current_pid:
                        print(f"Đang thử đóng phiên bản khác (PID: {pid})")
                        subprocess.call(['taskkill', '/F', '/PID', str(pid)], 
                                        shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        killed_count += 1
                except ValueError:
                    continue
                    
        print(f"Đã cố gắng đóng {killed_count} phiên bản khác")
    except Exception as e:
        print(f"Lỗi khi thử đóng các phiên bản khác: {e}")

# Đăng ký ứng dụng vào Startup (tự động khởi động)
def register_startup(register=True):
    """
    Đăng ký hoặc hủy đăng ký ứng dụng khởi động cùng Windows
    
    Args:
        register (bool): True để đăng ký, False để hủy đăng ký
    """
    # Chỉ thực hiện nếu đang chạy từ file exe
    if not getattr(sys, 'frozen', False):
        return
        
    app_name = "AutoOffice"
    exe_path = sys.executable
    
    try:
        # Mở registry key
        registry_key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, 
            r"Software\Microsoft\Windows\CurrentVersion\Run", 
            0, winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE)
        
        if register:
            # Đăng ký ứng dụng vào registry
            winreg.SetValueEx(registry_key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
            print(f"Đã đăng ký '{app_name}' vào Startup")
        else:
            # Hủy đăng ký
            try:
                winreg.DeleteValue(registry_key, app_name)
                print(f"Đã hủy đăng ký '{app_name}' khỏi Startup")
            except:
                # Key không tồn tại
                pass
        
        winreg.CloseKey(registry_key)
    except Exception as e:
        print(f"Lỗi khi thiết lập Startup: {e}")

# Kiểm tra xem ứng dụng có đang chạy với quyền admin không
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Kiểm tra nếu đang trong quá trình cập nhật
def is_updating():
    # Kiểm tra xem có tham số dòng lệnh nào chỉ ra rằng đây là quá trình cập nhật không
    if len(sys.argv) > 1 and sys.argv[1] == "--updating":
        return True
        
    # Kiểm tra xem có file đánh dấu cập nhật không
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
    update_marker = os.path.join(exe_dir, "updating.marker")
    return os.path.exists(update_marker)

# Thiết lập đường dẫn file log
def get_log_path():
    """Tạo đường dẫn đến file log"""
    # Tạo thư mục logs nếu chưa tồn tại
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # Tạo tên file log với thời gian
    current_date = datetime.datetime.now().strftime("%Y%m%d")
    log_file = os.path.join(log_dir, f"word_processor_{current_date}.txt")
    
    return log_file

# Hàm lấy đường dẫn tuyệt đối đến tài nguyên trong gói exe
def resource_path(relative_path):
    """Lấy đường dẫn tuyệt đối đến tài nguyên, hoạt động trong cả môi trường phát triển và PyInstaller"""
    try:
        # PyInstaller tạo một thư mục tạm thời và lưu đường dẫn trong _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# Kiểm tra và cài đặt các thư viện bắt buộc nếu chưa có
def check_and_install_dependencies():
    # Danh sách các thư viện cần thiết
    required_packages = ["python-docx", "docx2python", "pyinstaller"]
    
    # Thêm pywin32 nếu đang chạy trên Windows
    if os.name == 'nt':
        required_packages.append("pywin32")
    
    # Kiểm tra và cài đặt từng thư viện
    for package in required_packages:
        try:
            if package == "python-docx":
                __import__("docx")
            elif package == "pywin32":
                __import__("win32com")
            else:
                __import__(package.replace('-', '_').split('>=')[0])
        except ImportError:
            print(f"Đang cài đặt thư viện {package}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            except subprocess.CalledProcessError:
                print(f"Không thể cài đặt {package}. Vui lòng cài đặt thủ công.")
                sys.exit(1)

# Hàm đặt cửa sổ ở giữa màn hình
def center_window(window, width=300, height=100):
    # Lấy kích thước màn hình
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    
    # Tính toán vị trí x,y
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    # Đặt vị trí cho cửa sổ
    window.geometry(f"{width}x{height}+{x}+{y}")

# Hiển thị màn hình splash đơn giản
def show_splash(root):
    """
    Tạo một màn hình splash đơn giản
    
    Args:
        root (tk.Tk): Cửa sổ gốc Tkinter đã tạo
    """
    # Ẩn cửa sổ gốc trong khi hiển thị splash
    root.withdraw()
    
    # Tạo cửa sổ splash
    splash = tk.Toplevel(root)
    splash.title("")
    splash.overrideredirect(True)  # Không hiển thị thanh tiêu đề
    
    # Căn giữa cửa sổ
    center_window(splash, 300, 100)
    
    # Ngăn người dùng đóng cửa sổ
    splash.protocol("WM_DELETE_WINDOW", lambda: None)
    
    # Frame chính
    frame = tk.Frame(splash, bg="#f0f0f0", borderwidth=2, relief="groove")
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Tiêu đề
    tk.Label(frame, text="Ứng dụng xóa trang trắng trong Word", 
             font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(15, 5))
    
    # Thông báo
    message = tk.Label(frame, text="Đang khởi động...", bg="#f0f0f0")
    message.pack(pady=5)
    
    # Thanh tiến trình
    progress = tk.Canvas(frame, width=200, height=10, bg="white", highlightthickness=1)
    progress.pack(pady=5)
    progress.create_rectangle(0, 0, 0, 10, fill="#4CAF50", width=0, tags="progress")
    
    # Đặt lại foreground color của splash
    splash.attributes('-topmost', True)
    
    # Đảm bảo cập nhật giao diện
    splash.update()
    
    return splash, message, progress

# Hàm cập nhật splash
def update_splash(splash, message, progress, value, text=None):
    # Cập nhật thanh tiến trình
    progress.delete("progress")
    progress.create_rectangle(0, 0, 2 * value, 10, fill="#4CAF50", width=0, tags="progress")
    
    # Cập nhật thông báo nếu có
    if text:
        message.config(text=text)
    
    # Cập nhật giao diện
    splash.update()
    
    # Sleep để hiển thị animation
    time.sleep(0.02)

# Xóa file và thư mục tạm nếu tồn tại
def cleanup_temp_files():
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        
        # Xóa marker file nếu tồn tại
        update_marker = os.path.join(exe_dir, "updating.marker")
        if os.path.exists(update_marker):
            os.remove(update_marker)
            print(f"Đã xóa file đánh dấu cập nhật: {update_marker}")
        
        # Xóa thư mục temp_update nếu tồn tại
        temp_update_dir = os.path.join(exe_dir, "temp_update")
        if os.path.exists(temp_update_dir):
            import shutil
            shutil.rmtree(temp_update_dir)
            print(f"Đã xóa thư mục cập nhật tạm thời: {temp_update_dir}")
        
        # Xóa file backup cũ nếu cần
        backup_file = os.path.join(exe_dir, "AutoOffice_backup.exe")
        if os.path.exists(backup_file):
            try:
                os.remove(backup_file)
                print(f"Đã xóa file backup cũ: {backup_file}")
            except:
                pass
                
        # Xóa tất cả các tệp tạm thời từ quá trình cập nhật
        for temp_file in os.listdir(exe_dir):
            if temp_file.endswith(".new") or temp_file.endswith(".tmp") or temp_file.endswith(".old"):
                try:
                    full_path = os.path.join(exe_dir, temp_file)
                    os.remove(full_path)
                    print(f"Đã xóa tệp tạm thời: {full_path}")
                except Exception as e:
                    print(f"Không thể xóa tệp tạm thời {temp_file}: {e}")
                    
        # Xóa tất cả các thư mục tạm thời cập nhật trong thư mục temp
        temp_dir = tempfile.gettempdir()
        for item in os.listdir(temp_dir):
            if item.startswith("AutoOffice_Update_"):
                try:
                    full_path = os.path.join(temp_dir, item)
                    if os.path.isdir(full_path):
                        shutil.rmtree(full_path)
                    else:
                        os.remove(full_path)
                    print(f"Đã xóa tệp/thư mục tạm thời trong temp: {full_path}")
                except Exception as e:
                    print(f"Không thể xóa tệp/thư mục tạm thời {item}: {e}")
    except Exception as e:
        print(f"Lỗi khi dọn dẹp file tạm: {e}")

# Kiểm tra và xác minh cập nhật thành công
def verify_update_success():
    """Kiểm tra xem quá trình cập nhật đã thành công hay không sau khi khởi động lại"""
    try:
        if '--updated' not in sys.argv and '--from-update' not in sys.argv:
            return False  # Không phải khởi động sau cập nhật
            
        exe_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
        
        # Đọc nhật ký cập nhật nếu có
        update_logs = []
        temp_dir = tempfile.gettempdir()
        for item in os.listdir(temp_dir):
            if item.startswith("AutoOffice_Update_") and os.path.isdir(os.path.join(temp_dir, item)):
                log_file = os.path.join(temp_dir, item, "update_log.txt")
                if os.path.exists(log_file):
                    try:
                        with open(log_file, "r", encoding="ascii", errors="replace") as f:
                            update_logs.append(f.read())
                    except:
                        pass
        
        if update_logs:
            print("Đọc được nhật ký cập nhật:")
            for idx, log in enumerate(update_logs):
                print(f"--- Nhật ký {idx+1} ---")
                print(log[:500] + "..." if len(log) > 500 else log)  # Chỉ hiển thị 500 ký tự đầu tiên
                
        # Kiểm tra các tệp Python cần thiết
        required_files = ["main.py", "gui.py", "word_processor.py"]
        missing_files = []
        
        for file in required_files:
            file_path = os.path.join(exe_dir, file)
            if not os.path.exists(file_path):
                missing_files.append(file)
            elif os.path.getsize(file_path) == 0:
                missing_files.append(f"{file} (empty)")
        
        if missing_files:
            print(f"CẢNH BÁO: Thiếu hoặc lỗi các tệp sau khi cập nhật: {missing_files}")
            return False
        
        print("Xác minh cập nhật thành công: Tất cả các tệp đều tồn tại và không rỗng")
        return True
    except Exception as e:
        print(f"Lỗi khi xác minh cập nhật: {e}")
        return False

# Tạo GUI ứng dụng
def create_gui():
    """
    Tạo giao diện người dùng
    
    Returns:
        tuple: (tk.Tk, WordProcessor, WordCleanerApp) hoặc None nếu lỗi
    """
    # Nhập các module cần thiết
    WordProcessor, WordCleanerApp = import_modules()
    
    if not WordProcessor or not WordCleanerApp:
        # Hiển thị thông báo lỗi và thoát
        messagebox.showerror(
            "Lỗi khởi động",
            "Không thể nhập các module cần thiết. Hãy đảm bảo các file .py đã được cài đặt đúng."
        )
        return None
    
    # Tạo cửa sổ chính
    root = tk.Tk()
    root.title("AutoOffice - Tự động xử lý trang trắng trong file Word")
    
    # Thiết lập kích thước cửa sổ chính
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    
    # Thiết lập biểu tượng nếu có file
    try:
        # Kiểm tra xem ứng dụng đã được đóng gói chưa
        is_frozen = getattr(sys, 'frozen', False)
        
        # Lấy đường dẫn đến thư mục ứng dụng
        if is_frozen:
            app_dir = os.path.dirname(sys.executable)
            # Thử sử dụng Logo.ico trước
            icon_path = os.path.join(app_dir, "Logo.ico")
            if not os.path.exists(icon_path):
                # Nếu không có Logo.ico, sử dụng Logo.png
                icon_path = os.path.join(app_dir, "Logo.png")
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
            # Thử sử dụng Logo.ico trước
            icon_path = os.path.join(app_dir, "Logo.ico")
            if not os.path.exists(icon_path):
                # Nếu không có Logo.ico, sử dụng Logo.png
                icon_path = os.path.join(app_dir, "Logo.png")
        
        # Kiểm tra xem file có tồn tại không
        if os.path.exists(icon_path):
            # Nếu là file .ico, sử dụng trực tiếp
            if icon_path.lower().endswith('.ico'):
                root.iconbitmap(icon_path)
            # Nếu là file .png, tạo photo image và đặt làm icon
            elif icon_path.lower().endswith('.png'):
                icon_image = tk.PhotoImage(file=icon_path)
                root.iconphoto(True, icon_image)
                # Lưu biến vào root để tránh bị thu hồi bởi garbage collector
                root.icon_image = icon_image
    except Exception as e:
        print(f"Không thể thiết lập biểu tượng: {e}")
    
    # Tạo đối tượng WordProcessor
    word_processor = WordProcessor(log_file=get_log_path())
    
    # Tạo giao diện
    app = WordCleanerApp(root, word_processor)
    
    return root, word_processor, app

def cleanup_resources():
    """Giải phóng tài nguyên khi ứng dụng kết thúc"""
    global singleton_socket
    if singleton_socket:
        try:
            print("Đóng socket khi thoát ứng dụng")
            singleton_socket.close()
        except:
            pass
        singleton_socket = None

# Hàm tạo file batch an toàn (không dùng Unicode)
def create_safe_batch_file(content, file_path, encoding='ascii'):
    """
    Tạo file batch với mã hóa an toàn, loại bỏ ký tự Unicode nếu cần
    
    Args:
        content (str): Nội dung batch script
        file_path (str): Đường dẫn đến file batch
        encoding (str): Mã hóa sử dụng, mặc định là ascii
        
    Returns:
        bool: True nếu thành công, False nếu thất bại
    """
    try:
        # Loại bỏ dấu tiếng Việt nếu có
        import unicodedata
        normalized = unicodedata.normalize('NFKD', content)
        ascii_content = ''.join([c for c in normalized if not unicodedata.combining(c)])
        
        # Thay thế các ký tự không phải ASCII
        safe_content = ''
        for char in ascii_content:
            if ord(char) < 128:  # Chỉ giữ lại ký tự ASCII
                safe_content += char
            else:
                safe_content += '?'  # Thay thế bằng dấu hỏi
                
        # Ghi file với mã hóa ascii
        with open(file_path, 'w', encoding=encoding, errors='replace') as f:
            f.write(safe_content)
            
        return True
    except Exception as e:
        print(f"Lỗi khi tạo file batch an toàn: {e}")
        return False

# Hàm lấy đường dẫn ngắn Windows 8.3
def get_short_path(long_path):
    """
    Chuyển đổi đường dẫn thành định dạng ngắn 8.3 của Windows
    
    Args:
        long_path (str): Đường dẫn dài
        
    Returns:
        str: Đường dẫn ngắn định dạng 8.3, hoặc đường dẫn ban đầu nếu không thể chuyển đổi
    """
    try:
        if os.name != 'nt':  # Chỉ hoạt động trên Windows
            return long_path
            
        import win32api
        return win32api.GetShortPathName(long_path)
    except:
        return long_path

# Hàm mã hóa chuỗi với xử lý các ký tự tiếng Việt
def encode_vietnamese_safe(text, encoding='utf-8'):
    """
    Mã hóa chuỗi với xử lý đặc biệt cho tiếng Việt
    
    Args:
        text (str): Chuỗi cần mã hóa
        encoding (str): Mã hóa được sử dụng, mặc định là utf-8
        
    Returns:
        bytes: Chuỗi đã được mã hóa
    """
    try:
        # Thử mã hóa trực tiếp
        return text.encode(encoding)
    except UnicodeEncodeError:
        # Nếu lỗi, thử các phương pháp khác
        try:
            # Thử với 'replace'
            return text.encode(encoding, 'replace')
        except:
            # Loại bỏ dấu tiếng Việt
            import unicodedata
            normalized = unicodedata.normalize('NFKD', text)
            ascii_text = ''.join([c for c in normalized if not unicodedata.combining(c)])
            return ascii_text.encode(encoding, 'replace')

# Hàm giải mã chuỗi với xử lý lỗi
def decode_vietnamese_safe(binary_data, encodings=['utf-8', 'cp1252', 'ascii']):
    """
    Giải mã chuỗi với thử nhiều kiểu mã hóa
    
    Args:
        binary_data (bytes): Dữ liệu nhị phân cần giải mã
        encodings (list): Danh sách các mã hóa cần thử
        
    Returns:
        str: Chuỗi đã giải mã
    """
    for encoding in encodings:
        try:
            return binary_data.decode(encoding)
        except UnicodeDecodeError:
            continue
    
    # Nếu tất cả đều thất bại, sử dụng mã hóa có khả năng xử lý lỗi cao nhất
    return binary_data.decode('utf-8', 'replace')

# Hàm chính
def main():
    """Hàm chính điều khiển luồng chạy ứng dụng"""
    # Đăng ký hàm cleanup khi thoát
    atexit.register(cleanup_resources)
    
    # Kiểm tra xem có đang cập nhật không
    is_update_related = False
    if any(arg in sys.argv for arg in ['--updated', '--from-update', '--restore', '--from-failed-update']):
        print("Ứng dụng vừa được cập nhật hoặc phục hồi!")
        is_update_related = True
        
    # Kiểm tra nhiều bản sao
    if not prevent_multiple_instances():
        # Hiển thị thông báo và thoát
        messagebox.showinfo("Thông báo", "Ứng dụng đã đang chạy")
        sys.exit(0)
        
    try:
        # Dọn dẹp file tạm
        cleanup_temp_files()
        
        # Kiểm tra và cài đặt các thư viện cần thiết
        check_and_install_dependencies()
        
        # Tạo thư mục logs nếu chưa tồn tại
        if not os.path.exists('logs'):
            os.makedirs('logs')
            
        # Hiển thị splash screen
        root = tk.Tk()
        splash, splash_message, splash_progress = show_splash(root)
        
        # Cập nhật thanh tiến trình
        update_splash(splash, splash_message, splash_progress, 20, "Đang khởi tạo...")
        
        # Tạo giao diện
        update_splash(splash, splash_message, splash_progress, 40, "Đang tải các thành phần...")
        
        # Đóng splash trước khi tạo GUI chính
        update_splash(splash, splash_message, splash_progress, 90, "Đang hoàn tất...")
        root.destroy()
        
        # Xác minh cập nhật nếu cần
        if is_update_related:
            update_verified = verify_update_success()
            if not update_verified and '--from-failed-update' not in sys.argv:
                # Nếu quá trình cập nhật không thành công và không phải là chế độ phục hồi
                messagebox.showerror(
                    "Lỗi cập nhật",
                    "Phát hiện lỗi trong quá trình cập nhật: Một số tệp không được tải xuống chính xác.\n"
                    "Vui lòng thử cập nhật lại."
                )
        
        # Tạo GUI chính
        result = create_gui()
        if result is None:
            return
            
        main_root, word_processor, app = result
        
        # Kiểm tra xem vừa cập nhật hay không
        is_updated = '--updated' in sys.argv or '--from-update' in sys.argv
        is_restore = '--restore' in sys.argv or '--from-failed-update' in sys.argv
        
        # Thiết lập thông báo nếu vừa cập nhật
        if is_updated:
            word_processor.log("Ứng dụng vừa được cập nhật thành công")
            main_root.after(1000, lambda: messagebox.showinfo("Cập nhật thành công", "Ứng dụng đã được cập nhật lên phiên bản mới!"))
        
        # Thiết lập thông báo nếu vừa khôi phục sau lỗi cập nhật
        if is_restore:
            word_processor.log("Ứng dụng vừa được khôi phục sau lỗi cập nhật")
            main_root.after(1000, lambda: messagebox.showwarning("Khôi phục sau lỗi cập nhật", 
                                                            "Quá trình cập nhật gặp lỗi và ứng dụng đã được khôi phục từ bản sao lưu!\n\n"
                                                            "Vui lòng thử cập nhật lại sau."))
        
        # Khởi động vòng lặp chính
        main_root.mainloop()
    except Exception as e:
        # Ghi log lỗi và hiển thị thông báo
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        error_msg = f"[{timestamp}] [ERROR] Lỗi khởi động ứng dụng: {e}\n"
        
        # Ghi vào file log
        log_file = get_log_path()
        try:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(error_msg)
                f.write(traceback.format_exc())
        except:
            pass
            
        # Hiển thị hộp thoại lỗi
        messagebox.showerror("Lỗi khởi động", f"Đã xảy ra lỗi khi khởi động ứng dụng:\n{e}")

if __name__ == "__main__":
    # Kiểm tra xem có phiên bản khác đang chạy không
    instance_is_unique = prevent_multiple_instances()
    
    # Kiểm tra các tham số dòng lệnh
    is_updating = '--updating' in sys.argv
    is_updated = '--updated' in sys.argv
    is_restore = '--restore' in sys.argv
    is_from_update = '--from-update' in sys.argv
    is_from_failed_update = '--from-failed-update' in sys.argv
    
    if is_updating or is_updated or is_restore:
        # Nếu đang trong quá trình cập nhật, đóng các phiên bản cũ
        try_terminate_other_instances()
    
    # Ghi log các tham số khởi động
    print(f"Tham số khởi động: {sys.argv}")
    print(f"Phiên bản duy nhất: {instance_is_unique}")
    print(f"Đang cập nhật: {is_updating}")
    print(f"Đã cập nhật: {is_updated}")
    print(f"Khôi phục: {is_restore}")
    print(f"Từ quá trình cập nhật: {is_from_update}")
    print(f"Từ quá trình cập nhật thất bại: {is_from_failed_update}")
    
    # Nếu đang trong quá trình cập nhật hoặc đã có thông báo đang chạy phiên bản khác
    if not instance_is_unique and not (is_updating or is_updated or is_restore or is_from_update or is_from_failed_update):
        # Hiển thị thông báo nếu đây không phải là tình huống cập nhật
        messagebox.showinfo("Thông báo", "Một phiên bản khác của ứng dụng đang chạy!")
        sys.exit(0)  # Thoát mà không hiển thị thông báo
    
    # Khởi động ứng dụng
    main()
