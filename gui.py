import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
import os
from PIL import Image, ImageTk
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AutoOfficeGUI:
    def __init__(self, root, word_processor, updater=None):
        self.root = root
        self.word_processor = word_processor
        self.updater = updater
        
        # Thiết lập cửa sổ chính
        self.root.title("Auto Office - Xóa Trang Trắng")
        self.root.geometry("800x600")
        self.root.minsize(800, 600)
        
        # Sử dụng customtkinter appearance
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        
        # Tạo các biến chung
        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Sẵn sàng")
        self.progress_value = tk.DoubleVar(value=0)
        
        # Tạo logo nếu có
        self.logo_image = None
        self.setup_logo()
        
        # Tạo giao diện
        self.create_widgets()
        
        # Kiểm tra cập nhật nếu có
        if self.updater:
            self.check_for_updates()
    
    def setup_logo(self):
        """Thiết lập logo ứng dụng."""
        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logo.png")
            if os.path.exists(logo_path):
                original_image = Image.open(logo_path)
                # Thay đổi kích thước nếu cần
                resized_image = original_image.resize((100, 100), Image.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(resized_image)
                logger.info("Đã tải logo ứng dụng")
            else:
                logger.warning("Không tìm thấy file logo")
        except Exception as e:
            logger.error(f"Lỗi khi tải logo: {e}")
    
    def create_widgets(self):
        """Tạo các widget cho giao diện."""
        # Tạo frame chính
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tạo frame header và logo nếu có
        header_frame = ctk.CTkFrame(main_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        if self.logo_image:
            logo_label = tk.Label(header_frame, image=self.logo_image)
            logo_label.pack(side=tk.LEFT, padx=10)
        
        title_label = ctk.CTkLabel(
            header_frame, 
            text="Ứng Dụng Xóa Trang Trắng Trong Tệp Word",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Frame chọn tệp
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        file_label = ctk.CTkLabel(file_frame, text="Tệp Word:")
        file_label.pack(side=tk.LEFT, padx=10)
        
        file_entry = ctk.CTkEntry(file_frame, textvariable=self.file_path, width=500)
        file_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        browse_button = ctk.CTkButton(file_frame, text="Duyệt...", command=self.browse_file)
        browse_button.pack(side=tk.LEFT, padx=10)
        
        # Frame phân tích và kết quả
        result_frame = ctk.CTkFrame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        control_frame = ctk.CTkFrame(result_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        analyze_button = ctk.CTkButton(control_frame, text="Phân tích tệp", command=self.analyze_document)
        analyze_button.pack(side=tk.LEFT, padx=10)
        
        process_button = ctk.CTkButton(control_frame, text="Xử lý tự động", command=self.process_document)
        process_button.pack(side=tk.LEFT, padx=10)
        
        save_button = ctk.CTkButton(control_frame, text="Lưu tệp", command=self.save_document)
        save_button.pack(side=tk.LEFT, padx=10)
        
        # Khu vực hiển thị kết quả
        self.result_text = tk.Text(result_frame, height=15, wrap=tk.WORD)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = tk.Scrollbar(self.result_text, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        
        # Frame trạng thái và tiến trình
        status_frame = ctk.CTkFrame(main_frame)
        status_frame.pack(fill=tk.X, padx=10, pady=10)
        
        status_label = ctk.CTkLabel(status_frame, textvariable=self.status_text)
        status_label.pack(side=tk.LEFT, padx=10)
        
        progress_bar = ctk.CTkProgressBar(status_frame, variable=self.progress_value)
        progress_bar.pack(side=tk.RIGHT, padx=10, fill=tk.X, expand=True)
    
    def browse_file(self):
        """Mở hộp thoại chọn tệp Word."""
        file_path = filedialog.askopenfilename(
            title="Chọn tệp Word",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.file_path.set(file_path)
            self.status_text.set(f"Đã chọn tệp: {os.path.basename(file_path)}")
            logger.info(f"Đã chọn tệp: {file_path}")
    
    def analyze_document(self):
        """Phân tích tài liệu Word."""
        file_path = self.file_path.get()
        
        if not file_path:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp Word trước!")
            return
        
        self.result_text.delete(1.0, tk.END)
        self.status_text.set("Đang phân tích tệp...")
        self.progress_value.set(0.2)
        
        # Sử dụng thread để không làm treo giao diện
        def analyze_task():
            if self.word_processor.open_document(file_path):
                sections_info = self.word_processor.analyze_document()
                
                if sections_info:
                    self.root.after(0, self.update_analysis_results, sections_info)
                else:
                    self.root.after(0, lambda: self.status_text.set("Không thể phân tích tệp."))
            else:
                self.root.after(0, lambda: self.status_text.set("Không thể mở tệp."))
                self.root.after(0, lambda: messagebox.showerror("Lỗi", "Không thể mở tệp Word."))
        
        thread = threading.Thread(target=analyze_task)
        thread.daemon = True
        thread.start()
    
    def update_analysis_results(self, sections_info):
        """Cập nhật kết quả phân tích."""
        self.result_text.delete(1.0, tk.END)
        
        self.result_text.insert(tk.END, "Kết quả phân tích tài liệu Word:\n\n")
        
        doc_info = self.word_processor.get_document_info()
        if doc_info:
            self.result_text.insert(tk.END, f"Tài liệu có {doc_info['sections']} phần, {doc_info['paragraphs']} đoạn văn, {doc_info['tables']} bảng.\n\n")
        
        self.result_text.insert(tk.END, "Các ngắt phần được tìm thấy:\n")
        
        needs_conversion_count = 0
        
        for section in sections_info:
            section_desc = f"- Phần {section['index'] + 1}: Kiểu: {section['type_name']}"
            
            if section['needs_conversion']:
                section_desc += " (Cần chuyển đổi) ⚠️"
                needs_conversion_count += 1
                
            self.result_text.insert(tk.END, f"{section_desc}\n")
        
        if needs_conversion_count > 0:
            self.result_text.insert(tk.END, f"\nTìm thấy {needs_conversion_count} ngắt phần có thể gây ra trang trắng và cần chuyển đổi.")
            self.status_text.set(f"Phân tích hoàn tất: {needs_conversion_count} ngắt phần cần chuyển đổi")
        else:
            self.result_text.insert(tk.END, "\nKhông tìm thấy ngắt phần nào gây ra trang trắng.")
            self.status_text.set("Phân tích hoàn tất: Không có trang trắng")
        
        self.progress_value.set(1.0)
    
    def process_document(self):
        """Xử lý tài liệu để loại bỏ trang trắng."""
        if not self.word_processor.document:
            messagebox.showwarning("Cảnh báo", "Vui lòng phân tích tệp trước!")
            return
        
        self.status_text.set("Đang xử lý tệp...")
        self.progress_value.set(0.5)
        
        # Sử dụng thread để không làm treo giao diện
        def process_task():
            changes = self.word_processor.fix_empty_pages()
            
            if changes >= 0:
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n\nĐã xử lý {changes} ngắt phần."))
                self.root.after(0, lambda: self.status_text.set(f"Xử lý hoàn tất: Đã thay đổi {changes} ngắt phần"))
            else:
                self.root.after(0, lambda: self.status_text.set("Không thể xử lý tệp."))
                self.root.after(0, lambda: messagebox.showerror("Lỗi", "Không thể xử lý tệp."))
            
            self.root.after(0, lambda: self.progress_value.set(1.0))
        
        thread = threading.Thread(target=process_task)
        thread.daemon = True
        thread.start()
    
    def save_document(self):
        """Lưu tài liệu đã chỉnh sửa."""
        if not self.word_processor.document:
            messagebox.showwarning("Cảnh báo", "Vui lòng xử lý tệp trước!")
            return
        
        # Mở hộp thoại lưu tệp
        initial_file = os.path.basename(self.file_path.get())
        name, ext = os.path.splitext(initial_file)
        save_path = filedialog.asksaveasfilename(
            title="Lưu tệp Word",
            initialfile=f"{name}_fixed{ext}",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        
        if save_path:
            self.status_text.set("Đang lưu tệp...")
            self.progress_value.set(0.8)
            
            # Sử dụng thread để không làm treo giao diện
            def save_task():
                result = self.word_processor.save_document(save_path)
                
                if result:
                    self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n\nĐã lưu tệp vào: {result}"))
                    self.root.after(0, lambda: self.status_text.set(f"Đã lưu tệp thành công"))
                    self.root.after(0, lambda: messagebox.showinfo("Thành công", f"Đã lưu tệp vào:\n{result}"))
                else:
                    self.root.after(0, lambda: self.status_text.set("Không thể lưu tệp."))
                    self.root.after(0, lambda: messagebox.showerror("Lỗi", "Không thể lưu tệp Word."))
                
                self.root.after(0, lambda: self.progress_value.set(1.0))
            
            thread = threading.Thread(target=save_task)
            thread.daemon = True
            thread.start()
    
    def check_for_updates(self):
        """Kiểm tra cập nhật từ updater nếu có."""
        if not self.updater:
            return
            
        def update_task():
            has_update, version = self.updater.check_for_updates()
            
            if has_update:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Cập nhật mới", 
                    f"Có phiên bản mới: {version}\nVui lòng cập nhật để có trải nghiệm tốt nhất."
                ))
        
        thread = threading.Thread(target=update_task)
        thread.daemon = True
        thread.start()
