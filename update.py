import json
import os
import requests
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class AutoOfficeUpdater:
    def __init__(self, repo_url="https://github.com/truong-29/AutoOffice"):
        self.repo_url = repo_url
        self.api_url = "https://api.github.com/repos/truong-29/AutoOffice/contents/version.json"
        self.current_version = self._get_current_version()
        
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
