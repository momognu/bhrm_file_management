import os


class FileManager:
    def __init__(self):
        pass
        
    def load_files(self, directory):
        """加载目录中的所有文件"""
        files = []
        
        # 遍历目录及其子目录
        for root, dirs, filenames in os.walk(directory):
            for filename in filenames:
                file_path = os.path.join(root, filename)
                try:
                    stat = os.stat(file_path)
                    file_info = {
                        'name': filename,
                        'size': stat.st_size,
                        'modified': stat.st_mtime,
                        'path': file_path,
                        'type': os.path.splitext(filename)[1] if os.path.splitext(filename)[1] else '文件'
                    }
                    files.append(file_info)
                except Exception as e:
                    print(f"无法读取文件信息 {file_path}: {e}")
                    
        return files
        
    def format_size(self, size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} PB"