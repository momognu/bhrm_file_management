import os
import shutil
from datetime import datetime


class BackupTask:
    def __init__(self, files, backup_dir, start_time, end_time, frequency, name=None):
        self.files = files
        self.backup_dir = backup_dir
        self.start_time = start_time
        self.end_time = end_time
        self.frequency = frequency
        self.name = name if name else f"{datetime.now().strftime('%Y-%m-%d_%H_%M_%S')}_备份策略"
        self.last_backup = None
        
    def should_backup(self, current_time):
        """判断是否应该执行备份"""
        # 检查是否在时间范围内
        if not (self.start_time <= current_time <= self.end_time):
            return False
            
        # 检查是否满足频率要求
        if self.last_backup is None:
            return True
            
        # 根据频率判断
        if self.frequency == "每小时":
            return (current_time - self.last_backup).total_seconds() >= 3600
        elif self.frequency == "每6小时":
            return (current_time - self.last_backup).total_seconds() >= 21600
        elif self.frequency == "每天":
            return (current_time - self.last_backup).total_seconds() >= 86400
        elif self.frequency == "每周":
            return (current_time - self.last_backup).total_seconds() >= 604800
        elif self.frequency == "每月":
            return (current_time - self.last_backup).total_seconds() >= 2592000
            
        return False
        
    def execute_backup(self):
        """执行备份"""
        try:
            # 创建备份目录
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_subdir = os.path.join(self.backup_dir, f"backup_{timestamp}")
            os.makedirs(backup_subdir, exist_ok=True)
            
            # 复制文件
            for i, file_info in enumerate(self.files):
                src_path = file_info['path']
                filename = file_info['name']
                name, ext = os.path.splitext(filename)
                new_filename = f"{name}_{i:03d}{ext}"
                dst_path = os.path.join(backup_subdir, new_filename)
                
                # 确保目标目录存在
                os.makedirs(os.path.dirname(dst_path), exist_ok=True)
                
                # 复制文件
                shutil.copy2(src_path, dst_path)
                
            self.last_backup = datetime.now()
        except Exception as e:
            print(f"备份失败: {e}")