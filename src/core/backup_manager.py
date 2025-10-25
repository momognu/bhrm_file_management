from datetime import datetime


class BackupManager:
    def __init__(self):
        self.backup_tasks = []
        
    def add_task(self, task):
        """添加备份任务"""
        self.backup_tasks.append(task)
        
    def execute_tasks(self):
        """执行所有备份任务"""
        current_time = datetime.now()
        for task in self.backup_tasks:
            if task.should_backup(current_time):
                task.execute_backup()