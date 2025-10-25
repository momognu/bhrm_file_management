import os
import subprocess

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class BackupManagerDialog(QDialog):
    def __init__(self, backup_manager, parent=None):
        super().__init__(parent)
        self.backup_manager = backup_manager
        self.backup_config_file = "backup_config.json"
        self.setWindowTitle("备份策略管理")
        self.setModal(True)
        self.resize(900, 600)
        self.create_ui()
        self.update_task_list()
        
        # 定时更新任务列表
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_task_list)
        self.timer.start(5000)  # 每5秒更新一次
        
    def create_ui(self):
        """创建对话框界面"""
        layout = QVBoxLayout()
        
        # 任务列表
        self.task_table = QTableWidget()
        self.task_table.setColumnCount(5)
        self.task_table.setHorizontalHeaderLabels(["任务名称", "备份目录", "开始时间", "结束时间", "频率"])
        self.task_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 设置为只读
        
        # 设置列宽策略，允许手动调节
        self.task_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        # 设置默认列宽
        self.task_table.setColumnWidth(0, 250)  # 任务名称
        self.task_table.setColumnWidth(1, 220)  # 备份目录
        self.task_table.setColumnWidth(2, 160)  # 开始时间
        self.task_table.setColumnWidth(3, 160)  # 结束时间
        self.task_table.setColumnWidth(4, 100)  # 频率
        
        # 连接双击信号
        self.task_table.cellDoubleClicked.connect(self.on_cell_double_clicked)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("新增任务")
        self.remove_btn = QPushButton("删除任务")
        self.view_backup_btn = QPushButton("查看备份位置")
        self.close_btn = QPushButton("关闭")
        
        self.add_btn.clicked.connect(self.add_task)
        self.remove_btn.clicked.connect(self.remove_task)
        self.view_backup_btn.clicked.connect(self.view_backup_location)
        self.close_btn.clicked.connect(self.accept)
        
        button_layout.addWidget(self.add_btn)
        button_layout.addWidget(self.remove_btn)
        button_layout.addWidget(self.view_backup_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        
        # 添加到主布局
        layout.addWidget(QLabel("备份策略任务列表:"))
        layout.addWidget(self.task_table)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
    def update_task_list(self):
        """更新任务列表"""
        self.task_table.setRowCount(len(self.backup_manager.backup_tasks))
        
        for i, task in enumerate(self.backup_manager.backup_tasks):
            # 任务名称
            name_item = QTableWidgetItem(getattr(task, 'name', f'任务_{i+1}'))
            self.task_table.setItem(i, 0, name_item)
            
            # 备份目录
            dir_item = QTableWidgetItem(task.backup_dir)
            self.task_table.setItem(i, 1, dir_item)
            
            # 开始时间
            start_item = QTableWidgetItem(task.start_time.strftime("%Y-%m-%d %H:%M:%S"))
            self.task_table.setItem(i, 2, start_item)
            
            # 结束时间
            end_item = QTableWidgetItem(task.end_time.strftime("%Y-%m-%d %H:%M:%S"))
            self.task_table.setItem(i, 3, end_item)
            
            # 频率
            freq_item = QTableWidgetItem(task.frequency)
            self.task_table.setItem(i, 4, freq_item)
            
    def add_task(self):
        """新增任务"""
        from src.ui.backup_dialog import BackupDialog
        
        # 获取主窗口引用
        main_window = self.parent()
        if not main_window.selected_files:
            QMessageBox.warning(self, "警告", "请先在主窗口选择要备份的文件")
            return
            
        dialog = BackupDialog(main_window.selected_files, self)
        if dialog.exec_():
            backup_task = dialog.get_backup_task()
            self.backup_manager.add_task(backup_task)
            QMessageBox.information(self, "成功", "备份任务已创建")
            self.update_task_list()
        
    def remove_task(self):
        """删除任务"""
        selected_rows = self.task_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择要删除的任务")
            return
            
        reply = QMessageBox.question(self, "确认", "确定要删除选中的任务吗？", 
                                   QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            # 从后往前删除，避免索引问题
            for index in sorted([row.row() for row in selected_rows], reverse=True):
                del self.backup_manager.backup_tasks[index]
            self.update_task_list()
            
    def view_backup_location(self):
        """查看备份位置"""
        selected_rows = self.task_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择一个任务")
            return
            
        row = selected_rows[0].row()
        if row < len(self.backup_manager.backup_tasks):
            backup_dir = self.backup_manager.backup_tasks[row].backup_dir
            self.open_directory(backup_dir)
            
    def on_cell_double_clicked(self, row, column):
        """处理单元格双击事件"""
        # 如果双击的是备份目录列，则打开对应目录
        if column == 1 and row < len(self.backup_manager.backup_tasks):
            backup_dir = self.backup_manager.backup_tasks[row].backup_dir
            self.open_directory(backup_dir)
            
    def open_directory(self, directory):
        """打开目录"""
        try:
            if os.path.exists(directory):
                subprocess.Popen(['start', directory], shell=True)
            else:
                QMessageBox.warning(self, "错误", f"目录不存在: {directory}")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法打开目录: {e}")