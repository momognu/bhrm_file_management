import json
import os
from datetime import datetime

from PyQt5.QtWidgets import (
    QComboBox,
    QDateTimeEdit,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from src.core.backup_task import BackupTask


class BackupDialog(QDialog):
    def __init__(self, selected_files, parent=None):
        super().__init__(parent)
        self.selected_files = selected_files
        self.backup_config_file = "backup_config.json"
        self.setWindowTitle("设置备份策略")
        self.setModal(True)
        self.resize(400, 300)
        self.create_ui()
        self.load_backup_config()
        
    def create_ui(self):
        """创建对话框界面"""
        layout = QVBoxLayout()
        
        # 备份目录选择
        dir_layout = QHBoxLayout()
        self.backup_dir_label = QLabel("备份到:")
        self.backup_dir_edit = QLineEdit()
        self.backup_dir_browse = QPushButton("浏览")
        self.backup_dir_browse.clicked.connect(self.browse_backup_directory)
        
        dir_layout.addWidget(self.backup_dir_label)
        dir_layout.addWidget(self.backup_dir_edit)
        dir_layout.addWidget(self.backup_dir_browse)
        
        # 时间范围设置
        time_range_group = QGroupBox("备份时间范围")
        time_range_layout = QVBoxLayout()
        
        self.start_time_edit = QDateTimeEdit()
        self.start_time_edit.setDateTime(datetime.now())
        self.end_time_edit = QDateTimeEdit()
        self.end_time_edit.setDateTime(datetime.now())
        
        time_range_layout.addWidget(QLabel("开始时间:"))
        time_range_layout.addWidget(self.start_time_edit)
        time_range_layout.addWidget(QLabel("结束时间:"))
        time_range_layout.addWidget(self.end_time_edit)
        time_range_group.setLayout(time_range_layout)
        
        # 备份频率设置
        frequency_group = QGroupBox("备份频率")
        frequency_layout = QHBoxLayout()
        
        self.frequency_combo = QComboBox()
        self.frequency_combo.addItems(["每小时", "每6小时", "每天", "每周", "每月"])
        
        frequency_layout.addWidget(QLabel("备份间隔:"))
        frequency_layout.addWidget(self.frequency_combo)
        frequency_group.setLayout(frequency_layout)
        
        # 按钮
        button_layout = QHBoxLayout()
        self.ok_btn = QPushButton("确定")
        self.cancel_btn = QPushButton("取消")
        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)
        
        button_layout.addWidget(self.ok_btn)
        button_layout.addWidget(self.cancel_btn)
        
        # 添加到主布局
        layout.addLayout(dir_layout)
        layout.addWidget(time_range_group)
        layout.addWidget(frequency_group)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
    def load_backup_config(self):
        """加载备份配置"""
        try:
            if os.path.exists(self.backup_config_file):
                with open(self.backup_config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'last_backup_directory' in config:
                        self.backup_dir_edit.setText(config['last_backup_directory'])
        except Exception as e:
            print(f"加载备份配置失败: {e}")
            
    def save_backup_config(self):
        """保存备份配置"""
        try:
            config = {
                'last_backup_directory': self.backup_dir_edit.text()
            }
            with open(self.backup_config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存备份配置失败: {e}")
        
    def browse_backup_directory(self):
        """浏览备份目录"""
        directory = QFileDialog.getExistingDirectory(self, "选择备份目录")
        if directory:
            self.backup_dir_edit.setText(directory)
            self.save_backup_config()
            
    def get_backup_task(self):
        """获取备份任务配置"""
        return BackupTask(
            self.selected_files,
            self.backup_dir_edit.text(),
            self.start_time_edit.dateTime().toPyDateTime(),
            self.end_time_edit.dateTime().toPyDateTime(),
            self.frequency_combo.currentText()
        )
        
    def accept(self):
        """确认对话框"""
        if not self.backup_dir_edit.text():
            QMessageBox.warning(self, "警告", "请选择备份目录")
            return
            
        if self.start_time_edit.dateTime() >= self.end_time_edit.dateTime():
            QMessageBox.warning(self, "警告", "开始时间必须早于结束时间")
            return
            
        super().accept()