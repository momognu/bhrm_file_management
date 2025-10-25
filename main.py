import json
import os
import shutil
import sys
from datetime import datetime
from io import BytesIO

from PIL import Image, ImageDraw
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QComboBox,
    QDateTimeEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSystemTrayIcon,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class FileManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文件管理系统")
        self.setGeometry(100, 100, 1200, 800)
        
        # 配置文件路径
        self.config_file = "file_management_config.json"
        
        # 存储文件信息
        self.files = []
        self.selected_files = []
        self.backup_tasks = []
        
        # 创建系统托盘图标
        self.create_system_tray()
        
        # 创建主界面
        self.create_main_ui()
        
        # 加载配置
        self.load_config()
        
        # 初始化定时器
        self.backup_timer = QTimer()
        self.backup_timer.timeout.connect(self.check_backup_tasks)
        self.backup_timer.start(60000)  # 每分钟检查一次
        
    def create_system_tray(self):
        """创建系统托盘图标"""
        # 创建一个简单的图标
        image = Image.new('RGB', (64, 64), color=(73, 109, 137))
        # 绘制一个简单的文件图标
        draw = ImageDraw.Draw(image)
        draw.rectangle([10, 10, 54, 54], outline='white', width=2)
        draw.line([10, 10, 30, 10], fill='white', width=2)
        draw.line([30, 10, 30, 20], fill='white', width=2)
        draw.line([10, 54, 10, 40], fill='white', width=2)
        
        # 转换为PyQt图标
        byte_io = BytesIO()
        image.save(byte_io, format='PNG')
        pixmap = QPixmap()
        pixmap.loadFromData(byte_io.getvalue())
        
        self.tray_icon = QSystemTrayIcon(QIcon(pixmap), self)
        
        # 创建托盘菜单
        tray_menu = QMenu()
        restore_action = QAction("恢复窗口", self)
        restore_action.triggered.connect(self.showNormal)
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(QApplication.quit)
        
        tray_menu.addAction(restore_action)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 加载上次选择的目录
                    if 'last_directory' in config:
                        self.dir_path_edit.setText(config['last_directory'])
                        # 自动加载文件
                        if os.path.exists(config['last_directory']):
                            self.load_files()
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            
    def save_config(self):
        """保存配置文件"""
        try:
            config = {
                'last_directory': self.dir_path_edit.text()
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败: {e}")
        
    def create_main_ui(self):
        """创建主界面"""
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建目录选择区域
        dir_selection_layout = QHBoxLayout()
        self.dir_label = QLabel("选择目录:")
        self.dir_path_edit = QLineEdit()
        self.dir_browse_btn = QPushButton("浏览")
        self.dir_browse_btn.clicked.connect(self.browse_directory)
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.load_files)
        
        dir_selection_layout.addWidget(self.dir_label)
        dir_selection_layout.addWidget(self.dir_path_edit)
        dir_selection_layout.addWidget(self.dir_browse_btn)
        dir_selection_layout.addWidget(self.refresh_btn)
        
        # 创建文件表格
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(6)
        self.file_table.setHorizontalHeaderLabels(["文件名", "大小", "修改时间", "类型", "路径", "选择"])
        self.file_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.file_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.file_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.file_table.cellClicked.connect(self.on_file_selected)
        
        # 创建操作按钮区域
        button_layout = QHBoxLayout()
        self.backup_btn = QPushButton("设置备份策略")
        self.backup_btn.clicked.connect(self.open_backup_dialog)
        self.details_btn = QPushButton("查看详情")
        self.details_btn.clicked.connect(self.show_file_details)
        
        button_layout.addWidget(self.backup_btn)
        button_layout.addWidget(self.details_btn)
        button_layout.addStretch()
        
        # 添加到主布局
        main_layout.addLayout(dir_selection_layout)
        main_layout.addWidget(self.file_table)
        main_layout.addLayout(button_layout)
        
        # 创建文件详情区域
        self.create_details_panel(main_layout)
        
    def create_details_panel(self, parent_layout):
        """创建文件详情面板"""
        self.details_group = QGroupBox("文件详情")
        details_layout = QFormLayout()
        
        self.name_label = QLabel()
        self.size_label = QLabel()
        self.modified_label = QLabel()
        self.type_label = QLabel()
        self.path_label = QLabel()
        
        details_layout.addRow("文件名:", self.name_label)
        details_layout.addRow("大小:", self.size_label)
        details_layout.addRow("修改时间:", self.modified_label)
        details_layout.addRow("类型:", self.type_label)
        details_layout.addRow("路径:", self.path_label)
        
        self.details_group.setLayout(details_layout)
        parent_layout.addWidget(self.details_group)
        
    def browse_directory(self):
        """浏览目录"""
        directory = QFileDialog.getExistingDirectory(self, "选择目录")
        if directory:
            self.dir_path_edit.setText(directory)
            self.save_config()  # 保存配置
            self.load_files()
            
    def load_files(self):
        """加载文件列表"""
        directory = self.dir_path_edit.text()
        if not directory or not os.path.exists(directory):
            return
            
        self.files = []
        self.file_table.setRowCount(0)
        
        # 遍历目录及其子目录
        for root, dirs, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    stat = os.stat(file_path)
                    file_info = {
                        'name': file,
                        'size': stat.st_size,
                        'modified': stat.st_mtime,
                        'path': file_path,
                        'type': os.path.splitext(file)[1] if os.path.splitext(file)[1] else '文件'
                    }
                    self.files.append(file_info)
                except Exception as e:
                    print(f"无法读取文件信息 {file_path}: {e}")
        
        # 填充表格
        self.file_table.setRowCount(len(self.files))
        for i, file_info in enumerate(self.files):
            self.file_table.setItem(i, 0, QTableWidgetItem(file_info['name']))
            self.file_table.setItem(i, 1, QTableWidgetItem(self.format_size(file_info['size'])))
            self.file_table.setItem(i, 2, QTableWidgetItem(datetime.fromtimestamp(file_info['modified']).strftime('%Y-%m-%d %H:%M:%S')))
            self.file_table.setItem(i, 3, QTableWidgetItem(file_info['type']))
            self.file_table.setItem(i, 4, QTableWidgetItem(file_info['path']))
            
            # 添加复选框
            checkbox = QTableWidgetItem()
            checkbox.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            checkbox.setCheckState(Qt.Unchecked)
            self.file_table.setItem(i, 5, checkbox)
            
    def format_size(self, size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} PB"
        
    def on_file_selected(self, row, column):
        """处理文件选择事件"""
        if column == 5:  # 选择列
            checkbox = self.file_table.item(row, column)
            file_info = self.files[row]
            
            if checkbox.checkState() == Qt.Checked:
                if file_info not in self.selected_files:
                    self.selected_files.append(file_info)
            else:
                if file_info in self.selected_files:
                    self.selected_files.remove(file_info)
                    
    def show_file_details(self):
        """显示文件详情"""
        selected_rows = self.file_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "警告", "请先选择一个文件")
            return
            
        row = selected_rows[0].row()
        file_info = self.files[row]
        
        self.name_label.setText(file_info['name'])
        self.size_label.setText(self.format_size(file_info['size']))
        self.modified_label.setText(datetime.fromtimestamp(file_info['modified']).strftime('%Y-%m-%d %H:%M:%S'))
        self.type_label.setText(file_info['type'])
        self.path_label.setText(file_info['path'])
        
    def open_backup_dialog(self):
        """打开备份策略对话框"""
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要备份的文件")
            return
            
        dialog = BackupDialog(self.selected_files, self)
        if dialog.exec_():
            backup_task = dialog.get_backup_task()
            self.backup_tasks.append(backup_task)
            QMessageBox.information(self, "成功", "备份任务已创建")
            
    def check_backup_tasks(self):
        """检查并执行备份任务"""
        current_time = datetime.now()
        for task in self.backup_tasks:
            if task.should_backup(current_time):
                task.execute_backup()
                
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        # 最小化到系统托盘而不是关闭
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "文件管理系统",
            "程序已在后台运行",
            QSystemTrayIcon.Information,
            2000
        )

class BackupDialog(QDialog):
    def __init__(self, selected_files, parent=None):
        super().__init__(parent)
        self.selected_files = selected_files
        self.setWindowTitle("设置备份策略")
        self.setModal(True)
        self.resize(400, 300)
        self.create_ui()
        
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
        
    def browse_backup_directory(self):
        """浏览备份目录"""
        directory = QFileDialog.getExistingDirectory(self, "选择备份目录")
        if directory:
            self.backup_dir_edit.setText(directory)
            
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

class BackupTask:
    def __init__(self, files, backup_dir, start_time, end_time, frequency):
        self.files = files
        self.backup_dir = backup_dir
        self.start_time = start_time
        self.end_time = end_time
        self.frequency = frequency
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

def main():
    app = QApplication(sys.argv)
    window = FileManagementApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
