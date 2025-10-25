import json
import os
from datetime import datetime
from io import BytesIO

from PIL import Image, ImageDraw
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSystemTrayIcon,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from src.core.backup_manager import BackupManager
from src.core.file_manager import FileManager
from src.ui.backup_dialog import BackupDialog
from src.ui.backup_manager_dialog import BackupManagerDialog


class FileManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BHRM文件管理器")
        self.setGeometry(100, 100, 1200, 800)
        
        # 配置文件路径
        self.config_file = "file_management_config.json"
        
        # 核心管理器
        self.file_manager = FileManager()
        self.backup_manager = BackupManager()
        
        # 存储文件信息
        self.files_tree = {}
        self.selected_files = []
        
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
        
        # 创建文件树状视图
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderLabels(["文件名", "大小", "创建时间", "修改时间", "类型", "路径"])
        self.file_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.file_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.file_tree.customContextMenuRequested.connect(self.open_context_menu)
        self.file_tree.itemClicked.connect(self.on_file_selected)
        self.file_tree.setSortingEnabled(True)
        
        # 连接表头点击信号
        header = self.file_tree.header()
        header.sectionClicked.connect(self.on_header_clicked)
        
        # 初始化排序状态
        self.sort_column = 2  # 默认按创建时间排序
        self.sort_order = Qt.DescendingOrder  # 默认倒序
        
        # 创建操作按钮区域
        button_layout = QHBoxLayout()
        self.backup_btn = QPushButton("设置备份策略")
        self.backup_btn.clicked.connect(self.open_backup_dialog)
        self.details_btn = QPushButton("查看详情")
        self.details_btn.clicked.connect(self.show_file_details)
        
        self.manage_backup_btn = QPushButton("管理策略")
        self.manage_backup_btn.clicked.connect(self.open_backup_manager)
        
        button_layout.addWidget(self.backup_btn)
        button_layout.addWidget(self.details_btn)
        button_layout.addWidget(self.manage_backup_btn)
        button_layout.addStretch()
        
        # 添加到主布局
        main_layout.addLayout(dir_selection_layout)
        main_layout.addWidget(self.file_tree)
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
            
        self.files_tree = self.file_manager.load_files_tree(directory)
        
        # 清空树状视图
        self.file_tree.clear()
        
        # 填充树状视图
        self.populate_tree(self.files_tree, self.file_tree)
        
        # 展开根节点
        self.file_tree.expandAll()
        
        # 设置列宽
        self.file_tree.setColumnWidth(0, 300)  # 文件名
        self.file_tree.setColumnWidth(1, 80)   # 大小
        self.file_tree.setColumnWidth(2, 160)  # 创建时间
        self.file_tree.setColumnWidth(3, 160)  # 修改时间
        self.file_tree.setColumnWidth(4, 80)   # 类型
        self.file_tree.setColumnWidth(5, 300)  # 路径
        
        # 默认按创建时间倒序排列
        self.file_tree.sortItems(2, Qt.DescendingOrder)
        
    def populate_tree(self, node, parent_item):
        """填充树状视图"""
        if node['type'] == 'directory':
            # 创建目录节点
            tree_item = QTreeWidgetItem(parent_item)
            tree_item.setText(0, node['name'])
            tree_item.setText(3, "目录")
            tree_item.setText(4, node['path'])
            tree_item.setFlags(tree_item.flags() | Qt.ItemIsUserCheckable)
            tree_item.setCheckState(0, Qt.Unchecked)
            
            # 递归添加子节点
            for child in node['children']:
                self.populate_tree(child, tree_item)
        else:
            # 创建文件节点
            tree_item = QTreeWidgetItem(parent_item)
            tree_item.setText(0, node['name'])
            tree_item.setText(1, self.file_manager.format_size(node['size']))
            tree_item.setText(2, datetime.fromtimestamp(node['created']).strftime('%Y-%m-%d %H:%M:%S'))
            tree_item.setText(3, datetime.fromtimestamp(node['modified']).strftime('%Y-%m-%d %H:%M:%S'))
            tree_item.setText(4, node['extension'])
            tree_item.setText(5, node['path'])
            tree_item.setFlags(tree_item.flags() | Qt.ItemIsUserCheckable)
            tree_item.setCheckState(0, Qt.Unchecked)
            
    def on_file_selected(self, item, column):
        """处理文件选择事件"""
        if column == 0:  # 只有在第一列点击时才处理选择
            # 检查是否是目录
            is_directory = (item.text(3) == "目录")
            
            if item.checkState(0) == Qt.Checked:
                if is_directory:
                    # 如果是目录，选中所有子项
                    self.select_all_children(item, True)
                else:
                    # 添加到选中列表
                    file_info = {
                        'name': item.text(0),
                        'size': item.text(1),
                        'created': item.text(2),
                        'modified': item.text(3),
                        'type': item.text(4),
                        'path': item.text(5)
                    }
                    if file_info not in self.selected_files:
                        self.selected_files.append(file_info)
            else:
                if is_directory:
                    # 如果是目录，取消选中所有子项
                    self.select_all_children(item, False)
                else:
                    # 从选中列表移除
                    file_info = {
                        'name': item.text(0),
                        'size': item.text(1),
                        'created': item.text(2),
                        'modified': item.text(3),
                        'type': item.text(4),
                        'path': item.text(5)
                    }
                    if file_info in self.selected_files:
                        self.selected_files.remove(file_info)
                     
    def select_all_children(self, item, select):
        """递归选中或取消选中所有子项"""
        # 处理当前项的所有直接子项
        for i in range(item.childCount()):
            child = item.child(i)
            # 设置复选框状态
            if select:
                child.setCheckState(0, Qt.Checked)
                # 如果是文件，添加到选中列表
                if child.text(3) != "目录":
                    file_info = {
                        'name': child.text(0),
                        'size': child.text(1),
                        'created': child.text(2),
                        'modified': child.text(3),
                        'type': child.text(4),
                        'path': child.text(5)
                    }
                    if file_info not in self.selected_files:
                        self.selected_files.append(file_info)
            else:
                child.setCheckState(0, Qt.Unchecked)
                # 如果是文件，从选中列表移除
                if child.text(3) != "目录":
                    file_info = {
                        'name': child.text(0),
                        'size': child.text(1),
                        'created': child.text(2),
                        'modified': child.text(3),
                        'type': child.text(4),
                        'path': child.text(5)
                    }
                    if file_info in self.selected_files:
                        self.selected_files.remove(file_info)
                
            # 如果子项是目录，递归处理其子项
            if child.text(3) == "目录":
                self.select_all_children(child, select)
                    
    def show_file_details(self):
        """显示文件详情"""
        selected_items = self.file_tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先选择一个文件")
            return
            
        item = selected_items[0]
        self.name_label.setText(item.text(0))
        self.size_label.setText(item.text(1))
        self.modified_label.setText(item.text(3))  # 修改时间在第4列
        self.type_label.setText(item.text(4))      # 类型在第5列
        self.path_label.setText(item.text(5))      # 路径在第6列
        
    def open_context_menu(self, position):
        """打开右键菜单"""
        item = self.file_tree.itemAt(position)
        if not item:
            return
            
        menu = QMenu()
        
        # 添加菜单项
        open_action = QAction("打开", self)
        open_action.triggered.connect(lambda: self.open_file(item))
        
        details_action = QAction("查看详情", self)
        details_action.triggered.connect(lambda: self.show_item_details(item))
        
        menu.addAction(open_action)
        menu.addAction(details_action)
        
        # 添加"撤销选择"菜单项
        deselect_action = QAction("撤销选择", self)
        deselect_action.triggered.connect(self.deselect_all)
        menu.addAction(deselect_action)
        
        # 添加"勾选文件"菜单项
        select_files_action = QAction("勾选文件", self)
        select_files_action.triggered.connect(lambda: self.select_files_from_context_menu(item))
        menu.addAction(select_files_action)
        
        # 显示菜单
        menu.exec_(self.file_tree.viewport().mapToGlobal(position))
        
    def open_file(self, item):
        """打开文件"""
        try:
            file_path = item.text(5)  # 路径在第6列
            # 使用系统默认程序打开文件
            import subprocess
            subprocess.Popen(['start', file_path], shell=True)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法打开文件: {e}")
            
    def show_item_details(self, item):
        """显示项目详情"""
        self.name_label.setText(item.text(0))
        self.size_label.setText(item.text(1))
        self.modified_label.setText(item.text(3))  # 修改时间在第4列
        self.type_label.setText(item.text(4))      # 类型在第5列
        self.path_label.setText(item.text(5))      # 路径在第6列
        
    def on_header_clicked(self, column):
        """处理表头点击事件"""
        # 切换排序顺序
        if self.sort_column == column:
            # 如果点击的是当前排序列，则切换排序顺序
            self.sort_order = Qt.AscendingOrder if self.sort_order == Qt.DescendingOrder else Qt.DescendingOrder
        else:
            # 如果点击的是其他列，则默认使用降序排序
            self.sort_column = column
            self.sort_order = Qt.DescendingOrder
            
        # 执行排序
        self.file_tree.sortItems(self.sort_column, self.sort_order)
        
        
    def open_backup_dialog(self):
        """打开备份策略对话框"""
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要备份的文件")
            return
            
        dialog = BackupDialog(self.selected_files, self)
        if dialog.exec_():
            backup_task = dialog.get_backup_task()
            self.backup_manager.add_task(backup_task)
            QMessageBox.information(self, "成功", "备份任务已创建")
            
    def check_backup_tasks(self):
        """检查并执行备份任务"""
        self.backup_manager.execute_tasks()
        
    def closeEvent(self, event):
        """处理窗口关闭事件"""
        # 最小化到系统托盘而不是关闭
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "BHRM文件管理器",
            "程序已在后台运行",
            QSystemTrayIcon.Information,
            2000
        )
        
    def deselect_all(self):
        """撤销所有选择"""
        # 遍历所有顶级项目
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            self.deselect_item_and_children(item)
            
        # 清空选中文件列表
        self.selected_files.clear()
        
    def deselect_item_and_children(self, item):
        """递归取消选中项目及其所有子项"""
        # 取消当前项目的选中状态
        item.setCheckState(0, Qt.Unchecked)
        
        # 递归处理所有子项
        for i in range(item.childCount()):
            child = item.child(i)
            self.deselect_item_and_children(child)
            
    def open_backup_manager(self):
        """打开备份管理对话框"""
        dialog = BackupManagerDialog(self.backup_manager, self)
        dialog.exec_()
        
    def select_files_from_context_menu(self, item):
        """从右键菜单勾选文件"""
        # 获取当前选中的项目
        selected_items = self.file_tree.selectedItems()
        
        # 如果没有选中项目，则勾选右键点击的项目
        if not selected_items:
            selected_items = [item]
        
        # 遍历选中的项目，只勾选文件（不包括目录）
        for selected_item in selected_items:
            if selected_item.text(3) != "目录":  # 不是目录
                selected_item.setCheckState(0, Qt.Checked)
                # 添加到选中列表
                file_info = {
                    'name': selected_item.text(0),
                    'size': selected_item.text(1),
                    'created': selected_item.text(2),
                    'modified': selected_item.text(3),
                    'type': selected_item.text(4),
                    'path': selected_item.text(5)
                }
                if file_info not in self.selected_files:
                    self.selected_files.append(file_info)