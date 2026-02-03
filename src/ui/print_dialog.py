import os
import subprocess
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QMessageBox,
    QProgressBar,
    QGroupBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


class PrintThread(QThread):
    """打印线程，用于后台执行打印任务"""

    progress = pyqtSignal(int, str)  # 进度信号: (当前数量, 总数)
    finished = pyqtSignal(bool, str)  # 完成信号: (是否成功, 消息)

    def __init__(self, files, printer_name):
        super().__init__()
        self.files = files
        self.printer_name = printer_name

    def run(self):
        """执行打印任务"""
        success_count = 0
        failed_files = []
        total = len(self.files)

        for i, file_info in enumerate(self.files):
            file_path = file_info["path"]

            # 更新进度
            self.progress.emit(i + 1, f"正在打印: {file_info['name']}")

            try:
                # 检查文件是否存在
                if not os.path.exists(file_path):
                    failed_files.append(f"{file_info['name']} (文件不存在)")
                    continue

                # 使用Windows默认打印命令打印文件
                # 根据文件类型选择不同的打印方法
                ext = file_info["type"].lower()

                if ext in [
                    ".pdf",
                    ".doc",
                    ".docx",
                    ".xls",
                    ".xlsx",
                    ".ppt",
                    ".pptx",
                    ".txt",
                ]:
                    # 使用默认程序打印
                    if self.printer_name:
                        # 指定打印机
                        subprocess.run(
                            [
                                "powershell",
                                "-Command",
                                f'Start-Process -FilePath "{file_path}" -ArgumentList "/p" -Wait',
                            ],
                            check=True,
                            timeout=30,
                        )
                    else:
                        # 使用默认打印机
                        subprocess.run(
                            [
                                "powershell",
                                "-Command",
                                f'Start-Process -FilePath "{file_path}" -ArgumentList "/p" -Wait',
                            ],
                            check=True,
                            timeout=30,
                        )
                    success_count += 1
                elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
                    # 图片文件使用Windows照片查看器打印
                    subprocess.run(
                        [
                            "rundll32.exe",
                            "shimgvw.dll,ImageView_PrintTo",
                            f'/pt "{file_path}" "{self.printer_name}"'
                            if self.printer_name
                            else f'"{file_path}"',
                        ],
                        check=True,
                        timeout=30,
                    )
                    success_count += 1
                else:
                    # 其他文件类型尝试使用默认打印方法
                    try:
                        os.startfile(file_path, "print")
                        success_count += 1
                    except Exception as e:
                        failed_files.append(f"{file_info['name']} (不支持)")

            except Exception as e:
                failed_files.append(f"{file_info['name']} ({str(e)})")

        # 发送完成信号
        if success_count == total:
            self.finished.emit(True, f"成功打印 {success_count} 个文件")
        elif success_count > 0:
            self.finished.emit(
                True,
                f"成功打印 {success_count}/{total} 个文件\n失败: {', '.join(failed_files)}",
            )
        else:
            self.finished.emit(False, f"打印失败\n{', '.join(failed_files)}")


class PrintDialog(QDialog):
    """打印机选择和打印对话框"""

    def __init__(self, selected_files, parent=None):
        super().__init__(parent)
        self.selected_files = selected_files
        self.print_thread = None
        self.init_ui()
        self.load_printers()

    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("批量打印")
        self.setMinimumWidth(500)
        self.setModal(True)

        layout = QVBoxLayout()

        # 文件信息
        info_group = QGroupBox("打印信息")
        info_layout = QVBoxLayout()
        file_count_label = QLabel(f"已选择 {len(self.selected_files)} 个文件")
        info_layout.addWidget(file_count_label)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 打印机选择
        printer_group = QGroupBox("选择打印机")
        printer_layout = QVBoxLayout()

        printer_label_layout = QHBoxLayout()
        printer_label = QLabel("打印机:")
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(300)
        printer_label_layout.addWidget(printer_label)
        printer_label_layout.addWidget(self.printer_combo)
        printer_label_layout.addStretch()

        printer_layout.addLayout(printer_label_layout)
        printer_group.setLayout(printer_layout)
        layout.addWidget(printer_group)

        # 进度条
        self.progress_group = QGroupBox("打印进度")
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(len(self.selected_files))
        self.progress_bar.setValue(0)
        self.status_label = QLabel("准备打印...")
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_label)
        self.progress_group.setLayout(progress_layout)
        self.progress_group.setEnabled(False)
        layout.addWidget(self.progress_group)

        # 按钮
        button_layout = QHBoxLayout()
        self.print_btn = QPushButton("开始打印")
        self.print_btn.clicked.connect(self.start_print)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)

        button_layout.addStretch()
        button_layout.addWidget(self.print_btn)
        button_layout.addWidget(self.cancel_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def load_printers(self):
        """加载可用打印机列表"""
        try:
            # 使用PowerShell获取打印机列表
            result = subprocess.run(
                [
                    "powershell",
                    "-Command",
                    "Get-Printer | Select-Object -ExpandProperty Name",
                ],
                capture_output=True,
                text=True,
                check=True,
            )

            printers = [
                line.strip() for line in result.stdout.split("\n") if line.strip()
            ]

            if printers:
                self.printer_combo.addItems(printers)
                # 添加"默认打印机"选项
                self.printer_combo.insertItem(0, "默认打印机")
                self.printer_combo.setCurrentIndex(0)
            else:
                self.printer_combo.addItem("未检测到打印机")

        except Exception as e:
            self.printer_combo.addItem("无法获取打印机列表")
            print(f"获取打印机列表失败: {e}")

    def start_print(self):
        """开始打印"""
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "没有选择要打印的文件")
            return

        # 获取选择的打印机
        printer_index = self.printer_combo.currentIndex()
        if printer_index == 0:
            printer_name = None  # 使用默认打印机
        else:
            printer_name = self.printer_combo.currentText()

        # 禁用打印按钮和打印机选择
        self.print_btn.setEnabled(False)
        self.printer_combo.setEnabled(False)
        self.progress_group.setEnabled(True)

        # 创建并启动打印线程
        self.print_thread = PrintThread(self.selected_files, printer_name)
        self.print_thread.progress.connect(self.update_progress)
        self.print_thread.finished.connect(self.print_finished)
        self.print_thread.start()

    def update_progress(self, current, message):
        """更新进度"""
        self.progress_bar.setValue(current)
        self.status_label.setText(message)

    def print_finished(self, success, message):
        """打印完成"""
        self.progress_bar.setValue(len(self.selected_files))
        self.status_label.setText("打印完成")

        # 显示结果
        msg_box = QMessageBox(self)
        if success:
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("成功")
        else:
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("警告")
        msg_box.setText(message)
        msg_box.exec_()

        # 重新启用按钮
        self.print_btn.setEnabled(True)
        self.printer_combo.setEnabled(True)

        # 关闭对话框
        if success:
            self.accept()
