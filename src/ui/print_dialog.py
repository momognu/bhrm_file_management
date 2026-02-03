import os
import subprocess
import tempfile
from PIL import Image
from pdf2image import convert_from_path
import win32print
import win32api
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

    progress = pyqtSignal(int, str)  # 进度信号: (当前数量, 消息)
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

                # 根据文件类型选择不同的打印方法
                ext = file_info["type"].lower()

                if ext == ".pdf":
                    # PDF文件使用pikepdf + Pillow打印
                    result = self.print_pdf_with_pikepdf(file_path)
                    if result:
                        success_count += 1
                    else:
                        failed_files.append(f"{file_info['name']} (打印失败)")

                elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"]:
                    # 图片文件直接打印
                    if self.print_image_file(file_path):
                        success_count += 1
                    else:
                        failed_files.append(f"{file_info['name']} (打印失败)")

                elif ext in [".doc", ".docx"]:
                    # Word文档静默打印
                    if self.print_word_silent(file_path):
                        success_count += 1
                    else:
                        failed_files.append(f"{file_info['name']} (打印失败)")

                elif ext in [".xls", ".xlsx"]:
                    # Excel文档静默打印
                    if self.print_excel_silent(file_path):
                        success_count += 1
                    else:
                        failed_files.append(f"{file_info['name']} (打印失败)")

                elif ext in [".ppt", ".pptx"]:
                    # PowerPoint文档静默打印
                    if self.print_powerpoint_silent(file_path):
                        success_count += 1
                    else:
                        failed_files.append(f"{file_info['name']} (打印失败)")
                else:
                    # 其他文件类型使用Windows默认打印
                    try:
                        self.print_with_shellexecute(file_path)
                        success_count += 1
                    except Exception as e:
                        failed_files.append(f"{file_info['name']} ({str(e)})")

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

    def get_printer_name(self):
        """获取打印机名称"""
        try:
            if self.printer_name:
                return self.printer_name
            else:
                return win32print.GetDefaultPrinter()
        except Exception as e:
            print(f"获取打印机名称失败: {e}")
            return None

    def print_pdf_with_pikepdf(self, file_path):
        """使用pdf2image将PDF转换为图片后打印"""
        try:
            printer_name = self.get_printer_name()
            if not printer_name:
                print("无法获取打印机名称")
                return False

            # 查找poppler路径
            poppler_path = self.find_poppler_path()
            if not poppler_path:
                print("未找到poppler，无法使用pdf2image打印PDF")
                print("请安装poppler: https://github.com/oschwartz10612/poppler-windows/releases/")
                return False

            # 使用pdf2image将PDF转换为图片列表
            # dpi=300保证打印质量
            # thread_count=4提高转换速度
            try:
                images = convert_from_path(
                    file_path,
                    dpi=300,
                    thread_count=4,
                    fmt='png',
                    poppler_path=poppler_path
                )
            except Exception as e:
                print(f"pdf2image转换失败: {e}")
                return False

            # 遍历每一页并打印
            for page_num, img in enumerate(images):
                try:
                    if not self.print_image_to_printer(img, printer_name):
                        print(f"打印第{page_num + 1}页失败")
                        return False
                except Exception as e:
                    print(f"打印第{page_num + 1}页时出错: {e}")
                    return False

            return True

        except Exception as e:
            print(f"PDF打印失败: {e}")
            return False

    def find_poppler_path(self):
        """查找poppler的安装路径"""
        # 获取项目根目录
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

        # 常见的poppler安装路径
        possible_paths = [
            # 项目目录中的poppler
            os.path.join(project_root, "poppler", "poppler-24.02.0", "Library", "bin"),
            os.path.join(project_root, "poppler", "Library", "bin"),
            # 系统路径
            r"C:\Program Files\poppler\Library\bin",
            r"C:\Program Files (x86)\poppler\Library\bin",
            r"C:\poppler\Library\bin",
            r"C:\Program Files\poppler-24.02.0\Library\bin",
            r"C:\Program Files\poppler-23.12.0\Library\bin",
            r"C:\Program Files\poppler-23.11.0\Library\bin",
        ]

        # 检查常见路径
        for path in possible_paths:
            if os.path.exists(path):
                print(f"找到poppler路径: {path}")
                return path

        # 检查PATH环境变量
        import shutil
        if shutil.which("pdftoppm"):
            print("poppler在系统PATH中")
            return None  # poppler在PATH中，不需要指定路径

        print("未找到poppler")
        return None

    def print_image_file(self, file_path):
        """打印图片文件"""
        try:
            printer_name = self.get_printer_name()
            if not printer_name:
                return False

            # 打开图片
            img = Image.open(file_path)

            # 打印图片
            return self.print_image_to_printer(img, printer_name)

        except Exception as e:
            print(f"图片文件打印失败: {e}")
            return False

    def print_image_to_printer(self, img, printer_name):
        """将PIL Image打印到指定打印机"""
        try:
            # 确保图片是RGB模式
            if img.mode != 'RGB':
                img = img.convert('RGB')

            # 创建临时文件保存图片
            temp_file = tempfile.mktemp(suffix=".bmp")
            img.save(temp_file, format='BMP')

            try:
                # 使用win32print打印图片
                # 打开打印机
                hPrinter = win32print.OpenPrinter(printer_name)

                # 准备打印作业信息
                doc_name = os.path.basename(temp_file)
                pDocInfo = (doc_name, None, "RAW")

                # 启动打印作业
                win32print.StartDocPrinter(hPrinter, 1, pDocInfo)

                # 读取BMP文件并打印
                with open(temp_file, "rb") as f:
                    bmp_data = f.read()

                # 写入打印数据
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, bmp_data)
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                win32print.ClosePrinter(hPrinter)

                return True

            except Exception as e:
                print(f"图片打印失败: {e}")
                try:
                    win32print.ClosePrinter(hPrinter)
                except:
                    pass
                return False

            finally:
                # 清理临时文件
                if os.path.exists(temp_file):
                    os.remove(temp_file)

        except Exception as e:
            print(f"图片打印失败: {e}")
            return False

    def print_with_shellexecute(self, file_path):
        """使用Windows ShellExecute打印"""
        try:
            printer_name = self.get_printer_name()
            if not printer_name:
                return False

            win32api.ShellExecute(
                0,
                "printto",
                file_path,
                f'"{printer_name}"',
                ".",
                0,  # SW_HIDE = 0，隐藏窗口
            )
            return True
        except Exception as e:
            print(f"ShellExecute打印失败: {e}")
            return False

    def print_word_silent(self, file_path):
        """Word文档静默打印"""
        try:
            import win32com.client

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            doc = word.Documents.Open(file_path)

            # 获取打印机名称
            printer_name = self.get_printer_name()
            if printer_name:
                word.ActivePrinter = printer_name

            doc.PrintOut()
            doc.Close(False)
            word.Quit()

            return True
        except Exception as e:
            print(f"Word静默打印失败: {e}")
            return False

    def print_excel_silent(self, file_path):
        """Excel文档静默打印"""
        try:
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(file_path)

            # 获取打印机名称
            printer_name = self.get_printer_name()
            if printer_name:
                excel.ActivePrinter = printer_name

            workbook.PrintOut()
            workbook.Close(False)
            excel.Quit()

            return True
        except Exception as e:
            print(f"Excel静默打印失败: {e}")
            return False

    def print_powerpoint_silent(self, file_path):
        """PowerPoint文档静默打印"""
        try:
            import win32com.client

            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Visible = False

            presentation = ppt.Presentations.Open(file_path, WithWindow=False)

            # 获取打印机名称
            printer_name = self.get_printer_name()
            if printer_name:
                ppt.ActivePrinter = printer_name

            presentation.PrintOut()
            presentation.Close()
            ppt.Quit()

            return True
        except Exception as e:
            print(f"PowerPoint静默打印失败: {e}")
            return False


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