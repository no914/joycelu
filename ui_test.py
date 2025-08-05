import os
import re
import sys
import logging
import shutil
import time
import threading
import comtypes.client
from PIL import Image, ImageDraw
import tempfile
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel, QPushButton, 
                              QFileDialog, QComboBox, QTextEdit, QHBoxLayout, QVBoxLayout, 
                              QSplitter, QFrame, QSizePolicy, QMessageBox, QListWidget, 
                              QListWidgetItem, QProgressBar)
from PySide6.QtCore import (Qt, QTimer, Signal, QObject, QThread, Q_ARG, QProcess, Slot)
from PySide6.QtGui import QPixmap, QImage, QFont, QPainter, QPen, QColor

COLOR_SCHEME = {
    'bg': '#FAFAFA',
    'text': '#333333',
    'highlight': '#3F51B5',
    'button': '#5C6BC0',
    'panel': '#E0E0E0',
    'notes_bg': '#F5F5F5'
}

class WorkerSignals(QObject):
    status_update = Signal(str, float)
    progress_detail = Signal(str)
    error_occurred = Signal(str)
    conversion_complete = Signal(bool)
    images_loaded = Signal(list)
    notes_loaded = Signal(list)
    page_progress = Signal(int, int, float)

class PPTConverter(QObject):
    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()
        self.temp_dir = os.path.join(os.getenv("TEMP"), "ppt_viewer_debug")
        
    def convert_ppt(self, pptx_path):
        """在后台线程中执行PPT转换"""
        try:
            self.signals.status_update.emit("正在初始化PPT转换...", 0)
            time.sleep(0.1)
            
            if os.path.exists(self.temp_dir):
                self.signals.status_update.emit("清理临时目录...", 5)
                shutil.rmtree(self.temp_dir)
                
            os.makedirs(self.temp_dir, exist_ok=True)
            self.signals.status_update.emit("临时目录已创建", 10)
            
            self.signals.status_update.emit("正在连接PowerPoint...", 20)
            slide_images, slide_notes = self.ppt_to_images(pptx_path, self.temp_dir)
            
            if not slide_images:
                self.signals.error_occurred.emit("PPT转换失败")
                return
                
            self.signals.status_update.emit("正在准备显示...", 90)
            self.signals.images_loaded.emit(slide_images)
            self.signals.notes_loaded.emit(slide_notes)
            self.signals.status_update.emit("加载完成", 100)
            
        except Exception as e:
            self.signals.error_occurred.emit(f"处理错误: {str(e)}")
            logging.error(f"处理过程中发生异常: {str(e)}", exc_info=True)

    def ppt_to_images(self, pptx_path, output_dir):
        """将PPT转换为图片并获取备注（带独立页内进度）"""
        powerpoint = None
        presentation = None
        slide_images = []
        slide_notes = []
        
        try:
            self.signals.status_update.emit("正在连接PowerPoint...", 0)
            
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(os.path.abspath(pptx_path))
            total_slides = presentation.Slides.Count
            
            # 创建临时目录
            if os.path.exists(output_dir):
                shutil.rmtree(output_dir)
            os.makedirs(output_dir, exist_ok=True)
            
            # 获取幻灯片尺寸
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            dpi = 600  # 高分辨率导出
            width_px = int(slide_width * dpi / 72)
            height_px = int(slide_height * dpi / 72)
            
            for i in range(1, total_slides + 1):
                # 开始处理新页（重置页内进度）
                self.signals.page_progress.emit(i, total_slides, 0)
                time.sleep(0.1)  # 确保UI有足够时间更新
                
                # 步骤1: 准备导出路径
                img_path = os.path.join(output_dir, f"slide_{i:03d}.png")
                slide = presentation.Slides.Item(i)
                self.signals.page_progress.emit(i, total_slides, 25)
                time.sleep(0.1)
                
                # 步骤2: 导出幻灯片为图片
                try:
                    slide.Export(img_path, "PNG", width_px, height_px)
                    self.signals.page_progress.emit(i, total_slides, 50)
                    time.sleep(0.1)
                    
                    # 步骤3: 加载并处理图片
                    with Image.open(img_path) as img:
                        if img.mode != 'RGB':
                            img = img.convert('RGB')
                        slide_images.append(img.copy())
                    self.signals.page_progress.emit(i, total_slides, 75)
                    time.sleep(0.1)
                    
                    # 步骤4: 提取备注
                    notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                    slide_notes.append(notes if notes else "")
                    self.signals.page_progress.emit(i, total_slides, 100)
                    time.sleep(0.1)
                    
                except Exception as e:
                    self.signals.error_occurred.emit(f"第 {i} 页处理失败: {str(e)}")
                    slide_images.append(None)
                    slide_notes.append("")
                    continue
                
                # 计算整体进度
                overall_progress = (i / total_slides) * 100
                self.signals.status_update.emit(
                    f"已完成第 {i}/{total_slides} 页", 
                    overall_progress
                )
            
            return slide_images, slide_notes
            
        except Exception as e:
            logging.error(f"PPT转换失败: {str(e)}", exc_info=True)
            return None, None
        finally:
            if presentation:
                presentation.Close()
            if powerpoint:
                powerpoint.Quit()


class PPTViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self._setup_logging()
        self._log_system_info()
        self.current_ppt_path = None
        
        self.slide_images = []
        self.slide_notes = []
        self.current_slide = 0
        self.original_notes = []
        self.high_res_cache = {}
        self.last_update_time = 0
        
        self._initialize_ui()
        
        self.converter = PPTConverter()
        self.converter_thread = QThread()
        self.converter.moveToThread(self.converter_thread)
        
        self.converter.signals.status_update.connect(self._update_status)
        self.converter.signals.error_occurred.connect(self._show_error)
        self.converter.signals.images_loaded.connect(self._on_images_loaded)
        self.converter.signals.notes_loaded.connect(self._on_notes_loaded)
        self.converter.signals.page_progress.connect(self._update_page_progress)
        
        self.converter_thread.start()

        self.preview_process = None  # 用于存储预览进程
        self.is_preview_mode = False  # 标记是否在预览模式

        self._init_debug_log()  # 新增调试日志初始化

    def update_export_progress(self, message, progress):
        """处理来自转换器的进度更新"""
        if progress is not None:
            
            # 更新状态标签
            if progress >= 0:
                self.status_label.setText(f"{message} ({progress}%)")
            else:
                self.status_label.setText(message)
                
        # 强制刷新UI
        QApplication.processEvents()
        
    def _export_video(self):
        """改进的视频导出方法，增加错误处理"""
        try:
            if not hasattr(self, 'current_ppt_path') or not self.current_ppt_path:
                QMessageBox.warning(self, "警告", "没有可用的PPT文件")
                return
                
            self._update_status("准备生成视频...", 0)
            
            # 确保临时目录存在
            temp_dir = os.path.join(tempfile.gettempdir(), "ppt_video_temp")
            os.makedirs(temp_dir, exist_ok=True)
            self.temp_video_path = os.path.join(temp_dir, "temp_output.mp4")
            print(f"[DEBUG] 视频将输出到临时路径: {self.temp_video_path}")
            
            # 获取用户选择
            choices = self._get_user_choices()
            
            # 构建命令行参数
            script_path = os.path.abspath("test2.py")
            if not os.path.exists(script_path):
                raise Exception(f"找不到脚本文件: {script_path}")
                
            args = [
                "--language", choices["language"],
                "--width", str(choices["resolution"][0]),
                "--height", str(choices["resolution"][1]),
                "--speed", str(choices["speed"]),
                "--input", self.current_ppt_path,
                "--output", self.temp_video_path,
                "--progress_format", "percentage"  # 新增参数，要求脚本返回百分比进度
            ]
            
            # 创建并配置 QProcess
            self.process = QProcess()
            self.process.setProcessChannelMode(QProcess.MergedChannels)
            
            # 连接信号
            self.process.readyReadStandardOutput.connect(self._handle_process_output)
            self.process.finished.connect(self._on_export_finished)
            
            # 启动进程
            self.process.start(sys.executable, [script_path] + args)
            
            if not self.process.waitForStarted(5000):  # 5秒超时
                raise Exception("进程启动超时")
                
            self.export_btn.setEnabled(False)
            self.save_video_btn.setEnabled(False)
            self._update_status("视频生成中...", 1)
            
        except Exception as e:
            error_msg = f"导出失败: {str(e)}"
            if hasattr(self, 'process'):
                error_msg += f"\n进程状态: {self.process.state()}"
            self._show_error(error_msg)
            if hasattr(self, 'process'):
                self.process.kill()

    def _preview_single_slide(self):
        """预览当前幻灯片的视频"""
        if not self.slide_images or self.current_slide < 0 or self.current_slide >= len(self.slide_images):
            QMessageBox.warning(self, "警告", "没有可预览的幻灯片")
            return
        
        try:
            self._update_status("准备生成单页预览视频...", 0)
            self.preview_btn.setEnabled(False)
            self.is_preview_mode = True
            
            # 确保临时目录存在
            temp_dir = os.path.join(tempfile.gettempdir(), "ppt_single_preview")
            os.makedirs(temp_dir, exist_ok=True)
            single_preview_path = os.path.join(temp_dir, f"single_slide_{self.current_slide + 1:03d}.mp4")
            
            # 获取用户选择（使用当前设置）
            choices = self._get_user_choices()
            
            # 构建命令行参数 - 只处理当前页
            script_path = os.path.abspath("test2.py")
            if not os.path.exists(script_path):
                raise Exception(f"找不到脚本文件: {script_path}")
                
            args = [
                "--language", choices["language"],
                "--width", str(choices["resolution"][0]),
                "--height", str(choices["resolution"][1]),
                "--speed", str(choices["speed"]),
                "--input", self.current_ppt_path,
                "--output", single_preview_path,
                "--single_slide", str(self.current_slide + 1),  # 告诉脚本只处理当前页
                "--slide_number", str(self.current_slide + 1)   # 当前页码（从1开始）
            ]
            
            # 创建并配置 QProcess
            self.preview_process = QProcess()
            self.preview_process.setProcessChannelMode(QProcess.MergedChannels)
            
            # 连接信号
            self.preview_process.readyReadStandardOutput.connect(self._handle_preview_output)
            self.preview_process.finished.connect(self._on_preview_finished)
            
            # 启动进程
            self.preview_process.start(sys.executable, [script_path] + args)
            
            if not self.preview_process.waitForStarted(5000):  # 5秒超时
                raise Exception("预览进程启动超时")
                
            self._update_status("单页预览生成中...", 1)
            
        except Exception as e:
            error_msg = f"单页预览失败: {str(e)}"
            self._show_error(error_msg)
            if hasattr(self, 'preview_process') and self.preview_process:
                self.preview_process.kill()
            self.preview_btn.setEnabled(True)
            self.is_preview_mode = False

    def _handle_preview_output(self):
        """处理预览进程的输出"""
        if not self.preview_process:
            return
        
        try:
            while self.preview_process.canReadLine():
                raw_output = self.preview_process.readLine()
                output = raw_output.data().decode("utf-8", errors="replace").strip()
                self._log_to_debug_file(f"[PREVIEW] {output}")
                
                # 处理进度信息
                if output.startswith("PROGRESS|"):
                    parts = output.split("|")
                    if len(parts) == 4:
                        current_page = int(parts[1])
                        total_pages = int(parts[2])
                        progress_percent = float(parts[3])
                        self._update_preview_progress(current_page, total_pages, progress_percent)
        
        except Exception as e:
            self._log_to_debug_file(f"预览输出处理异常: {str(e)}")

    def _update_preview_progress(self, current_page, total_pages, progress_percent):
        """更新预览进度"""
        status_text = f"正在生成单页预览: 第 {current_page}/{total_pages} 页 ({progress_percent:.0f}%)"
        self.processing_status_label.setText(status_text)
        QApplication.processEvents()

    def _on_preview_finished(self, exitCode, exitStatus):
        """预览完成后的处理"""
        self.preview_btn.setEnabled(True)
        self.is_preview_mode = False
        
        if self.preview_process:
            self.preview_process.kill()
            self.preview_process = None
        
        if exitCode == 0:
            self._update_status("单页预览生成完成", 100)
            # 询问用户是否要播放预览视频
            reply = QMessageBox.question(
                self, 
                "预览完成", 
                "单页预览视频已生成，是否立即播放？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                self._play_preview_video()
        else:
            self._show_error("单页预览生成失败")
            self._update_status("单页预览失败", -1)

    def _play_preview_video(self):
        """播放预览视频"""
        try:
            temp_dir = os.path.join(tempfile.gettempdir(), "ppt_single_preview")
            preview_files = [
                os.path.join(temp_dir, f"single_slide_{self.current_slide + 1:03d}.mp4"),
                os.path.join(temp_dir, f"slide_{self.current_slide + 1:03d}.mp4"),
                os.path.join(temp_dir, "output.mp4")
            ]
            
            video_path = None
            for path in preview_files:
                if os.path.exists(path):
                    video_path = path
                    break
            
            if video_path and os.path.exists(video_path):
                # 使用系统默认播放器播放视频
                if sys.platform.startswith('darwin'):  # macOS
                    os.system(f'open "{video_path}"')
                elif sys.platform.startswith('win'):  # Windows
                    os.system(f'start "" "{video_path}"')
                else:  # Linux
                    os.system(f'xdg-open "{video_path}"')
            else:
                QMessageBox.information(self, "提示", "预览视频文件未找到")
        except Exception as e:
            self._show_error(f"播放预览视频失败: {str(e)}")

    def _init_debug_log(self):
        """初始化调试日志文件"""
        self.debug_log_path = os.path.join(os.getcwd(), "ppt_viewer_debug.log")
        # 清空已有日志（可选）
        open(self.debug_log_path, "w").close() 

    def load_ppt(self, file_path):
        """加载指定的PPT文件"""
        self.current_ppt_path = file_path
        try:
            if not os.path.exists(file_path):
                QMessageBox.warning(self, "错误", f"文件不存在: {file_path}")
                return
                
            self._update_status(f"正在加载 {os.path.basename(file_path)}...", 0)
            QTimer.singleShot(100, lambda: self.converter.convert_ppt(file_path))
            
        except Exception as e:
            self._show_error(f"加载PPT失败: {str(e)}")

    def _setup_logging(self):
        """配置日志记录系统"""
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            filename='ppt_viewer_debug.log',
            filemode='w'
        )
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        logging.getLogger().addHandler(console_handler)

    def _log_system_info(self):
        """记录系统环境信息"""
        logging.info("=== 系统环境 ===")
        logging.info(f"Python版本: {sys.version}")
        logging.info(f"操作系统: {sys.platform}")
        
        try:
            import PIL
            logging.info(f"Pillow版本: {PIL.__version__}")
        except Exception as e:
            logging.error(f"获取Pillow版本失败: {str(e)}")
            
        try:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            logging.info(f"PowerPoint版本: {powerpoint.Version}")
            powerpoint.Quit()
        except Exception as e:
            logging.error(f"获取PowerPoint信息失败: {str(e)}")

    def _initialize_ui(self):
        """初始化主窗口UI"""
        self.setWindowTitle("高清PPT查看器")
        self.resize(1200, 800)
        
        central_widget = QWidget()
        central_widget.setStyleSheet(f"background-color: {COLOR_SCHEME['bg']};")
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 左侧文件列表区域
        self._create_file_list_panel(main_layout)
        
        # 右侧主区域
        right_panel = QWidget()
        right_panel.setLayout(QVBoxLayout())
        right_panel.layout().setContentsMargins(0, 0, 0, 0)
        right_panel.layout().setSpacing(0)
        
        # 顶部控制栏
        top_frame = self._create_top_frame()
        right_panel.layout().addWidget(top_frame)
        
        # 主内容区域
        content_splitter = self._create_content_splitter()
        right_panel.layout().addWidget(content_splitter)
        
        # 创建底部控制栏
        self.bottom_frame = self._create_bottom_frame()
        
        # 创建页内进度框架
        self.page_progress_frame = QFrame()
        page_progress_layout = QHBoxLayout(self.page_progress_frame)
        page_progress_layout.setContentsMargins(0, 0, 0, 0)
        
        # 添加底部控制栏到右侧面板
        right_panel.layout().addWidget(self.bottom_frame)
        
        # 添加右侧面板到主布局
        main_layout.addWidget(right_panel)

    def _create_file_list_panel(self, main_layout):
        """创建左侧文件列表面板（已移除刷新按钮）"""
        self.file_list_panel = QFrame()
        self.file_list_panel.setStyleSheet(f"""
            QFrame {{
                background-color: {COLOR_SCHEME['panel']};
                border-right: 1px solid #BDBDBD;
            }}
        """)
        self.file_list_panel.setFixedWidth(250)
        
        file_list_layout = QVBoxLayout(self.file_list_panel)
        file_list_layout.setContentsMargins(5, 5, 5, 5)
        file_list_layout.setSpacing(5)
        
        # 标题和按钮（移除了刷新按钮）
        file_header = QFrame()
        file_header.setStyleSheet(f"background-color: {COLOR_SCHEME['panel']};")
        file_header_layout = QHBoxLayout(file_header)
        file_header_layout.setContentsMargins(5, 5, 5, 5)
        
        # 只保留打开按钮
        self.open_btn = QPushButton("打开")
        self.open_btn.setFixedSize(60, 25)
        self.open_btn.clicked.connect(self._open_ppt_file)

        file_header_layout.addWidget(self.open_btn)  # 只添加打开按钮
        file_header_layout.addStretch()
        
        # 搜索框
        self.search_box = QTextEdit()
        self.search_box.setMaximumHeight(30)
        self.search_box.setPlaceholderText("搜索...")
        self.search_box.textChanged.connect(self._filter_file_list)
        
        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #BDBDBD;
                padding: 2px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:hover {
                background-color: #E3F2FD;
            }
            QListWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
        """)
        self.file_list.itemDoubleClicked.connect(self._on_file_selected)
        
        # 添加到布局
        file_list_layout.addWidget(file_header)
        file_list_layout.addWidget(self.search_box)
        file_list_layout.addWidget(self.file_list)
        
        # 添加到主布局
        main_layout.addWidget(self.file_list_panel)
        
        # 初始加载文件列表（保留初始化加载）
        self._load_file_list()  # 将原来的_refresh_file_list改名为_load_file_list

    def _load_file_list(self):
        """初始化加载文件列表（原刷新功能）"""
        self.file_list.clear()
        
        # 获取当前目录下的PPT文件
        current_dir = os.getcwd()
        ppt_files = []
        
        for root, dirs, files in os.walk(current_dir):
            for file in files:
                if file.lower().endswith(('.ppt', '.pptx')):
                    ppt_files.append(os.path.join(root, file))
        
        # 按文件名排序
        ppt_files.sort()
        
        # 添加到列表控件
        for file_path in ppt_files:
            file_name = os.path.basename(file_path)
            item = QListWidgetItem(file_name)
            item.setData(Qt.UserRole, file_path)
            self.file_list.addItem(item)

    def _filter_file_list(self):
        """根据搜索框内容过滤文件列表"""
        search_text = self.search_box.toPlainText().lower()
        
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            item_text = item.text().lower()
            item.setHidden(search_text not in item_text)

    def _on_file_selected(self, item):
        """当文件被选中时加载PPT"""
        file_path = item.data(Qt.UserRole)
        self.load_ppt(file_path)

    def _open_ppt_file(self):
        """打开文件对话框选择PPT文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "打开PPT文件",
            os.getcwd(),
            "PPT文件 (*.ppt *.pptx);;所有文件 (*)"
        )
        
        if file_path:
            self.load_ppt(file_path)
            # 如果文件不在当前目录，添加到文件列表
            if not any(file_path == self.file_list.item(i).data(Qt.UserRole) 
                      for i in range(self.file_list.count())):
                file_name = os.path.basename(file_path)
                item = QListWidgetItem(file_name)
                item.setData(Qt.UserRole, file_path)
                self.file_list.addItem(item)

    def _create_top_frame(self):
        """创建顶部控制栏"""
        frame = QFrame()
        frame.setStyleSheet(f"background-color: {COLOR_SCHEME['panel']};")
        frame.setFixedHeight(40)
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(10)
        
        # 语言选择
        self.lang_label = QLabel("语言:")
        self.lang_combobox = QComboBox()
        self.lang_combobox.addItems(["中文", "英文"])
        self.lang_combobox.setFixedWidth(120)
        
        # 分辨率选择
        self.resolution_label = QLabel("分辨率:")
        self.resolution_combobox = QComboBox()
        self.resolution_combobox.addItems(["1920x1080 (全高清)", "1280x720 (高清)"])
        self.resolution_combobox.setFixedWidth(150)
        
        # 语速选择
        self.speed_label = QLabel("语速:")
        self.speed_combobox = QComboBox()
        self.speed_combobox.addItems(["慢速 (0.75x)", "正常 (1x)", "快速 (1.5x)"])
        self.speed_combobox.setFixedWidth(120)

        # 添加预览按钮
        self.preview_btn = QPushButton("单页预览")
        self.preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
            }
        """)
        self.preview_btn.setFixedWidth(100)
        self.preview_btn.clicked.connect(self._preview_single_slide)
        self.preview_btn.setEnabled(False)  # 初始禁用，直到有幻灯片加载
        layout.addWidget(self.preview_btn)  # 添加到其他按钮旁边
        
        # 输出视频按钮
        self.export_btn = QPushButton("输出视频")
        self.export_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_SCHEME['button']};
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }}
            QPushButton:hover {{
                background-color: {COLOR_SCHEME['highlight']};
            }}
        """)
        self.export_btn.setFixedWidth(100)
        self.export_btn.clicked.connect(self._export_video)
        
        # 保存视频按钮
        self.save_video_btn = QPushButton("保存视频")
        self.save_video_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
            }
        """)
        self.save_video_btn.setFixedWidth(100)
        self.save_video_btn.setEnabled(False)
        self.save_video_btn.clicked.connect(self._save_video_to_location)
        
        layout.addWidget(self.lang_label)
        layout.addWidget(self.lang_combobox)
        layout.addSpacing(10)
        layout.addWidget(self.resolution_label)
        layout.addWidget(self.resolution_combobox)
        layout.addSpacing(10)
        layout.addWidget(self.speed_label)
        layout.addWidget(self.speed_combobox)
        layout.addStretch()
        layout.addWidget(self.export_btn)
        layout.addWidget(self.save_video_btn)
        
        return frame

    def _create_content_splitter(self):
        """创建主内容区域"""
        splitter = QSplitter(Qt.Horizontal)
        
        # PPT显示区域
        self.ppt_frame = QFrame()
        self.ppt_frame.setStyleSheet("background-color: white;")
        
        ppt_layout = QVBoxLayout(self.ppt_frame)
        ppt_layout.setContentsMargins(5, 5, 5, 5)
        
        self.slide_label = PPTSlideLabel()
        self.slide_label.setAlignment(Qt.AlignCenter)
        self.slide_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.slide_label.setMouseTracking(True)  # ✅ 启用鼠标移动跟踪
        self.slide_label.setStyleSheet("""
            QLabel {
                border: 2px solid #A0A0A0;  /* 浅灰色边框，可换成 #000000 表示黑色 */
                border-radius: 4px;         /* 可选：圆角 */
                background-color: white;    /* 背景色，可选 */
            }
        """)
        self.slide_label.mouseMoveEvent = self._update_mouse_position
        ppt_layout.addWidget(self.slide_label)
        
        # 备注区域
        self.notes_frame = QFrame()
        self.notes_frame.setStyleSheet(f"background-color: {COLOR_SCHEME['notes_bg']};")
        
        notes_layout = QVBoxLayout(self.notes_frame)
        notes_layout.setContentsMargins(5, 5, 5, 5)
        notes_layout.setSpacing(5)
        
        notes_header = QFrame()
        notes_header.setStyleSheet(f"background-color: {COLOR_SCHEME['panel']};")
        
        notes_header_layout = QHBoxLayout(notes_header)
        notes_header_layout.setContentsMargins(5, 5, 5, 5)
        
        self.notes_title = QLabel("备注")
        self.notes_title.setStyleSheet(f"""
            color: {COLOR_SCHEME['text']};
            font-weight: bold;
        """)
        
        self.save_notes_btn = QPushButton("保存备注")
        self.save_notes_btn.setEnabled(False)
        self.save_notes_btn.clicked.connect(self._save_notes)
        
        notes_header_layout.addWidget(self.notes_title)
        notes_header_layout.addStretch()
        notes_header_layout.addWidget(self.save_notes_btn)
        
        self.notes_text = QTextEdit()
        self.notes_text.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #BDBDBD;
                padding: 5px;
            }
        """)
        self.notes_text.textChanged.connect(self._on_notes_modified)
        
        notes_layout.addWidget(notes_header)
        notes_layout.addWidget(self.notes_text)
        
        splitter.addWidget(self.ppt_frame)
        splitter.addWidget(self.notes_frame)
        
        splitter.setSizes([int(self.width() * 0.6), int(self.width() * 0.4)])
        
        return splitter

    def _create_bottom_frame(self):
        """创建底部控制栏"""
        frame = QFrame()
        frame.setStyleSheet(f"background-color: {COLOR_SCHEME['panel']};")
        frame.setFixedHeight(40)
        
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(10)

        # 导航按钮
        nav_frame = QFrame()
        nav_layout = QHBoxLayout(nav_frame)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(10)
        
        self.prev_btn = QPushButton("上一页")
        self.prev_btn.setEnabled(False)
        self.prev_btn.clicked.connect(self._prev_slide)
        
        self.next_btn = QPushButton("下一页")
        self.next_btn.setEnabled(False)
        self.next_btn.clicked.connect(self._next_slide)
        
        self.page_label = QLabel("0/0")
        
        # 鼠标位置标签
        self.mouse_pos_label = QLabel("位置: --% x --%")
        
        # 状态标签
        self.processing_status_label = QLabel("就绪")
        self.processing_status_label.setStyleSheet("color: #666666;")
        
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.next_btn)
        nav_layout.addWidget(self.page_label)
        
        layout.addWidget(nav_frame)
        layout.addStretch()
        layout.addWidget(self.mouse_pos_label)
        layout.addWidget(self.processing_status_label)
        
        return frame

    def _update_status(self, message, progress=None):
        """更新右下角状态文字显示"""
        try:
            if progress is not None:
                self.processing_status_label.setText(f"{message} ({int(progress)}%)")
            else:
                self.processing_status_label.setText(message)
            
            # 可选：强制刷新
            QApplication.processEvents()

        except Exception as e:
            print(f"状态更新失败: {e}")

    def _filter_status_message(self, raw_message):
        """过滤原始消息，返回用户友好的简洁信息"""
        # 定义需要过滤掉的技术性关键词
        technical_terms = [
            't:', 'chunk:', 'it/s', 'now=', 
            'frame=', 'fps=', 'size='
        ]
        
        # 如果是明确的进度信息
        if any(x in raw_message for x in ['%', '页', '进度']):
            # 提取最简洁的进度描述
            if '|' in raw_message:
                return raw_message.split('|')[0].strip()
            return raw_message
            
        # 如果是错误信息
        if 'error' in raw_message.lower():
            return "处理出错，请查看日志"
            
        # 过滤掉包含技术术语的消息
        if any(term in raw_message for term in technical_terms):
            return None
            
        # 默认返回原始消息（如果没有技术术语）
        return raw_message

    def _show_error(self, message):
        """显示错误信息"""
        self._update_status(f"错误: {message.split(':')[-1].strip()}", -1)
        QMessageBox.critical(self, "错误", message)

    def _on_images_loaded(self, images):
        """图片加载完成处理"""
        self._update_status("幻灯片加载完成", 100)
        self.slide_images = [img for img in images if img is not None]
        self.high_res_cache.clear()
        
        if self.slide_images:
            for i, img in enumerate(self.slide_images):
                self.high_res_cache[i] = img.copy()
                
        self._show_first_slide()
        self.page_label.setText(f"1/{len(self.slide_images)}")
        self.preview_btn.setEnabled(True)  # 启用预览按钮

    def _on_notes_loaded(self, notes):
        """备注加载完成处理"""
        self.slide_notes = notes
        self.original_notes = notes.copy()
        self._update_notes_content()

    def _show_first_slide(self):
        """显示第一张幻灯片"""
        self.current_slide = 0
        self._display_current_slide()
        self._update_nav_buttons()

    def _display_current_slide(self):
        if not self.slide_images or self.current_slide >= len(self.slide_images):
            return

        try:
            original_img = self.high_res_cache[self.current_slide]
            device_ratio = self.devicePixelRatio()
            phys_width = int(self.slide_label.width() * device_ratio)
            phys_height = int(self.slide_label.height() * device_ratio)
            
            img_ratio = original_img.width / original_img.height
            display_width = min(phys_width, original_img.width)
            display_height = int(display_width / img_ratio)
            
            if display_height > phys_height:
                display_height = phys_height
                display_width = int(display_height * img_ratio)

            if original_img.width / display_width > 2:
                intermediate_size = (display_width * 2, display_height * 2)
                temp_img = original_img.resize(intermediate_size, resample=Image.LANCZOS)
                resized_img = temp_img.resize((display_width, display_height), resample=Image.LANCZOS)
            else:
                resized_img = original_img.resize((display_width, display_height), resample=Image.LANCZOS)

            qimage = QImage(
                resized_img.tobytes(),
                resized_img.width,
                resized_img.height,
                resized_img.width * 3,
                QImage.Format_RGB888
            )
            pixmap = QPixmap.fromImage(qimage)
            pixmap.setDevicePixelRatio(device_ratio)

            self.slide_label.setPixmap(pixmap)
            
        except Exception as e:
            print(f"显示错误: {str(e)}")
            self._show_error_placeholder()

    def _update_notes_content(self):
        """更新备注内容"""
        if not self.slide_notes or self.current_slide >= len(self.slide_notes):
            self.notes_text.setPlainText("无备注内容")
            self.save_notes_btn.setEnabled(False)
            return
            
        try:
            note_content = self.slide_notes[self.current_slide]
            note_content = note_content.replace('\r\n', '\n').replace('\r', '\n')
            
            self.notes_text.setPlainText(note_content)
            self.save_notes_btn.setEnabled(False)
            
        except Exception as e:
            self.notes_text.setPlainText("备注加载失败")
            self.save_notes_btn.setEnabled(False)

    def _on_notes_modified(self):
        """备注内容修改处理"""
        if not hasattr(self, 'original_notes') or not self.original_notes:
            return
            
        if self.current_slide >= len(self.original_notes):
            return
            
        current_text = self.notes_text.toPlainText().strip()
        if current_text != self.original_notes[self.current_slide]:
            self.save_notes_btn.setEnabled(True)
        else:
            self.save_notes_btn.setEnabled(False)

    def _save_notes(self):
        """保存备注"""
        if not self.slide_notes or self.current_slide >= len(self.slide_notes):
            return
        
        try:
            modified_notes = self.notes_text.toPlainText().strip()
            self.slide_notes[self.current_slide] = modified_notes
            QMessageBox.information(self, "成功", "备注已保存")
            self.save_notes_btn.setEnabled(False)
            
            if hasattr(self, 'original_notes'):
                self.original_notes[self.current_slide] = modified_notes
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存备注时出错:\n{str(e)}")

    def _update_nav_buttons(self):
        """更新导航按钮状态"""
        prev_state = self.current_slide > 0
        next_state = self.current_slide < len(self.slide_images)-1
        
        self.prev_btn.setEnabled(prev_state)
        self.next_btn.setEnabled(next_state)

    def _prev_slide(self):
        """显示上一页"""
        if self.current_slide > 0:
            self.current_slide -= 1
            self._display_current_slide()
            self._update_notes_content()
            self.page_label.setText(f"{self.current_slide + 1}/{len(self.slide_images)}")
            self._update_nav_buttons()

    def _next_slide(self):
        """显示下一页"""
        if self.current_slide < len(self.slide_images)-1:
            self.current_slide += 1
            self._display_current_slide()
            self._update_notes_content()
            self.page_label.setText(f"{self.current_slide + 1}/{len(self.slide_images)}")
            self._update_nav_buttons()

    def _update_page_progress(self, current_page, total_pages, page_progress):
        """更新状态文字，显示当前页和整体进度信息（不再使用进度条）"""
        try:
            # 构造用户友好的进度文字，显示在 status_label 上
            overall_progress = ((current_page - 1) / total_pages * 100) + (page_progress / total_pages)
            overall_pct = int(overall_progress)
            page_pct = int(page_progress)

            status_text = f"第 {current_page}/{total_pages} 页 | 整体: {overall_pct}% | 当前页: {page_pct}%"
            
            # 将进度信息以文字方式显示到右下角 status label
            self.processing_status_label.setText(status_text)

            # 强制刷新 UI（可选）
            QApplication.processEvents()

        except Exception as e:
            print(f"更新进度文字失败: {e}")


    def _show_error_placeholder(self):
        """显示错误占位图"""
        width = max(100, self.slide_label.width())
        height = max(100, self.slide_label.height())
        
        img = Image.new('RGB', (width, height), color='red')
        draw = ImageDraw.Draw(img)
        draw.text((50, 50), "幻灯片显示错误", fill='white')
        
        qimage = QImage(img.tobytes(), img.width, img.height, 
                      img.width * 3, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(qimage)
        self.slide_label.setPixmap(pixmap)

    @Slot(int, QProcess.ExitStatus)
    def _on_export_finished(self, exitCode, exitStatus):
        """视频导出完成后的处理"""
        self.export_btn.setEnabled(True)
        
        if exitCode == 0:
            self._update_status("视频生成完成，请点击保存按钮")
            self.save_video_btn.setEnabled(True)
        else:
            self._show_error("视频导出失败")
            if hasattr(self, 'temp_video_path'):
                os.remove(self.temp_video_path)

    def _export_video(self):
        """改进的视频导出方法，增加错误处理"""
        try:
            if not hasattr(self, 'current_ppt_path') or not self.current_ppt_path:
                QMessageBox.warning(self, "警告", "没有可用的PPT文件")
                return
                
            self._update_status("准备生成视频...", 0)
            
            # 确保临时目录存在
            temp_dir = os.path.join(tempfile.gettempdir(), "ppt_video_temp")
            os.makedirs(temp_dir, exist_ok=True)
            self.temp_video_path = os.path.join(temp_dir, "temp_output.mp4")
            
            # 获取用户选择
            choices = self._get_user_choices()
            
            # 构建命令行参数
            script_path = os.path.abspath("test2.py")
            if not os.path.exists(script_path):
                raise Exception(f"找不到脚本文件: {script_path}")
                
            args = [
                "--language", choices["language"],
                "--width", str(choices["resolution"][0]),
                "--height", str(choices["resolution"][1]),
                "--speed", str(choices["speed"]),
                "--input", self.current_ppt_path,
                "--output", self.temp_video_path
            ]
            
            # 创建并配置 QProcess
            self.process = QProcess()
            self.process.setProcessChannelMode(QProcess.MergedChannels)
            
            # 连接信号
            self.process.readyReadStandardOutput.connect(self._handle_process_output)
            self.process.finished.connect(self._on_export_finished)
            
            # 启动进程
            self.process.start(sys.executable, [script_path] + args)
            
            if not self.process.waitForStarted(5000):  # 5秒超时
                raise Exception("进程启动超时")
                
            self.export_btn.setEnabled(False)
            self.save_video_btn.setEnabled(False)
            self._update_status("视频生成中...", 1)
            
        except Exception as e:
            error_msg = f"导出失败: {str(e)}"
            if hasattr(self, 'process'):
                error_msg += f"\n进程状态: {self.process.state()}"
            self._show_error(error_msg)
            if hasattr(self, 'process'):
                self.process.kill()

    def _extract_progress_info(self, output):
        """解析标准化的进度信息格式"""
        if output.startswith("PAGE_PROGRESS|"):
            parts = output.split("|")
            if len(parts) == 4:
                try:
                    current_page = int(parts[1])
                    total_pages = int(parts[2])
                    page_progress = float(parts[3])
                    
                    return {
                        "type": "page",
                        "current_page": current_page,
                        "total_pages": total_pages,
                        "page_progress": page_progress,
                        "overall_progress": ((current_page - 1) / total_pages * 100) + 
                                        (page_progress / total_pages)
                    }
                except:
                    return None
        return None

    def _generate_status_text(self, progress_info):
        """生成正确的状态文本"""
        if not progress_info:
            return "正在处理..."
        
        if "current_page" in progress_info and "total_pages" in progress_info:
            # 当有页数信息时，显示"第X/Y页 (Z%)"
            base_text = f"正在处理第 {progress_info['current_page']}/{progress_info['total_pages']} 页"
            if "page_progress" in progress_info:
                return f"{base_text} ({progress_info['page_progress']:.0f}%)"
            return base_text
        elif "page_progress" in progress_info:
            # 只有百分比时，显示"处理进度: X%"
            return f"处理进度: {progress_info['page_progress']:.0f}%"
        else:
            return "正在处理..."

    @Slot()
    def _handle_process_output(self):
        while self.process.canReadLine():
            try:
                raw_output = self.process.readLine()
                output = raw_output.data().decode("utf-8", errors="replace").strip()
                
                self._log_to_debug_file(f"[PROCESS] {output}")  # ✅ 有写入调试日志吗？

                if output.startswith("PROGRESS|"):
                    parts = output.split("|")
                    if len(parts) == 4:
                        current_page = int(parts[1])
                        total_pages = int(parts[2])
                        progress_percent = float(parts[3])
                        self._update_video_progress(current_page, total_pages, progress_percent)
                    
            except Exception as e:
                self._log_to_debug_file(f"输出处理异常: {str(e)}")

    def _update_video_progress(self, current_page, total_pages, progress_percent):
        """更新视频导出进度"""
        
        # 更新状态标签
        status_text = f"正在导出视频: 第 {current_page}/{total_pages} 页 ({progress_percent:.0f}%)"
        self.processing_status_label.setText(status_text)
        
        # 强制刷新UI
        QApplication.processEvents()

    def _log_to_debug_file(self, message):
        """专用方法记录调试信息"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(self.debug_log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")

    @Slot()
    def _handle_process_error(self):
        """简化错误信息显示"""
        error = self.process.readAllStandardError().data().decode('utf-8').strip()
        if error:
            # 显示简化错误信息
            if not any(x in error for x in ['it/s', 'now=None', 'chunk:', 't:']):
                self._update_status("处理出错，请查看日志")

    def _get_user_choices(self):
        """从界面获取用户选择"""
        lang_map = {"中文": "zh-cn", "英文": "en"}
        lang = lang_map.get(self.lang_combobox.currentText(), "zh-cn")
        
        res_text = self.resolution_combobox.currentText().split()[0]
        resolution = tuple(map(int, res_text.split("x")))
        
        speed_map = {
            "慢速 (0.75x)": 0.75,
            "正常 (1x)": 1.0,
            "快速 (1.5x)": 1.5
        }
        speed = speed_map.get(self.speed_combobox.currentText(), 1.0)
        
        return {
            "language": lang,
            "resolution": resolution,
            "speed": speed
        }
    
    def _update_mouse_position(self, event):
        # --------------------------------------------------
        # 功能：仅当鼠标位于PPT图片内容区域（即幻灯片显示部分）时，显示百分比位置；
        #       如果鼠标在QLabel中的空白区域（比如左右/上下），则显示 "位置: --% x --%"
        # --------------------------------------------------

        # ---- 1. 检查当前是否有幻灯片且图片有效 ----
        if (not self.slide_images or
            self.current_slide < 0 or
            self.current_slide >= len(self.slide_images) or
            self.slide_images[self.current_slide] is None):
            self.mouse_pos_label.setText("位置: --% x --%")
            return

        try:
            # ---- 2. 获取鼠标在 QLabel 中的坐标 ----
            if hasattr(event, 'position'):  # Qt 6
                mouse_x = event.position().x()
                mouse_y = event.position().y()
            else:  # Qt 5 / 兼容
                mouse_x = event.pos().x()
                mouse_y = event.pos().y()

            label_width = self.slide_label.width()
            label_height = self.slide_label.height()

            # ---- 3. QLabel 尺寸无效，不处理 ----
            if label_width <= 0 or label_height <= 0:
                self.mouse_pos_label.setText("位置: --% x --%")
                return

            # ---- 4. 当前幻灯片图片信息 ----
            img = self.slide_images[self.current_slide]
            img_width = img.width
            img_height = img.height

            # ---- 5. 计算图片缩放后的尺寸（保持宽高比） ----
            img_ratio = img_width / img_height
            label_ratio = label_width / label_height

            if img_ratio > label_ratio:
                # 图片更宽（相对于 QLabel），以宽度为准进行缩放，高度按比例缩小
                scaled_width = label_width
                scaled_height = int(label_width / img_ratio)
            else:
                # 图片更高，以高度为准进行缩放，宽度按比例缩小
                scaled_height = label_height
                scaled_width = int(label_height * img_ratio)

            # ---- 6. 计算图片在 QLabel 中的居中偏移量 ----
            offset_x = (label_width - scaled_width) // 2
            offset_y = (label_height - scaled_height) // 2

            # ---- 7. 判断鼠标是否在【图片实际显示区域】内 ----
            is_inside_image = (
                mouse_x >= offset_x and
                mouse_x < offset_x + scaled_width and
                mouse_y >= offset_y and
                mouse_y < offset_y + scaled_height
            )

            if is_inside_image:
                # ---- 8. 计算鼠标在图片内的相对坐标，并转为百分比 ----
                rel_x = mouse_x - offset_x
                rel_y = mouse_y - offset_y

                x_percent = int((rel_x / scaled_width) * 100)
                y_percent = int((rel_y / scaled_height) * 100)

                self.mouse_pos_label.setText(f"位置: {x_percent}% x {y_percent}%")
            else:
                # ---- 9. 鼠标不在图片区域（可能在左右/上下空白处），显示 --% x --% ----
                self.mouse_pos_label.setText("位置: --% x --%")

        except Exception as e:
            print(f"[ERROR] _update_mouse_position: {e}")
            self.mouse_pos_label.setText("位置: 计算错误")

    def _save_video_to_location(self):
        """让用户选择保存位置"""
        if not hasattr(self, 'temp_video_path') or not os.path.exists(self.temp_video_path):
            self._show_error("没有可保存的视频文件")
            return
        
        try:
            file_path, selected_filter = QFileDialog.getSaveFileName(
                self,
                "保存视频文件",
                os.path.join(os.path.expanduser("~"), "Desktop", "output.mp4"),
                "MP4视频文件 (*.mp4);;所有文件 (*)"
            )
            
            if file_path:
                if not file_path.lower().endswith('.mp4'):
                    file_path += '.mp4'
                
                shutil.copy2(self.temp_video_path, file_path)
                
                QMessageBox.information(
                    self,
                    "保存成功",
                    f"视频已保存到:\n{file_path}",
                    QMessageBox.Ok
                )
                
                os.remove(self.temp_video_path)
                del self.temp_video_path
                
        except Exception as e:
            self._show_error(f"保存失败: {str(e)}")

    def closeEvent(self, event):
        """确保所有临时资源被释放"""
        if hasattr(self, 'temp_video_path') and os.path.exists(self.temp_video_path):
            os.remove(self.temp_video_path)
        
        if hasattr(self, 'converter_thread'):
            self.converter_thread.quit()
            self.converter_thread.wait(1000)  # 等待1秒
        
        event.accept()

class PPTSlideLabel(QLabel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMouseTracking(True)
        # 可选：设置背景色为白色，便于观察
        self.setStyleSheet("background-color: white;")

    def paintEvent(self, event):
        # Step 1: 先让 QLabel 正常绘制图片内容
        super().paintEvent(event)

        # Step 2: 如果没有图片、或当前没有幻灯片，不绘制框
        if (not hasattr(self, 'parent') or 
            not hasattr(self.parent(), 'slide_images') or
            not hasattr(self.parent(), 'current_slide')):
            return

        parent = self.parent()
        slide_images = getattr(parent, 'slide_images', [])
        current_slide = getattr(parent, 'current_slide', -1)

        if (not slide_images or 
            current_slide < 0 or 
            current_slide >= len(slide_images) or 
            slide_images[current_slide] is None):
            return

        try:
            # Step 3: 获取当前显示的图片对象及其原始尺寸
            img = slide_images[current_slide]
            img_width = img.width
            img_height = img.height

            # Step 4: 获取 QLabel 当前尺寸
            label_width = self.width()
            label_height = self.height()

            if label_width <= 0 or label_height <= 0:
                return

            # Step 5: 计算图片缩放后的尺寸（保持宽高比，居中显示）
            img_ratio = img_width / img_height
            label_ratio = label_width / label_height

            if img_ratio > label_ratio:
                # 图片更宽，以宽度为准缩放
                scaled_width = label_width
                scaled_height = int(label_width / img_ratio)
            else:
                # 图片更高，以高度为准缩放
                scaled_height = label_height
                scaled_width = int(label_height * img_ratio)

            # Step 6: 计算居中偏移
            offset_x = (label_width - scaled_width) // 2
            offset_y = (label_height - scaled_height) // 2

            # Step 7: 图片实际显示区域
            image_rect = (
                offset_x,
                offset_y,
                scaled_width,
                scaled_height
            )

            # Step 8: 使用 QPainter 绘制边框（仅围绕图片内容区域）
            painter = QPainter(self)
            pen = QPen(QColor(0, 0, 0), 2)  # 黑色，2px 线宽，可自定义颜色/粗细
            painter.setPen(pen)
            painter.drawRect(
                image_rect[0], image_rect[1], image_rect[2], image_rect[3]
            )

        except Exception as e:
            print(f"[PPTSlideLabel] 绘制边框时出错: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    
    try:
        window = PPTViewer()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.critical(f"程序崩溃: {str(e)}", exc_info=True)
        QMessageBox.critical(None, "致命错误", f"程序发生致命错误:\n{str(e)}")