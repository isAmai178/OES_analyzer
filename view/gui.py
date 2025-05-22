import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QSpinBox,
    QDoubleSpinBox, QTableWidget, QTableWidgetItem, QMessageBox, QMenu,
    QTextEdit, QGroupBox , QHeaderView,  QCheckBox, QGridLayout, QComboBox,
    QDialog, QListWidget, QSizePolicy, QProgressDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from controller.controller import OESController
import pandas as pd
import os
from typing import List, Dict
import logging
import matplotlib
matplotlib.use('Qt5Agg')  # 設置 matplotlib 後端
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class OESAnalyzerGUI(QMainWindow):
    """
    Main GUI class for the OES Analyzer application.
    This View interacts with the Controller to display results and handle user input.
    """

    def __init__(self):
        super().__init__()
        self.controller = OESController()
        self.start_index = 0
        self.end_index = 0
        self.analysis_results = {}  # Initialize analysis_results to avoid AttributeError
        self.base_name = ""  # 初始化 base_name
        self.base_names = {}  # 用於存儲每個資料夾的 base_name
        self.start_indices = {}  # 用於存儲每個資料夾的 start_index
        self.end_indices = {}  # 用於存儲每個資料夾的 end_index
        self.selected_folders = []  # 用於存儲選擇的資料夾
        self.setWindowTitle("OES Analyzer")

        # 获取屏幕分辨率
        screen = QApplication.primaryScreen()
        screen_rect = screen.availableGeometry()
        screen_width = screen_rect.width()
        screen_height = screen_rect.height()

        # 设置初始窗口大小为屏幕宽度的80%和高度的80%
        self.resize(int(screen_width * 0.8), int(screen_height * 0.8))

        # 设置最小尺寸
        self.setMinimumSize(800, 600)

        self._init_ui()

    def _init_ui(self):
        """Initialize the GUI layout and components."""
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)

        self.main_layout.setSpacing(10)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        
        #最上方檔案功能
        # self._setup_OES_file_section(self.main_layout)  # 檔案路徑設定

        # 創建一個水平佈局
        self.horizontal_layout = QHBoxLayout()
        self.main_layout.addLayout(self.horizontal_layout)

        # 左側功能區域
        spectrum_group = QGroupBox("光譜分析")
        # spectrum_group.setStyleSheet("QGroupBox { font-size: 14px; font-weight: bold;}")  # 設置標題字體大小和粗體
        left_layout = QVBoxLayout(spectrum_group)
        self.horizontal_layout.addWidget(spectrum_group)

        # 右側功能區域
        stability_group = QGroupBox("穩定度分析")
        # stability_group.setStyleSheet("QGroupBox { font-size: 14px; font-weight: bold;}")  # 設置標題字體大小和粗體
        right_layout = QVBoxLayout(stability_group)
        self.horizontal_layout.addWidget(stability_group)

        #左邊為光譜分析功能
        self._setup_OES_file_section(left_layout)  # 檔案路徑設定
        self._setup_save_directory_selection(left_layout) #保存檔案路徑設定
        self._setup_OES_parameters_section(left_layout)  # 參數設定
        self._setup_waveband_settings(left_layout)  # 波段設定
        # 創建一個新的水平佈局來放置按鈕
        button_layout = QHBoxLayout()
        # 光譜分析按鈕
        self._setup_OES_analysis_section(button_layout)  # 分析按鈕
        # 擷取特定波段數據的按鈕
        self._setup_extract_waveband_section(button_layout)  # 擷取特定波段數據
        left_layout.addLayout(button_layout)
        self._setup_image_display(left_layout)  # 圖像顯示

        # 右側為穩定度分析功能
        self._setup_stability_file_section(right_layout) # 檔案路徑設定
        self._setup_Stability_analysis_parameters_section(right_layout)  # 參數設定
        self._setup_Stability_analysis_section(right_layout)  # 分析按鈕
        self._setup_results_section(right_layout)  # 結果顯示

    def _setup_OES_file_section(self, parent_layout):
        """Create file selection section."""
        # 光譜分析 GroupBox
        spectrum_group = QGroupBox("檔案設定")
        spectrum_layout = QVBoxLayout()

        folder_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("選擇資料夾路徑...")

        browse_button = QPushButton("瀏覽")
        browse_button.clicked.connect(self._browse_folder)

        folder_layout.addWidget(QLabel("資料夾路徑:"))
        folder_layout.addWidget(self.path_edit)
        folder_layout.addWidget(browse_button)

        spectrum_layout.addLayout(folder_layout)
        spectrum_group.setLayout(spectrum_layout)
        parent_layout.addWidget(spectrum_group)
        # spectrum_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

    def _setup_stability_file_section(self, parent_layout):
        # 穩定度 GroupBox
        stability_group = QGroupBox("檔案設定")
        stability_layout = QVBoxLayout()

        folder_browse_button = QPushButton("選擇資料夾")
        folder_browse_button.clicked.connect(self._browse_folders)
        stability_layout.addWidget(folder_browse_button)

        self.folder_selector = QComboBox()
        self.folder_selector.currentIndexChanged.connect(self._update_results_display)
        stability_layout.addWidget(self.folder_selector)

        stability_layout.addLayout(stability_layout)
        stability_group.setLayout(stability_layout)
        parent_layout.addWidget(stability_group)

    def _setup_Stability_analysis_parameters_section(self ,parent_layout):
        """Create parameters input section."""
        group = QGroupBox("穩定度參數設定")
        layout = QVBoxLayout()
        
        # Parameters grid
        params_grid = QHBoxLayout()
        params_grid.setSpacing(20)
        
        #Wave detection
        wave_layout = QVBoxLayout()
        wave_label = QLabel('檢測波長:')
        self.detect_wave_spin = QDoubleSpinBox()
        self.detect_wave_spin.setRange(0, 1000)
        self.detect_wave_spin.setValue(657)
        self.detect_wave_spin.setFixedWidth(125)
        self.detect_wave_spin.setSuffix(" nm")
        wave_layout.addWidget(wave_label)
        wave_layout.addWidget(self.detect_wave_spin)
        params_grid.addLayout(wave_layout)
        
        # Threshold
        threshold_layout = QVBoxLayout()
        threshold_label = QLabel('光譜強度:')
        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(0, 10000)
        self.threshold_spin.setValue(1000)
        self.threshold_spin.setSuffix(" a.u.")
        self.threshold_spin.setFixedWidth(125)
        threshold_layout.addWidget(threshold_label)
        threshold_layout.addWidget(self.threshold_spin)
        params_grid.addLayout(threshold_layout)
        
        # Section count
        section_layout = QVBoxLayout()
        section_label = QLabel('解離區段數:')
        self.section_spin = QSpinBox()
        self.section_spin.setRange(2, 10)
        self.section_spin.setValue(3)
        self.section_spin.setFixedWidth(125)
        section_layout.addWidget(section_label)
        section_layout.addWidget(self.section_spin)
        params_grid.addLayout(section_layout)
        
        layout.addLayout(params_grid)
        group.setLayout(layout)
        parent_layout.addWidget(group)

    def _setup_OES_parameters_section(self, parent_layout):
        """Create parameters input section."""
        group = QGroupBox("光譜分析參數設定")
        layout = QVBoxLayout()

        # 添加跳過範圍設定
        skip_layout = QHBoxLayout()
        self.skip_range = QLineEdit("10")
        skip_layout.addWidget(QLabel("最高峰值跳過範圍(nm):"))
        skip_layout.addWidget(self.skip_range)
        layout.addLayout(skip_layout)
        
        # 初始範圍設定
        range_layout = QHBoxLayout()
        self.initial_start = QLineEdit("10")
        self.initial_end = QLineEdit("70")
        range_layout.addWidget(QLabel("光譜預估解離時間(秒數):"))
        range_layout.addWidget(self.initial_start)
        range_layout.addWidget(QLabel("到"))
        range_layout.addWidget(self.initial_end)
        layout.addLayout(range_layout)

        # 添加過濾強度勾選框
        self.filter_checkbox = QCheckBox("過濾低於指定強度")
        self.intensity_threshold = QLineEdit("1000")  # 默認強度閾值
        intensity_layout = QHBoxLayout()
        intensity_layout.addWidget(self.filter_checkbox)
        intensity_layout.addWidget(QLabel("想過濾的強度值:"))
        intensity_layout.addWidget(self.intensity_threshold)
        layout.addLayout(intensity_layout)

        group.setLayout(layout)
        parent_layout.addWidget(group)
        # self.main_layout.addLayout(params_layout)

    def _setup_waveband_settings(self, parent_layout):
        group = QGroupBox("波段設定")
        layout = QVBoxLayout()
        
        # 特定波段設定
        self.wavebands = QLineEdit("486.0, 612.0, 656.0, 777.0")
        layout.addWidget(QLabel("特定波段值(必填) (用逗號分隔):"))
        layout.addWidget(self.wavebands)
        
        # 變化量設定
        self.thresholds = QLineEdit("250, 350, 450, 550")
        layout.addWidget(QLabel("變化量(必填) (用逗號分隔):"))
        layout.addWidget(self.thresholds)
        
        group.setLayout(layout)
        parent_layout.addWidget(group)

    def _setup_image_display(self, parent_layout):
        group = QGroupBox("波長比較圖")
        layout = QVBoxLayout()

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image_label.setMinimumHeight(300)
        self.image_label.setScaledContents(True)
        layout.addWidget(self.image_label)

        # 啟用右鍵選單
        self.image_label.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.image_label.customContextMenuRequested.connect(self._show_image_context_menu)

        group.setLayout(layout)
        parent_layout.addWidget(group)

    def _show_image_context_menu(self, pos):
        """顯示圖片的右鍵選單"""
        menu = QMenu()
        zoom_action = menu.addAction("放大圖片")
        action = menu.exec(self.image_label.mapToGlobal(pos))

        if action == zoom_action:
            self._zoom_image()
    
    def _zoom_image(self):
        """放大圖片的窗口"""
        if hasattr(self, 'output_path'):
            try:
                # 將窗口設為類的屬性
                self.zoom_window = QWidget()
                self.zoom_window.setWindowTitle("放大圖片")
                layout = QVBoxLayout(self.zoom_window)

                zoom_image_label = QLabel()
                zoom_image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

                # 嘗試加載圖片
                pixmap = QPixmap(self.output_path)
                if pixmap.isNull():
                    raise ValueError("無法加載圖片，請檢查路徑。")

                zoom_image_label.setPixmap(pixmap.scaled(800, 600, Qt.AspectRatioMode.KeepAspectRatio))
                layout.addWidget(zoom_image_label)

                self.zoom_window.setLayout(layout)
                self.zoom_window.resize(800, 600)
                self.zoom_window.show()
            except Exception as e:
                print(f"錯誤: {e}")  # 在控制台輸出錯誤信息

    def update_image_display(self, image_path: str):
        """更新圖片顯示"""
        pixmap = QPixmap(image_path)
        scaled_pixmap = pixmap.scaled(self.image_label.size(), 
                                    Qt.AspectRatioMode.KeepAspectRatio,
                                    Qt.TransformationMode.SmoothTransformation)
        self.image_label.setPixmap(scaled_pixmap)

    def _setup_Stability_analysis_section(self, parent_layout):
        """Setup the analysis button section with folder browsing and selection."""
        # # Add a folder browsing button
        # folder_browse_button = QPushButton("選擇資料夾")
        # folder_browse_button.clicked.connect(self._browse_folders)
        # parent_layout.addWidget(folder_browse_button)

        # # Add a dropdown for selecting a folder
        # self.folder_selector = QComboBox()
        # self.folder_selector.currentIndexChanged.connect(self._update_results_display)
        # parent_layout.addWidget(self.folder_selector)

        analyze_button = QPushButton("穩定度分析")
        analyze_button.clicked.connect(self._analyze_data)
        analyze_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        parent_layout.addWidget(analyze_button)

    def _setup_OES_analysis_section(self ,parent_layout):
        """Setup the analysis button section."""
        analyze_button = QPushButton("光譜分析")
        analyze_button.clicked.connect(self._OES_analyze_data)
        analyze_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        parent_layout.addWidget(analyze_button)

    def _setup_results_section(self ,parent_layout):
        """Create results display section."""
        # 創建一個水平佈局來放置表格和圖表
        results_layout = QHBoxLayout()
        
        # 左側表格
        table_layout = QVBoxLayout()
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.setHorizontalHeaderLabels(["區段", "平均值", "標準差", "變異數", "穩定度"])
        
        # 設置表格大小策略
        self.results_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # 設置表格的最小大小
        self.results_table.setMinimumWidth(400)
        self.results_table.setMinimumHeight(300)
        
        # 設置表格的調整策略
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)  # 允許手動調整列寬
        self.results_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)    # 允許手動調整行高
        
        # 設置初始列寬
        self.results_table.setColumnWidth(0, 100)  # 區段
        self.results_table.setColumnWidth(1, 100)  # 平均值
        self.results_table.setColumnWidth(2, 100)  # 標準差
        self.results_table.setColumnWidth(3, 100)  # 變異數
        self.results_table.setColumnWidth(4, 100)  # 穩定度
        
        # 啟用表格的拖動調整功能
        self.results_table.horizontalHeader().setSectionsMovable(True)  # 允許拖動調整列順序
        
        table_layout.addWidget(self.results_table)
        
        # 添加時間信息標籤
        self.time_info_label = QLabel()
        self.time_info_label.setStyleSheet("QLabel { color: #666; font-size: 12px; }")
        table_layout.addWidget(self.time_info_label)
        
        # 啟用右鍵選單
        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._show_context_menu)
        
        save_btn = QPushButton('儲存結果')
        save_btn.clicked.connect(self._save_results)
        table_layout.addWidget(save_btn)
        
        # 右側圖表
        plot_layout = QVBoxLayout()
        self.plot_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        plot_layout.addWidget(self.plot_canvas)
        
        plot_btn = QPushButton('更新Error Bar圖表')
        plot_btn.clicked.connect(self._update_error_bar_plot)
        plot_layout.addWidget(plot_btn)
        
        # 將左右兩側添加到水平佈局
        results_layout.addLayout(table_layout)
        results_layout.addLayout(plot_layout)
        
        # 將水平佈局添加到父佈局
        parent_layout.addLayout(results_layout)

    def _setup_save_directory_selection(self, parent_layout):
            group = QGroupBox("保存路徑設定")
            layout = QVBoxLayout()
            
            # 保存路徑選擇
            save_folder_layout = QHBoxLayout()
            self.save_folder_path = QLineEdit()
            save_browse_button = QPushButton("選擇保存路徑")
            save_browse_button.clicked.connect(self._browse_save_folder)
            save_folder_layout.addWidget(QLabel("保存路徑:"))
            save_folder_layout.addWidget(self.save_folder_path)
            save_folder_layout.addWidget(save_browse_button)
            
            layout.addLayout(save_folder_layout)
            group.setLayout(layout)
            parent_layout.addWidget(group)

    def _setup_extract_waveband_section(self, parent_layout):
        """Setup the extract waveband button section."""
        extract_button = QPushButton("擷取特定波段數據")
        extract_button.clicked.connect(self._extract_specific_waveband_data)
        extract_button.setStyleSheet(""" 
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        parent_layout.addWidget(extract_button)

    def _extract_specific_waveband_data(self):
        """Trigger extraction of specific waveband data."""
        try:
            folder_path = self.path_edit.text()
            save_folder_path = self.save_folder_path.text()
            wavebands = [float(x.strip()) for x in self.wavebands.text().split(",")]
            

            if not folder_path or not save_folder_path:
                QMessageBox.warning(self, "警告", "請選擇資料夾路徑和保存路徑")
                return

            self.controller.extract_specific_waveband_data(
                folder_path=folder_path,
                base_name=self.base_name,
                wavebands=wavebands,
                save_folder_path=save_folder_path
            )

            QMessageBox.information(self, "成功", "特定波段數據已擷取並儲存！")

        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def _browse_save_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "選擇保存路徑")
        if folder:
            self.save_folder_path.setText(folder)

    def _browse_folder(self):
        """Handle folder browsing action."""
        folder_path = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if folder_path:
            self.path_edit.setText(folder_path)

            # 自動掃描 start_index 和 end_index
            try:
                base_name, start_index, end_index = self.controller.scan_file_indices(folder_path)
                self.start_index = start_index
                self.end_index = end_index
                self.base_name = base_name

                QMessageBox.information(
                    self,
                    "成功",
                    f"檢測到檔案範圍：起始索引 {start_index}, 結束索引 {end_index}"
                )

            except Exception as e:
                QMessageBox.critical(self, "錯誤", str(e))

    def _analyze_data(self):
        """Trigger stability analysis through the controller."""
        try:
            if not hasattr(self, 'selected_folders') or not self.selected_folders:
                raise ValueError("請選擇至少一個資料夾進行分析")

            detect_wave = self.detect_wave_spin.value()
            threshold = self.threshold_spin.value()
            section_count = self.section_spin.value()

            # 初始化結果字典
            self.analysis_results = {}
            self.time_info = {}  # 新增：用於存儲時間信息
            failed_folders = []  # 用於存儲分析失敗的資料夾
            
            # 創建進度對話框
            progress = QProgressDialog("正在分析資料夾...", "取消", 0, len(self.selected_folders), self)
            progress.setWindowTitle("分析進度")
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.setMinimumDuration(0)  # 立即顯示進度對話框
            
            # 對每個資料夾進行分析
            for index, folder in enumerate(self.selected_folders):
                # 更新進度對話框
                progress.setValue(index)
                progress.setLabelText(f"正在分析第 {index + 1}/{len(self.selected_folders)} 個資料夾:\n{os.path.basename(folder)}")
                
                # 檢查是否取消
                if progress.wasCanceled():
                    QMessageBox.warning(self, "警告", "分析已被使用者取消")
                    return
                
                try:
                    base_path = folder
                    if not base_path:
                        raise ValueError("請選擇資料夾路徑")
                    logger.info(f"分析第 {index + 1} 筆資料夾: {folder}")
                    # 獲取 base_name, start_index, end_index
                    base_name, start_index, end_index = self.controller.scan_file_indices(folder)
                    # 使用原有的方法加載和處理資料
                    self.controller.load_and_process_data(
                        base_path, 
                        base_name=base_name, 
                        start_index=start_index, 
                        end_index=end_index)

                    # 調用原有的 analyze_data 方法進行分析
                    results_df, activate_time, end_time = self.controller.analyze_data(
                        detect_wave=detect_wave,
                        threshold=threshold,
                        section_count=section_count,
                        base_name=base_name,
                        base_path=base_path,
                        start_index=start_index
                    )
                    # 存儲每個資料夾的結果
                    self.analysis_results[folder] = results_df
                    # 存儲時間信息
                    self.time_info[folder] = (activate_time, end_time)
                except Exception as e:
                    logger.error(f"分析資料夾 {folder} 時發生錯誤: {str(e)}")
                    failed_folders.append((folder, str(e)))
                    continue

            # 完成進度對話框
            progress.setValue(len(self.selected_folders))

            # 更新下拉式選單，只顯示成功分析的資料夾
            self.folder_selector.blockSignals(True)
            self.folder_selector.clear()
            self.folder_selector.addItems([os.path.basename(folder) for folder in self.selected_folders if folder not in [f[0] for f in failed_folders]])
            self.folder_selector.blockSignals(False)

            # 更新結果表格，顯示第一個成功分析的資料夾的結果
            if self.selected_folders:
                first_successful_folder = next((folder for folder in self.selected_folders if os.path.basename(folder) == self.folder_selector.currentText()), None)
                if first_successful_folder:
                    self._update_results_table(self.analysis_results[first_successful_folder])
                    # 更新時間信息
                    if first_successful_folder in self.time_info:
                        activate_time, end_time = self.time_info[first_successful_folder]
                        self.time_info_label.setText(f"Activation Time: {activate_time}, End Time: {end_time}")
            
            # 顯示分析結果摘要
            success_count = len(self.selected_folders) - len(failed_folders)
            fail_count = len(failed_folders)
            
            message = f"分析完成！\n成功分析: {success_count} 個資料夾\n"
            if fail_count > 0:
                message += f"\n分析失敗: {fail_count} 個資料夾\n"
                message += "\n失敗的資料夾列表：\n"
                for folder, error in failed_folders:
                    message += f"- {os.path.basename(folder)}: {error}\n"
            
            QMessageBox.information(self, "分析結果摘要", message)
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def _OES_analyze_data(self):
        try:
            # 獲取所有輸入值
            folder_path = self.path_edit.text()
            save_folder_path = self.save_folder_path.text()
            
            # 檢查輸入的有效性
            if not folder_path or not save_folder_path:
                QMessageBox.warning(self, "警告", "請選擇資料夾路徑和保存路徑")
                return
            
            self.controller.load_and_process_data(
                base_path=folder_path,
                base_name=self.base_name,
                start_index=self.start_index,
                end_index=self.end_index
            )

            base_name = self.base_name
            initial_start = int(self.initial_start.text())
            initial_end = int(self.initial_end.text())
            wavebands = [float(x.strip()) for x in self.wavebands.text().split(",")]
            thresholds = [float(x.strip()) for x in self.thresholds.text().split(",")]
            skip_range_nm = float(self.skip_range.text())
            

            #使用用戶選擇的保存路徑新增資料夾名為
            if os.path.basename(save_folder_path) == "OES光譜分析結果":
                output_directory = save_folder_path
            else:
                output_directory = os.path.join(save_folder_path, "OES光譜分析結果")
                os.makedirs(output_directory,exist_ok=True)
            
            file_paths = [
                os.path.join(folder_path, f"{base_name}_S{str(i).zfill(4)}.txt")
                for i in range(initial_start, initial_end + 1)
            ]
            
            # 調用 Controller 進行分析
            _, _, self.output_path, _ = self.controller.execute_OES_analysis(
                folder_path,
                save_folder_path,
                base_name,
                file_paths,
                initial_start,
                initial_end,
                wavebands,
                thresholds,
                skip_range_nm,
                self.filter_checkbox.isChecked(),
                float(self.intensity_threshold.text()) if self.filter_checkbox.isChecked() else None
            )
            # 檢查 output_path 是否為 None
            if self.output_path is None:
                raise ValueError("分析過程中未生成有效的輸出路徑。")
            
            # 修改圖片名稱以顯示過濾狀態
            if self.filter_checkbox.isChecked():
                filtered_output_path = self.output_path.replace(".png", "_filtered.png")
                # 如果檔案已存在，直接覆蓋
                if os.path.exists(filtered_output_path):
                    os.remove(filtered_output_path)
                os.rename(self.output_path, filtered_output_path)
                self.output_path = filtered_output_path

            self.update_image_display(self.output_path)
            result_message = (
                f"分析完成！結果已保存至：{os.path.basename(save_folder_path)}\n"
            )
            QMessageBox.information(self, "完成", result_message)
        except ValueError as e:
            QMessageBox.critical(self, "輸入錯誤", str(e))
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"分析過程發生錯誤: {str(e)}")

    def _save_results(self):
        """Save analysis results through the controller."""
        try:
            # 新增選擇儲存目錄的對話框
            save_dir = QFileDialog.getExistingDirectory(self, '選擇儲存位置')
            if not save_dir:
                raise ValueError("請選擇資料夾路徑")        

            self.controller.save_results_to_excel(save_dir, self.threshold_spin.value(), self.selected_folders)
            QMessageBox.information(self, "成功", "結果已儲存！")

        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))

    def _update_results_table(self, result):
        """Update the results table with analysis data."""
        self.results_table.setRowCount(0)  # 清空現有的行
        for index, row in result.iterrows():
            row_position = self.results_table.rowCount()
            self.results_table.insertRow(row_position)
            for column, data in enumerate(row):
                self.results_table.setItem(row_position, column, QTableWidgetItem(str(data)))

    def _show_context_menu(self, pos):
        """顯示右鍵選單"""
        menu = QMenu()
        copy_cell_action = menu.addAction("複製當前儲存格")
        copy_row_action = menu.addAction("複製當前行")
        copy_all_action = menu.addAction("複製全部")
        # plot_intensity_action = menu.addAction("儲存單波強度圖")
        view_plot_action = menu.addAction("查看強度圖")
        
        action = menu.exec(self.results_table.mapToGlobal(pos))
        
        if action == copy_cell_action:
            self._copy_cell()
        elif action == copy_row_action:
            self._copy_row()
        elif action == copy_all_action:
            self._copy_all()
        # elif action == plot_intensity_action:
        #     self._generate_intensity_plot()
        elif action == view_plot_action:
            self._view_intensity_plot()

    def _copy_cell(self):
        """複製選中儲存格"""
        if self.results_table.currentItem() is not None:
            clipboard = QApplication.clipboard()
            clipboard.setText(self.results_table.currentItem().text())
            QMessageBox.information(self, "複製成功", "已複製選中儲存格內容")

    def _copy_row(self):
        """複製整行"""
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            row_data = []
            for col in range(self.results_table.columnCount()):
                item = self.results_table.item(current_row, col)
                if item is not None:
                    # 只添加數值內容
                    try:
                        # 嘗試將文本轉換為浮點數
                        value = float(item.text())
                        row_data.append(str(value))
                    except ValueError:
                        # 如果轉換失敗，跳過該項
                        continue
            
            clipboard = QApplication.clipboard()
            clipboard.setText('\t'.join(row_data))
            QMessageBox.information(self, "複製成功", "已複製整行數值內容")

    def _copy_all(self):
        """複製全部內容"""
        all_data = []
        # 添加表頭
        headers = []
        for col in range(self.results_table.columnCount()):
            headers.append(self.results_table.horizontalHeaderItem(col).text())
        all_data.append('\t'.join(headers))
        
        # 添加數據
        for row in range(self.results_table.rowCount()):
            row_data = []
            for col in range(self.results_table.columnCount()):
                item = self.results_table.item(row, col)
                if item is not None:
                    row_data.append(item.text())
            all_data.append('\t'.join(row_data))
        
        clipboard = QApplication.clipboard()
        clipboard.setText('\n'.join(all_data))
        QMessageBox.information(self, "複製成功", "已複製全部內容")

    def _generate_intensity_plot(self):
        """生成選中波段的強度圖"""
        try:
            # 獲取當前選中的資料夾
            selected_folder = self.folder_selector.currentText()
            folder_path = next((folder for folder in self.selected_folders if os.path.basename(folder) == selected_folder), None)
            
            if not folder_path:
                QMessageBox.warning(self, "警告", "請先選擇一個資料夾")
                return

            # 獲取檢測波長
            detect_wave = self.detect_wave_spin.value()
            
            # 獲取保存路徑
            save_dir = QFileDialog.getExistingDirectory(self, '選擇保存位置')
            if not save_dir:
                return

            # 生成圖表
            plt.figure(figsize=(10, 6))
            
            # 讀取所有檔案並繪製強度圖
            files = os.listdir(folder_path)
            spectrum_files = [f for f in files if f.endswith('.txt') and f.startswith(self.base_names[folder_path])]
            
            times = []
            intensities = []
            
            for file_name in sorted(spectrum_files):
                try:
                    file_path = os.path.join(folder_path, file_name)
                    with open(file_path, 'r', encoding='utf-8') as file:
                        for line in file:
                            parts = line.strip().split(';')
                            if len(parts) > 1:
                                try:
                                    wave = float(parts[0])
                                    if abs(wave - detect_wave) < 0.1:  # 允許0.1nm的誤差
                                        time_point = int(file_name.split('_S')[-1].split('.')[0])
                                        intensity = float(parts[1])
                                        times.append(time_point)
                                        intensities.append(intensity)
                                        break
                                except ValueError:
                                    continue
                except Exception as e:
                    logger.error(f"Error processing file {file_name}: {e}")
                    continue

            if not times:
                QMessageBox.warning(self, "警告", f"在波長 {detect_wave}nm 處未找到數據")
                return

            # 繪製圖表
            plt.plot(times, intensities, 'b-', linewidth=2)
            plt.title(f'{detect_wave}nm intensity change')
            plt.xlabel('Time point')
            plt.ylabel('Intensity (a.u.)')
            plt.grid(True)

            # 保存圖表
            output_path = os.path.join(save_dir, f'intensity_plot_{detect_wave}nm.png')
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()

            QMessageBox.information(self, "成功", f"強度圖已保存至：{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"生成強度圖時發生錯誤: {str(e)}")

    def _view_intensity_plot(self):
        """直接查看強度圖"""
        try:
            # 獲取當前選中的資料夾
            selected_folder = self.folder_selector.currentText()
            folder_path = next((folder for folder in self.selected_folders if os.path.basename(folder) == selected_folder), None)
            
            if not folder_path:
                QMessageBox.warning(self, "警告", "請先選擇一個資料夾")
                return

            # 獲取檢測波長
            detect_wave = self.detect_wave_spin.value()
            
            # 讀取所有檔案並繪製強度圖
            files = os.listdir(folder_path)
            spectrum_files = [f for f in files if f.endswith('.txt') and f.startswith(self.base_names[folder_path])]
            
            times = []
            intensities = []
            
            for file_name in sorted(spectrum_files):
                try:
                    file_path = os.path.join(folder_path, file_name)
                    with open(file_path, 'r', encoding='utf-8') as file:
                        for line in file:
                            parts = line.strip().split(';')
                            if len(parts) > 1:
                                try:
                                    wave = float(parts[0])
                                    if abs(wave - detect_wave) < 0.1:  # 允許0.1nm的誤差
                                        time_point = int(file_name.split('_S')[-1].split('.')[0])
                                        intensity = float(parts[1])
                                        times.append(time_point)
                                        intensities.append(intensity)
                                        break
                                except ValueError:
                                    continue
                except Exception as e:
                    logger.error(f"Error processing file {file_name}: {e}")
                    continue

            if not times:
                QMessageBox.warning(self, "警告", f"在波長 {detect_wave}nm 處未找到數據")
                return

            # 創建一個新的窗口來顯示圖表
            plot_window = QDialog(self)  # 使用 QDialog 而不是 QWidget
            plot_window.setWindowTitle(f'Intensity Plot - {detect_wave}nm')
            plot_window.setModal(False)  # 設置為非模態對話框
            layout = QVBoxLayout(plot_window)

            # 創建新的figure
            fig = Figure(figsize=(10, 6))
            canvas = FigureCanvas(fig)
            ax = fig.add_subplot(111)
            
            # 繪製圖表
            ax.plot(times, intensities, 'b-', linewidth=2)
            ax.set_title(f'{detect_wave}nm intensity change')
            ax.set_xlabel('Time point')
            ax.set_ylabel('Intensity (a.u.)')
            ax.grid(True)
            
            # 調整布局
            fig.tight_layout()
            
            # 添加畫布到布局
            layout.addWidget(canvas)

            # 創建按鈕布局
            button_layout = QHBoxLayout()
            
            # 添加保存按鈕
            save_button = QPushButton("保存圖表")
            save_button.clicked.connect(lambda: self._save_plot(fig))
            button_layout.addWidget(save_button)
            
            # 添加關閉按鈕
            close_button = QPushButton("關閉")
            close_button.clicked.connect(plot_window.close)
            button_layout.addWidget(close_button)
            
            # 添加按鈕布局到主布局
            layout.addLayout(button_layout)

            # 設置窗口大小
            plot_window.resize(800, 600)
            
            # 保持對figure的引用，防止被垃圾回收
            plot_window.figure = fig
            plot_window.canvas = canvas
            
            # 顯示窗口
            plot_window.show()
            
            # 確保窗口在最前面
            plot_window.raise_()
            plot_window.activateWindow()

        except Exception as e:
            logger.error(f"生成強度圖時發生錯誤: {str(e)}")
            QMessageBox.critical(self, "錯誤", f"生成強度圖時發生錯誤: {str(e)}")

    def _save_plot(self, figure):
        """保存圖表"""
        try:
            save_dir = QFileDialog.getExistingDirectory(self, '選擇保存位置')
            if not save_dir:
                return

            # 獲取當前選中的資料夾名稱
            selected_folder = self.folder_selector.currentText()
            folder_name = os.path.basename(selected_folder)
            
            # 獲取檢測波長
            detect_wave = self.detect_wave_spin.value()
            
            # 生成包含更多信息的文件名
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            output_path = os.path.join(save_dir, f'{folder_name}_{detect_wave}nm_{timestamp}.png')
            
            figure.savefig(output_path, dpi=300, bbox_inches='tight')
            QMessageBox.information(self, "成功", f"強度圖已保存至：{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"保存圖表時發生錯誤: {str(e)}")

    def _browse_folders(self):
        """Handle folder browsing action for stability analysis using custom dialog."""
        dialog = MultiFolderDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_folders = dialog.get_selected_folders()
            self.folder_selector.blockSignals(True)
            self.folder_selector.clear()
            self.folder_selector.addItems([os.path.basename(folder) for folder in self.selected_folders])
            self.folder_selector.blockSignals(False)
            
            # 對每個選擇的資料夾進行掃描以獲取 base_name, start_index, end_index
            for folder in self.selected_folders:
                try:
                    base_name, start_index, end_index = self.controller.scan_file_indices(folder)
                    # 可以將這些值存儲在一個字典中以便後續使用
                    self.base_names[folder] = base_name
                    self.start_indices[folder] = start_index
                    self.end_indices[folder] = end_index
                except Exception as e:
                    QMessageBox.critical(self, "錯誤", f"在資料夾 {folder} 中掃描檔案時發生錯誤: {str(e)}")

            QMessageBox.information(self, "成功", f"已選擇 {len(self.selected_folders)} 個資料夾進行分析")

    def _update_results_display(self):
        selected_folder = self.folder_selector.currentText()
        folder_path = next((folder for folder in self.selected_folders if os.path.basename(folder) == selected_folder), None)
        # 新增判斷：如果 analysis_results 還沒準備好，直接 return，不跳警告
        if not hasattr(self, 'analysis_results') or not self.analysis_results:
            return
        if folder_path and folder_path in self.analysis_results and self.analysis_results[folder_path] is not None:
            result = self.analysis_results[folder_path]
            self._update_results_table(result)
            if folder_path in self.time_info:
                activate_time, end_time = self.time_info[folder_path]
                self.time_info_label.setText(f"Activation Time: {activate_time}, End Time: {end_time}")
        else:
            # 只有在分析結果真的為空時才跳警告
            if self.analysis_results:  # 只有有分析結果時才跳
                QMessageBox.warning(self, "警告", "未找到選擇的資料夾分析結果")

    def _update_error_bar_plot(self):
        """同時顯示多個實驗的 Error bar 圖表（以穩定度為基準），並支援右鍵放大"""
        try:
            if not hasattr(self, 'analysis_results') or not self.analysis_results:
                QMessageBox.warning(self, "警告", "請先進行分析")
                return

            exp_labels = []
            total_stabilities = []
            lowers = []
            uppers = []

            for idx, folder_path in enumerate(self.selected_folders):
                df = self.analysis_results.get(folder_path)
                if df is None or df.empty:
                    continue
                stabilities = df['穩定度'].tolist()
                total_stability = stabilities[-1]  # 總區段
                min_stability = min(stabilities)
                max_stability = max(stabilities)
                lower = total_stability - min_stability
                upper = max_stability - total_stability
                exp_labels.append(f"Exp.{idx+1}")
                total_stabilities.append(total_stability)
                lowers.append(lower)
                uppers.append(upper)

            if not exp_labels:
                QMessageBox.warning(self, "警告", "沒有可用的分析結果")
                return

            yerr = np.array([lowers, uppers])

            self.plot_canvas.figure.clear()
            ax = self.plot_canvas.figure.add_subplot(111)
            ax.errorbar(range(len(exp_labels)), total_stabilities, yerr=yerr, fmt='o', capsize=10, capthick=2, elinewidth=2, color='blue')
            ax.set_xticks(range(len(exp_labels)))
            ax.set_xticklabels(exp_labels, rotation=0)
            ax.set_title('Total Stability with Section Range (All Experiments)')
            ax.set_xlabel('Experiment')
            ax.set_ylabel('Stability')
            ax.grid(True)
            self.plot_canvas.figure.tight_layout()
            self.plot_canvas.draw()
            self.plot_canvas.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            self.plot_canvas.customContextMenuRequested.connect(self._show_errorbar_context_menu)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"更新圖表時發生錯誤: {str(e)}")

    def _show_errorbar_context_menu(self, pos):
        """顯示 error bar 圖的右鍵選單"""
        menu = QMenu()
        zoom_action = menu.addAction("放大圖表")
        action = menu.exec(self.plot_canvas.mapToGlobal(pos))
        if action == zoom_action:
            self._zoom_errorbar_plot()

    def _zoom_errorbar_plot(self):
        """放大 error bar 圖表的窗口（以穩定度為基準）"""
        try:
            self.zoom_errorbar_window = QWidget()
            self.zoom_errorbar_window.setWindowTitle("放大 Error Bar 圖表")
            layout = QVBoxLayout(self.zoom_errorbar_window)
            fig = Figure(figsize=(10, 6))
            ax = fig.add_subplot(111)
            exp_labels = []
            total_stabilities = []
            lowers = []
            uppers = []
            for idx, folder_path in enumerate(self.selected_folders):
                df = self.analysis_results.get(folder_path)
                if df is None or df.empty:
                    continue
                stabilities = df['穩定度'].tolist()
                total_stability = stabilities[-1]
                min_stability = min(stabilities)
                max_stability = max(stabilities)
                lower = total_stability - min_stability
                upper = max_stability - total_stability
                exp_labels.append(f"Exp.{idx+1}")
                total_stabilities.append(total_stability)
                lowers.append(lower)
                uppers.append(upper)
            yerr = np.array([lowers, uppers])
            ax.errorbar(range(len(exp_labels)), total_stabilities, yerr=yerr, fmt='o', capsize=10, capthick=2, elinewidth=2, color='blue')
            ax.set_xticks(range(len(exp_labels)))
            ax.set_xticklabels(exp_labels, rotation=0)
            ax.set_title('Total Stability with Section Range (All Experiments)')
            ax.set_xlabel('Experiment')
            ax.set_ylabel('Stability')
            ax.grid(True)
            fig.tight_layout()
            canvas = FigureCanvas(fig)
            layout.addWidget(canvas)
            self.zoom_errorbar_window.setLayout(layout)
            self.zoom_errorbar_window.resize(900, 600)
            self.zoom_errorbar_window.show()
            self.zoom_errorbar_window.raise_()
            self.zoom_errorbar_window.activateWindow()
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"放大圖表時發生錯誤: {str(e)}")

class MultiFolderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("選擇多個資料夾")
        self.setMinimumSize(400, 300)

        self.layout = QVBoxLayout(self)

        # List to display selected folders
        self.folder_list = QListWidget()
        self.folder_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.layout.addWidget(self.folder_list)

        # Button to add folders
        self.add_button = QPushButton("添加資料夾")
        self.add_button.clicked.connect(self.add_folder)
        self.layout.addWidget(self.add_button)

        # Button to remove selected folders
        self.remove_button = QPushButton("刪除選中資料夾")
        self.remove_button.clicked.connect(self.remove_selected_folders)
        self.layout.addWidget(self.remove_button)

        # Button to confirm selection
        self.confirm_button = QPushButton("確認")
        self.confirm_button.clicked.connect(self.accept)
        self.layout.addWidget(self.confirm_button)

    def add_folder(self):
        """Open a dialog to select a folder and add its parent folder's contents to the list."""
        folder = QFileDialog.getExistingDirectory(self, "選擇資料夾", "", QFileDialog.Option.ShowDirsOnly | QFileDialog.Option.DontUseNativeDialog)
        if folder:
            parent_folder = os.path.dirname(folder)
            try:
                subfolders = [os.path.join(parent_folder, f) for f in os.listdir(parent_folder) if os.path.isdir(os.path.join(parent_folder, f))]
                for subfolder in subfolders:
                    if subfolder not in [self.folder_list.item(i).text() for i in range(self.folder_list.count())]:
                        self.folder_list.addItem(subfolder)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"無法讀取資料夾內容: {str(e)}")

    def remove_selected_folders(self):
        """Remove selected folders from the list."""
        selected_items = self.folder_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "請選擇要刪除的資料夾")
            return
        
        for item in selected_items:
            self.folder_list.takeItem(self.folder_list.row(item))

    def get_selected_folders(self):
        """Return a list of selected folders."""
        return [self.folder_list.item(i).text() for i in range(self.folder_list.count())]
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = OESAnalyzerGUI()
    gui.show()
    sys.exit(app.exec())