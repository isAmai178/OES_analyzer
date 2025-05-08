import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QSpinBox,
    QDoubleSpinBox, QTableWidget, QTableWidgetItem, QMessageBox, QMenu,
    QTextEdit, QGroupBox , QHeaderView,  QCheckBox, QGridLayout, QComboBox,
    QDialog, QListWidget, QSizePolicy
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from controller.controller import OESController
import pandas as pd
import os
from typing import List, Dict
import logging
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
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["區段", "平均值", "標準差", "穩定度"])
        self.results_table.setSizeAdjustPolicy(QTableWidget.SizeAdjustPolicy.AdjustToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        parent_layout.addWidget(self.results_table)
        
        # 啟用右鍵選單
        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self._show_context_menu)
        
        save_btn = QPushButton('儲存結果')
        save_btn.clicked.connect(self._save_results)
        parent_layout.addWidget(save_btn)

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
            
            # 對每個資料夾進行分析
            for index, folder in enumerate(self.selected_folders):
                base_path = folder
                if not base_path:
                    raise ValueError("請選擇資料夾路徑")
                logger.info(f"分析第 {index + 1} 筆資料夾: {folder}")
                # 獲取 base_name, start_index, end_index
                base_name, start_index, end_index = self.controller.scan_file_indices(folder)
                # logger.info(base_name, start_index, end_index)
                # 使用原有的方法加載和處理資料
                self.controller.load_and_process_data(
                    base_path, 
                    base_name=base_name, 
                    start_index=start_index, 
                    end_index=end_index)

                # 調用原有的 analyze_data 方法進行分析
                results_df = self.controller.analyze_data(
                    detect_wave=detect_wave,
                    threshold=threshold,
                    section_count=section_count,
                    base_name=base_name,
                    base_path=base_path,
                    start_index=start_index
                )
                # print(results_df)
                # 存儲每個資料夾的結果
                self.analysis_results[folder] = results_df

             # 更新下拉式選單
            self.folder_selector.clear()
            self.folder_selector.addItems([os.path.basename(folder) for folder in self.selected_folders])

            # 更新結果表格，顯示第一個資料夾的結果
            if self.selected_folders:
                self._update_results_table(self.analysis_results[self.selected_folders[0]])
            
            QMessageBox.information(self, "成功", "所有選擇的資料夾已分析完成！")
            print(f"分析結果: {self.analysis_results}")
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

            self.controller.save_results_to_excel(save_dir, self.threshold_spin.value(),self.base_name)
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
        
        action = menu.exec(self.results_table.mapToGlobal(pos))
        
        if action == copy_cell_action:
            self._copy_cell()
        elif action == copy_row_action:
            self._copy_row()
        elif action == copy_all_action:
            self._copy_all()

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

    def _browse_folders(self):
        """Handle folder browsing action for stability analysis using custom dialog."""
        dialog = MultiFolderDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_folders = dialog.get_selected_folders()
            self.folder_selector.clear()
            self.folder_selector.addItems([os.path.basename(folder) for folder in self.selected_folders])
            
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
        """Update the results table based on the selected folder."""
        selected_folder = self.folder_selector.currentText()  # 獲取選擇的資料夾名稱
        folder_path = next((folder for folder in self.selected_folders if os.path.basename(folder) == selected_folder), None)
        if folder_path and folder_path in self.analysis_results:
            result = self.analysis_results[folder_path]  # 獲取對應的分析結果
            self._update_results_table(result)  # 更新結果表格
        else:
            QMessageBox.warning(self, "警告", "未找到選擇的資料夾分析結果")

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