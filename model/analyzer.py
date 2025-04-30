import os
from typing import List, Dict, Tuple, Optional, Callable
import logging
from dataclasses import dataclass
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@dataclass
class SpectralData:
    """Data class for storing spectral measurement data."""
    time_point: float
    intensity: float

class OESAnalyzer:
    """
    Analyzer class for processing OES data.

    This class provides methods for reading, processing, and analyzing
    spectral data from OES measurements.
    """

    def __init__(self):
        """Initialize the OES Analyzer."""
        self._all_data: Dict[float, List[float]] = {}
        logger.info("OES Analyzer initialized")

    @staticmethod
    def generate_file_names(base_name: str, start: int, end: int, extension: str = '.txt') -> List[str]:
        """
        Generate a list of file names based on the parameters.

        Args:
            base_name: Base name for the files
            start: Starting index
            end: Ending index
            extension: File extension (default: '.txt')

        Returns:
            List of generated file names
        """
        return [f"{base_name}_S{str(i).zfill(4)}{extension}" for i in range(start, end + 1)]

    def read_data(self, file_path: str) -> List[SpectralData]:
        """
        Read data from a given file.

        Args:
            file_path: Path to the data file

        Returns:
            List of SpectralData objects containing time points and intensities
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                contents = file.readlines()

            data = []
            for line in contents:
                if ';' in line:
                    try:
                        time_point, intensity = map(float, line.strip().split(';'))
                        data.append(SpectralData(time_point, intensity))
                    except ValueError as e:
                        logger.warning(f"Skipping invalid line in {file_path}: {e}")
                        continue

            logger.debug(f"Successfully read {len(data)} data points from {file_path}")
            return data

        except Exception as e:
            logger.error(f"Error reading file {file_path}: {e}")
            raise

    def read_file_to_data(self, file_names: List[str], base_path: str) -> Dict[float, List[float]]:
        """
        Read all files and store data.

        Args:
            file_names: List of file names to process
            base_path: Base path for the files

        Returns:
            Dictionary mapping time points to lists of intensity values
        """
        self._all_data.clear()

        for file_name in file_names:
            try:
                file_path = os.path.join(base_path, file_name)
                data = self.read_data(file_path)

                for spectral_data in data:
                    if spectral_data.time_point not in self._all_data:
                        self._all_data[spectral_data.time_point] = []
                    self._all_data[spectral_data.time_point].append(spectral_data.intensity)
            
            except Exception as e:
                logger.error(f"Error processing file {file_name}: {e}")
                continue

        logger.info(f"Processed {len(file_names)} files with {len(self._all_data)} time points")
        return self._all_data
    
    def set_files(self, file_paths: List[str]):
        """設置要分析的文件列表"""
        self.selected_files = file_paths

    def read_values_by_line(self, file_path: str) -> Dict[float, float]:
        """讀取單個文件中的value和測量值"""
        values = {}
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                for line in file:
                    parts = line.strip().split(';')
                    if len(parts) > 1:
                        try:
                            value = float(parts[0])
                            if value >= 195.0:
                                values[value] = float(parts[1])
                        except ValueError:
                            logger.info(f"Skipping line due to ValueError: {line.strip()}")
        except FileNotFoundError:
            logger.info(f"The file at {file_path} was not found.")
        except Exception as e:
            logger.info(f"An error occurred: {e}")
        return values
    
    def gather_values(self) -> Dict:
        """收集所有文件的數據"""
        self.all_values = {}
        for file_path in self.selected_files:
            file_values = self.read_values_by_line(file_path)
            if not file_values:
                logger.info(f"No valid data found in {file_path}")
                continue
            for value, measurement in file_values.items():
                if value not in self.all_values:
                    self.all_values[value] = []
                self.all_values[value].append((os.path.basename(file_path), measurement))
        # logger.info(self.all_values)
        return self.all_values

    def find_peak_points(self, data: Dict[float, List[Tuple[str, float]]]) -> List[dict]:
        """找出每個波段的最高點"""
        peak_points = []
        for value, measurements in data.items():
            measurements_only = [m[1] for m in measurements]
            max_value = max(measurements_only)
            max_index = measurements_only.index(max_value)
            file_name = measurements[max_index][0]
            
            peak_points.append({
                '波段': value,
                '最大值': max_value,
                '檔案名': file_name,
                '時間點': file_name.split('S')[-1].split('.')[0]
            })
        # logger.info(peak_points)
        # 按最大值排序
        return sorted(peak_points, key=lambda x: x['最大值'], reverse=True)
    
    def find_specific_wavebands_differences(self, wavebands: List[float], threshold: float = 200) -> Dict:
        """分析特定波段的差異"""
        specific_differences = {}
        for value, measurements in self.all_values.items():
            # logger.info(value,measurements)
            if value in wavebands:
                measurements_only = [m[1] for m in measurements]
                min_measurement = min(measurements_only)
                max_measurement = max(measurements_only)
                if abs(max_measurement - min_measurement) > threshold:
                    largest_diff = max(measurements, key=lambda x: abs(x[1] - min_measurement))
                    filename = largest_diff[0]
                    file_second = filename.split('S')[-1].split('.')[0]
                    specific_differences[value] = (min_measurement, max_measurement, largest_diff, file_second)
            
        return specific_differences

    def find_significant_differences(self, threshold: float = 200) -> Dict:
        """分析所有波段的顯著差異"""
        significant_differences = {}
        for value, measurements in self.all_values.items():
            measurements_only = [m[1] for m in measurements]
            min_measurement = min(measurements_only)
            max_measurement = max(measurements_only)
            if abs(max_measurement - min_measurement) > threshold:
                largest_diff = max(measurements, key=lambda x: abs(x[1] - min_measurement))
                significant_differences[value] = (min_measurement, max_measurement, largest_diff)
        return significant_differences

    def allSpectrum_plot(self, data1, skip_range_nm, output_directory, file_name, intensity_threshold=None):
        """繪製全波段圖形並標記出最高波段"""
        try:
            # 過濾低於指定強度的波型
            if intensity_threshold is not None:
                data1 = {k: v for k, v in data1.items() if any(m[1] > intensity_threshold for m in v)}
            
            # 找出每個數據集的最大值點
            peaks1 = self.find_peak_points(data1)

            # 取得最大值的波長
            max_peak1 = peaks1[0]  # 已經按最大值排序，所以第一個就是最大的
            sorted_peaks = sorted(peaks1, key=lambda x: x['最大值'], reverse=True)

            # 準備數據
            wavelengths1 = sorted(data1.keys())
            y1 = [max(m[1] for m in data1[w]) for w in wavelengths1]
        
            # 創建圖表
            plt.figure(figsize=(10, 6))

            # 添加最大值波段信息到標題
            title_text = (f'ALL_Spectrum & Higher Peaks \n'
                        f'Max_peak: {max_peak1["波段"]:.1f}nm')
            plt.title(title_text)
            
            # 繪製線條
            plt.plot(wavelengths1, y1, color='red', label='Highest_data', linewidth=1)

            marked_peaks = []
            for index, peak in enumerate(sorted_peaks):
                if len(marked_peaks) >= 5:
                    break
                
                # 檢查是否需要跳過範圍
                if not any(abs(peak['波段'] - marked_peak['波段']) <= skip_range_nm for marked_peak in marked_peaks):
                    # 調整標註位置以避免重疊
                    offset = len(marked_peaks) * 10  # 根據已標註的數量調整偏移量
                    rotation_angle = 0
                    # Adjust annotation to connect to the left side of the x-axis with dashed lines
                    xytext_offset = (-50, -50)  # Position to the left of the x-axis
                    plt.annotate(f'intensity: {peak["最大值"]:.1f}',
                                xy=(peak['波段'], peak['最大值']),
                                xytext=xytext_offset, textcoords='offset points', 
                                arrowprops=dict(arrowstyle='->', lw=1.5, linestyle='dashed'),
                                rotation=rotation_angle)  # 根據條件設置旋轉角度
                    marked_peaks.append(peak)

            x_ticks = [peak['波段'] for peak in marked_peaks]
            
            plt.xticks(ticks=x_ticks, labels=[f'{wavelengths1:.1f}nm' for wavelengths1 in x_ticks], rotation=75)        
            # Add annotations for the top three peak intensities on the y-axis
            
            # 設置圖表屬性
            plt.xlabel('Wavelength(nm)')
            plt.ylabel('Intensity(Cts)')
            # plt.legend()
            
            # 建構檔案名稱
            output_file_name = f"{file_name}_allspectrum_highestPeaks.png"

            # 保存圖表
            output_path = os.path.join(output_directory, output_file_name)
            plt.savefig(output_path, dpi=300, bbox_inches='tight')
            plt.close()
        
            logger.info(f"已生成最大值比較圖：{output_path}")
            return output_path
        
        except Exception as e:
            logger.info(f"生成比較圖時發生錯誤: {str(e)}")
            return None

    def analyze_sections(self, wave_data: List[float], section: int) -> Dict[str, Dict[str, float]]:
        """
        Analyze wave data in sections.

        Args:
            wave_data: List of wave data points
            section: Number of sections

        Returns:
            Dictionary containing analysis results for each section
        """
        section_size = len(wave_data) // section
        sectioned_data = {}

        # Analyze individual sections
        for i in range(section):
            start_idx = i * section_size
            end_idx = start_idx + section_size if i < section - 1 else len(wave_data)
            section_wave_data = wave_data[start_idx:end_idx]

            std_value = np.std(section_wave_data)
            average_value = np.mean(section_wave_data)
            stability = round((std_value / average_value) * 100, 3)

            sectioned_data[f'區段{i+1}'] = {
                'mean': average_value,
                'std': std_value,
                '穩定度': stability
            }

        # Analyze total section
        total_mean = np.mean(wave_data)
        total_std = np.std(wave_data)
        total_stability = round((total_std / total_mean) * 100, 3)

        sectioned_data['總區段'] = {
            'mean': total_mean,
            'std': total_std,
            '穩定度': total_stability
        }

        return sectioned_data

    def OES_analyze_and_export(self, wavebands: List[float], thresholds: List[float], 
                           base_name, skip_range_nm: float, output_directory: str) -> Tuple[str, str]:
        """執行分析並導出結果"""
        self.gather_values()
        # 使用傳遞的 output_directory
        os.makedirs(output_directory, exist_ok=True)
        # 處理特定波段數據
        specific_excel_name = os.path.join(output_directory, f"{base_name}_特定波段解離情況.xlsx")
        # logger.info(specific_excel_name)
        with pd.ExcelWriter(specific_excel_name) as specific_writer:
            for threshold in thresholds:
                specific_differences = self.find_specific_wavebands_differences(wavebands, threshold)
                if specific_differences:
                    specific_data = []
                    for value, (min_measurement, max_measurement, largest_diff, _) in sorted(specific_differences.items()):
                        specific_data.append({
                            '波段': value,
                            '最小值': min_measurement,
                            '最大值': max_measurement,
                            '差值': max_measurement - min_measurement
                        })
                    specific_df = pd.DataFrame(specific_data)
                    specific_df.to_excel(specific_writer, sheet_name=f"threshold_{threshold}", index=False)
                else:
                    # Add a default sheet if no data is available
                    pd.DataFrame({'Message': ['No data available for this threshold']}).to_excel(specific_writer, sheet_name=f"threshold_{threshold}", index=False)

        # 處理所有波段數據
        excel_name = os.path.join(output_directory, f"{base_name}_全部解離波段.xlsx")
        with pd.ExcelWriter(excel_name) as writer:
            for threshold in thresholds:
                significant_differences = self.find_significant_differences(threshold)
                if significant_differences:
                    data = []
                    for value, (min_measurement, max_measurement, largest_diff) in sorted(significant_differences.items()):
                        data.append({
                            '波段': value,
                            '最小值': min_measurement,
                            '最大值': max_measurement,
                            '差值': max_measurement - min_measurement
                        })
                    df = pd.DataFrame(data)
                    df.to_excel(writer, sheet_name=f"threshold_{threshold}", index=False)
                else:
                    # Add a default sheet if no data is available
                    pd.DataFrame({'Message': ['No data available for this threshold']}).to_excel(writer, sheet_name=f"threshold_{threshold}", index=False)

        return excel_name, specific_excel_name

    def filter_low_intensity(self, threshold: float):
        """將低於指定強度的波段設置為0"""
        if not self.all_values:
            print("all_values is empty")
            return
        
        for value, measurements in self.all_values.items():
            # print(f"Processing value: {value}, measurements: {measurements}")
            for i, (file_name, intensity) in enumerate(measurements):
                # print(f"File: {file_name}, Intensity: {intensity}")
                if intensity < threshold:
                    self.all_values[value][i] = (file_name, 0.0)

    def prepare_results_dataframe(self, sectioned_data: Dict[str, Dict[str, float]]) -> pd.DataFrame:
        """
        Prepare results DataFrame from sectioned data.

        Args:
            sectioned_data: Dictionary containing analysis results

        Returns:
            DataFrame containing formatted results
        """
        results = []
        for section_name, stats in sectioned_data.items():
            results.append([
                section_name,
                stats['mean'],
                stats['std'],
                stats['穩定度']
            ])

        return pd.DataFrame(results, columns=['區段', '平均值', '標準差', '穩定度'])

    def detect_activate_time(self, max_wave: float, threshold: float, start_index: int) -> Tuple[Optional[int], Optional[int]]:
        """
        Find activation time points.

        Args:
            max_wave: Wave length to analyze
            threshold: Threshold for activation detection
            start_index: Starting index for the analysis

        Returns:
            Tuple of activation start and end times
        """
        if max_wave not in self._all_data:
            logger.error(f"Wave length {max_wave} not found in data")
            return None, None

        time_series = self._all_data[max_wave]
        activated = False
        activate_time = None
        end_time = None
        
        for i in range(len(time_series) - 1):
            diff = time_series[i + 1] - time_series[i]
            
            if not activated and diff > threshold:
                activate_time = i + 1 + start_index
                activated = True
                logger.debug(f"Activation detected at index {activate_time}")
            elif activated and diff < -threshold:
                end_time = i + 1 + start_index
                logger.debug(f"Deactivation detected at index {end_time}")
                break
                
        return activate_time, end_time