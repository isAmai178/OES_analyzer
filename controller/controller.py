import logging
from model.analyzer import OESAnalyzer
import pandas as pd
import os
from typing import Tuple, Optional, List
# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class OESController:
    """
    Controller class for coordinating the interaction between the Model (OESAnalyzer)
    and the View (GUI or other output mechanisms).
    """

    def __init__(self):
        """Initialize the OES Controller with the OESAnalyzer instance."""
        self.analyzer = OESAnalyzer()
        self.analysis_results = None  # To store analysis results

    def load_and_process_data(self, base_path: str, base_name: str, start_index: int, end_index: int) -> None:
        """
        Load data from files and process them.

        Args:
            base_path: Base directory where the data files are located.
            base_name: Base name of the files to process.
            start_index: Starting index of the files.
            end_index: Ending index of the files.

        Returns:
            None
        """
        try:
            logger.info("Generating file names...")
            file_names = self.analyzer.generate_file_names(base_name, start_index, end_index)

            logger.info("Reading and processing data...")
            self.analyzer.read_file_to_data(file_names, base_path)
            logger.info("Data successfully loaded and processed.")

        except Exception as e:
            logger.error(f"Error during data loading and processing: {e}")
            raise
    
    def execute_OES_analysis(self, folder_path, save_folder_path, base_name, file_paths,initial_start,
                initial_end, wavebands, thresholds, skip_range_nm, filter_enabled, intensity_threshold):
            try:
                                
                # activate_time, end_time = self.analyzer.detect_activate_time(detect_wave, thresholds, start_index)
                # print(activate_time,end_time)
                # if activate_time is None or end_time is None:
                #     raise ValueError("Could not detect activation time.")

                self.analyzer.set_files(file_paths)
                # 執行分析
                logger.info("開始分析...")
                output_directory = self.prepare_output_directory(save_folder_path)

                excel_file, specific_excel_file = self.analyzer.OES_analyze_and_export(
                    wavebands=wavebands,
                    thresholds=thresholds,
                    base_name=base_name,
                    skip_range_nm=skip_range_nm,
                    output_directory=output_directory
                )

                # 檢查是否需要過濾低強度波段
                if filter_enabled:
                    self.analyzer.filter_low_intensity(intensity_threshold)

                # 找出並顯示峰值點
                peak_points = self.analyzer.find_peak_points(self.analyzer.all_values)

                # 生成全波段圖
                output_path = self.analyzer.allSpectrum_plot(
                    self.analyzer.all_values,
                    skip_range_nm,
                    output_directory,
                    base_name.split('_')[1]  # 取得檔案前段名稱
                )
                
                return excel_file, specific_excel_file, output_path, peak_points

            except Exception as e:
                logger.error(e)
                raise RuntimeError(f"分析過程發生錯誤: {str(e)}")
                

    def scan_file_indices(self, folder_path: str) -> Tuple[Optional[str], Optional[int], Optional[int]]:
        """
        Scan the folder to find the range of indices for the given base name.

        Args:
            folder_path: Path to the folder containing the files.

        Returns:
            A tuple containing the start and end indices.
        """
        try:
            files = os.listdir(folder_path)
            spectrum_files = [f for f in files if f.endswith('.txt') and '_S' in f]
            
            if not spectrum_files:
                return None, None, None
            
            base_name = spectrum_files[0].split('_S')[0]            
            indices = []
            for file in spectrum_files:
                if file.startswith(base_name):
                    try:
                        index = int(file.split('_S')[-1].replace('.txt', ''))
                        indices.append(index)
                    except ValueError:
                        continue
            
            if not indices:
                return None, None, None
                
            return base_name, min(indices), max(indices)
        
        except Exception as e:
            logger.error(f"Error finding spectrum files: {e}")
            return None, None, None
    def analyze_data(self, detect_wave: float, threshold: float, section_count: int,base_name: str, base_path: str, start_index: int ) -> pd.DataFrame:
        """
        Analyze the processed data and return a DataFrame of results.

        Args:
            detect_wave: Wave length to analyze.
            threshold: Threshold for activation detection.
            section_count: Number of sections for analysis.

        Returns:
            DataFrame containing the analysis results.
        """
        try:
            logger.info("Detecting activation and analyzing data...")

            # Ensure the data for the specific wave exists
            if detect_wave not in self.analyzer._all_data:
                raise ValueError(f"Wave length {detect_wave} not found in the data.")
            # 1. find active time point and end time point
            activate_time, end_time = self.analyzer.detect_activate_time(detect_wave, threshold, start_index)
            print(activate_time,end_time)
            if activate_time is None or end_time is None:
                raise ValueError("Could not detect activation time.")

            # 2. read data of active time period 
            activate_time_file = self.analyzer.generate_file_names(base_name, activate_time + 10, end_time - 10)
            activate_time_data = self.analyzer.read_file_to_data(activate_time_file, base_path)

            wave_data = activate_time_data[detect_wave]
            sectioned_data = self.analyzer.analyze_sections(wave_data, section_count)

            # Store and return results
            self.analysis_results = self.analyzer.prepare_results_dataframe(sectioned_data)
            logger.info("Data analysis completed successfully.")
            return self.analysis_results

        except Exception as e:
            logger.error(f"Error during data analysis: {e}")
            raise

    def save_results_to_excel(self, base_path: str, threshold: float, base_name: str) -> None:
        """
        Save the analysis results to an Excel file.

        Args:
            base_path: Directory where the results should be saved.
            threshold: Threshold value used in the analysis (included in file naming).

        Returns:
            None
        """
        try:
            if self.analysis_results is None:
                raise ValueError("No analysis results to save. Please run the analysis first.")

            excel_path = os.path.join(base_path, f'{base_name}.xlsx')

            with pd.ExcelWriter(excel_path) as writer:
                self.analysis_results.to_excel(writer, sheet_name=f"Threshold_{threshold}", index=False)

            logger.info(f"Results successfully saved to {excel_path}")

        except Exception as e:
            logger.error(f"Error saving results to Excel: {e}")
            raise
    def prepare_output_directory(self, save_folder_path):
        if os.path.basename(save_folder_path) == "OES光譜分析結果":
            return save_folder_path
        else:
            output_directory = os.path.join(save_folder_path, "OES光譜分析結果")
            os.makedirs(output_directory, exist_ok=True)
            return output_directory
        
    def extract_specific_waveband_data(self, folder_path: str, base_name: str, wavebands: List[float], save_folder_path: str) -> None:
        """
        Extract specific waveband data from all files and save to Excel.

        Args:
            folder_path: Path to the folder containing the data files.
            base_name: Base name of the files to process.
            wavebands: List of specific wavebands to extract.
            save_folder_path: Directory where the results should be saved.

        Returns:
            None
        """
        all_data = {}
        
        # 獲取所有檔案
        try:
            files = os.listdir(folder_path)
            spectrum_files = [f for f in files if f.endswith('.txt') and f.startswith(base_name)]
            
            if not spectrum_files:
                logger.warning("No spectrum files found.")
                return
            
            logger.info(f"Found {len(spectrum_files)} files to process.")

            for file_name in spectrum_files:
                file_path = os.path.join(folder_path, file_name)
                try:
                    data = self.analyzer.read_data(file_path)  # 使用 OESAnalyzer 的 read_data 方法
                    for spectral_data in data:
                        if spectral_data.time_point not in all_data:
                            all_data[spectral_data.time_point] = {}
                        for waveband in wavebands:
                            if waveband not in all_data[spectral_data.time_point]:
                                all_data[spectral_data.time_point][waveband] = spectral_data.intensity
                except Exception as e:
                    logger.error(f"Error processing file {file_name}: {e}")
            
            # Save to Excel
            output_file = os.path.join(save_folder_path, f"{base_name}_特定波段數據.xlsx")
            df = pd.DataFrame.from_dict(all_data, orient='index').reset_index()
            df.columns = ['Time Point'] + [f'{wb} nm' for wb in wavebands]
            df.to_excel(output_file, index=False)
            logger.info(f"特定波段數據已被存至 {output_file}")

        except Exception as e:
            logger.error(f"Error scanning files in {folder_path}: {e}")