import sys
import os
import pandas as pd
import numpy as np
from scipy.fft import rfft, rfftfreq
from PySide6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QWidget, QTextEdit, QFileDialog, QMessageBox, QGridLayout, 
                             QLabel, QLineEdit, QFrame, QHBoxLayout)
from PySide6.QtCore import QThread, Signal, Qt


# ==========================================
# [FFT 연산 스레드]
# ==========================================
class FFTWorker(QThread):
    log_signal = Signal(str)
    finished_signal = Signal(str)
    error_signal = Signal(str)

    def __init__(self, eng_path, vib_paths, output_path, vib1_title, vib2_title):
        super().__init__()
        self.eng_path = eng_path
        self.vib_paths = vib_paths
        self.output_path = output_path
        self.vib1_title = vib1_title
        self.vib2_title = vib2_title
        self.target_cols = ['Vib1_X', 'Vib1_Y', 'Vib1_Z', 'Vib2_X', 'Vib2_Y', 'Vib2_Z']
        self.freq_limit = 500

    def log(self, message):
        self.log_signal.emit(message)

    def read_csv_with_fallback(self, file_path, encodings=['utf-8', 'cp949', 'euc-kr'], **kwargs):
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding, **kwargs)
                return df
            except Exception:
                continue
        raise Exception(f"🚫 {file_path} 파일을 읽을 수 없습니다.")

    def parse_time_to_seconds(self, series):
        def convert_val(val):
            try:
                s = str(val).strip()
                if s in ['nan', 'None', '', 'NaN']: return np.nan
                parts = s.split(':')
                if len(parts) == 2: return float(parts[0]) * 60 + float(parts[1])
                elif len(parts) == 3: return float(parts[0]) * 3600 + float(parts[1]) * 60 + float(parts[2])
                else: return float(s)
            except: return np.nan
        return series.apply(convert_val)

    def run(self):
        try:
            self.log("📂 Step 1: 데이터 로딩 및 병합 중...")
            all_vib_dfs = []
            for path in self.vib_paths:
                self.log(f"  -> {os.path.basename(path)} 읽는 중...")
                df_vib_raw = self.read_csv_with_fallback(path, header=None)
                df_temp = df_vib_raw.iloc[5:, :9].copy()
                df_temp = df_temp.iloc[:, [0, 3, 4, 5, 6, 7, 8]]
                df_temp.columns = ['Timeline', 'Vib1_X', 'Vib1_Y', 'Vib1_Z', 'Vib2_X', 'Vib2_Y', 'Vib2_Z']
                all_vib_dfs.append(df_temp)
            
            df_vib = pd.concat(all_vib_dfs, axis=0).reset_index(drop=True)
            df_vib['Timeline_sec'] = self.parse_time_to_seconds(df_vib['Timeline'])
            df_vib = df_vib.dropna(subset=['Timeline_sec']).reset_index(drop=True)
            self.log(f"✅ 진동 데이터 통합 완료 (총 {len(df_vib)} 행)")

            self.log(f"  -> {os.path.basename(self.eng_path)} 읽는 중...")
            df_eng_raw = self.read_csv_with_fallback(self.eng_path)
            df_eng = df_eng_raw.iloc[4:, [0, df_eng_raw.columns.get_loc("Engine_Speed"), df_eng_raw.columns.get_loc("PLA Control - A_POS1")]].copy()
            df_eng.columns = ['Timeline', 'Engine_Speed', 'PLA_Control']
            df_eng['Timeline_sec'] = self.parse_time_to_seconds(df_eng['Timeline'])
            df_eng = df_eng.dropna(subset=['Timeline_sec']).reset_index(drop=True)

            df_combined = pd.merge_asof(df_vib, df_eng[['Timeline_sec', 'Engine_Speed', 'PLA_Control']], on='Timeline_sec', direction='nearest')
            df_combined['Engine_Speed'] = pd.to_numeric(df_combined['Engine_Speed'], errors='coerce')
            df_combined = df_combined.dropna(subset=['Engine_Speed']).reset_index(drop=True)

            df_combined.to_csv("df_combined.csv", index=False)

            avg_dt = df_combined['Timeline_sec'].diff().mean()
            if avg_dt <= 0 or pd.isna(avg_dt): raise ValueError("시간 간격 계산 실패")
            fs = 1.0 / avg_dt
            self.log(f"✅ 샘플링 주파수: {fs:.2f} Hz")

            self.log("🔍 Step 2: RPM 구간별 데이터 분할 중...")
            target_sizes = [4096, 2048, 1024, 512, 256]
            segmented_data = {}
            for base_rpm in range(1500, 4300, 100):
                mask = (df_combined['Engine_Speed'] >= base_rpm - 50) & (df_combined['Engine_Speed'] <= base_rpm + 50)
                sub_df = df_combined[mask].copy().reset_index(drop=True)
                if len(sub_df) >= 256:
                    chosen_size = next((s for s in target_sizes if len(sub_df) >= s), 0)
                    if chosen_size > 0:
                        center_idx = (sub_df['Engine_Speed'] - base_rpm).abs().idxmin()
                        start_pos = max(0, center_idx - (chosen_size // 2))
                        end_pos = min(len(sub_df), start_pos + chosen_size)
                        segmented_data[base_rpm] = sub_df.iloc[start_pos:end_pos].copy()

            self.log("🚀 Step 3: FFT 연산 시작...")
            fft_results = {}
            for rpm, sub_df in segmented_data.items():
                n = len(sub_df)
                freqs = rfftfreq(n, d=1/fs)
                rpm_fft_df = pd.DataFrame(index=freqs)
                rpm_fft_df.index.name = 'Frequency(Hz)'
                for col in self.target_cols:
                    yf = rfft(sub_df[col].values)
                    rpm_fft_df[f'{col}_Amp'] = np.abs(yf) * (2.0 / n)
                fft_results[rpm] = rpm_fft_df

            self.log(f"📊 Step 4: 엑셀 파일 및 네이티브 차트 생성 중...")
            with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:
                workbook = writer.book
                for rpm, res_df in fft_results.items():
                    sheet_name = f"{rpm}RPM"
                    plot_df = res_df[res_df.index <= self.freq_limit]
                    if len(plot_df) == 0: continue
                    plot_df.to_excel(writer, sheet_name=sheet_name)
                    worksheet = writer.sheets[sheet_name]
                    num_rows = len(plot_df)
                    cat_range = f"'{sheet_name}'!$A$2:$A${num_rows + 1}"

                    chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                    lp_cols = ['Vib1_X_Amp', 'Vib1_Y_Amp', 'Vib1_Z_Amp']
                    for col_name in lp_cols:
                        if col_name in plot_df.columns:
                            col_idx = plot_df.columns.get_loc(col_name) + 1 
                            col_letter = chr(65 + col_idx) 
                            chart1.add_series({
                                'name':       f"'{sheet_name}'!${col_letter}$1",
                                'categories': cat_range,
                                'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                            })
                    chart1.set_title({'name': self.vib1_title}) 
                    chart1.set_x_axis({'name': 'Frequency (Hz)', 'num_format': '0', 'min':0, 'max':500, 'major_unit': 50})
                    chart1.set_y_axis({'name': 'Amplitude'})
                    chart1.set_size({'width': 525, 'height': 300})
                    worksheet.insert_chart('H2', chart1)

                    chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                    gb_cols = ['Vib2_X_Amp', 'Vib2_Y_Amp', 'Vib2_Z_Amp']
                    for col_name in gb_cols:
                        if col_name in plot_df.columns:
                            col_idx = plot_df.columns.get_loc(col_name) + 1 
                            col_letter = chr(65 + col_idx)
                            chart2.add_series({
                                'name':       f"'{sheet_name}'!${col_letter}$1",
                                'categories': cat_range,
                                'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                            })
                    chart2.set_title({'name': self.vib2_title}) 
                    chart2.set_x_axis({'name': 'Frequency (Hz)', 'num_format': '0', 'min':0, 'max':500, 'major_unit': 50})
                    chart2.set_y_axis({'name': 'Amplitude'})
                    chart2.set_size({'width': 525, 'height': 300})
                    worksheet.insert_chart('H18', chart2)

            self.finished_signal.emit(self.output_path)
        except Exception as e:
            self.error_signal.emit(str(e))


# ==========================================
# [메인 GUI 창]
# ==========================================
class FFTApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.eng_path = None
        self.vib_paths = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("FFT Analysts Tool ver1.0 (26.09.09)")
        self.resize(500, 500)

        main_layout = QVBoxLayout()
        grid_layout = QGridLayout()

        # 버튼 생성
        self.btn_eng = QPushButton("🛠️ 엔진데이터")
        self.btn_vib = QPushButton("📊 진동데이터(복수)")
        self.btn_fft = QPushButton("⚡ FFT 실행")
        self.btn_reset = QPushButton("🔄 초기화")

        btn_size = (100, 30)
        for btn in [self.btn_eng, self.btn_vib, self.btn_fft, self.btn_reset]:
            btn.setMinimumSize(*btn_size)

        grid_layout.addWidget(self.btn_eng, 0, 0)
        grid_layout.addWidget(self.btn_vib, 0, 1)
        grid_layout.addWidget(self.btn_fft, 1, 0)
        grid_layout.addWidget(self.btn_reset, 1, 1)

        # 입력 필드 레이아웃
        input_layout = QVBoxLayout()
        
        vib1_row = QHBoxLayout()
        vib1_row.addWidget(QLabel("그래프 1 타이틀 :"))
        self.edit_vib1 = QLineEdit("VIB1")
        vib1_row.addWidget(self.edit_vib1)
        input_layout.addLayout(vib1_row)

        vib2_row = QHBoxLayout()
        vib2_row.addWidget(QLabel("그래프 2 타이틀 :"))
        self.edit_vib2 = QLineEdit("VIB2")
        vib2_row.addWidget(self.edit_vib2)
        input_layout.addLayout(vib2_row)

        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)

        # 메시지 창
        self.msg_window = QTextEdit()
        self.msg_window.setReadOnly(True)
        self.msg_window.setPlaceholderText("메시지 창입니다...")
        self.btn_fft.setEnabled(False)

        # 전체 레이아웃 결합
        main_layout.addLayout(grid_layout)
        main_layout.addLayout(input_layout)
        main_layout.addWidget(line)
        main_layout.addWidget(self.msg_window)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # 시그널 연결
        self.btn_eng.clicked.connect(self.load_eng_file)
        self.btn_vib.clicked.connect(self.load_vib_files_multi)
        self.btn_fft.clicked.connect(self.run_fft_process)
        self.btn_reset.clicked.connect(self.reset_all)

    def log(self, text):
        self.msg_window.append(text)

    def load_eng_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "엔진 CSV 파일 선택", "", "CSV Files (*.csv);;All Files (*)")
        if file_path:
            self.eng_path = file_path
            self.log(f"📂 엔진 데이터 설정됨: {os.path.basename(file_path)}")
            self.check_ready()

    def load_vib_files_multi(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "진동 CSV 파일들 선택 (파일명 순으로 합쳐짐)", "", "CSV Files (*.csv);;All Files (*)")
        if file_paths:
            file_paths.sort()
            self.vib_paths = file_paths
            names = ", ".join([os.path.basename(p) for p in file_paths])
            self.log(f"📂 진동 데이터(복수) 설정됨:\n{names}")
            self.check_ready()

    def check_ready(self):
        if self.eng_path and self.vib_paths:
            self.btn_fft.setEnabled(True)

    def reset_all(self):
        self.eng_path = None
        self.vib_paths = []
        self.btn_fft.setEnabled(False)
        self.edit_vib1.setText("VIB1")
        self.edit_vib2.setText("VIB2")
        self.msg_window.clear()
        self.log("🔄 모든 설정이 초기화되었습니다.")

    def run_fft_process(self):
        # 1. 먼저 저장할 경로와 파일명을 물어봅니다.
        default_name = "FFT_Analysis_Report.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "결과 저장 경로 선택", default_name, "Excel Files (*.xlsx)")
        
        # 사용자가 취소를 누르면 중단합니다.
        if not save_path:
            return

        # 2. 입력 필드에서 타이틀 정보를 가져옵니다.
        vib1_title = self.edit_vib1.text()
        vib2_title = self.edit_vib2.text()

        # 3. 분석 시작
        self.set_buttons_enabled(False)
        self.worker = FFTWorker(self.eng_path, self.vib_paths, save_path, vib1_title, vib2_title)
        self.worker.log_signal.connect(self.log)
        self.worker.error_signal.connect(self.on_error)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def set_buttons_enabled(self, enabled):
        self.btn_eng.setEnabled(enabled)
        self.btn_vib.setEnabled(enabled)
        self.btn_fft.setEnabled(enabled)
        self.btn_reset.setEnabled(enabled)

    def on_error(self, error_msg):
        self.log(f"❌ 에러 발생: {error_msg}")
        QMessageBox.critical(self, "Error", error_msg)
        self.set_buttons_enabled(True)

    def on_finished(self, final_path):
        self.log(f"✨ 완료! 파일 저장됨: {final_path}")
        QMessageBox.information(self, "Success", f"분석이 완료되었습니다!\n{final_path}")
        self.set_buttons_enabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FFTApp()
    window.show()
    sys.exit(app.exec())
