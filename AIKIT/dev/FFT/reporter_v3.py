import pandas as pd
import re
import numpy as np
import argparse
import os

def main():
    # 1. Argument Parser 설정 (파일 경로만 받음)
    parser = argparse.ArgumentParser(description="FFT Analysis Report Generator")
    parser.add_argument('--file', type=str, required=True, help="입력할 엑셀 파일 경로 (예: test.xlsx)")
    args = parser.parse_args()

    file_name = args.file

    # 파일 존재 여부 확인
    if not os.path.exists(file_name):
        print(f"에러: 파일을 찾을 수 없습니다: {file_name}")
        return

    base, ext = os.path.splitext(file_name)
    output_file = f"{base}_REPORT{ext}"

    # --- 설정 ---
    # 모든 시트를 읽어옴
    all_sheets = pd.read_excel(file_name, engine='openpyxl', sheet_name=None)

    def extract_rpm(sheet_name):
        match = re.search(r'\d+', sheet_name)
        return int(match.group()) if match else None

    # 분석 대상 컬럼 정의
    target_cols = ['Vib1_X_Amp', 'Vib1_Y_Amp', 'Vib1_Z_Amp', 'Vib2_X_Amp', 'Vib2_Y_Amp', 'Vib2_Z_Amp']
    mod_cols = [f"{col}_mod" for col in target_cols]
    updated_sheets = {}
    
    # Summary를 위한 데이터 리스트
    summary_data = []

    # 3. 데이터 처리
    for sheet_name, df in all_sheets.items():
        rpm = extract_rpm(sheet_name)
        if rpm is None:
            print(f"경고: '{sheet_name}' 시트에서 RPM 숫자를 찾을 수 없어 건너뜁니다.")
            continue

        # Frequency(Hz) 컬럼이 있는지 확인
        if 'Frequency(Hz)' not in df.columns:
            print(f"경고: '{sheet_name}' 시트에 'Frequency(Hz)' 컬럼이 없습니다. 건너뜁니다.")
            continue

        # 26.09.15 4200RPM 확장하면서 130에서 140으로 변경
        # max_freq = 130
        max_freq = 140

        temp_df = df[df['Frequency(Hz)'] <= max_freq].copy()

        if temp_df.empty:
            print(f"경고: '{sheet_name}' 시트는 필터링 후 데이터가 없어 건너뜁니다.")
            continue

        # Order 계산 및 _mod 컬럼 생성
        order_col_name = 'temp_order_calc'
        temp_df[order_col_name] = (temp_df['Frequency(Hz)'] * 60) / rpm

        for col in target_cols:
            if col in temp_df.columns:
                new_col_name = f"{col}_mod"
                condition = (temp_df[order_col_name] < 2.7)
                temp_df[new_col_name] = np.where(condition, temp_df[col], 0)
            else:
                # 컬럼이 없는 경우 0으로 채움 (에러 방지)
                temp_df[f"{col}_mod"] = 0
        
        temp_df = temp_df.drop(columns=[order_col_name])
        updated_sheets[sheet_name] = temp_df

        # --- Peak 값 추출 (Summary용) ---
        peak_values = {'RPM': rpm}
        for m_col in mod_cols:
            if m_col in temp_df.columns:
                non_zero_vals = temp_df[m_col][temp_df[m_col] > 0]
                peak_values[m_col] = non_zero_vals.max() if not non_zero_vals.empty else 0
            else:
                peak_values[m_col] = 0
        summary_data.append(peak_values)

    if not summary_data:
        print("처리할 유효한 데이터가 없습니다.")
        return

    # Summary 데이터프레임 생성 및 RPM 기준 정렬
    summary_df = pd.DataFrame(summary_data).set_index('RPM').sort_index().reset_index()

    # 4. 엑셀 저장 및 차트 생성
    print(f"Summary 생성 및 파일 저장 중: {output_file}")

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- [Step 1] Summary 시트 작성 ---
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        summary_ws = writer.sheets['Summary']
        
        num_summary_rows = len(summary_df)
        
        # 그룹 정의 (Summary 차트용)
        vib1_mod_cols = ['Vib1_X_Amp_mod', 'Vib1_Y_Amp_mod', 'Vib1_Z_Amp_mod']
        vib2_mod_cols = ['Vib2_X_Amp_mod', 'Vib2_Y_Amp_mod', 'Vib2_Z_Amp_mod']

        def get_col_letter_from_df(df, col_name):
            """데이터프레임의 컬럼명을 엑셀 열 문자로 변환"""
            try:
                idx = df.columns.get_loc(col_name)
                if idx < 26: return chr(65 + idx)
                else: return f"'{chr(65 + (idx // 26) - 1)}{chr(65 + (idx % 26))}'"
            except: return None

        # --- Chart 1: VIB1 Peak Trend ---
        chart1 = workbook.add_chart({'type': 'line'})
        has_vib1_chart = False
        for col_name in vib1_mod_cols:
            if col_name in summary_df.columns:
                col_letter = get_col_letter_from_df(summary_df, col_name)
                if col_letter:
                    chart1.add_series({
                        'name':       f"'{col_name}'",
                        'categories': f"Summary!$A$2:$A${num_summary_rows + 1}",
                        'values':     f"Summary!${col_letter}$2:${col_letter}${num_summary_rows + 1}",
                    })
                    has_vib1_chart = True
        if has_vib1_chart:
            chart1.set_title({'name': 'VIB1(LP TC) Peak Amplitude Trend (MOD Filtered)'})
            chart1.set_x_axis({'name': 'RPM'})
            chart1.set_y_axis({'name': 'Amplitude'})
            chart1.set_size({'width': 600, 'height': 300})
            summary_ws.insert_chart('H2', chart1)

        # --- Chart 2: VIB2 Peak Trend ---
        chart2 = workbook.add_chart({'type': 'line'})
        has_vib2_chart = False
        for col_name in vib2_mod_cols:
            if col_name in summary_df.columns:
                col_letter = get_col_letter_from_df(summary_df, col_name)
                if col_letter:
                    chart2.add_series({
                        'name':       f"'{col_name}'",
                        'categories': f"Summary!$A$2:$A${num_summary_rows + 1}",
                        'values':     f"Summary!${col_letter}$2:${col_letter}${num_summary_rows + 1}",
                    })
                    has_vib2_chart = True
        if has_vib2_chart:
            chart2.set_title({'name': 'VIB2(GEAR BOX) Peak Amplitude Trend (MOD Filtered)'})
            chart2.set_x_axis({'name': 'RPM'})
            chart2.set_y_axis({'name': 'Amplitude'})
            chart2.set_size({'width': 600, 'height': 300})
            summary_ws.insert_chart('H22', chart2)

        # --- [Step 2] 개별 시트 작성 ---
        for sheet_name, df in updated_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            
            num_rows = len(df)
            cat_range = f"'{sheet_name}'!$A$2:$A${num_rows + 1}"

            def get_col_letter(df, col_name):
                try:
                    idx = df.columns.get_loc(col_name)
                    if idx < 26: return chr(65 + idx)
                    else: return f"'{chr(65 + (idx // 26) - 1)}{chr(65 + (idx % 26))}'"
                except KeyError: return None

            # Chart 1: LP TC (Vib1_X, Y, Z) Original
            lp_cols = [c for c in ['Vib1_X_Amp', 'Vib1_Y_Amp', 'Vib1_Z_Amp'] if c in df.columns]
            if lp_cols:
                chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                for col_name in lp_cols:
                    col_letter = get_col_letter(df, col_name)
                    if col_letter:
                        chart1.add_series({
                            'name':       f"'{sheet_name}'!${col_letter}$1",
                            'categories': cat_range,
                            'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                        })
                chart1.set_title({'name': 'VIB1(LP TC) (Original)'})
                chart1.set_x_axis({'name': 'Frequency (Hz)'})
                chart1.set_y_axis({'name': 'Amplitude'})
                chart1.set_size({'width': 500, 'height': 300})
                worksheet.insert_chart('N2', chart1)

            # Chart 2: LP TC MOD
            lp_mod_cols = [f"{c}_mod" for c in lp_cols]
            if lp_mod_cols:
                chart1_mod = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                for col_name in lp_mod_cols:
                    col_letter = get_col_letter(df, col_name)
                    if col_letter:
                        chart1_mod.add_series({
                            'name':       f"'{sheet_name}'!${col_letter}$1",
                            'categories': cat_range,
                            'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                        })
                chart1_mod.set_title({'name': 'VIB1(LP TC) (MOD Filtered)'})
                chart1_mod.set_x_axis({'name': 'Frequency (Hz)'})
                chart1_mod.set_y_axis({'name': 'Amplitude'})
                chart1_mod.set_size({'width': 500, 'height': 300})
                worksheet.insert_chart('V2', chart1_mod)

            # Chart 3: GEARBOX (Vib2_X, Y, Z) Original
            gb_cols = [c for c in ['Vib2_X_Amp', 'Vib2_Y_Amp', 'Vib2_Z_Amp'] if c in df.columns]
            if gb_cols:
                chart2 = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                for col_name in gb_cols:
                    col_letter = get_col_letter(df, col_name)
                    if col_letter:
                        chart2.add_series({
                            'name':       f"'{sheet_name}'!${col_letter}$1",
                            'categories': cat_range,
                            'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                        })
                chart2.set_title({'name': 'VIB2(GEAR BOX) (Original)'})
                chart2.set_x_axis({'name': 'Frequency (Hz)'})
                chart2.set_y_axis({'name': 'Amplitude'})
                chart2.set_size({'width': 500, 'height': 300})
                worksheet.insert_chart('N18', chart2)

            # Chart 4: GEARBOX MOD
            gb_mod_cols = [f"{c}_mod" for c in gb_cols]
            if gb_mod_cols:
                chart2_mod = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
                for col_name in gb_mod_cols:
                    col_letter = get_col_letter(df, col_name)
                    if col_letter:
                        chart2_mod.add_series({
                            'name':       f"'{sheet_name}'!${col_letter}$1",
                            'categories': cat_range,
                            'values':     f"'{sheet_name}'!${col_letter}$2:${col_letter}${num_rows + 1}",
                        })
                chart2_mod.set_title({'name': 'VIB2(GEAR BOX) (MOD Filtered)'})
                chart2_mod.set_x_axis({'name': 'Frequency (Hz)'})
                chart2_mod.set_y_axis({'name': 'Amplitude'})
                chart2_mod.set_size({'width': 500, 'height': 300})
                worksheet.insert_chart('V18', chart2_mod)

    print(f"\n모든 작업이 완료되었습니다: {output_file}")

if __name__ == "__main__":
    main()
