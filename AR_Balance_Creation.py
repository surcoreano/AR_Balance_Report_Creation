import os
import threading
import platform
import pandas as pd
import numpy as np
import tkinter as tk

from datetime import datetime
from tkinter import filedialog, messagebox, ttk

# ─────────────────────────────
# Theme & Typography (UI 디자인 설정)
# ─────────────────────────────
if platform.system() == "Darwin":
    FN = "Arial"
    FM = "Menlo"
else:
    FN = "Arial"
    FM = "Consolas"

APP_BG = "#f4f6f9"
CARD_BG = "#ffffff"
PRIMARY = "#1a73e8"
PRIMARY_DARK = "#1557b0"
TEXT = "#202124"
SUBTEXT = "#5f6368"
BORDER = "#dadce0"
LOG_BG = "#f8f9fa"

# ─────────────────────────────
# 다국어 지원 텍스트
# ─────────────────────────────
TEXTS = {
    "EN": {
        "title": "Excel Report Generator",
        "header": "AR Balance Report Creator",
        "subtitle": "Create AR Balance Excel File automatically by partners.",
        "language": "Language",
        "excel_file": "Excel file",
        "output_folder": "Output folder",
        "due_date": "Payment Due Date",
        "due_date_hint": "Example: 260625 (YYMMDD)",
        "browse": "Browse",
        "start": "Start process",
        "ready": "Select an Excel file and output folder.",
        "log": "Result Log",
        "excel_error": "Please select a valid Excel file.",
        "folder_error": "Please select a valid Output folder.",
        "date_error": "Please enter the date in YYMMDD format.",
        "confirm_title": "Confirmation",
        "confirm_msg": "Do you want to start generating the report?",
        "cancelled": "Process cancelled.",
        "error": "Error",
        "done": "Completed",
        "reading_excel": "Reading and analyzing Excel data...",
        "processing": "Formatting and filtering data...",
        "saving": "Applying styles and saving Excel file...",
        "completed_status": "Completed successfully. File saved.",
        "completed_msg": "Excel report generation is complete.\n\nSaved location:\n{}"
    },
    "ES": {
        "title": "Generador de Reportes Excel",
        "header": "AR Balance Report Creator",
        "subtitle": "Create AR Balance Excel File automatically by partners.",
        "language": "Idioma",
        "excel_file": "Archivo Excel",
        "output_folder": "Carpeta de destino",
        "due_date": "Fecha de Vencimiento",
        "due_date_hint": "Ejemplo: 260625 (YYMMDD)",
        "browse": "Buscar",
        "start": "Iniciar proceso",
        "ready": "Seleccione un archivo Excel y una carpeta de destino.",
        "log": "Registro de resultados",
        "excel_error": "Seleccione un archivo Excel válido.",
        "folder_error": "Seleccione una carpeta de destino válida.",
        "date_error": "Ingrese la fecha en formato YYMMDD.",
        "confirm_title": "Confirmación",
        "confirm_msg": "¿Desea iniciar la generación del reporte?",
        "cancelled": "Proceso cancelado.",
        "error": "Error",
        "done": "Completado",
        "reading_excel": "Leyendo y analizando datos de Excel...",
        "processing": "Formateando y filtrando datos...",
        "saving": "Aplicando estilos y guardando archivo Excel...",
        "completed_status": "Completado con éxito. Archivo guardado.",
        "completed_msg": "La generación del reporte Excel ha finalizado.\n\nUbicación:\n{}"
    },
    "KR": {
        "title": "Excel 리포트 생성기",
        "header": "AR Balance Report Creator",
        "subtitle": "Create AR Balance Excel File automatically by partners.",
        "language": "언어",
        "excel_file": "원본 Excel 파일",
        "output_folder": "결과물 저장 폴더",
        "due_date": "Payment Due Date",
        "due_date_hint": "예시: 260625 (YYMMDD 형식)",
        "browse": "찾아보기",
        "start": "실행 시작",
        "ready": "Excel 파일과 저장할 폴더를 선택해주세요.",
        "log": "진행 로그",
        "excel_error": "올바른 Excel 파일을 선택해주세요.",
        "folder_error": "올바른 저장 폴더를 선택해주세요.",
        "date_error": "날짜를 정확히 YYMMDD 형식으로 입력해주세요.",
        "confirm_title": "확인",
        "confirm_msg": "리포트 생성을 시작하시겠습니까?",
        "cancelled": "작업이 취소되었습니다.",
        "error": "오류",
        "done": "완료",
        "reading_excel": "Excel 데이터를 읽고 분석하는 중...",
        "processing": "데이터를 포맷팅하고 필터링하는 중...",
        "saving": "서식을 지정하고 정밀 저장하는 중...",
        "completed_status": "성공적으로 완료되었습니다.",
        "completed_msg": "엑셀 리포트 생성이 완료되었습니다.\n\n저장 위치:\n{}"
    }
}

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.lang = tk.StringVar(value="ES")  # 기본 언어 스페인어로 지정 가능
        
        self.excel_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.due_date = tk.StringVar()

        self.title("Excel Report Generator")
        self.geometry("750x700")
        self.configure(bg=APP_BG)

        self._setup_style()
        self._build_ui()
        self._apply_language()

    def t(self, key):
        return TEXTS[self.lang.get()][key]

    def _setup_style(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except tk.TclError: pass

        style.configure("Google.Horizontal.TProgressbar", troughcolor="#e8f0fe", background=PRIMARY, thickness=6)
        style.configure("Google.TCombobox", fieldbackground=APP_BG, background=APP_BG, padding=4)

    def _build_ui(self):
        top_frame = tk.Frame(self, bg=APP_BG, padx=30, pady=20)
        top_frame.pack(fill="x")

        title_inner = tk.Frame(top_frame, bg=APP_BG)
        title_inner.pack(side="left")
        self.header_lbl = tk.Label(title_inner, font=(FN, 20, "bold"), bg=APP_BG, fg=TEXT)
        self.header_lbl.pack(anchor="w")
        self.subtitle_lbl = tk.Label(title_inner, font=(FN, 10), bg=APP_BG, fg=SUBTEXT)
        self.subtitle_lbl.pack(anchor="w", pady=(5, 0))

        lang_inner = tk.Frame(top_frame, bg=APP_BG)
        lang_inner.pack(side="right", anchor="ne")
        self.language_lbl = tk.Label(lang_inner, font=(FN, 9), bg=APP_BG, fg=SUBTEXT)
        self.language_lbl.pack(anchor="w")
        self.language_combo = ttk.Combobox(lang_inner, textvariable=self.lang, values=["EN", "ES", "KR"], state="readonly", width=6, style="Google.TCombobox")
        self.language_combo.pack(pady=(2, 0))
        self.language_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_language())

        main_frame = tk.Frame(self, bg=CARD_BG, padx=30, pady=30, highlightbackground=BORDER, highlightthickness=1)
        main_frame.pack(fill="x", padx=20)

        self.excel_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.excel_lbl.pack(anchor="w")
        excel_row = tk.Frame(main_frame, bg=CARD_BG)
        excel_row.pack(fill="x", pady=(5, 15))
        self.excel_entry = tk.Entry(excel_row, textvariable=self.excel_file, font=(FN, 10), relief="solid", bd=1)
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.excel_btn = tk.Button(excel_row, font=(FN, 9, "bold"), bg=CARD_BG, fg=PRIMARY, width=10, relief="solid", bd=1, command=self._browse_excel)
        self.excel_btn.pack(side="left", padx=(10, 0), ipady=4)

        self.folder_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.folder_lbl.pack(anchor="w")
        folder_row = tk.Frame(main_frame, bg=CARD_BG)
        folder_row.pack(fill="x", pady=(5, 15))
        self.folder_entry = tk.Entry(folder_row, textvariable=self.output_folder, font=(FN, 10), relief="solid", bd=1)
        self.folder_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.folder_btn = tk.Button(folder_row, font=(FN, 9, "bold"), bg=CARD_BG, fg=PRIMARY, width=10, relief="solid", bd=1, command=self._browse_folder)
        self.folder_btn.pack(side="left", padx=(10, 0), ipady=4)

        self.date_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.date_lbl.pack(anchor="w")
        date_row = tk.Frame(main_frame, bg=CARD_BG)
        date_row.pack(fill="x", pady=(5, 2))
        self.date_entry = tk.Entry(date_row, textvariable=self.due_date, font=(FN, 10), relief="solid", bd=1, width=20)
        self.date_entry.pack(side="left", ipady=6)
        self.date_hint_lbl = tk.Label(date_row, font=(FN, 8), bg=CARD_BG, fg=SUBTEXT)
        self.date_hint_lbl.pack(side="left", padx=(10, 0))

        log_frame = tk.Frame(self, bg=CARD_BG, padx=20, pady=20, highlightbackground=BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.log_lbl = tk.Label(log_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.log_lbl.pack(anchor="w", pady=(0, 5))
        
        self.log = tk.Text(log_frame, height=6, font=(FM, 9), bg=LOG_BG, fg=TEXT, relief="solid", bd=1, padx=10, pady=10, state="disabled")
        self.log.pack(fill="both", expand=True)

        bottom_frame = tk.Frame(self, bg=APP_BG, padx=20, pady=10)
        bottom_frame.pack(fill="x", side="bottom")

        self.progress_var = tk.IntVar(value=0)
        status_inner = tk.Frame(bottom_frame, bg=APP_BG)
        status_inner.pack(side="left", fill="x", expand=True, padx=(0, 20))
        
        self.progress = ttk.Progressbar(status_inner, variable=self.progress_var, maximum=100, style="Google.Horizontal.TProgressbar")
        self.progress.pack(fill="x")
        self.status_lbl = tk.Label(status_inner, font=(FN, 9), bg=APP_BG, fg=PRIMARY, anchor="w")
        self.status_lbl.pack(fill="x", pady=(5, 0))

        self.run_btn = tk.Button(bottom_frame, font=(FN, 10, "bold"), bg=PRIMARY, fg="white", activebackground=PRIMARY_DARK, activeforeground="white", relief="flat", padx=20, pady=10, command=self._run)
        self.run_btn.pack(side="right")

    def _apply_language(self):
        self.title(self.t("title"))
        self.header_lbl.config(text=self.t("header"))
        self.subtitle_lbl.config(text=self.t("subtitle"))
        self.language_lbl.config(text=self.t("language"))
        self.excel_lbl.config(text=self.t("excel_file"))
        self.folder_lbl.config(text=self.t("output_folder"))
        self.date_lbl.config(text=self.t("due_date"))
        self.date_hint_lbl.config(text=self.t("due_date_hint"))
        self.excel_btn.config(text=self.t("browse"))
        self.folder_btn.config(text=self.t("browse"))
        self.run_btn.config(text=self.t("start"))
        self.status_lbl.config(text=self.t("ready"))
        self.log_lbl.config(text=self.t("log"))

    def _browse_excel(self):
        path = filedialog.askopenfilename(title=self.t("excel_file"), filetypes=[("Excel files", "*.xlsx *.xls")])
        if path: self.excel_file.set(path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title=self.t("output_folder"))
        if path: self.output_folder.set(path)

    def _log(self, message):
        self.log.config(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, pct, text):
        self.progress_var.set(pct)
        self.status_lbl.config(text=text)
        self.update_idletasks()

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _run(self):
        excel_path = self.excel_file.get().strip()
        output_dir = self.output_folder.get().strip()
        due_date_str = self.due_date.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror(self.t("error"), self.t("excel_error"))
            return
        if not output_dir or not os.path.isdir(output_dir):
            messagebox.showerror(self.t("error"), self.t("folder_error"))
            return
        
        try:
            payment_due_date = datetime.strptime(due_date_str, "%y%m%d")
        except ValueError:
            messagebox.showerror(self.t("error"), self.t("date_error"))
            return

        if not messagebox.askyesno(self.t("confirm_title"), self.t("confirm_msg")):
            self.status_lbl.config(text=self.t("cancelled"))
            return

        self.run_btn.config(state="disabled", bg="#9aa0a6")
        self.progress_var.set(0)
        self._clear_log()
        self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Started processing...")

        def worker():
            try:
                self._set_status(20, self.t("reading_excel"))
                self._log(f"Reading target file: {os.path.basename(excel_path)}")
                
                # 시스템 유도 타이틀 행 무시 처리 자동화 스캔
                temp_df = pd.read_excel(excel_path, header=None, nrows=30)
                header_idx = 0
                for i, row in temp_df.iterrows():
                    if any(pd.notna(val) and 'Invoice No' in str(val) for val in row.values):
                        header_idx = i
                        break
                        
                original_df = pd.read_excel(excel_path, header=header_idx)
                df = original_df.copy()

                self._set_status(40, self.t("processing"))
                self._log("Processing Reference No tracking...")
                
                fill_cols = ['Reference No', 'PO No', 'Comments', 'Invoice Remark']
                for col in fill_cols:
                    if col in df.columns:
                        df[col] = df[col].replace(r'^\s*$', np.nan, regex=True)

                if 'Reference No' not in df.columns:
                    df['Reference No'] = np.nan

                if 'PO No' in df.columns: df['Reference No'] = df['Reference No'].fillna(df['PO No'])
                if 'Comments' in df.columns: df['Reference No'] = df['Reference No'].fillna(df['Comments'])
                if 'Invoice Remark' in df.columns: df['Reference No'] = df['Reference No'].fillna(df['Invoice Remark'])

                self._log("Filtering and renaming columns to Spanish...")
                df.columns = [' '.join(str(c).split()) for c in df.columns]

                columns_to_keep = [
                    'Invoice No.', 'AR Class', 'Trx Date', 'Due Date', 
                    'Original Amount (Entered Curr.)', 'Balance Total', 
                    'Bill To Code', 'Bill To Name', 'Reference No'
                ]
                existing_cols = [col for col in columns_to_keep if col in df.columns]
                df = df[existing_cols]

                rename_dict = {
                    'Invoice No.': 'Factura', 'AR Class': 'Tipología', 'Trx Date': 'Fecha emisión',
                    'Due Date': 'Fecha Vencimiento', 'Original Amount (Entered Curr.)': 'Importe',
                    'Balance Total': 'Balance', 'Bill To Code': 'Código Cliente',
                    'Bill To Name': 'Nombre Cliente', 'Reference No': 'Referencia'
                }
                df = df.rename(columns=rename_dict)

                ar_class_mapping = {'CM': 'Abono', 'DM': 'Devolución recibo', 'INV': 'Factura', 'PMT': 'Importe a cuenta'}
                if 'Tipología' in df.columns:
                    df['Tipología'] = df['Tipología'].replace(ar_class_mapping)

                self._log("Filtering dates and sorting data...")
                if 'Fecha Vencimiento' in df.columns:
                    df['Fecha Vencimiento'] = pd.to_datetime(df['Fecha Vencimiento'], errors='coerce')
                    df = df[df['Fecha Vencimiento'] <= payment_due_date]

                sort_cols = [col for col in ['Nombre Cliente', 'Fecha Vencimiento'] if col in df.columns]
                if sort_cols:
                    df = df.sort_values(by=sort_cols, ascending=[True]*len(sort_cols))

                # 💡 [요구사항 3] 날짜 데이터 형태 정제 (시간 분초 강제 삭제 및 YYYY-MM-DD 화)
                for date_col in ['Fecha emisión', 'Fecha Vencimiento']:
                    if date_col in df.columns:
                        df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')

                # 💡 [요구사항 2] 정해진 스페인어 컬럼 순서 배열 강제 조정
                target_order = ['Código Cliente', 'Nombre Cliente', 'Tipología', 'Factura', 'Fecha emisión', 'Fecha Vencimiento', 'Importe', 'Balance', 'Referencia']
                final_order = [col for col in target_order if col in df.columns]
                df = df[final_order]

                # 숫자 데이터 강제 파싱 및 결측치 수렴
                for num_col in ['Importe', 'Balance']:
                    if num_col in df.columns:
                        df[num_col] = pd.to_numeric(df[num_col], errors='coerce').fillna(0)

                # 💡 [요구사항 5] Summary Sheet 피벗 데이터 빌드
                self._log("Generating Summary Data (Group by Client)...")
                if 'Código Cliente' in df.columns and 'Nombre Cliente' in df.columns and 'Balance' in df.columns:
                    df_summary = df.groupby(['Código Cliente', 'Nombre Cliente'], as_index=False)['Balance'].sum()
                    df_summary = df_summary.sort_values(by='Nombre Cliente', ascending=True)
                else:
                    df_summary = pd.DataFrame(columns=['Código Cliente', 'Nombre Cliente', 'Balance'])

                self._set_status(70, self.t("saving"))
                
                filename = os.path.basename(excel_path)
                name, ext = os.path.splitext(filename)
                output_filepath = os.path.join(output_dir, f"{name}_Report{ext}")

                self._log("Applying strict fonts and colors styles via openpyxl...")

                # 💡 [요구사항 1, 4, 5] 서식 통합 레이아웃 엔진 가동
                with pd.ExcelWriter(output_filepath, engine='openpyxl') as writer:
                    df_summary.to_excel(writer, sheet_name='Summary', index=False)
                    df.to_excel(writer, sheet_name='Report Format', index=False)
                    original_df.to_excel(writer, sheet_name='Original Data', index=False)
                    
                    workbook = writer.book
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter

                    # 폰트 구성 객체 생성 (Arial Narrow, Size 10 고정)
                    font_regular = Font(name='Arial Narrow', size=10)
                    font_bold = Font(name='Arial Narrow', size=10, bold=True)
                    font_red = Font(name='Arial Narrow', size=10, color='FF0000')
                    font_red_bold = Font(name='Arial Narrow', size=10, color='FF0000', bold=True)
                    
                    header_fill = PatternFill(start_color='E8F0FE', end_color='E8F0FE', fill_type='solid')
                    total_fill = PatternFill(start_color='F1F3F4', end_color='F1F3F4', fill_type='solid')
                    
                    thin_side = Side(style='thin', color='DADCE0')
                    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                    
                    # 💡 회계 금융 규격 커스텀 넘버 포맷 코드 (양수, 음수 괄호형, 0은 대시)
                    num_format = '#,##0.00;[Red](#,##0.00);"-"'

                    for sheet_name in ['Summary', 'Report Format', 'Original Data']:
                        if sheet_name not in workbook.sheetnames:
                            continue
                        ws = workbook[sheet_name]
                        col_names = [str(cell.value).strip() for cell in ws[1]]
                        
                        # 헤더 바인딩 서식
                        for cell in ws[1]:
                            cell.font = font_bold
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                            cell.border = thin_border
                        
                        # 전 데이터 셀 순회 적용
                        for row in range(2, ws.max_row + 1):
                            for col in range(1, ws.max_column + 1):
                                cell = ws.cell(row=row, column=col)
                                cell.font = font_regular
                                cell.border = thin_border
                                
                                col_name = col_names[col-1] if col-1 < len(col_names) else ''
                                
                                # 수치 컬럼 탐색 서식 할당
                                if col_name in ['Importe', 'Balance', 'Original Amount (Entered Curr.)', 'Balance Total']:
                                    cell.number_format = num_format
                                    if isinstance(cell.value, (int, float)) and cell.value < 0:
                                        cell.font = font_red
                                    cell.alignment = Alignment(horizontal='right')
                                elif col_name in ['Código Cliente', 'Factura', 'Tipología', 'Fecha emisión', 'Fecha Vencimiento', 'Bill To Code', 'Invoice No.', 'AR Class']:
                                    cell.alignment = Alignment(horizontal='center')
                                else:
                                    cell.alignment = Alignment(horizontal='left')
                        
                        # Summary 하단 토탈 마감 빌드
                        if sheet_name == 'Summary' and ws.max_row > 1:
                            total_row = ws.max_row + 1
                            
                            c_lbl = ws.cell(row=total_row, column=1, value='Total')
                            c_lbl.font = font_bold
                            c_lbl.alignment = Alignment(horizontal='center')
                            c_lbl.fill = total_fill
                            c_lbl.border = thin_border
                            
                            c_emp = ws.cell(row=total_row, column=2, value='')
                            c_emp.fill = total_fill
                            c_emp.border = thin_border
                            
                            if 'Balance' in col_names:
                                b_idx = col_names.index('Balance') + 1
                                b_letter = get_column_letter(b_idx)
                                formula = f"=SUM({b_letter}2:{b_letter}{total_row-1})"
                                
                                c_sum = ws.cell(row=total_row, column=b_idx, value=formula)
                                c_sum.font = font_bold
                                c_sum.number_format = num_format
                                c_sum.alignment = Alignment(horizontal='right')
                                c_sum.fill = total_fill
                                c_sum.border = thin_border
                        
                        # 가시성을 위한 자동 열 너비 계산 조절 알고리즘
                        for col in ws.columns:
                            max_len = 0
                            for cell in col:
                                val_str = str(cell.value or '')
                                if val_str.startswith('='):
                                    val_str = "1,000,000.00"
                                max_len = max(max_len, len(val_str))
                            col_letter = get_column_letter(col[0].column)
                            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

                self._set_status(100, self.t("completed_status"))
                self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Task completed successfully.")
                messagebox.showinfo(self.t("done"), self.t("completed_msg").format(output_filepath))

            except Exception as e:
                self._set_status(0, f"{self.t('error')}: {e}")
                self._log(f"❌ ERROR: {str(e)}")
                messagebox.showerror(self.t("error"), f"An error occurred:\n{str(e)}")
            finally:
                self.run_btn.config(state="normal", bg=PRIMARY)

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()
