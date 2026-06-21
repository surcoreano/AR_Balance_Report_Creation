import os
import re
import platform
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime


# ─────────────────────────────
# Theme
# ─────────────────────────────
if platform.system() == "Darwin":
    FN = "Apple SD Gothic Neo"
    FM = "Menlo"
else:
    FN = "Segoe UI"
    FM = "Consolas"

APP_BG = "#f8fafd"
CARD_BG = "#ffffff"
PRIMARY = "#1a73e8"
PRIMARY_DARK = "#1558b0"
TEXT = "#202124"
SUBTEXT = "#5f6368"
BORDER = "#dadce0"
LOG_BG = "#f1f3f4"


TEXTS = {
    "EN": {
        "title": "Netting PDF Splitter",
        "header": "Netting PDF Splitter",
        "subtitle": "Split PDF files automatically by trading partner.",
        "select_pdf": "PDF file",
        "select_folder": "Output folder",
        "date_label": "File name date",
        "date_hint": "YYMMDD",
        "example": "Example: ES001627001B_DAROCA CASTILLO_260420.pdf",
        "browse": "Browse",
        "ready": "Select a PDF file and output folder.",
        "start": "Start split",
        "analyzing": "Analyzing... ({}/{}) pages",
        "saving": "Saving... ({}/{}) {}",
        "error": "Error",
        "select_pdf_error": "Please select a PDF file.",
        "select_folder_error": "Please select an output folder.",
        "date_error": "Please enter a date.",
        "completed_status": "Completed. {} files saved.",
        "completed_title": "Completed",
        "completed_msg": "Split into {} files.\n\nSaved at:\n{}",
        "filename": "Filename",
        "pages": "Pages",
        "language": "Language",
        "log": "Result log",
    },
    "ES": {
        "title": "Divisor de PDF Netting",
        "header": "Divisor de PDF Netting",
        "subtitle": "Divide archivos PDF automáticamente por trading partner.",
        "select_pdf": "Archivo PDF",
        "select_folder": "Carpeta de destino",
        "date_label": "Fecha del nombre del archivo",
        "date_hint": "YYMMDD",
        "example": "Ejemplo: ES001627001B_DAROCA CASTILLO_260420.pdf",
        "browse": "Buscar",
        "ready": "Seleccione un PDF y una carpeta de destino.",
        "start": "Iniciar división",
        "analyzing": "Analizando... ({}/{}) páginas",
        "saving": "Guardando... ({}/{}) {}",
        "error": "Error",
        "select_pdf_error": "Seleccione un archivo PDF.",
        "select_folder_error": "Seleccione una carpeta de destino.",
        "date_error": "Introduzca una fecha.",
        "completed_status": "Completado. {} archivos guardados.",
        "completed_title": "Completado",
        "completed_msg": "División completada en {} archivos.\n\nUbicación:\n{}",
        "filename": "Archivo",
        "pages": "Páginas",
        "language": "Idioma",
        "log": "Registro de resultados",
    },
    "KR": {
        "title": "Netting PDF 분할기",
        "header": "Netting PDF 분할기",
        "subtitle": "Trading Partner 기준으로 PDF를 자동 분할합니다.",
        "select_pdf": "PDF 파일",
        "select_folder": "저장 폴더",
        "date_label": "파일명 날짜",
        "date_hint": "YYMMDD",
        "example": "예: ES001627001B_DAROCA CASTILLO_260420.pdf",
        "browse": "찾아보기",
        "ready": "PDF 파일과 저장 폴더를 선택하세요.",
        "start": "분할 시작",
        "analyzing": "분석 중... ({}/{}) 페이지",
        "saving": "저장 중... ({}/{}) {}",
        "error": "오류",
        "select_pdf_error": "PDF 파일을 선택해주세요.",
        "select_folder_error": "저장 폴더를 선택해주세요.",
        "date_error": "날짜를 입력해주세요.",
        "completed_status": "완료. 총 {}개 파일이 저장되었습니다.",
        "completed_title": "완료",
        "completed_msg": "총 {}개 파일로 분할 완료되었습니다.\n\n저장 위치:\n{}",
        "filename": "파일명",
        "pages": "페이지",
        "language": "언어",
        "log": "결과 로그",
    },
}


def parse_reports(pdf_path, progress_callback=None):
    import pdfplumber

    reports = []
    current = None

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)

        for i, page in enumerate(pdf.pages):
            if progress_callback:
                progress_callback("analyzing", i + 1, total)

            text = page.extract_text()
            if not text:
                continue

            tp_match = re.search(r'Trading Partner Reference\s+(.+?)\s+Transaction Due Date', text)
            batch_match = re.search(r'Batch Number\s+(\d+)', text)
            loc_match = re.search(r'Location\s+([A-Z0-9]{12})', text)

            if tp_match and batch_match:
                if current:
                    current['end_page'] = i
                    reports.append(current)

                current = {
                    'start_page': i,
                    'trading_partner': tp_match.group(1).strip(),
                    'batch_number': batch_match.group(1).strip(),
                    'location': loc_match.group(1) if loc_match else None,
                }

            elif loc_match and current and not current.get('location'):
                current['location'] = loc_match.group(1)

        if current:
            current['end_page'] = total
            reports.append(current)

    return reports


def split_pdf(pdf_path, output_dir, date_str, progress_callback=None):
    from pypdf import PdfReader, PdfWriter

    reports = parse_reports(pdf_path, progress_callback)
    reader = PdfReader(pdf_path)
    total = len(reports)
    results = []

    for idx, r in enumerate(reports):
        if progress_callback:
            progress_callback("saving", idx + 1, total, r['trading_partner'])

        writer = PdfWriter()

        for page_num in range(r['start_page'], r['end_page']):
            writer.add_page(reader.pages[page_num])

        loc = r['location'] or 'UNKNOWN'
        tp_clean = re.sub(r'[\\/:*?"<>|]', '', r['trading_partner'])
        filename = f"{loc}_{tp_clean}_{date_str}.pdf"
        out_path = os.path.join(output_dir, filename)

        with open(out_path, 'wb') as f:
            writer.write(f)

        results.append({
            'filename': filename,
            'pages': r['end_page'] - r['start_page']
        })

    return results


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.lang = tk.StringVar(value="EN")
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.date_str = tk.StringVar(value=datetime.today().strftime("%d%m%y"))

        self.configure(bg=APP_BG)
        self.resizable(True, True)

        self._setup_style()
        self._build_ui()
        self._apply_language()

        self.update_idletasks()
        self.geometry("820x680")
        self.minsize(760, 640)

    def t(self, key):
        return TEXTS[self.lang.get()][key]

    def _setup_style(self):
        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Google.Horizontal.TProgressbar",
            troughcolor="#e8f0fe",
            background=PRIMARY,
            bordercolor="#e8f0fe",
            lightcolor=PRIMARY,
            darkcolor=PRIMARY,
            thickness=8
        )

        style.configure(
            "Google.TCombobox",
            fieldbackground=CARD_BG,
            background=CARD_BG,
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=6
        )

    def _build_ui(self):
        # Main container
        outer = tk.Frame(self, bg=APP_BG)
        outer.pack(fill="both", expand=True, padx=32, pady=28)

        # Header
        top = tk.Frame(outer, bg=APP_BG)
        top.pack(fill="x")

        title_area = tk.Frame(top, bg=APP_BG)
        title_area.pack(side="left", fill="x", expand=True)

        self.header_lbl = tk.Label(
            title_area,
            font=(FN, 22, "bold"),
            bg=APP_BG,
            fg=TEXT,
            anchor="w"
        )
        self.header_lbl.pack(anchor="w")

        self.subtitle_lbl = tk.Label(
            title_area,
            font=(FN, 11),
            bg=APP_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        self.subtitle_lbl.pack(anchor="w", pady=(4, 0))

        lang_area = tk.Frame(top, bg=APP_BG)
        lang_area.pack(side="right", anchor="ne")

        self.language_lbl = tk.Label(
            lang_area,
            font=(FN, 10),
            bg=APP_BG,
            fg=SUBTEXT
        )
        self.language_lbl.pack(anchor="w")

        self.language_combo = ttk.Combobox(
            lang_area,
            textvariable=self.lang,
            values=["EN", "ES", "KR"],
            state="readonly",
            width=8,
            style="Google.TCombobox"
        )
        self.language_combo.pack(pady=(4, 0))
        self.language_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_language())

        # Card
        card = tk.Frame(
            outer,
            bg=CARD_BG,
            highlightbackground=BORDER,
            highlightthickness=1
        )
        card.pack(fill="x", pady=(28, 20))

        form = tk.Frame(card, bg=CARD_BG, padx=28, pady=26)
        form.pack(fill="x")
        form.columnconfigure(0, weight=1)

        self.pdf_lbl, self.pdf_entry, self.pdf_btn = self._input_row(
            form, 0, self.pdf_path, self._browse_pdf
        )

        self.folder_lbl, self.folder_entry, self.folder_btn = self._input_row(
            form, 2, self.output_dir, self._browse_dir
        )

        # Date field
        self.date_lbl = tk.Label(
            form,
            font=(FN, 10, "bold"),
            bg=CARD_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        self.date_lbl.grid(row=4, column=0, columnspan=2, sticky="w", pady=(18, 6))

        date_line = tk.Frame(form, bg=CARD_BG)
        date_line.grid(row=5, column=0, columnspan=2, sticky="ew")
        date_line.columnconfigure(0, weight=0)
        date_line.columnconfigure(1, weight=1)

        self.date_entry = tk.Entry(
            date_line,
            textvariable=self.date_str,
            font=(FN, 12),
            bg="#ffffff",
            fg=TEXT,
            relief="solid",
            bd=1,
            width=12,
            insertbackground=TEXT
        )
        self.date_entry.grid(row=0, column=0, ipady=8, sticky="w")

        self.date_hint_lbl = tk.Label(
            date_line,
            font=(FN, 10),
            bg=CARD_BG,
            fg=SUBTEXT
        )
        self.date_hint_lbl.grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.example_lbl = tk.Label(
            form,
            font=(FN, 9),
            bg=CARD_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        self.example_lbl.grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # Progress
        progress_card = tk.Frame(
            outer,
            bg=CARD_BG,
            highlightbackground=BORDER,
            highlightthickness=1
        )
        progress_card.pack(fill="x", pady=(0, 20))

        progress_inner = tk.Frame(progress_card, bg=CARD_BG, padx=28, pady=22)
        progress_inner.pack(fill="x")

        self.progress_var = tk.IntVar(value=0)

        self.progress = ttk.Progressbar(
            progress_inner,
            variable=self.progress_var,
            maximum=100,
            style="Google.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x")

        self.status_lbl = tk.Label(
            progress_inner,
            font=(FN, 10),
            bg=CARD_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        self.status_lbl.pack(fill="x", pady=(10, 0))

        # Button
        self.run_btn = tk.Button(
            outer,
            command=self._run,
            font=(FN, 12, "bold"),
            bg=PRIMARY,
            fg="white",
            activebackground=PRIMARY_DARK,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=34,
            pady=12,
            cursor="hand2"
        )
        self.run_btn.pack(anchor="e", pady=(0, 20))

        # Log area
        log_card = tk.Frame(
            outer,
            bg=CARD_BG,
            highlightbackground=BORDER,
            highlightthickness=1
        )
        log_card.pack(fill="both", expand=True)

        log_inner = tk.Frame(log_card, bg=CARD_BG, padx=24, pady=18)
        log_inner.pack(fill="both", expand=True)

        self.log_lbl = tk.Label(
            log_inner,
            font=(FN, 10, "bold"),
            bg=CARD_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        self.log_lbl.pack(fill="x", pady=(0, 8))

        log_wrap = tk.Frame(log_inner, bg=LOG_BG)
        log_wrap.pack(fill="both", expand=True)

        self.log = tk.Text(
            log_wrap,
            height=9,
            font=(FM, 10),
            bg=LOG_BG,
            fg=TEXT,
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            state="disabled"
        )

        sb = tk.Scrollbar(log_wrap, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)

        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)

    def _input_row(self, parent, row, variable, command):
        label = tk.Label(
            parent,
            font=(FN, 10, "bold"),
            bg=CARD_BG,
            fg=SUBTEXT,
            anchor="w"
        )
        label.grid(row=row, column=0, columnspan=2, sticky="w", pady=(0 if row == 0 else 18, 6))

        entry = tk.Entry(
            parent,
            textvariable=variable,
            font=(FN, 12),
            bg="#ffffff",
            fg=TEXT,
            relief="solid",
            bd=1,
            insertbackground=TEXT
        )
        entry.grid(row=row + 1, column=0, sticky="ew", ipady=8, padx=(0, 12))

        button = tk.Button(
            parent,
            command=command,
            font=(FN, 10, "bold"),
            bg="#ffffff",
            fg=PRIMARY,
            activebackground="#e8f0fe",
            activeforeground=PRIMARY,
            relief="solid",
            bd=1,
            padx=18,
            pady=8,
            cursor="hand2"
        )
        button.grid(row=row + 1, column=1, sticky="ew")

        return label, entry, button

    def _apply_language(self):
        self.title(self.t("title"))
        self.header_lbl.config(text=self.t("header"))
        self.subtitle_lbl.config(text=self.t("subtitle"))
        self.language_lbl.config(text=self.t("language"))
        self.pdf_lbl.config(text=self.t("select_pdf"))
        self.folder_lbl.config(text=self.t("select_folder"))
        self.date_lbl.config(text=self.t("date_label"))
        self.date_hint_lbl.config(text=self.t("date_hint"))
        self.example_lbl.config(text=self.t("example"))
        self.pdf_btn.config(text=self.t("browse"))
        self.folder_btn.config(text=self.t("browse"))
        self.status_lbl.config(text=self.t("ready"))
        self.run_btn.config(text=self.t("start"))
        self.log_lbl.config(text=self.t("log"))

    def _browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])

        if path:
            self.pdf_path.set(path)

            m = re.search(r'(\d{6})', os.path.basename(path))
            if m:
                self.date_str.set(m.group(1))

            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(path))

    def _browse_dir(self):
        path = filedialog.askdirectory()

        if path:
            self.output_dir.set(path)

    def _log(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _progress_callback(self, mode, current, total, name=None):
        if total == 0:
            pct = 0
            status = ""
        elif mode == "analyzing":
            pct = int((current / total) * 50)
            status = self.t("analyzing").format(current, total)
        elif mode == "saving":
            pct = 50 + int((current / total) * 50)
            status = self.t("saving").format(current, total, name)
        else:
            pct = 0
            status = ""

        self.progress_var.set(pct)
        self.status_lbl.config(text=status)
        self.update_idletasks()

    def _run(self):
        pdf = self.pdf_path.get().strip()
        out = self.output_dir.get().strip()
        date = self.date_str.get().strip()

        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror(self.t("error"), self.t("select_pdf_error"))
            return

        if not out or not os.path.isdir(out):
            messagebox.showerror(self.t("error"), self.t("select_folder_error"))
            return

        if not date:
            messagebox.showerror(self.t("error"), self.t("date_error"))
            return

        self.run_btn.config(state="disabled", bg="#9aa0a6")
        self.progress_var.set(0)

        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

        def worker():
            try:
                results = split_pdf(pdf, out, date, self._progress_callback)

                self.progress_var.set(100)
                self.status_lbl.config(
                    text=self.t("completed_status").format(len(results), out)
                )

                self._log(f"{self.t('filename'):<58} {self.t('pages'):>4}")
                self._log("─" * 64)

                for r in results:
                    self._log(f"{r['filename']:<58} {r['pages']:>3}p")

                messagebox.showinfo(
                    self.t("completed_title"),
                    self.t("completed_msg").format(len(results), out)
                )

            except Exception as e:
                self.progress_var.set(0)
                self.status_lbl.config(text=f"{self.t('error')}: {e}")
                messagebox.showerror(self.t("error"), str(e))

            finally:
                self.run_btn.config(state="normal", bg=PRIMARY)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
