import os
import re
import smtplib
import threading
import platform
import pandas as pd
import tkinter as tk

from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ─────────────────────────────
# Theme & Typography (디자인 설정)
# ─────────────────────────────
if platform.system() == "Darwin":
    FN = "Arial" # 맥 호환
    FM = "Menlo"
else:
    FN = "Arial" # 윈도우 호환
    FM = "Consolas"

APP_BG = "#f4f6f9"      # 배경 회색
CARD_BG = "#ffffff"     # 메인 카드 흰색
PRIMARY = "#1a73e8"     # 메인 포인트 컬러 (파랑)
PRIMARY_DARK = "#1557b0"
TEXT = "#202124"
SUBTEXT = "#5f6368"
BORDER = "#dadce0"
LOG_BG = "#f8f9fa"

# ─────────────────────────────
# SMTP Config
# ─────────────────────────────
SMTP_SERVER = "lgekrhqmh01.lge.com"
SMTP_PORT = 25
REMITE_EMAIL = "noresponder@lge.com"
REMITE_NAME = "LG Validacion"

TEXTS = {
    "EN": {
        "title": "Compensación Email Sender",
        "header": "Compensación Email Sender",
        "subtitle": "Send each PDF file automatically based on the Excel email list.",
        "language": "Language",
        "excel_file": "Excel file",
        "pdf_folder": "PDF folder",
        "bcc": "BCC recipients",
        "bcc_hint": "Example: manager@lge.com; team@lge.com (Separate with ;)",
        "browse": "Browse",
        "start": "Start sending",
        "ready": "Select the Excel file and PDF folder, then click Start.",
        "log": "Result Log",
        "excel_error": "Please select a valid Excel file.",
        "folder_error": "Please select a valid PDF folder.",
        "confirm_title": "Confirmation",
        "confirm_msg": "Do you want to start sending emails?",
        "cancelled": "Process cancelled.",
        "error": "Error",
        "done": "Completed",
        "reading_excel": "Reading Excel file...",
        "processing": "Processing PDF files...",
        "sending": "Sending... ({}/{}) {}",
        "completed_status": "Completed. Processed: {} / Sent: {} / Failed: {}",
        "completed_msg": "Processed: {}\nSent: {}\nFailed: {}\nLog file:\n{}",
        "no_pdf": "No PDF files were found in the selected folder.",
        "sent": "SENT",
        "not_sent": "NOT SENT",
        "no_match": "NO MATCH IN EXCEL",
    },
    "ES": {
        "title": "Envío de Emails Compensación",
        "header": "Envío de Emails Compensación",
        "subtitle": "Envía automáticamente cada PDF según la lista de emails del Excel.",
        "language": "Idioma",
        "excel_file": "Archivo Excel",
        "pdf_folder": "Carpeta de PDFs",
        "bcc": "Correos BCC",
        "bcc_hint": "Ejemplo: manager@lge.com; team@lge.com (Separe con ;)",
        "browse": "Buscar",
        "start": "Iniciar envío",
        "ready": "Seleccione el Excel y la carpeta de PDFs, luego pulse Iniciar.",
        "log": "Registro de resultados",
        "excel_error": "Seleccione un archivo Excel válido.",
        "folder_error": "Seleccione una carpeta de PDFs válida.",
        "confirm_title": "Confirmación",
        "confirm_msg": "¿Desea iniciar el proceso de envío de emails?",
        "cancelled": "Proceso cancelado.",
        "error": "Error",
        "done": "Completado",
        "reading_excel": "Leyendo archivo Excel...",
        "processing": "Procesando archivos PDF...",
        "sending": "Enviando... ({}/{}) {}",
        "completed_status": "Completado. Procesados: {} / Enviados: {} / Fallidos: {}",
        "completed_msg": "Procesados: {}\nEnviados: {}\nFallidos: {}\nArchivo log:\n{}",
        "no_pdf": "No se encontraron archivos PDF en la carpeta seleccionada.",
        "sent": "ENVIADO",
        "not_sent": "NO ENVIADO",
        "no_match": "NO MATCH EN EXCEL",
    },
    "KR": {
        "title": "Compensación 이메일 자동 발송기",
        "header": "Compensación 이메일 자동 발송기",
        "subtitle": "Excel 이메일 리스트를 기준으로 각 PDF를 자동 발송합니다.",
        "language": "언어",
        "excel_file": "Excel 파일",
        "pdf_folder": "PDF 폴더",
        "bcc": "숨은참조(BCC)",
        "bcc_hint": "예시: manager@lge.com; team@lge.com (세미콜론으로 구분)",
        "browse": "찾아보기",
        "start": "발송 시작",
        "ready": "Excel 파일과 PDF 폴더를 선택한 후 발송을 시작하세요.",
        "log": "결과 로그",
        "excel_error": "올바른 Excel 파일을 선택해주세요.",
        "folder_error": "올바른 PDF 폴더를 선택해주세요.",
        "confirm_title": "확인",
        "confirm_msg": "이메일 발송을 시작하시겠습니까?",
        "cancelled": "작업이 취소되었습니다.",
        "error": "오류",
        "done": "완료",
        "reading_excel": "Excel 파일을 읽는 중...",
        "processing": "PDF 파일을 처리하는 중...",
        "sending": "발송 중... ({}/{}) {}",
        "completed_status": "완료. 처리: {} / 발송: {} / 실패: {}",
        "completed_msg": "처리: {}\n발송: {}\n실패: {}\n로그 파일:\n{}",
        "no_pdf": "선택한 폴더에서 PDF 파일을 찾을 수 없습니다.",
        "sent": "발송 완료",
        "not_sent": "발송 실패",
        "no_match": "Excel 매칭 없음",
    },
}

EMAIL_ES = {
    "email_subject": "Envío de PDF Compensación: {}",
    "email_title": "Envío del Resultado de Compensación",
    "email_greeting": "Buenos días.",
    "email_footer": "Este correo ha sido generado automáticamente.<br>Por favor, no responda a esta dirección."
}

def parse_emails(cell_value):
    emails = []
    if not cell_value or str(cell_value).lower() == "nan":
        return emails
    for part in str(cell_value).split(";"):
        part = part.strip()
        match = re.search(r"<(.+?)>", part)
        email = match.group(1) if match else part
        if email:
            emails.append(email)
    return emails

def build_html_body():
    current_date = datetime.now().strftime("%Y/%m/%d")
    return f"""
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0; padding:40px 0; background-color:#f4f5f7; font-family:'Segoe UI', Arial, Helvetica, sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
        <td align="center">
            <table width="700" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff; border:1px solid #e0e0e0; text-align:left;">
                <tr>
                    <td style="background-color:#A50034; padding:30px; color:#ffffff;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="color:#ffffff;">
                            <tr>
                                <td style="font-size:14px; font-weight:bold; vertical-align:middle;">
                                    <span style="background-color:#ffffff; color:#A50034; border-radius:50%; padding:4px 7px; margin-right:6px; font-size:12px;">LG</span> 
                                    LG Electronics &middot; Compensación
                                </td>
                                <td align="right" style="font-size:13px; color:#ffccd5;">Madrid, España</td>
                            </tr>
                            <tr>
                                <td colspan="2" style="padding-top:25px; font-size:12px; font-weight:bold; letter-spacing:1px; color:#ffccd5;">
                                    NOTIFICACIÓN DE COMPENSACIÓN
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" style="padding-top:4px; font-size:28px; font-weight:bold;">
                                    {EMAIL_ES["email_title"]}
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="background-color:#f8f9fa; padding:12px 30px; border-bottom:1px solid #eeeeee; font-size:13px; color:#555555;">
                        <strong>Fecha:</strong> {current_date} &nbsp;&nbsp;&nbsp;&nbsp; <strong>Remitente:</strong> {REMITE_NAME}
                    </td>
                </tr>
                <tr>
                    <td style="padding:35px 30px; color:#333333; font-size:15px; line-height:1.6;">
                        <p style="margin-top:0; font-size:16px;"><strong>{EMAIL_ES["email_greeting"]}</strong></p>
                        <p>Adjunto les remitimos el resultado del último proceso de <strong>Compensación</strong> realizado.</p>
                        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:25px 0;">
                            <tr>
                                <td style="background-color:#fff0f2; border-left:4px solid #A50034; padding:18px; font-size:14px; color:#202124;">
                                    En caso de que tengan cualquier consulta o necesiten alguna aclaración adicional, quedamos a su disposición a través del correo 
                                    <strong>lg.validacion@lgepartner.com</strong>.
                                </td>
                            </tr>
                        </table>
                        <p>Con el fin de garantizar una gestión más eficiente y una respuesta coherente para todos, les rogamos su colaboración para que las consultas se canalicen únicamente por esta vía.</p>
                        <p>Quedamos a su disposición y agradecemos de antemano su comprensión y colaboración.</p>
                        <p>Un saludo cordial.</p>
                    </td>
                </tr>
                <tr>
                    <td style="padding:25px 30px; border-top:1px solid #eeeeee; background-color:#ffffff;">
                        <table width="100%" cellpadding="0" cellspacing="0" border="0">
                            <tr>
                                <td style="font-size:12px; color:#9aa0a6; line-height:1.5;">
                                    {EMAIL_ES["email_footer"]}<br>
                                    Gestión de Compensación &middot; {current_date}
                                </td>
                                <td align="right" style="font-size:14px; font-weight:bold; color:#A50034; opacity:0.6; vertical-align:bottom;">
                                    LG Electronics
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>
</body>
</html>
"""

def send_email(destinations, pdf_file, bcc_list):
    try:
        msg = MIMEMultipart()
        msg["From"] = f"{REMITE_NAME} <{REMITE_EMAIL}>"
        msg["To"] = ", ".join(destinations)
        msg["Subject"] = EMAIL_ES["email_subject"].format(os.path.basename(pdf_file))

        if bcc_list:
            msg["Bcc"] = "; ".join(bcc_list)

        msg.attach(MIMEText(build_html_body(), "html"))

        with open(pdf_file, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")
            attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_file))
            msg.attach(attachment)

        final_recipients = destinations + bcc_list
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.sendmail(REMITE_EMAIL, final_recipients, msg.as_string())
        return True, ""
    except Exception as e:
        return False, str(e)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.lang = tk.StringVar(value="ES")
        self.excel_file = tk.StringVar()
        self.pdf_folder = tk.StringVar()
        self.bcc = tk.StringVar()

        self.title("Compensación Email Sender")
        self.geometry("750x750")
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
        # 1. 상단 타이틀 및 언어 선택 (회색 배경 영역)
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

        # 2. 메인 입력 폼 (흰색 카드 영역)
        main_frame = tk.Frame(self, bg=CARD_BG, padx=30, pady=30, highlightbackground=BORDER, highlightthickness=1)
        main_frame.pack(fill="x", padx=20)

        # 엑셀 파일
        self.excel_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.excel_lbl.pack(anchor="w")
        excel_row = tk.Frame(main_frame, bg=CARD_BG)
        excel_row.pack(fill="x", pady=(5, 15))
        self.excel_entry = tk.Entry(excel_row, textvariable=self.excel_file, font=(FN, 10), relief="solid", bd=1)
        self.excel_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.excel_btn = tk.Button(excel_row, font=(FN, 9, "bold"), bg=CARD_BG, fg=PRIMARY, width=10, relief="solid", bd=1, command=self._browse_excel)
        self.excel_btn.pack(side="left", padx=(10, 0), ipady=4)

        # PDF 폴더
        self.folder_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.folder_lbl.pack(anchor="w")
        folder_row = tk.Frame(main_frame, bg=CARD_BG)
        folder_row.pack(fill="x", pady=(5, 15))
        self.folder_entry = tk.Entry(folder_row, textvariable=self.pdf_folder, font=(FN, 10), relief="solid", bd=1)
        self.folder_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self.folder_btn = tk.Button(folder_row, font=(FN, 9, "bold"), bg=CARD_BG, fg=PRIMARY, width=10, relief="solid", bd=1, command=self._browse_folder)
        self.folder_btn.pack(side="left", padx=(10, 0), ipady=4)

        # 숨은참조(BCC)
        self.bcc_lbl = tk.Label(main_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.bcc_lbl.pack(anchor="w")
        self.bcc_entry = tk.Entry(main_frame, textvariable=self.bcc, font=(FN, 10), relief="solid", bd=1)
        self.bcc_entry.pack(fill="x", pady=(5, 2), ipady=6)
        self.bcc_hint_lbl = tk.Label(main_frame, font=(FN, 8), bg=CARD_BG, fg=SUBTEXT)
        self.bcc_hint_lbl.pack(anchor="w")

        # 3. 로그 창 및 진행 상황 (흰색 카드 영역)
        log_frame = tk.Frame(self, bg=CARD_BG, padx=20, pady=20, highlightbackground=BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.log_lbl = tk.Label(log_frame, font=(FN, 9, "bold"), bg=CARD_BG, fg=TEXT)
        self.log_lbl.pack(anchor="w", pady=(0, 5))
        
        self.log = tk.Text(log_frame, height=8, font=(FM, 9), bg=LOG_BG, fg=TEXT, relief="solid", bd=1, padx=10, pady=10, state="disabled")
        self.log.pack(fill="both", expand=True)

        # 4. 하단 상태 및 실행 버튼 (회색 배경 영역)
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
        self.folder_lbl.config(text=self.t("pdf_folder"))
        self.bcc_lbl.config(text=self.t("bcc"))
        self.bcc_hint_lbl.config(text=self.t("bcc_hint"))
        self.excel_btn.config(text=self.t("browse"))
        self.folder_btn.config(text=self.t("browse"))
        self.run_btn.config(text=self.t("start"))
        self.status_lbl.config(text=self.t("ready"))
        self.log_lbl.config(text=self.t("log"))

    def _browse_excel(self):
        path = filedialog.askopenfilename(title=self.t("excel_file"), filetypes=[("Excel files", "*.xlsx *.xls")])
        if path: self.excel_file.set(path)

    def _browse_folder(self):
        path = filedialog.askdirectory(title=self.t("pdf_folder"))
        if path: self.pdf_folder.set(path)

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

    def _read_excel(self, excel_file):
        df = pd.read_excel(excel_file, sheet_name="ASC emailing List", header=0, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        df = df[[df.columns[0], df.columns[2]]]
        df.columns = ["BillTo", "Contacto"]
        df["BillTo"] = df["BillTo"].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        df["Contacto"] = df["Contacto"].fillna("").astype(str).str.strip()
        return df

    def _parse_bcc(self):
        raw = self.bcc.get().strip()
        if not raw: return []
        return [e.strip() for e in raw.split(";") if e.strip()]

    def _run(self):
        excel_file = self.excel_file.get().strip()
        pdf_folder = self.pdf_folder.get().strip()

        if not excel_file or not os.path.isfile(excel_file):
            messagebox.showerror(self.t("error"), self.t("excel_error"))
            return
        if not pdf_folder or not os.path.isdir(pdf_folder):
            messagebox.showerror(self.t("error"), self.t("folder_error"))
            return
        if not messagebox.askyesno(self.t("confirm_title"), self.t("confirm_msg")):
            self.status_lbl.config(text=self.t("cancelled"))
            return

        self.run_btn.config(state="disabled", bg="#9aa0a6")
        self.progress_var.set(0)
        self._clear_log()

        def worker():
            try:
                bcc_list = self._parse_bcc()
                self._set_status(5, self.t("reading_excel"))
                df = self._read_excel(excel_file)
                pdfs = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]

                if not pdfs:
                    messagebox.showwarning(self.t("error"), self.t("no_pdf"))
                    self._set_status(0, self.t("no_pdf"))
                    return

                log_file = os.path.join(os.getcwd(), f"log_envio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                processed, sent, failed = 0, 0, 0
                success_logs, failed_logs = [], []

                self._set_status(10, self.t("processing"))

                for idx, pdf in enumerate(pdfs, start=1):
                    pct = 10 + int((idx / len(pdfs)) * 85)
                    self._set_status(pct, self.t("sending").format(idx, len(pdfs), pdf))
                    processed += 1

                    raw_bill_to = pdf.split("_")[0]
                    bill_to = re.sub(r'[^a-zA-Z0-9]', '', raw_bill_to).upper()
                    row = df[df["BillTo"] == bill_to]

                    if not row.empty:
                        emails = parse_emails(row["Contacto"].values[0])
                        pdf_path = os.path.join(pdf_folder, pdf)

                        if emails:
                            ok, err = send_email(emails, pdf_path, bcc_list)
                            if ok:
                                sent += 1
                                success_logs.append(f"{pdf} (→ {', '.join(emails)})")
                            else:
                                failed += 1
                                failed_logs.append(f"{pdf} (→ {', '.join(emails)}) | Error: {err}")
                        else:
                            failed += 1
                            failed_logs.append(f"{pdf} | Error: NO EMAIL")
                    else:
                        failed += 1
                        failed_logs.append(f"{pdf} | Error: {self.t('no_match')}")

                final_log_text = f"==== RESULT LOG ====\n{datetime.now()}\n\n[Success] ({len(success_logs)})\n"
                final_log_text += "\n".join([f"✅ {log}" for log in success_logs]) + "\n" if success_logs else "- None\n"
                final_log_text += f"\n[Failed] ({len(failed_logs)})\n"
                final_log_text += "\n".join([f"❌ {log}" for log in failed_logs]) + "\n" if failed_logs else "- None\n"

                self._log(final_log_text)
                with open(log_file, "w", encoding="utf-8") as log:
                    log.write(final_log_text)

                self._set_status(100, self.t("completed_status").format(processed, sent, failed))
                messagebox.showinfo(self.t("done"), self.t("completed_msg").format(processed, sent, failed, log_file))

            except Exception as e:
                self._set_status(0, f"{self.t('error')}: {e}")
                messagebox.showerror(self.t("error"), str(e))
            finally:
                self.run_btn.config(state="normal", bg=PRIMARY)

        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()
