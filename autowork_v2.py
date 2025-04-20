from fpdf import FPDF
import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button, Label, Toplevel
import datetime
import os
import subprocess

FONT_PATH = "C:/Windows/Fonts/malgun.ttf"

class StyledPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.add_font("Malgun", "", FONT_PATH, uni=True)
        self.set_font("Malgun", "", 9)
        self.set_auto_page_break(auto=True, margin=10)

    def draw_title(self, start_date, end_date):
        self.set_font("Malgun", "", 20)
        self.cell(0, 8, "근무 및 출장명령서", ln=True, align='C')
        self.ln(3)
        self.set_font("Malgun", "", 9)
        self.cell(0, 5, f"□ 기간: {start_date} ~ {end_date}", ln=True)
        self.ln(1)

    def draw_table(self, df):
        weekdays = ['월', '화', '수', '목', '금', '토', '일']
        for _, row in df.iterrows():
            date_obj = pd.to_datetime(row['날짜'])
            date_str = date_obj.strftime('%m.%d')
            weekday = f"({weekdays[date_obj.weekday()]})"

            # 날짜 및 장소
            self.set_font("Malgun", "", 9)
            self.cell(25, 6, f"{date_str} {weekday}", border=1, align='C')
            self.cell(0, 6, f"장  소: {row['장소']}", border=1)
            self.ln(6)

            # 출발/검진 시간
            self.cell(25, 6, "", border=1)
            self.cell(0, 6, f"출발시간: {row['출발시간']} / 검진시간: {row['검진시간']}", border=1)
            self.ln(6)

            # 스태프
            x0, y0 = self.get_x(), self.get_y()
            self.cell(25, 12, "", border=1)
            self.set_xy(x0 + 25, y0)
            line1 = f"의사: {row['의사']} / 행정: {row['행정']} / 병리: {row['병리']}"
            line2 = f"방사선: {row['방사선']} / 간호: {row['간호']}"
            self.multi_cell(0, 6, f"{line1}\n{line2}", border=1)
            self.ln(1)

    def draw_footer(self):
        self.set_font("Malgun", "", 10)
        self.cell(0, 5, "※ 상황에 따라 교차근무 가능", ln=True)
        self.ln(4)

        today = datetime.datetime.today().strftime("%Y년 %m월 %d일")
        self.cell(0, 5, today, ln=True, align='C')
        self.ln(2)

        self.set_font("Malgun", "", 10)
        col_titles = ["담당", "건강증진과장", "인구사업과장", "행정지원과장", "원장", "본부장"]
        box_width = 160 / len(col_titles)

        x0, y0 = self.get_x(), self.get_y()
        self.multi_cell(15, 9, "결\n재", border=1, align='C')
        self.set_xy(x0 + 15, y0)
        for title in col_titles:
            self.cell(box_width, 9, title, border=1, align='C')
        self.ln()

        self.set_x(x0 + 15)
        for _ in col_titles:
            self.cell(box_width, 9, "", border=1)
        self.ln(14)

        self.set_font("Malgun", "", 13)
        self.cell(0, 8, "인구보건복지협회 충북세종지회", ln=True, align='C')


def get_non_conflicting_filename(folder_path, base_name, extension=".pdf"):
    counter = 0
    while True:
        filename = f"{base_name}{'' if counter == 0 else f'_{counter}'}{extension}"
        full_path = os.path.join(folder_path, filename)
        if not os.path.exists(full_path):
            return full_path
        counter += 1


def show_success_window(output_path):
    folder_path = os.path.dirname(output_path)
    win = Toplevel()
    win.title("PDF 생성 완료")
    win.geometry("360x100")
    Label(win, text="✅ PDF 생성이 완료되었습니다!", font=("Arial", 12)).pack(pady=10)
    Button(win, text="PDF 저장된 폴더 열기", command=lambda: subprocess.Popen(f'explorer "{folder_path}"'), width=30).pack()


def generate_pdf_from_excel(filepath):
    try:
        df = pd.read_excel(filepath)
        start_dt = pd.to_datetime(df['날짜'].iloc[0])
        end_dt = pd.to_datetime(df['날짜'].iloc[-1])
        start_str = start_dt.strftime('%Y.%m.%d(%a)')
        end_str = end_dt.strftime('%Y.%m.%d(%a)')

        filename_dt_str = start_dt.strftime('%y년%m월%d일')
        base_filename = f"{filename_dt_str} 이후 근무 및 출장명령서"

        # 스크립트 실행 폴더 내에 '근무일지' 디렉터리 생성
        script_dir = os.path.dirname(os.path.abspath(__file__))
        target_dir = os.path.join(script_dir, '근무일지')
        os.makedirs(target_dir, exist_ok=True)

        output_path = get_non_conflicting_filename(target_dir, base_filename)

        pdf = StyledPDF()
        pdf.add_page()
        pdf.draw_title(start_str, end_str)
        pdf.draw_table(df)
        pdf.draw_footer()

        pdf.output(output_path)

        show_success_window(output_path)
    except Exception as e:
        messagebox.showerror("오류", f"PDF 생성 중 오류 발생:\n{e}")


def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        generate_pdf_from_excel(filepath)

# GUI 실행
root = Tk()
root.title("근무 및 출장명령서 자동 생성기")
root.geometry("420x200")

Label(root, text="근무 및 출장명령서 PDF 자동 생성기", font=("Arial", 14)).pack(pady=20)
Button(root, text="일주일치 근무일정 업로드", command=open_file, width=30, height=2).pack(pady=10)

root.mainloop()
