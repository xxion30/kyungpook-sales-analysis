import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import platform

# =========================
# 한글 폰트 설정 (그래프)
# =========================
if platform.system() == "Windows":
    plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False

# =========================
# 전역 변수
# =========================
df = None
current_fig = None
current_mode = None   # "top_n" or "season"

VACATION_MONTHS = [1, 2, 7, 8]
NON_VACATION_MONTHS = [3, 4, 5, 6, 9, 10, 11, 12]

# =========================
# CSV 데이터 로드
# =========================
def load_csv():
    global df
    try:
        path = filedialog.askopenfilename(
            filetypes=[("CSV 파일", "*.csv")]
        )
        if not path:
            return

        df = pd.read_csv(path, encoding="euc-kr")

        required_cols = {"가맹점명", "연월", "매출액"}
        if not required_cols.issubset(df.columns):
            raise ValueError("CSV 컬럼 구조 오류")

        df["연월"] = pd.to_datetime(df["연월"], errors="coerce")
        if df["연월"].isna().any():
            raise ValueError("날짜 파싱 실패")

        df["월"] = df["연월"].dt.month

        messagebox.showinfo("성공", "데이터 로드 완료")

    except UnicodeDecodeError:
        messagebox.showerror("인코딩 오류", "euc-kr 인코딩 CSV 파일이 아닙니다.")
    except Exception as e:
        messagebox.showerror("오류", str(e))

# =========================
# 상위 N개 매출 분석
# =========================
def analyze_top_n():
    global current_fig, current_mode
    current_mode = "top_n"

    if df is None:
        messagebox.showwarning("경고", "CSV 데이터를 먼저 불러오세요.")
        return

    try:
        n = int(entry_n.get())
        if n <= 0:
            raise ValueError
    except:
        messagebox.showerror("입력 오류", "N은 1 이상의 정수여야 합니다.")
        return

    store_count = df["가맹점명"].nunique()
    if n > store_count:
        n = store_count
        messagebox.showinfo("안내", f"N이 가맹점 수보다 커서 {n}으로 조정됨")

    result = (
        df.groupby("가맹점명")["매출액"]
        .sum()
        .sort_values(ascending=False)
        .head(n)
        .reset_index()
    )

    fig, ax = plt.subplots(figsize=(8, 5))
    sns.barplot(data=result, x="매출액", y="가맹점명", ax=ax)
    ax.set_title(f"총 매출 기준 상위 {n}개 가맹점")

    show_graph(fig)

# =========================
# 시즌 분석
# =========================
def analyze_season():
    global current_fig, current_mode
    current_mode = "season"

    if df is None:
        messagebox.showwarning("경고", "데이터를 먼저 불러오세요.")
        return

    vac_df = df[df["월"].isin(VACATION_MONTHS)]
    non_df = df[df["월"].isin(NON_VACATION_MONTHS)]

    if vac_df.empty or non_df.empty:
        messagebox.showwarning("데이터 부족", "일부 시즌 데이터가 부족합니다.")
        return

    vac_sum = vac_df["매출액"].sum()
    non_sum = non_df["매출액"].sum()
    rate = ((vac_sum - non_sum) / non_sum * 100) if non_sum != 0 else 0

    season_df = pd.DataFrame({
        "시즌": ["방학 시즌", "비방학 시즌"],
        "총 매출액": [vac_sum, non_sum]
    })

    fig, ax = plt.subplots(figsize=(6, 4))
    sns.barplot(data=season_df, x="시즌", y="총 매출액", ax=ax)
    ax.set_title(f"시즌별 매출 비교 (증감률 {rate:.2f}%)")

    show_graph(fig)

# =========================
# 그래프 표시
# =========================
def show_graph(fig):
    global current_fig
    current_fig = fig

    for w in graph_frame.winfo_children():
        w.destroy()

    canvas_tk = FigureCanvasTkAgg(fig, master=graph_frame)
    canvas_tk.draw()
    canvas_tk.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# =========================
# PNG 저장
# =========================
def save_png():
    if current_fig is None:
        messagebox.showwarning("경고", "저장할 그래프가 없습니다.")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("PNG", "*.png")]
    )
    if path:
        current_fig.savefig(path)
        messagebox.showinfo("완료", "PNG 저장 완료")

# =========================
# PDF 저장 (분석 모드별 분기)
# =========================
def save_pdf():
    if df is None or current_mode is None:
        messagebox.showwarning("경고", "저장할 분석 결과가 없습니다.")
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF 파일", "*.pdf")]
    )
    if not path:
        return

    try:
        pdfmetrics.registerFont(
            TTFont("Malgun", "C:/Windows/Fonts/malgun.ttf")
        )

        c = canvas.Canvas(path, pagesize=A4)
        text = c.beginText(40, 800)
        text.setFont("Malgun", 12)

        text.textLine("경대북문지기 매출 분석 보고서")
        text.textLine("")

        if current_mode == "top_n":
            text.textLine("📌 상위 가맹점 매출 분석")
            text.textLine("")

            result = (
                df.groupby("가맹점명")["매출액"]
                .sum()
                .sort_values(ascending=False)
                .head(int(entry_n.get()))
            )

            for i, (store, sales) in enumerate(result.items(), 1):
                text.textLine(f"{i}. {store} : {sales:,.0f}원")

        elif current_mode == "season":
            vac = df[df["월"].isin(VACATION_MONTHS)]["매출액"].sum()
            non = df[df["월"].isin(NON_VACATION_MONTHS)]["매출액"].sum()
            rate = ((vac - non) / non * 100) if non != 0 else 0

            text.textLine("📌 시즌별 매출 분석")
            text.textLine("")
            text.textLine(f"방학 시즌 총 매출액: {vac:,.0f}원")
            text.textLine(f"비방학 시즌 총 매출액: {non:,.0f}원")
            text.textLine(f"매출 증감률: {rate:.2f}%")

        c.drawText(text)
        c.save()

        messagebox.showinfo("완료", "PDF 저장 완료")

    except Exception as e:
        messagebox.showerror("오류", f"PDF 생성 실패: {e}")

# =========================
# 초기화
# =========================
def reset_all():
    global df, current_fig, current_mode
    df = None
    current_fig = None
    current_mode = None

    entry_n.delete(0, tk.END)
    entry_n.insert(0, "10")

    for w in graph_frame.winfo_children():
        w.destroy()

    messagebox.showinfo("초기화", "초기화 완료")

# =========================
# GUI 구성
# =========================
root = tk.Tk()
root.title("경대북문지기 매출 분석 시스템")
root.geometry("900x600")

top_frame = tk.Frame(root)
top_frame.pack(pady=10)

tk.Button(top_frame, text="CSV 불러오기", command=load_csv).pack(side=tk.LEFT, padx=5)

tk.Label(top_frame, text="상위 N개").pack(side=tk.LEFT)
entry_n = tk.Entry(top_frame, width=5)
entry_n.insert(0, "10")
entry_n.pack(side=tk.LEFT, padx=5)

tk.Button(top_frame, text="상위 매출 분석", command=analyze_top_n).pack(side=tk.LEFT, padx=5)
tk.Button(top_frame, text="시즌 분석", command=analyze_season).pack(side=tk.LEFT, padx=5)
tk.Button(top_frame, text="PNG 저장", command=save_png).pack(side=tk.LEFT, padx=5)
tk.Button(top_frame, text="PDF 저장", command=save_pdf).pack(side=tk.LEFT, padx=5)
tk.Button(top_frame, text="초기화", command=reset_all).pack(side=tk.LEFT, padx=5)

graph_frame = tk.Frame(root)
graph_frame.pack(fill=tk.BOTH, expand=True)

root.mainloop()
