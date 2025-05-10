import sys, os
import json
import time

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import openpyxl            # Excel 읽기 엔진
import pyautogui
import pyperclip
import pygetwindow as gw
import mss
from PIL import Image
import cv2, numpy as np

# ——— 유틸 함수들 —————————————————————————————




def resource_path(rel_path):
    """
    PyInstaller 번들링 시 _MEIPASS 폴더를,
    일반 실행 시 스크립트 폴더를 기준으로 경로를 반환.
    """
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    real_path = os.path.join(base, rel_path)
    print(real_path)
    return real_path

# def click_friend_tab_with_mss(win, template_path, confidence=0.7):
#     # 0) Unicode 경로도 잘 여는 PIL로 미리 열기
#     try:
#         template = Image.open(template_path)
#     except Exception as e:
#         log.insert(tk.END, f"❌ 템플릿 열기 실패: {template_path} ({e})\n")
#         log.see(tk.END)
#         return False

#     # 1) 창 활성화/최대화
#     win.activate(); time.sleep(0.2)
#     win.maximize(); time.sleep(0.5)
#     win = get_real_kakao_window()

#     # 2) mss로 스크린샷 떠서 PIL 이미지로 변환
#     mon   = mss.mss().monitors[0]
#     left  = max(win.left, mon["left"])
#     top   = max(win.top,  mon["top"])
#     width = min(win.width,  mon["width"]  - (left - mon["left"]))
#     height= min(win.height, mon["height"] - (top  - mon["top"]))
#     region= {"left": left, "top": top, "width": width, "height": height}

#     with mss.mss() as sct:
#         sct_img  = sct.grab(region)
#         haystack = Image.frombytes("RGB",
#                        (sct_img.width, sct_img.height),
#                         sct_img.rgb)
#         haystack.save("debug_haystack.png")  # ← 이 줄 추가

#     # 3) PIL.Image 객체로 매칭 → cv2.imread 호출 안 함
#     match = pyautogui.locate(template, haystack,
#                              confidence=confidence,
#                              grayscale=True)
#     if not match:
#         log.insert(tk.END, f"❌ 아이콘 '{template_path}' 못 찾음\n"); log.see(tk.END)
#         return False

#     # 4) 화면 절대 좌표로 클릭
#     cx, cy = pyautogui.center(match)
#     px = region["left"] + cx
#     py = region["top"]  + cy
#     pyautogui.click(px, py)
#     time.sleep(0.2)
#     return True

def click_friend_tab_with_mss(win, template_path, confidence=0.7):
    # 0) PIL로 템플릿 열기
    try:
        pil_tpl = Image.open(template_path)
    except Exception as e:
        log.insert(tk.END, f"❌ 템플릿 열기 실패: {template_path} ({e})\n")
        log.see(tk.END)
        return False

    # 1) 창 활성화/최대화
    win.activate(); time.sleep(0.2)
    win.maximize(); time.sleep(0.5)
    # win = get_real_kakao_window()

    # 2) 화면 캡처 영역 계산
    mon    = mss.mss().monitors[0]
    left   = max(win.left,  mon["left"])
    top    = max(win.top,   mon["top"])
    width  = min(win.width,  mon["width"]  - (left - mon["left"]))
    height = min(win.height, mon["height"] - (top  - mon["top"]))
    region = {"left": left, "top": top, "width": width, "height": height}

    # 3) mss → PIL → OpenCV 그레이스케일
    with mss.mss() as sct:
        sct_img = sct.grab(region)
        hay_pil = Image.frombytes("RGB", (sct_img.width, sct_img.height), sct_img.rgb)
    
    hay_cv = cv2.cvtColor(np.array(hay_pil), cv2.COLOR_RGB2GRAY)

    # 4) 템플릿도 OpenCV 그레이스케일로
    tpl_cv = cv2.cvtColor(np.array(pil_tpl), cv2.COLOR_RGB2GRAY)
    h_tpl, w_tpl = tpl_cv.shape[:2]

    # 5) 멀티스케일 매칭
    best = None  # (confidence, x, y, scale)
    scales = np.linspace(0.6, 1.4, 17)  # 60%~140% 범위를 17단계로
    for scale in scales:
        nw, nh = int(w_tpl*scale), int(h_tpl*scale)
        if nw < 10 or nh < 10:
            continue
        tpl_rs = cv2.resize(tpl_cv, (nw, nh), interpolation=cv2.INTER_AREA)
        res    = cv2.matchTemplate(hay_cv, tpl_rs, cv2.TM_CCOEFF_NORMED)
        _, max_v, _, max_loc = cv2.minMaxLoc(res)

        if max_v >= confidence and (best is None or max_v > best[0]):
            best = (max_v, max_loc[0], max_loc[1], scale)

    if not best:
        log.insert(tk.END, f"❌ 아이콘 '{template_path}' 못 찾음 (멀티스케일)\n"); log.see(tk.END)
        return False

    val, x, y, scale = best
    # 6) 화면 절대 좌표 계산 & 클릭
    cx = region["left"] + x + int(w_tpl*scale/2)
    cy = region["top"]  + y + int(h_tpl*scale/2)
    pyautogui.click(cx, cy)
    time.sleep(0.2)
    return True

def get_chat_window(name, timeout=3.0, interval=0.1):
    elapsed = 0
    while elapsed < timeout:
        for w in gw.getWindowsWithTitle(name):
            if w.title == name and w.width > 200 and w.height > 200:
                return w
        time.sleep(interval); elapsed += interval
    return None

def get_real_kakao_window():
    for w in gw.getWindowsWithTitle("카카오톡"):
        if w.title == "카카오톡" and w.width > 300:
            return w
    return None

MAX_MSG_LEN = 20
def shorten_message(msg):
    return msg if len(msg) <= MAX_MSG_LEN else msg[:MAX_MSG_LEN] + "..." 

def show_full_message(event):
    sel = tree.focus()
    if not sel: return
    idx = int(sel)
    full = message_data.at[idx, "메시지"]
    messagebox.showinfo("전체 메시지", full)

def load_json(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}
def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ——— Tooltip (마우스 오버 전체 메시지) ——————————————————————
class ToolTip:
    def __init__(self, widget):
        self.widget = widget; self.tipwindow = None
    def showtip(self, text):
        if self.tipwindow or not text: return
        x = self.widget.winfo_pointerx()+20; y = self.widget.winfo_pointery()+20
        tw = tk.Toplevel(self.widget); tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        msg = tk.Message(tw, text=text, width=200, justify="left",
                         font=("tahoma",8), background="#ffffe0",
                         relief="solid", borderwidth=1)
        msg.pack(ipadx=1, ipady=1); self.tipwindow = tw
    def hidetip(self):
        if self.tipwindow: self.tipwindow.destroy(); self.tipwindow = None

# ——— Excel 로드 —————————————————————————————————————
def load_excel():
    global message_data
    path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xls")])
    if not path: return
    try: message_data = pd.read_excel(path, engine="openpyxl")
    except: message_data = pd.read_excel(path)
    cols = message_data.columns.tolist()
    if "이름" not in cols or "메시지" not in cols:
        message_data = message_data.rename(columns={cols[0]:"이름", cols[1]:"메시지"})
    tree.delete(*tree.get_children())
    for idx, row in message_data.iterrows():
        tree.insert("", "end", iid=idx, values=(row["이름"], shorten_message(row["메시지"])))
    messagebox.showinfo("로드 완료", "Excel 로드 성공!")

# ——— 자동응답 Q&A ————————————————————————————————————
def add_qa():
    q,a = question_var.get(), answer_var.get()
    if not (q and a): messagebox.showwarning("입력 필요","질문·답변 모두 입력"); return
    auto_reply[q] = a; qa_tree.insert("", "end", values=(q,a))
    save_json(QA_FILE, auto_reply); question_var.set(""); answer_var.set("")
def delete_qa():
    sel = qa_tree.selection()
    if not sel: return
    q = qa_tree.item(sel, "values")[0]
    del auto_reply[q]; qa_tree.delete(sel); save_json(QA_FILE, auto_reply)

# ——— 자동응답 대상 ————————————————————————————————————
def add_user():
    u = user_var.get().strip()
    if not u or u in users: return
    users[u]=""; user_list.insert(tk.END,u)
    save_json(USERS_FILE, users); user_var.set("")
def delete_user():
    sel = user_list.curselection()
    if not sel: return
    u=user_list.get(sel); del users[u]; user_list.delete(sel)
    save_json(USERS_FILE, users)

# ——— 카톡 메시지 전송 (boolean 반환) ——————————————————————
def send_kakao_message(name, msg):
    try:
        print(f"[send_kakao_message] 시작 → {name}")  # ★
        win = get_real_kakao_window()
        if not win:
            print("  ❌ 카카오톡 창을 못 찾았습니다.")
            return False

        # 친구탭 클릭
        print("  → 친구탭 클릭 시도")
        if not click_friend_tab_with_mss(win, resource_path('img/friend_tab.png')):
            print("  ❌ 친구탭 클릭 실패")
            return False
        time.sleep(0.2)

        # 검색
        print(f"  → '{name}' 검색")
        pyautogui.hotkey("ctrl","f"); time.sleep(0.1)
        for _ in range(20): pyautogui.press("backspace")
        pyperclip.copy(name); pyautogui.hotkey("ctrl","v"); time.sleep(0.1)
        pyautogui.press("enter"); time.sleep(0.2)

        # 채팅창 팝업 검증
        chat_win = get_chat_window(name)
        if not chat_win or chat_win.title != name:
            print(f"  ⚠️ 팝업창 검증 실패 (찾은: {chat_win and chat_win.title})")
            return False

        # 메시지 입력창 클릭
        print("  → 메시지 입력창 활성화")
        clicked_message = click_friend_tab_with_mss(chat_win, resource_path('img/message_insert.png'))
        if not clicked_message:
            print(f"  ⚠️ 메시지 입력창 선택 실패")
            return False

        # 실제 전송
        print(f"  → 메시지 전송: {msg[:20]}{'...' if len(msg)>20 else ''}")
        pyautogui.hotkey("ctrl","a"); time.sleep(0.1)
        pyautogui.press("backspace")
        pyperclip.copy(msg); pyautogui.hotkey("ctrl","v"); time.sleep(0.1)
        pyautogui.press("enter"); pyautogui.press("esc")

        print(f"✅ [send_kakao_message] {name} 전송 성공")  # ★
        return True

    except Exception as e:
        print(f"❌ [send_kakao_message] {name} 전송 실패: {e}")
        return False


def send_selected():
    sel = tree.selection()
    if not sel:
        messagebox.showwarning("선택 필요", "전송할 메시지를 선택하세요.")
        return

    total = len(sel)
    print(f"[send_selected] 총 {total}명 전송 시작")  # ★

    for i, iid in enumerate(sel, start=1):
        name = tree.item(iid, "values")[0]
        msg  = message_data.at[int(iid), "메시지"]

        print(f"[{i}/{total}] → {name}님 전송 시도")  # ★

        ok = send_kakao_message(name, msg)
        if ok:
            print(f"   ✅ {name}님 전송 완료")  # ★
            tree.item(iid, tags=('success',))
        else:
            print(f"   ❌ {name}님 전송 실패")  # ★
            tree.item(iid, tags=('failure',))

    print("[send_selected] 전송 작업 완료")  # ★
    messagebox.showinfo("완료", f"{total}명에게 전송 작업이 완료되었습니다.")


# ——— 기타 액션 ————————————————————————————————————
def select_all():
    for i in tree.get_children(): tree.selection_add(i)

def delete_selected():
    sel = tree.selection()
    if not sel: return
    if not messagebox.askyesno("삭제 확인","선택된 메시지 삭제?"): return
    positions = [int(iid) for iid in sel]
    message_data.drop(message_data.index[positions], inplace=True)
    message_data.reset_index(drop=True, inplace=True)
    tree.delete(*tree.get_children())
    for idx,row in message_data.iterrows():
        tree.insert("", "end", iid=idx,
                    values=(row["이름"], shorten_message(row["메시지"])))
    messagebox.showinfo("삭제 완료","삭제되었습니다")

# ——— 초기 데이터 로드 —————————————————————————————
QA_FILE    = resource_path("auto_reply.json")
USERS_FILE = resource_path("auto_reply_users.json")
auto_reply = load_json(QA_FILE)
users      = load_json(USERS_FILE)

# ——— GUI 세팅 ————————————————————————————————————
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
app = ctk.CTk()
app.title("💬 카톡 자동 메시지 매크로")
app.geometry("800x900")

# 메시지 자동전송 프레임
ctk.CTkLabel(app, text="💬 메시지 자동전송", font=("Arial",16,"bold")).pack(pady=5)
fm = ctk.CTkFrame(app); fm.pack(fill="x", padx=20)
ctk.CTkButton(fm, text="📂 Excel 로드",        command=load_excel,     width=120).pack(side="left", padx=5)
ctk.CTkButton(fm, text="📩 선택 메시지 전송",  command=send_selected,  width=180).pack(side="left", padx=5)
ctk.CTkButton(fm, text="✔ 전체 선택",        command=select_all,     width=120).pack(side="left", padx=5)
ctk.CTkButton(fm, text="❌ 선택 삭제",        command=delete_selected,width=120).pack(side="left", padx=5)

# 메시지 Treeview
tree_frame = ctk.CTkScrollableFrame(app, height=150)
tree_frame.pack(fill="x", padx=20, pady=5)
tree = ttk.Treeview(tree_frame, columns=("이름","메시지"), show="headings")
tree.heading("이름", text="이름");    tree.column("이름", width=120, anchor="center")
tree.heading("메시지", text="메시지"); tree.column("메시지", width=500, anchor="w")
tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", show_full_message)

# 툴팁
tooltip = ToolTip(tree)
def on_motion(e):
    region = tree.identify("region", e.x, e.y)
    if region=="cell":
        row, col = tree.identify_row(e.y), tree.identify_column(e.x)
        if row and col=="#2":
            tooltip.showtip(message_data.at[int(row),"메시지"])
        else: tooltip.hidetip()
    else: tooltip.hidetip()
tree.bind("<Motion>", on_motion)
tree.bind("<Leave>", lambda e: tooltip.hidetip())

# 태그 스타일
tree.tag_configure('success', background='#d0f0c0')
tree.tag_configure('failure', background='#f0d0d0')

# 진행 상황 표시 위젯
progress_label = ctk.CTkLabel(app, text="진행 상황: 0/0", anchor="w")
progress_label.pack(fill="x", padx=20, pady=(5,0))
progress_bar = ctk.CTkProgressBar(app, width=760)
progress_bar.set(0)
progress_bar.pack(fill="x", padx=20, pady=(0,10))

# 🤖 자동응답 Q&A 프레임
ctk.CTkLabel(app, text="🤖 자동응답 매크로", font=("Arial",16,"bold")).pack(pady=5)
qf = ctk.CTkFrame(app); qf.pack(fill="x", padx=20)
question_var = tk.StringVar(); answer_var = tk.StringVar()
ctk.CTkEntry(qf, placeholder_text="질문",  textvariable=question_var, width=200).pack(side="left", padx=5)
ctk.CTkEntry(qf, placeholder_text="답변",  textvariable=answer_var,   width=300).pack(side="left", padx=5)
ctk.CTkButton(qf, text="➕ 추가", command=add_qa).pack(side="left", padx=5)
ctk.CTkButton(qf, text="❌ 삭제", command=delete_qa).pack(side="left", padx=5)

qa_frame = ctk.CTkScrollableFrame(app, height=100); qa_frame.pack(fill="both", padx=20, pady=5)
qa_tree = ttk.Treeview(qa_frame, columns=("질문","답변"), show="headings", height=6)
qa_tree.heading("질문", text="질문"); qa_tree.heading("답변", text="답변")
qa_tree.column("질문", width=200, anchor="center"); qa_tree.column("답변", width=500, anchor="w")
qa_tree.pack(fill="both", expand=True)
for q,a in auto_reply.items():
    qa_tree.insert("", "end", values=(q,a))

# 👤 자동응답 대상 프레임
ctk.CTkLabel(app, text="👤 자동응답 대상", font=("Arial",16,"bold")).pack(pady=5)
uf = ctk.CTkFrame(app); uf.pack(fill="x", padx=20)
user_var = tk.StringVar()
ctk.CTkEntry(uf, placeholder_text="대상 이름", textvariable=user_var, width=250).pack(side="left", padx=5)
ctk.CTkButton(uf, text="➕ 추가", command=add_user).pack(side="left", padx=5)
ctk.CTkButton(uf, text="❌ 삭제", command=delete_user).pack(side="left", padx=5)

user_frame = ctk.CTkScrollableFrame(app, height=100); user_frame.pack(fill="both", padx=20, pady=5)
user_list = tk.Listbox(user_frame); user_list.pack(fill="both", expand=True)
for u in users: user_list.insert(tk.END, u)

# 📝 로그창
log_frame = ctk.CTkFrame(app); log_frame.pack(fill="both", expand=True, padx=20, pady=10)
ctk.CTkLabel(log_frame, text="📝 로그", font=("나눔고딕",14,"bold")).pack(pady=5)
# log = ctk.CTkTextbox(log_frame, height=100, state="normal")
log = ctk.CTkTextbox(
    log_frame,
    height=100,
    state="normal",
    text_color="white",    # ← 여기 추가
    fg_color="#2b2b2b"     # (옵션) 배경색도 약간 밝게
)
log.pack(fill="both", expand=True, padx=10, pady=5)

app.mainloop()


class StdoutRedirector:
    def __init__(self, widget):
        self.widget = widget
    def write(self, s):
        s = s.replace('\r','')
        self.widget.insert(tk.END, s)
        self.widget.see(tk.END)
    def flush(self):
        pass

# ↓ log 위젯을 만든 직후
sys.stdout = StdoutRedirector(log)
sys.stderr = StdoutRedirector(log)