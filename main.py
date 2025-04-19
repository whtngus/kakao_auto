import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyautogui
import pyperclip
import pygetwindow as gw
import time
import json
import os


def get_real_kakao_window():
    windows = gw.getWindowsWithTitle("카카오톡")
    for win in windows:
        if win.title == "카카오톡" and win.width > 300:
            return win
    return None

# 메시지 최대 길이 지정 및 메시지 축약 함수
MAX_MSG_LEN = 20
def shorten_message(msg):
    return msg if len(msg) <= MAX_MSG_LEN else msg[:MAX_MSG_LEN] + "..."

def show_full_message(event):
    selected_item = tree.focus()
    if selected_item:
        full_message = message_data.loc[tree.index(selected_item), "메시지"]
        messagebox.showinfo("전체 메시지", full_message)
        
        
# 🌟 자동응답 및 메시지 데이터 로드
def load_json(filename):
    if os.path.exists(filename):
        with open(filename, "r", encoding="utf-8") as file:
            return json.load(file)
    return {}

def save_json(filename, data):
    with open(filename, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=4)
        
# Excel 로드 함수 최적화(인덱스 처리)
def load_excel():
    global message_data
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        message_data = pd.read_excel(file_path)
        tree.delete(*tree.get_children())
        for idx, row in message_data.iterrows():
            short_msg = shorten_message(row["메시지"])
            tree.insert("", "end", iid=idx, values=(row["이름"], short_msg))
        messagebox.showinfo("로드 완료", "Excel 로드 성공!")

# QA 추가/삭제 최적화
def add_qa():
    q, a = question_entry_var.get(), answer_entry_var.get()
    if q and a:
        auto_reply_dict[q] = a
        qa_tree.insert("", "end", values=(q, a))
        save_json(AUTO_REPLY_FILE, auto_reply_dict)
        question_entry_var.set("")
        answer_entry_var.set("")
    else:
        messagebox.showwarning("입력 필요", "질문/답변 모두 입력!")

def delete_qa():
    selected = qa_tree.selection()
    if selected:
        q = qa_tree.item(selected, "values")[0]
        qa_tree.delete(selected)
        del auto_reply_dict[q]
        save_json(AUTO_REPLY_FILE, auto_reply_dict)

def add_user():
    user = user_entry_var.get()
    if user and user not in auto_reply_users:
        auto_reply_users[user] = ''
        user_listbox.insert(tk.END, f"👤 {user}")
        save_json(AUTO_REPLY_USERS_FILE, auto_reply_users)
        user_entry_var.set("")

def delete_selected_messages():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("선택 필요", "삭제할 항목을 선택하세요!")
        return

    if messagebox.askyesno("삭제 확인", "정말 선택한 메시지를 삭제하시겠습니까?"):
        global message_data
        indexes_to_drop = []
        for item in selected_items:
            index = int(item)
            indexes_to_drop.append(index)
            tree.delete(item)

        message_data.drop(indexes_to_drop, inplace=True)
        message_data.reset_index(drop=True, inplace=True)

        messagebox.showinfo("삭제 완료", "선택한 항목이 삭제되었습니다.")
        
        
def delete_user():
    sel = user_listbox.curselection()
    if sel:
        user = user_listbox.get(sel).replace("👤 ", "")
        user_listbox.delete(sel)
        del auto_reply_users[user]
        save_json(AUTO_REPLY_USERS_FILE, auto_reply_users)

# 📢 카카오톡 메시지 전송 함수
def send_kakao_message(name, message):
    try:
        kakao_window = get_real_kakao_window()
        if kakao_window:
            kakao_window.activate()
            time.sleep(0.1)
            kakao_window.maximize()  # 창 최대화 (더 안정적인 좌표 클릭)
            time.sleep(0.1)
        else:
            print("카카오톡 창이 발견되지 않았습니다.")

        kakao_window.activate()
        time.sleep(0.1)
        friend_tab_position = (50, 70)  # 좌측 상단 친구탭 아이콘 위치 (조정 필요)
        pyautogui.click(friend_tab_position)
        time.sleep(0.2)
        try:
            close_search_bar()
        except:
            pass
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.1)
        for _ in range(15):
            pyautogui.press('backspace')

        pyperclip.copy(name)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.1)
        pyautogui.press('enter')
        time.sleep(0.1)

        pyperclip.copy(message)
        for _ in range(15):
            pyautogui.press('backspace')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)
        pyautogui.press('enter')

        log_text.insert(tk.END, f"✅ {name}님에게 전송 완료\n")
        log_text.see(tk.END)

    except Exception as e:
        log_text.insert(tk.END, f"❌ {name}님에게 전송 실패: {e}\n")
        log_text.see(tk.END)
        
def send_selected_message():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("선택 필요", "메시지를 보낼 항목을 선택하세요!")
        return
    
    for item in selected_items:
        values = tree.item(item, "values")
        idx = int(item)  # Treeview의 아이템 ID는 인덱스임
        name = values[0]
        full_message = message_data.loc[idx, "메시지"]
        send_kakao_message(name, full_message)
        
def select_all():
    for item in tree.get_children():
        tree.selection_add(item)

def deselect_all():
    for item in tree.selection():
        tree.selection_remove(item)


def use_chatgpt():
    selected = qa_tree.selection()
    if selected:
        question = qa_tree.item(selected, "values")[0]
        messagebox.showinfo("ChatGPT", f"'{question}' 에 대한 ChatGPT 답변 생성")
    else:
        messagebox.showwarning("선택 필요", "질문을 선택하세요!")
        
        
        
# 🌟 JSON 로드 후 데이터 불러오기
AUTO_REPLY_FILE = "auto_reply.json"
AUTO_REPLY_USERS_FILE = "auto_reply_users.json"
auto_reply_dict = load_json(AUTO_REPLY_FILE)
auto_reply_users = load_json(AUTO_REPLY_USERS_FILE)

# === GUI 구성 시작 ===
# 🌙 예쁜 Dark 테마 스킨 적용
ctk.set_appearance_mode("dark")        # Dark 모드
ctk.set_default_color_theme("dark-blue")  # dark-blue 테마 (세련된 색상)
app = ctk.CTk()
app.title("💬 카카오톡 자동 메시지 & 자동응답 매크로")
app.geometry("800x900")

question_entry_var = tk.StringVar()
answer_entry_var = tk.StringVar()
user_entry_var = tk.StringVar()

# 📊 메시지 자동전송 프레임
ctk.CTkLabel(app, text="💬 카카오톡 메시지 자동전송", font=("Arial", 16, "bold")).pack(pady=5)

frame_msg = ctk.CTkFrame(app)
frame_msg.pack(fill="x", padx=20)

ctk.CTkButton(frame_msg, text="📂 Excel 로드", command=load_excel, width=120).pack(side="left", padx=5)
ctk.CTkButton(frame_msg, text="📩 선택 메시지 전송", command=send_selected_message, width=180).pack(side="left", padx=5)
ctk.CTkButton(frame_msg, text="✔ 전체 선택", command=select_all, width=120).pack(side="left", padx=5)
ctk.CTkButton(frame_msg, text="❌ 선택 삭제", command=delete_selected_messages, width=120).pack(side="left", padx=5)

tree_frame = ctk.CTkScrollableFrame(app, height=150)
tree_frame.pack(fill="x", padx=20, pady=5)

tree = ttk.Treeview(tree_frame, columns=("이름", "메시지"), show="headings")
tree.heading("이름", text="이름")
tree.heading("메시지", text="메시지 내용")
tree.column("이름", width=120, anchor="center")
tree.column("메시지", width=500, anchor="w")
tree.pack(fill="both", expand=True)
tree.bind("<Double-1>", show_full_message)

# 🤖 자동응답 매크로 (질문-답변 형식 & ChatGPT 버튼 추가)
ctk.CTkLabel(app, text="🤖 자동응답 매크로", font=("Arial", 16, "bold")).pack(pady=5)

qa_input_frame = ctk.CTkFrame(app)
qa_input_frame.pack(fill="x", padx=20)

ctk.CTkEntry(qa_input_frame, placeholder_text="질문 입력", width=200, textvariable=question_entry_var).pack(side="left", padx=5)
ctk.CTkEntry(qa_input_frame, placeholder_text="답변 입력", width=250, textvariable=answer_entry_var).pack(side="left", padx=5)
ctk.CTkButton(qa_input_frame, text="➕ 추가", command=add_qa).pack(side="left", padx=5)
ctk.CTkButton(app, text="❌ 선택 삭제", command=delete_qa).pack(pady=5)
ctk.CTkButton(qa_input_frame, text="🤖 ChatGPT 사용", command=use_chatgpt).pack(side="right", padx=5)
qa_scroll_frame = ctk.CTkScrollableFrame(app, height=100)
qa_scroll_frame.pack(fill="both", padx=20, pady=5)

qa_tree = ttk.Treeview(qa_scroll_frame, columns=("질문", "답변"), show="headings", height=10)
qa_tree.heading("질문", text="질문")
qa_tree.heading("답변", text="답변")
qa_tree.column("질문", width=200, anchor="center")
qa_tree.column("답변", width=500, anchor="w")
qa_tree.pack(fill="both", expand=True)

for q, a in auto_reply_dict.items():
    qa_tree.insert("", "end", values=(q, a))



# 👤 자동응답 대상
ctk.CTkLabel(app, text="👤 자동응답 대상", font=("Arial", 16, "bold")).pack(pady=5)

user_input_frame = ctk.CTkFrame(app)
user_input_frame.pack(fill="x", padx=20)

ctk.CTkEntry(user_input_frame, placeholder_text="자동응답 대상 추가", width=250, textvariable=user_entry_var).pack(side="left", padx=5)
ctk.CTkButton(user_input_frame, text="➕ 추가", command=add_user).pack(side="left", padx=5)
ctk.CTkButton(user_input_frame, text="❌ 선택 삭제", command=delete_user).pack(side="left", padx=5)

user_scroll_frame = ctk.CTkScrollableFrame(app, height=100)
user_scroll_frame.pack(fill="both", padx=20, pady=5)

user_listbox = tk.Listbox(user_scroll_frame)
user_listbox.pack(fill="both", expand=True)

for user in auto_reply_users:
    user_listbox.insert(tk.END, f"👤 {user}")

    
# 📜 로그창 추가 (하단 로그 메시지 표시용)
log_frame = ctk.CTkFrame(app)
log_frame.pack(fill="both", expand=True, padx=20, pady=10)

log_label = ctk.CTkLabel(log_frame, text="📝 로그 기록", font=("나눔고딕", 14, "bold"))
log_label.pack(pady=5)

log_text = ctk.CTkTextbox(log_frame, height=100, state="normal")
log_text.pack(fill="both", expand=True, padx=10, pady=5)

# 로그 추가 함수
def add_log(message):
    log_text.configure(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.configure(state="disabled")
    log_text.see(tk.END)
    

def close_search_bar():
    btn_location = pyautogui.locateCenterOnScreen('close_btn.png', confidence=0.9)
    if btn_location:
        pyautogui.click(btn_location)
        time.sleep(0.3)
    else:
        print("검색창 닫기 버튼을 찾지 못했습니다.")
        
# 🏁 GUI 실행
app.mainloop()
   