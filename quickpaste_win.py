# quickpaste_win.py
import threading
import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import simpledialog, messagebox
import keyring
import pyperclip
import time
import pystray
from PIL import Image
import os
import sys


# define total count of hotkey (>7要再補VK_LETTERS)
def ask_hotkey_count():
    root = tk.Tk()
    root.withdraw() # 不顯示主視窗
    while True:
        val = simpledialog.askstring(" ", "請輸入字詞數量 (1~7):")
        if val is None:
            sys.exit(0)  # 使用者按取消就結束程式
        if val.isdigit() and 1 <= int(val) <= 7:
            root.destroy()
            return int(val)
        else:
            messagebox.showerror("Be Careful", "僅能輸入1~7")

TOTAL_HOTKEY_COUNT = ask_hotkey_count()

# --------------------
# 常數與 Win32 API
# --------------------
user32 = ctypes.windll.user32
WM_HOTKEY = 0x0312

# virtual key codes
VK_CONTROL = 0x11
VK_V = 0x56

KEYEVENTF_KEYUP = 0x0002

# hotkey modifiers for RegisterHotKey
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
VK_LETTERS = [0x5A, 0x58, 0x31, 0x32, 0x33, 0x34, 0x35]  # Z, X, 1, 2, 3, 4, 5

SERVICE_NAME = "QuickPasteLocal"  # 用於 keyring

# --------------------
# 安全/行為設定
# --------------------
CLEAR_CLIP_AFTER = 5.0  # 貼上後多少秒清空剪貼簿（若為 0 則不清空）
# 若你想恢復剪貼簿原本內容，可改寫下面的邏輯；但為簡單起見這裡清空（安全優先）

# --------------------
# keyring 存取
# --------------------
def save_secret(key: str, value: str):
    """把 value 存到 Windows Credential Manager (keyring)"""
    # keyring.set_password(service, username, password)
    keyring.set_password(f"{SERVICE_NAME}_{key}", key, value)

def load_secret(key: str):
    return keyring.get_password(f"{SERVICE_NAME}_{key}", key)

# --------------------
# Clipboard 與貼上
# --------------------
def put_on_clipboard_and_paste(text: str):
    """把文字放到剪貼簿，模擬 Ctrl+V 貼上，然後視設定清空剪貼簿"""
    if not text:
        return
    try:
        # 儲存目前剪貼簿（如果你想要恢復可在此擴充）
        # prev = pyperclip.paste()
        pyperclip.copy(text)
        # 很短的延遲確保剪貼簿已更新
        time.sleep(0.05)
        # 模擬 Ctrl+V
        # keybd_event(VK_CONTROL, ...); keybd_event(VK_V,...)
        user32.keybd_event(VK_CONTROL, 0, 0, 0)  # ctrl down
        user32.keybd_event(VK_V, 0, 0, 0)        # v down
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)  # v up
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)  # ctrl up
        # 清空或恢復剪貼簿
        if CLEAR_CLIP_AFTER and CLEAR_CLIP_AFTER > 0:
            def clear_clip():
                # 簡單地清空剪貼簿（安全考量）
                pyperclip.copy("")
            t = threading.Timer(CLEAR_CLIP_AFTER, clear_clip)
            t.daemon = True
            t.start()
    except Exception as e:
        print("paste err:", e)

# --------------------
# Hotkey message loop
# --------------------
def message_loop(stop_event, on_str_hotkey):
    # register hotkeys
    for i in range(TOTAL_HOTKEY_COUNT):
        hotkey_id = i + 1
        vk = VK_LETTERS[i]
        if not user32.RegisterHotKey(None, hotkey_id, MOD_CONTROL | MOD_SHIFT, vk):
            print(f"RegisterHotKey str{i+1} failed")

    try:
        msg = wintypes.MSG()
        while not stop_event.is_set():
            has_msg = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE=1, non-blocking
            if has_msg:
                if msg.message == WM_HOTKEY:
                    idx = msg.wParam - 1
                    if 0 <= idx < TOTAL_HOTKEY_COUNT:
                        on_str_hotkey(f"str{idx+1}")
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                # sleep a short while to avoid busy loop
                time.sleep(0.05)
    finally:
        # unregister hotkeys
        for i in range(TOTAL_HOTKEY_COUNT):
            user32.UnregisterHotKey(None, i + 1)

def resource_path(relative_path):
    """獲取資源在執行檔或開發環境中的正確路徑"""
    if hasattr(sys, '_MEIPASS'):  # PyInstaller 打包後執行時會有這個屬性
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# --------------------
# GUI (Tkinter)
# --------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("神奇小工具")
        width = 300
        height = 100 + TOTAL_HOTKEY_COUNT*60
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        root.geometry(f"{width}x{height}+{x}+{y}")
        # icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        # root.iconbitmap(icon_path)
        root.iconbitmap(resource_path("icon.ico"))
        root.resizable(False, False)

        self.labels = []
        self.buttons = []
        for i in range(TOTAL_HOTKEY_COUNT):
            name = f"str{i+1}"
            lbl = tk.Label(root, text=f"字詞{i+1}: (未設定)")
            lbl.pack(pady=(12 if i == 0 else 6, 4))
            btn = tk.Button(root, text=f"設定字詞{i+1}", command=lambda i=i: self.set_str(i+1))
            btn.pack()
            self.labels.append(lbl)
            self.buttons.append(btn)

        tk.Button(root, text="☆★ 清除所有字詞資料 ★☆", command=self.clear_storage).pack(pady=(10,0))

        tk.Label(root, text="組合鍵: Ctrl+Shift", fg="gray").pack(pady=(8,0))
        TEXTS = ["z", "x", "1", "2", "3", "4", "5"]
        hotkey_text = "熱鍵: " + ", ".join(TEXTS[:TOTAL_HOTKEY_COUNT])
        tk.Label(root, text=hotkey_text, fg="gray").pack(pady=(0,0))

        # load current
        self.refresh_labels()

        # start hotkey message loop in background thread
        self.stop_event = threading.Event()
        t = threading.Thread(target=message_loop, args=(self.stop_event, self.on_str_hotkey), daemon=True)
        t.start()

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Unmap>", self.on_minimize)

    def refresh_labels(self):
        for i, lbl in enumerate(self.labels):
            key = f"str{i+1}"
            val = load_secret(key)
            lbl.config(text=f"字詞{i+1}: {val[:6] + '...' if val else '(未設定)'}")

            if val:
                self.buttons[i].config(state="disabled", bg="lightgray")
            else:
                self.buttons[i].config(state="normal", bg="SystemButtonFace")

    def set_str(self, target):
        val = simpledialog.askstring("設定字詞", f"請輸入字詞{target}", parent=self.root)
        if val is not None:
            save_secret(f"str{target}", val)
            self.refresh_labels()

    def clear_storage(self):
        # keyring delete_password may raise depending backend; handle gracefully
        for i in range(TOTAL_HOTKEY_COUNT):
            key = f"str{i+1}"
            try:
                keyring.delete_password(f"{SERVICE_NAME}_{key}", key)
            except Exception:
                pass
        # messagebox.showinfo("Oh My God", "已從 Credential Manager 清除所有字詞資料")
        self.refresh_labels()

    def on_str_hotkey(self, target):
        val = load_secret(target)
        if not val:
            # optional: 在前景顯示提示
            print("字詞未設定")
            return
        put_on_clipboard_and_paste(val)

    def on_close(self):
        # stop message loop
        self.stop_event.set()
        # 清空所有儲存的字詞資料
        self.clear_storage()
        # wait briefly for thread to end
        time.sleep(0.1)
        self.root.destroy()

    def on_minimize(self, event):
        """當使用者最小化視窗時，隱藏到托盤"""
        if self.root.state() == "iconic":
            self.root.withdraw()
            self.show_tray_icon()

    def show_tray_icon(self):
        """建立托盤圖示"""
        # icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        # image = Image.open(icon_path)

        def show_window(icon=None, item=None):
            """顯示主視窗"""
            self.root.deiconify()
            self.root.after(10, self.root.lift)

        def quit_app(icon, item):
            icon.stop()
            self.root.after(0, self.root.destroy)

        menu = pystray.Menu(
            pystray.MenuItem("開啟 QuickPaste", show_window),
            pystray.MenuItem("退出", quit_app)
        )

        self.icon = pystray.Icon("QuickPaste", Image.open(resource_path("icon.ico")), "QuickPaste", menu)

        t = threading.Thread(target=self.icon.run, daemon=True)
        t.start()

if __name__ == "__main__":
    root = tk.Tk() # 建立UI視窗
    app = App(root) # 綁定function
    root.mainloop() # 啟動GUI事件循環
