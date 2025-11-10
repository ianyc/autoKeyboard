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
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
# we will register Ctrl+Alt+1 and Ctrl+Alt+2
HOTKEY_ID_ACCOUNT = 1
HOTKEY_ID_PASSWORD = 2
VK_1 = 0x31
VK_2 = 0x32
VK_Z = 0x5A
VK_X = 0x58

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
    keyring.set_password(SERVICE_NAME, key, value)

def load_secret(key: str):
    return keyring.get_password(SERVICE_NAME, key)

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
def message_loop(stop_event, on_account_hotkey, on_password_hotkey):
    # register hotkeys
    if not user32.RegisterHotKey(None, HOTKEY_ID_ACCOUNT, MOD_CONTROL | MOD_SHIFT, VK_Z):
        print("RegisterHotKey account failed")
    if not user32.RegisterHotKey(None, HOTKEY_ID_PASSWORD, MOD_CONTROL | MOD_SHIFT, VK_X):
        print("RegisterHotKey password failed")

    try:
        msg = wintypes.MSG()
        while not stop_event.is_set():
            has_msg = user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1)  # PM_REMOVE=1, non-blocking
            if has_msg:
                if msg.message == WM_HOTKEY:
                    if msg.wParam == HOTKEY_ID_ACCOUNT:
                        on_account_hotkey()
                    elif msg.wParam == HOTKEY_ID_PASSWORD:
                        on_password_hotkey()
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                # sleep a short while to avoid busy loop
                time.sleep(0.05)
    finally:
        # unregister hotkeys
        user32.UnregisterHotKey(None, HOTKEY_ID_ACCOUNT)
        user32.UnregisterHotKey(None, HOTKEY_ID_PASSWORD)

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
        root.geometry("320x200")
        # icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        # root.iconbitmap(icon_path)
        root.iconbitmap(resource_path("icon.ico"))
        root.resizable(False, False)

        self.lbl_acc = tk.Label(root, text="帳號: (未設定)")
        self.lbl_acc.pack(pady=(12,4))
        btn_set_acc = tk.Button(root, text="設定帳號", command=self.set_account)
        btn_set_acc.pack()

        self.lbl_pwd = tk.Label(root, text="密碼: (未設定)")
        self.lbl_pwd.pack(pady=(12,4))
        btn_set_pwd = tk.Button(root, text="設定密碼", command=self.set_password)
        btn_set_pwd.pack()

        # tk.Label(root, text="熱鍵：Ctrl+shitf+z = 帳號，Ctrl+Alt+2 = 密碼", fg="gray").pack(pady=(8,0))

        tk.Button(root, text="清空儲存 (從 Credential Manager 移除)", command=self.clear_storage).pack(pady=(8,0))

        # load current
        self.refresh_labels()

        # start hotkey message loop in background thread
        self.stop_event = threading.Event()
        t = threading.Thread(target=message_loop, args=(self.stop_event, self.on_account_hotkey, self.on_password_hotkey), daemon=True)
        t.start()

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Unmap>", self.on_minimize)

    def refresh_labels(self):
        acc = load_secret("account")
        pwd = load_secret("password")
        self.lbl_acc.config(text=f"帳號: {acc[:6] + '...' if acc else '(未設定)'}")
        self.lbl_pwd.config(text=f"密碼: {'已設定' if pwd else '(未設定)'}")

    def set_account(self):
        val = simpledialog.askstring("設定帳號", "請輸入帳號：", parent=self.root)
        if val is not None:
            save_secret("account", val)
            # messagebox.showinfo("已儲存", "帳號已儲存至 Windows Credential Manager（使用者專屬）")
            self.refresh_labels()

    def set_password(self):
        val = simpledialog.askstring("設定密碼", "請輸入密碼：", parent=self.root, show="*")
        if val is not None:
            save_secret("password", val)
            # messagebox.showinfo("已儲存", "密碼已儲存至 Windows Credential Manager（使用者專屬）")
            self.refresh_labels()

    def clear_storage(self):
        # keyring delete_password may raise depending backend; handle gracefully
        try:
            keyring.delete_password(SERVICE_NAME, "account")
        except Exception:
            pass
        try:
            keyring.delete_password(SERVICE_NAME, "password")
        except Exception:
            pass
        messagebox.showinfo("已清除", "已從 Credential Manager 移除儲存的帳號/密碼（若存在）")
        self.refresh_labels()

    def on_account_hotkey(self):
        acc = load_secret("account")
        if not acc:
            # optional: 在前景顯示提示
            print("帳號未設定")
            return
        put_on_clipboard_and_paste(acc)

    def on_password_hotkey(self):
        pwd = load_secret("password")
        if not pwd:
            print("密碼未設定")
            return
        put_on_clipboard_and_paste(pwd)

    def on_close(self):
        # stop message loop
        self.stop_event.set()
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

    def on_close(self):
        self.stop_event.set()
        time.sleep(0.1)
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk() # 建立UI視窗
    app = App(root) # 綁定function
    root.mainloop() # 啟動GUI事件循環
