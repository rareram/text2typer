import tkinter as tk
from tkinter import ttk
import pyperclip
import pyautogui
import keyboard
import threading
import time
import win32gui
import win32api

class DualMonitorTypingMacro:
    def __init__(self, master):
        self.master = master
        master.title("고오급 복사 텍스트 타이핑 매크로")

        self.frame = ttk.Frame(master, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.status_label = ttk.Label(self.frame, text="대기 중")
        self.status_label.grid(column=0, row=0, columnspan=2, pady=5)

        self.preview_label = ttk.Label(self.frame, text="복사된 내용: ")
        self.preview_label.grid(column=0, row=1, columnspan=2, pady=5)

        self.copy_button = ttk.Button(self.frame, text="텍스트 복사", command=self.copy_text)
        self.copy_button.grid(column=0, row=2, pady=5)

        self.paste_button = ttk.Button(self.frame, text="타이핑 시작", command=self.start_typing)
        self.paste_button.grid(column=1, row=2, pady=5)

        self.monitor_check_button = ttk.Button(self.frame, text="모니터 확인", command=self.check_monitor)
        self.monitor_check_button.grid(column=0, row=3, columnspan=2, pady=5)

        self.progress = ttk.Progressbar(self.frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.grid(column=0, row=4, columnspan=2, pady=5)

        self.monitor_label = ttk.Label(self.frame, text="", font=("Arial", 16, "bold"))
        self.monitor_label.grid(column=0, row=5, columnspan=2, pady=10)

        self.clipboard_content = ""
        self.is_typing = False

        # ESC 키 이벤트 리스너 추가
        keyboard.on_press_key("esc", self.on_esc_press)

    def copy_text(self):
        keyboard.press_and_release('ctrl+c')
        time.sleep(0.1)  # 복사가 완료될 때까지 잠시 대기
        self.clipboard_content = pyperclip.paste()
        preview = self.clipboard_content[:10] + "..." if len(self.clipboard_content) > 10 else self.clipboard_content
        self.preview_label.config(text=f"복사된 내용: {preview}")
        self.status_label.config(text="텍스트가 복사되었습니다.")

    def start_typing(self):
        if not self.clipboard_content:
            self.status_label.config(text="복사된 내용이 없습니다.")
            return

        self.is_typing = True
        self.paste_button.config(state=tk.DISABLED)
        threading.Thread(target=self._type_on_second_monitor, daemon=True).start()

    def on_esc_press(self, e):
        if self.is_typing:
            self.is_typing = False
            self.paste_button.config(state=tk.NORMAL)
            self.status_label.config(text="타이핑이 중지되었습니다.")

    def check_monitor(self):
        window_handle = win32gui.GetForegroundWindow()
        monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromWindow(window_handle))
        monitor_number = 1 if monitor_info['Monitor'][0] == 0 else 2
        self.monitor_label.config(text=f"{monitor_number}번 모니터")

    def _type_on_second_monitor(self):
        self.status_label.config(text="2번 모니터로 이동 중...")
        
        # 2번 모니터로 마우스 이동
        screen_width, screen_height = pyautogui.size()
        pyautogui.moveTo(screen_width + 100, screen_height // 2)  # 2번 모니터의 중앙으로 이동
        
        # 메모장 활성화를 위해 클릭
        pyautogui.click()
        time.sleep(0.5)  # 메모장이 활성화될 때까지 대기

        self.status_label.config(text="타이핑 중...")
        self.progress['maximum'] = len(self.clipboard_content)

        for i, char in enumerate(self.clipboard_content):
            if not self.is_typing:
                break
            keyboard.write(char)  # pyautogui.write 대신 keyboard.write 사용
            time.sleep(0.01)  # 타이핑 속도 조절
            self.progress['value'] = i + 1
            self.master.update_idletasks()

        self.progress['value'] = 0
        self.paste_button.config(state=tk.NORMAL)
        self.status_label.config(text="타이핑 완료" if self.is_typing else "타이핑 중지됨")
        self.is_typing = False

root = tk.Tk()
app = DualMonitorTypingMacro(root)
root.mainloop()