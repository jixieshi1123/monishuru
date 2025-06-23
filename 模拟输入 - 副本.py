import tkinter as tk
from tkinter import ttk, scrolledtext
import win32api
import win32con
import time
import threading

class RealKeyboardSimulator:
    def __init__(self, root):
        self.root = root
        self.root.title("真实键盘输入模拟器 - Windows 版")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 添加窗口始终置顶选项
        self.always_on_top = tk.BooleanVar(value=True)
        self.set_window_topmost()
        
        self.running = False
        self.thread = None
        self.create_widgets()

    def set_window_topmost(self):
        """设置窗口是否始终置顶"""
        self.root.attributes('-topmost', self.always_on_top.get())
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="输入要模拟的内容:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.text_input = scrolledtext.ScrolledText(main_frame, width=60, height=10)
        self.text_input.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        delay_frame = ttk.LabelFrame(main_frame, text="输入设置", padding="10")
        delay_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(delay_frame, text="开始延迟(秒):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.start_delay = ttk.Scale(delay_frame, from_=1, to=10, orient=tk.HORIZONTAL, length=200)
        self.start_delay.set(3)
        self.start_delay.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(delay_frame, text="字符间隔(毫秒):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.char_delay = ttk.Scale(delay_frame, from_=10, to=500, orient=tk.HORIZONTAL, length=200)
        self.char_delay.set(80)
        self.char_delay.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        self.start_delay_value = ttk.Label(delay_frame, text="3秒")
        self.start_delay_value.grid(row=0, column=1, sticky=tk.E, pady=5)
        
        self.char_delay_value = ttk.Label(delay_frame, text="80毫秒")
        self.char_delay_value.grid(row=1, column=1, sticky=tk.E, pady=5)
        
        self.start_delay.bind("<Motion>", lambda e: self.update_delay_label(self.start_delay, self.start_delay_value, "秒"))
        self.char_delay.bind("<Motion>", lambda e: self.update_delay_label(self.char_delay, self.char_delay_value, "毫秒"))
        
        # 添加窗口置顶选项
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Checkbutton(
            top_frame, 
            text="窗口始终置顶", 
            variable=self.always_on_top,
            command=self.set_window_topmost
        ).pack(side=tk.LEFT)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="开始模拟", command=self.start_simulation)
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        self.stop_button = ttk.Button(button_frame, text="停止", command=self.stop_simulation, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=10)
        
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        ttk.Label(main_frame, textvariable=self.status_var).grid(row=5, column=0, sticky=tk.W, pady=5)
        
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(status_frame, text="提示: 开始前请将光标定位到目标位置").pack(side=tk.LEFT, padx=10, pady=5)

    def update_delay_label(self, scale, label, unit):
        value = int(scale.get())
        label.config(text=f"{value}{unit}")

    def start_simulation(self):
        text = self.text_input.get("1.0", tk.END).strip()
        start_delay = int(self.start_delay.get())
        char_delay = int(self.char_delay.get()) / 1000
        
        if not text:
            self.status_var.set("错误：请输入要模拟的内容")
            return
        
        self.status_var.set(f"准备开始，倒计时: {start_delay}秒")
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.running = True
        
        self.thread = threading.Thread(target=self.simulate_typing, args=(text, start_delay, char_delay))
        self.thread.daemon = True
        self.thread.start()

    def simulate_typing(self, text, start_delay, char_delay):
        try:
            for i in range(start_delay, 0, -1):
                if not self.running:
                    break
                self.status_var.set(f"准备开始，倒计时: {i}秒")
                time.sleep(1)
            
            if self.running:
                self.status_var.set("正在输入...")
                for char in text:
                    if not self.running:
                        break
                    self.press_key(char)
                    time.sleep(char_delay)
                
                if self.running:
                    self.status_var.set("输入完成")
        
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
        finally:
            self.root.after(0, self.reset_ui)

    def press_key(self, char):
        # 完整的虚拟键码映射表（用数字替代 win32con 常量）
        special_chars = {
            ' ': 32,       # 空格
            '\n': 13,      # 回车
            '\t': 9,       # Tab
            '!': (49, True),  # 1 + Shift
            '@': (50, True),  # 2 + Shift
            '#': (51, True),  # 3 + Shift
            '$': (52, True),  # 4 + Shift
            '%': (53, True),  # 5 + Shift
            '^': (54, True),  # 6 + Shift
            '&': (55, True),  # 7 + Shift
            '*': (56, True),  # 8 + Shift
            '(': (57, True),  # 9 + Shift
            ')': (48, True),  # 0 + Shift
            '-': 189,       # 减号
            '_': (189, True), # 下划线
            '=': 187,       # 等号
            '+': (187, True), # 加号
            '[': 219,       # 左方括号
            '{': (219, True), # 左花括号
            ']': 221,       # 右方括号
            '}': (221, True), # 右花括号
            '\\': 220,      # 反斜杠
            '|': (220, True), # 竖线
            ';': 186,       # 分号
            ':': (186, True), # 冒号
            "'": 222,       # 单引号
            '"': (222, True), # 双引号
            ',': 188,       # 逗号
            '<': (188, True), # 小于号
            '.': 190,       # 句号
            '>': (190, True), # 大于号
            '/': 191,       # 斜杠
            '?': (191, True), # 问号
            '`': 192,       # 波浪号
            '~': (192, True), # 波浪号
        }

        if char in special_chars:
            key_info = special_chars[char]
            if isinstance(key_info, tuple):
                key_code, need_shift = key_info
                if need_shift:
                    self._send_key_with_shift(key_code)
            else:
                self._send_key(key_info)
            return

        if char.isalnum():
            vk_code = ord(char.upper()) if char.isalpha() else ord(char)
            if char.isupper() or (char in '!@#$%^&*()_+{}|:"<>?~`'):
                self._send_key_with_shift(vk_code)
            else:
                self._send_key(vk_code)
        else:
            vk_code = win32api.VkKeyScan(char)
            if vk_code != -1:
                if vk_code & 0x100:
                    self._send_key_with_shift(vk_code & 0xFF)
                else:
                    self._send_key(vk_code & 0xFF)

    def _send_key(self, key_code):
        win32api.keybd_event(key_code, 0, 0, 0)
        time.sleep(0.01)
        win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)

    def _send_key_with_shift(self, key_code):
        win32api.keybd_event(16, 0, 0, 0)  # Shift 键码
        time.sleep(0.01)
        win32api.keybd_event(key_code, 0, 0, 0)
        time.sleep(0.01)
        win32api.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.01)
        win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)

    def stop_simulation(self):
        self.running = False
        self.status_var.set("已停止")
        self.reset_ui()

    def reset_ui(self):
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.running = False

if __name__ == "__main__":
    import platform
    if platform.system() != "Windows":
        print("此程序仅支持 Windows 系统")
    else:
        try:
            import win32api
            root = tk.Tk()
            app = RealKeyboardSimulator(root)
            root.mainloop()
        except ImportError:
            print("缺少 win32api 模块，请安装 pywin32: pip install pywin32")
