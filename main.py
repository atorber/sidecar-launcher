import tkinter as tk
from tkinter import ttk
import subprocess
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Injector Controller")
        self.root.geometry("300x200")
        self.root.resizable(False, False)

        # 设置样式
        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 12), padding=10)
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TFrame", background="#f0f0f0")
        
        # 创建一个框架
        frame = ttk.Frame(root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 添加标题
        title_label = ttk.Label(frame, text="Injector Controller", font=("Helvetica", 16, "bold"))
        title_label.pack(pady=10)

        # 添加启动按钮
        self.start_button = ttk.Button(frame, text="启动", command=self.start_injection)
        self.start_button.pack(pady=5)

        # 添加停止按钮
        self.stop_button = ttk.Button(frame, text="停止", command=self.stop_injection)
        self.stop_button.pack(pady=5)
        self.stop_button.config(state=tk.DISABLED)  # 初始状态禁用停止按钮

        # 添加状态标签
        self.status_label = ttk.Label(frame, text="状态：未启动", font=("Helvetica", 12))
        self.status_label.pack(pady=10)

        self.process = None

    def start_injection(self):
        if self.process is None:
            # 获取wechat.exe的路径的进程号,运行指令：tasklist | findstr WeChat.exe
            process = os.popen("tasklist | findstr WeChat.exe").read()
            print("WeChat.exe的进程信息：", process)
            if "WeChat.exe" not in process:
                self.status_label.config(text="状态：未找到 WeChat.exe 进程")
                return

            # 获取进程号
            process_id = process.split()[1]
            print("WeChat.exe的进程号为：", process_id)
            # 获取当前根目录
            current_path = os.getcwd()
            # 拼接Injector.exe的路径
            injector_path = os.path.join(current_path, "assets", "Injector.exe")
            # 拼接wxhelper.dll的路径
            wxhelper_path = os.path.join(current_path, "assets", "wxhelper-native.dll")
            
            command = [injector_path, "-p", process_id, "--inject", wxhelper_path]
            print("注入指令", command)
            self.process = subprocess.Popen(command)
            self.status_label.config(text="状态：已启动")
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            print("注入已启动")
        else:
            self.status_label.config(text="状态：已注入")

    def stop_injection(self):
        if self.process is not None:
            self.process.terminate()
            self.process = None
            self.status_label.config(text="状态：已停止")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            print("注入已停止")
        else:
            self.status_label.config(text="状态：未注入")
            print("注入未启动")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
