"""
  <AUTO CCTV>
    Copyright (C) 2023-2024  李明锐

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import webbrowser
import schedule
import time
import pyautogui
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import win32com.client
import sys
import win32con
import win32gui

# 设置全局变量
url = "https://tv.cctv.com/live/index.shtml"
start_time = "19:00"  # 设置开始时间
end_time = "19:30"  # 设置结束时间
running = True

# 打开网页并双击全屏播放
def open_and_play():
    webbrowser.open(url)
    time.sleep(5)  # 等待页面加载

    # 判断窗口是否最大化，如果不是最大化则最大化窗口
    browser_hwnd = win32gui.GetForegroundWindow()  # 获取当前活动窗口句柄
    placement = win32gui.GetWindowPlacement(browser_hwnd)  # 获取窗口放置信息
    if placement[1] == win32con.SW_SHOWNORMAL:  # 判断窗口是否在正常位置
        win32gui.ShowWindow(browser_hwnd, win32con.SW_MAXIMIZE)  # 最大化窗口

    # 添加一些间隔
    time.sleep(0.5)

    # 屏幕中心双击
    screen_width, screen_height = pyautogui.size()
    center_x, center_y = screen_width / 2, screen_height / 2

    # 第一次点击
    pyautogui.click(center_x, center_y, button='left')

    # 添加一些间隔
    time.sleep(0.05)

    # 第二次点击，实现双击效果
    pyautogui.click(center_x, center_y, button='left')

# 关闭网页
def close_browser():
    pyautogui.hotkey('ctrl', 'w')  # 关闭当前标签页

# 定义定时任务
def job():
    open_and_play()
    schedule.every().day.at(end_time).do(close_browser)

# 开启定时任务
schedule.every().day.at(start_time).do(job)

# 创建GUI控制界面
def exit_program():
    root.destroy()

class App:
    start_hour_var = None
    start_minute_var = None
    end_hour_var = None
    end_minute_var = None

    def __init__(self, root):
        self.root = root
        root.title("AutoCCTV 新闻联播自动放映系统")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)  # 关闭窗口事件处理
        self.root.iconify()  # 最小化窗口

        # 设置窗口背景颜色为深蓝色
        self.root.configure(background='navy')

        # 设置主题样式为"clam"，并使用深蓝色
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure('.', background='navy', foreground='yellow')

        # 创建顶部菜单
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # 添加帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)

        # 添加说明按钮
        help_menu.add_command(label="说明", command=self.show_instructions)

        # 添加关于按钮
        help_menu.add_command(label="关于", command=self.show_about)

        # 添加一个Frame
        status_frame = ttk.Frame(root)
        status_frame.pack(pady=10)

        time_frame = ttk.Frame(root)
        time_frame.pack(anchor=tk.W)

        # 添加开始时间框架
        start_time_frame = ttk.Frame(time_frame)
        start_time_frame.grid(row=0, column=0, padx=2, pady=2)
        ttk.Label(start_time_frame, text="开始时间：").grid(row=0, column=0)
        App.start_hour_var = tk.StringVar()
        App.start_minute_var = tk.StringVar()
        App.start_hour_var.set(start_time.split(":")[0])
        App.start_minute_var.set(start_time.split(":")[1])
        self.start_hour_spinbox = ttk.Spinbox(start_time_frame, from_=0, to=23, textvariable=App.start_hour_var,
                                              width=2)
        self.start_hour_spinbox.grid(row=0, column=1)
        ttk.Label(start_time_frame, text="时").grid(row=0, column=2)
        self.start_minute_spinbox = ttk.Spinbox(start_time_frame, from_=0, to=59, textvariable=App.start_minute_var,
                                                width=2)
        self.start_minute_spinbox.grid(row=0, column=3)
        ttk.Label(start_time_frame, text="分").grid(row=0, column=4)

        # 添加结束时间框架
        end_time_frame = ttk.Frame(time_frame)
        end_time_frame.grid(row=1, column=0, padx=2, pady=2)
        ttk.Label(end_time_frame, text="  结束时间：").grid(row=0, column=0)
        App.end_hour_var = tk.StringVar()
        App.end_minute_var = tk.StringVar()
        App.end_hour_var.set(end_time.split(":")[0])
        App.end_minute_var.set(end_time.split(":")[1])
        self.end_hour_spinbox = ttk.Spinbox(end_time_frame, from_=0, to=23, textvariable=App.end_hour_var, width=2)
        self.end_hour_spinbox.grid(row=0, column=1)
        ttk.Label(end_time_frame, text="时").grid(row=0, column=2)
        self.end_minute_spinbox = ttk.Spinbox(end_time_frame, from_=0, to=59, textvariable=App.end_minute_var, width=2)
        self.end_minute_spinbox.grid(row=0, column=3)
        ttk.Label(end_time_frame, text="分").grid(row=0, column=4)

        update_time_button = ttk.Button(root, text="更新时间", command=self.update_time)
        update_time_button.pack(pady=5)

        # 添加开机自启动按钮
        startup_button = ttk.Button(root, text="添加开机自启动", command=self.create_startup_shortcut)
        startup_button.pack(pady=5)

        # 读取开始和结束时间
        self.load_time_from_file()

        exit_button = ttk.Button(root, text="退出", command=exit_program)
        exit_button.pack(pady=10)

        # 添加定时任务调度
        self.root.after(1000, self.check_schedule)

    def show_about(self):
        # 弹出关于对话框
        about_info = ("AutoCCTV 新闻联播自动放映系统\n版本号：v0.3\n时间：2024/3/10\n项目地址：https://github.com/Return-Log/AutoCCTV"
                      "\nCopyright © 2023-2024 李明锐. All Rights Reserved.\n本软件根据GNU通用公共许可证第3版（GPLv3）发布。")
        tk.messagebox.showinfo("关于", about_info)

    def show_instructions(self):
        # 弹出说明对话框
        instructions_text = """
        AutoCCTV说明\n·程序会在指定时间自动打开默认浏览器全屏播放新闻联播并定时关闭\n\nv0.3更新：1.增加自定义播放时间功能 2.自动将未处于屏幕中心的窗口最大化
        """
        self.show_text_dialog("使用说明", instructions_text)

    def show_text_dialog(self, title, text):
        # 弹出一个文本对话框
        text_dialog = tk.Toplevel(self.root)
        text_dialog.title(title)

        text_widget = tk.Text(text_dialog, wrap=tk.WORD, width=40, height=10)
        text_widget.insert(tk.END, text)
        text_widget.config(state=tk.DISABLED)  # 设置文本框为不可编辑
        text_widget.pack()

    def check_schedule(self):
        schedule.run_pending()
        self.root.after(1000, self.check_schedule)

    def add_startup_button(self):
        startup_button = ttk.Button(self.root, text="添加开机自启动", command=self.create_startup_shortcut)
        startup_button.pack(pady=10)

    def create_startup_shortcut(self):
        # 获取当前脚本的路径
        script_path = sys.argv[0]  # 获取封装后的exe文件的路径

        # 创建快捷方式
        startup_folder = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")
        shortcut_path = os.path.join(startup_folder, "AutoCCTV.lnk")

        # 如果快捷方式已存在，先删除
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)

        # 创建快捷方式
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = script_path
        shortcut.WorkingDirectory = os.path.dirname(script_path)
        shortcut.save()

        messagebox.showinfo("提示", "已添加开机自启动")

    def on_close(self):
        # 窗口关闭事件处理
        self.root.iconify()  # 最小化窗口

    def update_time(self):
        global start_time, end_time

        # 获取小时和分钟的值
        start_hour = self.start_hour_var.get().zfill(2)  # 补零，确保是两位数
        start_minute = self.start_minute_var.get().zfill(2)  # 补零，确保是两位数
        end_hour = self.end_hour_var.get().zfill(2)  # 补零，确保是两位数
        end_minute = self.end_minute_var.get().zfill(2)  # 补零，确保是两位数

        # 格式化时间
        start_time = f"{start_hour}:{start_minute}"
        end_time = f"{end_hour}:{end_minute}"

        messagebox.showinfo("提示", "时间已更新")

        # 将新的开始和结束时间保存到文件中
        self.save_time_to_file()

        # 重新计划定时任务
        self.reschedule_job()

    def reschedule_job(self):
        # 清除之前的定时任务
        schedule.clear()

        # 定义新的定时任务
        def job():
            open_and_play()
            schedule.every().day.at(end_time).do(close_browser)

        # 开启定时任务
        schedule.every().day.at(start_time).do(job)

    def load_time_from_file(self):
        filename = "time_config.txt"
        if os.path.exists(filename):
            with open(filename, 'r') as f:
                lines = f.readlines()
                for line in lines:
                    key, value = line.strip().split('=')
                    if key == "start_time":
                        start_time = value
                        self.start_hour_var.set(start_time.split(":")[0])
                        self.start_minute_var.set(start_time.split(":")[1])
                    elif key == "end_time":
                        end_time = value
                        self.end_hour_var.set(end_time.split(":")[0])
                        self.end_minute_var.set(end_time.split(":")[1])

    def save_time_to_file(self):
        filename = "time_config.txt"
        with open(filename, 'w') as f:
            f.write(f"start_time={start_time}\n")
            f.write(f"end_time={end_time}\n")


# 运行GUI
# 获取当前脚本的目录
script_directory = os.path.dirname(sys.argv[0])
icon_path = os.path.join(script_directory, 'icon.ico')

root = tk.Tk()

# 设置图标
if os.path.exists(icon_path):
    root.iconbitmap(icon_path)
else:
    print("图标文件不存在：", icon_path)

app = App(root)
root.mainloop()
