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
import sys  # 导入sys模块

# 设置全局变量
url = "https://tv.cctv.com/live/index.shtml"
start_time = "18:59"  # 设置开始时间
end_time = "19:30"  # 设置结束时间
running = True

# 打开网页并双击全屏播放
def open_and_play():
    webbrowser.open(url)
    time.sleep(5)  # 等待页面加载

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

        # 添加开机自启动按钮
        self.add_startup_button()

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

        # 添加一个Label来表示运行状态
        self.status_label = ttk.Label(status_frame, text="●", foreground="green", font=("Arial", 16, "bold"))
        self.status_label.grid(row=0, column=0)

        # 添加一个Label来表示"程序运行中"
        self.status_text_label = ttk.Label(status_frame, text="程序运行中", font=("Arial", 12))
        self.status_text_label.grid(row=0, column=1, padx=5)

        exit_button = ttk.Button(root, text="退出", command=exit_program)
        exit_button.pack(pady=10)

        # 添加定时任务调度
        self.root.after(1000, self.check_schedule)
        # 开始绿点的闪烁
        self.blink_green_dot()

    def toggle_running(self):
        global running
        running = not running
        if running:
            schedule.resume()
            self.status_text_label.config(text="程序运行中")  # 显示"程序运行中"
        else:
            schedule.pause()
            self.status_text_label.config(text="")  # 隐藏"程序运行中"

    def blink_green_dot(self):
        # 闪烁绿点
        current_state = self.status_label.cget("text")
        if current_state == "":
            self.status_label.config(text="●")
        else:
            self.status_label.config(text="")
        # 递归调用自身，以达到不断闪烁效果
        self.root.after(500, self.blink_green_dot)

    def show_about(self):
        # 弹出关于对话框
        about_info = "AutoCCTV 新闻联播自动放映系统\n版本号：v0.2\n时间：2024/3/3\n项目地址：https://github.com/Return-Log/AutoCCTV\nCopyright © 2023-2024 李明锐. All Rights Reserved.\n本软件根据GNU通用公共许可证第3版（GPLv3）发布。"
        tk.messagebox.showinfo("关于", about_info)

    def show_instructions(self):
        # 弹出说明对话框
        instructions_text = """
        AutoCCTV说明\n·程序运行时会在18：59自动打开默认浏览器全屏播放新闻联播并在19：30自动关闭。\n\nv0.2更新：1.添加开机自启动 2.运行时自动缩小至任务栏 3.更改主题颜色
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

# 运行GUI
root = tk.Tk()
root.iconbitmap('icon.ico')  # 设置自定义图标
app = App(root)
root.mainloop()