import webbrowser
import schedule
import time
import pyautogui
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox


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
        about_info = "AutoCCTV 新闻联播自动放映系统\n版本号：v0.1\n作者：李明锐\n时间：2023/12/30\n项目地址：https://github.com/Return-Log/AutoCCTV"
        tk.messagebox.showinfo("关于", about_info)

    def show_instructions(self):
        # 弹出说明对话框
        instructions_text = """
        AutoCCTV说明\n·程序运行时会在18：59自动打开默认浏览器全屏播放新闻联播并在19：30自动关闭。\n·详细文档及添加自启动方法： https://github.com/Return-Log/AutoCCTV/blob/master/readme.md
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


# 运行GUI
root = tk.Tk()
app = App(root)
root.mainloop()
