import glob2
import os
import tkinter as tk
from tkinter import *
from tkinter import Label, ttk, scrolledtext, StringVar
# import tkinter.messagebox as messagebox
import pyautogui
# import win32con
# import win32gui

from system_hotkey import SystemHotkey
from tiresomeRpa import Mylog, allwork, jumpout  # 调用主要程序

WorkPagePath = ''
redotimes = 0
windowNameTire = 'tiresoonneRpa'
windowclass = 'TkTopLevel'
theme = '本色'
LogMethod = '输出到文本框'
hk = SystemHotkey()
StartKey = []  # 列表
StopKey = []


def tobedeveloped():
    pyautogui.alert('该功能还未开发')


def themechange(themechoose):  # 感觉这个函数没啥意义，但是又不知道该怎么让点击的时候改变theme
    global theme
    theme = themechoose


def FThemePath(name):  # find theme path
    global theme
    themeimgPath = 'theme'
    if theme == '本色':
        themeimgPath = 'theme_base'

    name = str(name)
    path = '.\\theme\\' + themeimgPath + '\\' + name  # 找到某一张图的路径
    global img
    img = tk.PhotoImage(file=path)
    return img


#  @ 功能：拿到文件夹列表
#  @ 参数：[I] :p 当前要查看的目录
def getDirList(p):  # 其实这个还没看太懂
    p = str(p)
    if p == "":
        return []
    p = p.replace("/", "\\")
    if p[-1] != "\\":
        p = p + "\\"
    a = os.listdir(p)
    b = [x for x in a if os.path.isdir(p + x)]
    return b


TotalTaskList = getDirList('../Source')  #


def workbegin():  # 开始执行更新一下数据读入，引入主程序功能
    global WorkPagePath, redotimes, LogMethod
    # print(redotimes)  # 报错检查
    allwork(redotimes, WorkPagePath, LogMethod)


def finish_working():  # 快捷键调用
    exit(0)


def windowbegin():
    global WorkPagePath, redotimes, LogMethod
    window = tk.Tk()  # 创建根窗口
    window.title(windowNameTire)  # 窗口命名
    window.iconphoto(False, FThemePath('document.png'))  # 第一参数False 表示该图标图像仅适用于该特定窗口，而不适用于将来创建的 toplevels 窗口
    window.geometry("1400x600+200+200")  # 窗口大小和边距

    # -----------设置菜单栏-------------
    menu = tk.Menu(window)
    submenu_1 = tk.Menu(menu, tearoff=0)  # tearoff默认是否下拉
    submenu_1.add_command(label='账户', command=tobedeveloped)
    submenu_1.add_command(label='文件包位置', command=tobedeveloped)
    submenu_1.add_cascade(label='快捷键', command=tobedeveloped)
    menu.add_cascade(label='设置', menu=submenu_1)

    submenu_2 = tk.Menu(submenu_1)
    submenu_2.add_command(label='本色', command=themechange('本色'))  # command基本按键菜单
    submenu_2.add_command(label='其他颜色', command=tobedeveloped)
    menu.add_cascade(label='主题', menu=submenu_2)  # cascade 创建能继续下拉的菜单

    window.config(menu=menu)
    # ---------------------------------

    # global StartKey, StopKey
    # StartKey = ['control', 'b']  # 列表
    StopKey = ['control', 'q']  # quit
    # hk.register(StartKey, callback=workbegin)  # 设置开始进程快捷键
    hk.register(StopKey, callback=jumpout)  # 设置终止快捷键

    Frame_1 = tk.Frame(window, width=600, height=600, relief='groove', bd=1)  # 创建一个window的分区
    Frame_1.pack(side='left')  # 分区放在

    # ------------选择文件包-------------
    lbl_1 = Label(Frame_1, text="选择文件包", font=("Arial Bold", 10))  # 标签1 放在frm_package 中
    lbl_1.place(x=270, y=80, width=150)
    lbl_4 = Label(Frame_1, image=FThemePath('document.png'))
    lbl_4.place(x=250, y=70, width=50)
    Combobox_1 = tk.ttk.Combobox(Frame_1, values=TotalTaskList, width=24, height=30)  # 创建下拉菜单
    Combobox_1.place(x=180, y=150)
    # ---------------------------------

    # -----------打开文件包进行编辑------------
    def openpage():
        global WorkPagePath
        print(WorkPagePath)
        os.startfile(WorkPagePath)
    btn_2 = Button(Frame_1, text="打开文件包", bg="lightgreen", fg="black",
                   command=lambda: [renew(), openpage()])  # 按钮
    btn_2.place(x=100, y=400, width=120)
    # ---------------------------------

    # ------------选择执行次数-----------
    # r = tk.StringVar()  # 生成字符串变量
    # r.set('1')  # 初始化变量值
    rad1 = Radiobutton(Frame_1, text="执行", value=2)
    # rad1.variable = r,  # 单选按钮关联的变量
    # rad1.value = '1',  # 设置选中单选按钮时的变量值
    rad1.place(x=350, y=300, width=80)

    spin_1 = Spinbox(Frame_1, from_=1, to=100, width=5)  # spinbox
    spin_1.place(x=430, y=300)

    lbl_2 = Label(Frame_1, text="次", font=("Arial Bold", 10))  # 标签2
    lbl_2.place(x=500, y=300)

    rad2 = Radiobutton(Frame_1, text="无限执行", value=1)  # 单选框
    rad2.place(x=350, y=350, width=120)
    # ---------------------------------

    # -------------开始执行-------------
    def renew():
        global WorkPagePath, redotimes, LogMethod
        WorkPagePath = '..\\Source\\' + Combobox_1.get()
        redotimes = spin_1.get()
        if rad2 is True:
            redotimes = -1
        textinf('设置已读取')
        LogMethod = Combobox_2.get()
        if LogMethod == '输出到文本框':  # 更改print为向文本框输出
            sys.stdout = RDToWText(info_scltext)

    btn_2 = Button(Frame_1, text="开始执行", bg="orange", fg="black",
                   command=lambda: [renew(), workbegin()])  # 按钮
    btn_2.place(x=380, y=400, width=120)
    # ---------------------------------

    # ----------停止进程-----------
    btn_3 = Button(Frame_1, text="停止进程", bg="red", fg="black",
                   command=jumpout)  # 按钮
    btn_3.place(x=380, y=500, width=120)
    # ---------------------------------

    Frame_2 = tk.Frame(window, width=800, height=600, relief='groove', bd=1)  # 创建一个window的分区
    Frame_2.pack(side='right')  # 分区放在

    # -------显示文本框并将进程输出到文本框--------  # 会输出主要的进展，日志选择文本框则日志也会输出在这里
    info_scltext = scrolledtext.ScrolledText(Frame_2, relief="solid", width=60, height=13, bg="black")  # 创建滚动文本框
    info_scltext.place(x=50, y=100, width=700, height=400)  # 放置

    info_scltext.tag_add('tagwhite', '1.0', 'end')  # 申明一个tag,在1到end位置使用
    info_scltext.tag_config('tagwhite', foreground='white', font=("宋体", 14))  # 设置tag即插入文字的大小,颜色等

    def textinf(*Buf):
        for i in Buf:
            info_scltext.insert(END, i, 'tagwhite')  # end在末尾插入    INSERT 在光标处插入
            info_scltext.insert(END, '\n')
    # ---------------------------------

    # -------------日志记录选择-------------
    lbl_3 = Label(Frame_2, text="日志记录选择", font=("Arial Bold", 10))  # 标签1
    lbl_3.place(x=50, y=60, width=150)
    logmethodlist = ['输出到文本框', '输入到文件', 'Debug', '不记录']
    Combobox_2 = tk.ttk.Combobox(Frame_2, values=logmethodlist, width=24, height=30)
    Combobox_2.current(0)
    Combobox_2.place(x=250, y=60)

    def readlog():  # 将系统记录的日志文件输出到文本框
        filename = glob2.glob('..\\Software_files\\' + '\\*.txt')[0]
        with open(filename) as log:  # 没写完
            textinf(log.read())

    btn_3 = Button(Frame_2, text="查看日志文件", bg="lightgreen", fg="black", command=readlog)
    btn_3.place(x=550, y=500, width=170, height=50)
    # ---------------------------------

    window.mainloop()  # 窗口进入消息主循环，能随时接受命令


class RDToWText(object):  # 重定向输出类  为了能调用write所以弄个类装
    def __init__(self, text_widget):
        self.text_space = text_widget  # 将文本框其备份
        self.stdoutbak = sys.stdout  # 输出位置 备份
        self.stderrbak = sys.stderr  # 错误信息输出位置 备份

    def write(self, str):
        self.text_space.insert(END, str)  # 在最后一行插入
        self.text_space.see(END)  # 看最后一行
        self.text_space.update()

    def restoreStd(self):
        # 恢复标准输出
        sys.stdout = self.stdoutbak
        sys.stderr = self.stderrbak

    def flush(self):  # 关闭程序时会调用flush刷新缓冲区，没有该函数关闭时会报错
        pass


if __name__ == '__main__':
    windowbegin()
    # hk.unregister(StartKey)  # 取消快捷键 放在程序退出时
