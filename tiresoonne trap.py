# http://c.biancheng.net/tkinter/menu-widget.html  菜单栏设置
# https://zhuanlan.zhihu.com/p/92017216  tkinter 基础
# https://pypi.org/project/PyAutoGUI/  下载库
# https://summer.blog.csdn.net/article/details/84650938  # pyautogui
# https://blog.csdn.net/zhuan_long/article/details/120953194  # 对窗口的控制
# https://www.cnblogs.com/turtle-fly/p/3280519.html  # 重定向print    __redirection__
'''
hwnd = win32gui.FindWindow(lpClassName=None, lpWindowName=None)  # 查找窗口，不找子窗口，返回值为0表示未找到窗口
hwnd = win32gui.FindWindowEx(hwndParent=0, hwndChildAfter=0, lpszClass=None, lpszWindow=None)  # 查找子窗口，返回值为0表示未找到子窗口
win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
SW_HIDE：隐藏窗口并激活其他窗口。nCmdShow=0。
SW_SHOWNORMAL：激活并显示一个窗口。如果窗口被最小化或最大化，系统将其恢复到原来的尺寸和大小。应用程序在第一次显示窗口的时候应该指定此标志。nCmdShow=1。
SW_SHOWMINIMIZED：激活窗口并将其最小化。nCmdShow=2。
SW_SHOWMAXIMIZED：激活窗口并将其最大化。nCmdShow=3。
SW_SHOWNOACTIVATE：以窗口最近一次的大小和状态显示窗口。激活窗口仍然维持激活状态。nCmdShow=4。
SW_SHOW：在窗口原来的位置以原来的尺寸激活和显示窗口。nCmdShow=5。
SW_MINIMIZE：最小化指定的窗口并且激活在Z序中的下一个顶层窗口。nCmdShow=6。
SW_SHOWMINNOACTIVE：窗口最小化，激活窗口仍然维持激活状态。nCmdShow=7。
SW_SHOWNA：以窗口原来的状态显示窗口。激活窗口仍然维持激活状态。nCmdShow=8。
SW_RESTORE：激活并显示窗口。如果窗口最小化或最大化，则系统将窗口恢复到原来的尺寸和位置。在恢复最小化窗口时，应用程序应该指定这个标志。nCmdShow=9。

'''

import sys

sys.stdout.write('hello'+'\n')
print('hello')  # 等价


class __redirection__(object):

    def __init__(self):
        self.buff = ''  # 输出缓存
        self.stdoutbak = sys.stdout  # 输出路径备份
        self.stderrbak = sys.stderr  # 错误信息输出位置 备份

    def write(self, output_stream):  # 输出读取
        self.buff += output_stream  # buff 等于 buff+output_stream

    def to_console(self):  # 定向到工作台并将buff输出
        sys.stdout = self.stdoutbak  # 又把备份路径还回去了
        print(self.buff)

    def to_file(self, file_path):
        with open(file_path, 'w') as f:
            sys.stdout = f  # f本身有write 的功能
            print(self.buff)

    def to_window(self, text_widget):
        text_space = text_widget
        # 但是窗口文本框并没有write
        print(self.buff)

    def flush(self):
        self.buff = ''  # 释放buff

    def reset(self):
        sys.stdout = self.stdoutbak


if __name__ == "__main__":
    # redirection
    r_obj = __redirection__()
    sys.stdout = r_obj

    # get output stream
    print('hello')
    print('there')

    # redirect to console
    r_obj.to_console()

    # redirect to file
    # r_obj.to_file('out.log')

    # flush buffer
    r_obj.flush()

    # reset
    r_obj.reset()

# class StdoutRedirector(object):  # 重定向输出类  这部分copy来的没看太懂，为什么非要弄个class？为了有个write吗
#     def __init__(self, text_widget):
#         self.text_space = text_widget  # 将其备份
#         self.stdoutbak = sys.stdout
#         self.stderrbak = sys.stderr
#
#     def write(self, str):
#         self.text_space.insert(END, str)
#         self.text_space.insert(END, '\n')
#         self.text_space.see(END)
#         self.text_space.update()
#
#     def restoreStd(self):
#         # 恢复标准输出
#         sys.stdout = self.stdoutbak
#         sys.stderr = self.stderrbak
#
#     def flush(self):  # 关闭程序时会调用flush刷新缓冲区，没有该函数关闭时会报错
#         pass


# # -*- coding: utf-8 -*-
#
#
# import tkinter as tk  # 使用Tkinter前需要先导入
#
# # 第1步，实例化object，建立窗口window
# window = tk.Tk ( )
#
# # 第2步，给窗口的可视化起名字
# window.title ( 'My Window' )
#
# # 第3步，设定窗口的大小(长 * 宽)
# window.geometry ( '500x300' )  # 这里的乘是小x
#
# # 第4步，在图形界面上创建一个标签用以显示内容并放置
# l = tk.Label ( window , text = '      ' , bg = 'green' )
# l.pack ( )
#
# # 第10步，定义一个函数功能，用来代表菜单选项的功能，这里为了操作简单，定义的功能比较简单
# counter = 0
#
#
# def do_job () :
#     global counter
#     l.config ( text = 'do ' + str ( counter ) )
#     counter += 1
#
#
# # 第5步，创建一个菜单栏，这里我们可以把他理解成一个容器，在窗口的上方
# menubar = tk.Menu ( window )
#
# # 第6步，创建一个File菜单项（默认不下拉，下拉内容包括New，Open，Save，Exit功能项）
# filemenu = tk.Menu ( menubar , tearoff = 0 )
# # 将上面定义的空菜单命名为File，放在菜单栏中，就是装入那个容器中
# menubar.add_cascade ( label = 'File' , menu = filemenu )
#
# # 在File中加入New、Open、Save等小菜单，即我们平时看到的下拉菜单，每一个小菜单对应命令操作。
# filemenu.add_command ( label = 'New' , command = do_job )
# filemenu.add_command ( label = 'Open' , command = do_job )
# filemenu.add_command ( label = 'Save' , command = do_job )
# filemenu.add_separator ( )  # 添加一条分隔线
# filemenu.add_command ( label = 'Exit' , command = window.quit )  # 用tkinter里面自带的quit()函数
#
# # 第7步，创建一个Edit菜单项（默认不下拉，下拉内容包括Cut，Copy，Paste功能项）
# editmenu = tk.Menu ( menubar , tearoff = 0 )
# # 将上面定义的空菜单命名为 Edit，放在菜单栏中，就是装入那个容器中
# menubar.add_cascade ( label = 'Edit' , menu = editmenu )
#
# # 同样的在 Edit 中加入Cut、Copy、Paste等小命令功能单元，如果点击这些单元, 就会触发do_job的功能
# editmenu.add_command ( label = 'Cut' , command = do_job )
# editmenu.add_command ( label = 'Copy' , command = do_job )
# editmenu.add_command ( label = 'Paste' , command = do_job )
#
# # 第8步，创建第二级菜单，即菜单项里面的菜单
# submenu = tk.Menu ( filemenu )  # 和上面定义菜单一样，不过此处实在File上创建一个空的菜单
# filemenu.add_cascade ( label = 'Import' , menu = submenu , underline = 0 )  # 给放入的菜单submenu命名为Import
#
# # 第9步，创建第三级菜单命令，即菜单项里面的菜单项里面的菜单命令（有点拗口，笑~~~）
# submenu.add_command ( label = 'Submenu_1' , command = do_job )  # 这里和上面创建原理也一样，在Import菜单项中加入一个小菜单命令Submenu_1
#
# # 第11步，创建菜单栏完成后，配置让菜单栏menubar显示出来
# window.config ( menu = menubar )
#
# # 第12步，主窗口循环显示
# window.mainloop ( )

# 版权声明：本文为CSDN博主「他的固执，我的幼稚」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
# 原文链接：https://blog.csdn.net/m0_51915132/article/details/113496574

# def windowctrl(WC, WN, action):
#     hwnd = win32gui.FindWindow(WC, WN)
#     print(hwnd)
#     if hwnd == 0:
#         if action == 0:  # 窗口最小化
#             print('窗口最小化')
#             if win32gui.IsIconic(hwnd) is not True:
#                 win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)  #
#         if action == 1:
#             if win32gui.IsIconic(hwnd):  # 窗口恢复原来大小
#                 win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
#
# class a(object):
#     def __init__(self):
#         print('这是一个类')
#
#     def bbk(self):
#         print('this is bbk')
#
#
# b = a()
