import re  # split
# from pyautogui import FAILSAFE, PAUSE
import pyautogui
import time
import glob2
import win32api
import win32con
import xlrd
import pyperclip
import os
# import datetime as dt


pyautogui.FAILSAFE = True  # 保护措施，避免失控`鼠标移到左上角产生异常而中断
pyautogui.PAUSE = 0  # 默认最小操作周期

log_file = glob2.glob('..\\Software_files\\' + '\\*.txt')[0]
LogMethod = 2
sheet1 = None
nrows = 2
WorkPath = ''
run = 1


def jumpout():
    global run
    Mylog('停止进程')
    run = 0


def Mylog(*Buf):  # *是让Buf接受任意多参数并放在一个元组中
    global LogMethod
    if LogMethod == '输入到文件':  # 输出到文件tirelog
        with open(log_file, 'a') as log:  # a 表示追加
            # print(dt.datetime.now().strftime('%F %T:%f'), end=' ')  # 那么长一嘟噜
            for i in Buf:
                print(i, file=log, end=' ')
            print('', file=log)
    elif LogMethod == '输出到文本框':  # 输出到文本框/pycharm
        # print(dt.datetime.now().strftime('%F %T:%f'), end=' ')
        for i in Buf:
            print(i, end=' ')
        print(end='\n')
    elif LogMethod == 'Debug':
        # print(dt.datetime.now().strftime('%F %T:%f'), end=' ')
        for i in Buf:
            print(i, end=' ')
        print(end='\n')
    elif LogMethod == '不记录':
        pass


class Cmd(object):  # 定义一个类，用来装一行指令，检查，工作
    global run

    def __init__(self, cmd1, cmd2, cmd3, cmd4, cmd5, cmd6, cmd7, cmd8, cmd9, cmd10, nowrow):
        self.permission = cmd1  # 注意这里都是包含类型.ctype和内容.value的
        self.pic_name = cmd2
        self.cmd_detail = cmd3
        self.out_time_act = cmd4
        self.out_time = cmd5
        self.redo = cmd6
        self.notes = cmd7
        self.confidence = cmd8
        self.interval = cmd9
        self.move_duration = cmd10
        self.now_row = nowrow + 1  # 标定是第几行（在这里把系统计数改为同excel的计数），在默认图片序号初有用到

        if cmd1.ctype == 0:  # ctype类型，0为空
            self.permission.value = 1
        if cmd2.ctype == 0:
            self.pic_name.value = str(self.now_row) + '.png'
        if cmd3.ctype == 0:
            self.cmd_detail.value = 'left=1'
        if cmd4.ctype == 0:
            self.out_time_act.value = -1
        if cmd5.ctype == 0:
            self.out_time.value = -1
        if cmd5.ctype == 0 and cmd4.ctype != 0:
            self.out_time.value = '次数=5'
        if cmd6.ctype == 0:
            self.redo.value = 1
        if cmd8.ctype == 0:
            self.confidence.value = 0.9
        if cmd9.ctype == 0:
            self.interval.value = 0
        if cmd10.ctype == 0:
            self.move_duration.value = 0.01  # 还没检测就先赋默认值了，是不是有点问题

        self.CmdName = []  # 声明两个字符串，用来读取命令名和要求
        self.CmdDetail = []
        self.CmdAmount = 0
        # 对字符串赋值，感觉直接在类里弄方便点
        if '=' in str(self.cmd_detail.value):
            sourceStr = str(self.cmd_detail.value)
            replaceStr = sourceStr.replace('，', ',')
            replaceStr = replaceStr.replace(',', '=')
            SplitStr = re.split('=', replaceStr)
            i = 0
            count = 0
            while count < len(SplitStr):
                self.CmdName.append(SplitStr[count])
                self.CmdDetail.append(SplitStr[count + 1])
                count += 2
                i += 1
            self.CmdAmount = i
        else:
            self.CmdName.append(str(self.cmd_detail.value))

    # ctype  空0 字符串1 数字2 日期3 布尔4 error5
    def rowdataCheck(self):  # 调用时在类中，仅检查一行的数据
        check = 1
        if self.permission.value != 1 and self.permission.value != 0:
            Mylog('第', self.now_row, '行，第1列执行开关不为0或1')
            check = 0
        if self.pic_name.ctype != 0 and self.pic_name.ctype != 1:
            Mylog('第', self.now_row, '行，第2列图片名不为空且不为字符串')
            check = 0
        if self.cmd_detail.ctype != 0:  # 第三列，命令名称和细节不为空的话,开始判断
            sourceStr = str(self.cmd_detail.value)  # ？为什么要加str
            replaceStr = sourceStr.replace('，', ',')
            if replaceStr.count('=') != replaceStr.count(',') + 1:
                Mylog('第', self.now_row, '行，第3列命令细节=不等于，')
                check = 0
        if self.out_time_act.ctype != 0 and self.out_time_act.ctype != 1:
            Mylog('第', self.now_row, '行，第4列超时行为不为空且不为字符串')
            check = 0
        if self.out_time.ctype != 0 and self.out_time.ctype != 1 and self.out_time.ctype != 2 \
                or self.out_time.ctype == 2 and self.out_time.ctype != -1:
            Mylog('第', self.now_row, '行，第5列小于0且不等于-1')
            check = 0
        if self.redo.value < 0 and self.redo.value != -1:
            Mylog('第', self.now_row, '行，第6列重做次数小于0且不为-1')
            check = 0
        # 第七列为备注，不检查
        if self.confidence.ctype != 0 and 2 or self.confidence.value < 0 \
                or self.confidence.value > 1:
            Mylog('第', self.now_row, '行，第8列置信度不为0且不为数字或不在0到1范围内')
            check = 0
        if self.interval.ctype != 0 and 2 or self.interval.value < 0:
            Mylog('第', self.now_row, '行，第9列找图间隔不为0且不为数字或小于0')
            check = 0
        if self.move_duration.ctype != 0 and self.move_duration.ctype != 2 \
                or self.move_duration.value < 0:
            Mylog('第', self.now_row, '行，第10列鼠标移动时间不为0且不为数字,或小于0')
            check = 0
        return check

    def work(self):  # 调用时在类中，仅运行一行的数据
        icmd = 0  # 为了让第三列事件和细节顺序执行
        while icmd < self.CmdAmount and run == 1:
            Mylog('开始执行', self.CmdName[icmd], self.CmdDetail[icmd])
            if self.CmdName[icmd] == '移动':
                if self.CmdDetail[icmd] == '0':  # 找图
                    if self.redoFAMPic() == 0:  # 如果没找到
                        Mylog('没找到，查看超时行为')
                        return self.OtAct()  # 调用并return超时行为，跳出了work
                else:  # 最上下左右和绝对移动
                    width, height = pyautogui.size()  # 屏幕的宽度和高度
                    if self.CmdDetail[icmd] == '最上':
                        pyautogui.moveTo(width / 2, 0, duration=0.25)
                    elif self.CmdDetail[icmd] == '最下':
                        pyautogui.moveTo(width / 2, height, duration=0.25)
                    elif self.CmdDetail[icmd] == '最左':
                        pyautogui.moveTo(0, height / 2, duration=0.25)
                    elif self.CmdDetail[icmd] == '最右':
                        pyautogui.moveTo(width, height / 2, duration=0.25)
                    else:
                        # moved = True  # ?
                        Split = re.split('/', self.CmdDetail[icmd])
                        Mylog('移动鼠标到', Split)
                        pyautogui.moveTo(int(Split[0]), int(Split[1]), 0)
            elif self.CmdName[icmd] == 'left' or self.CmdName[icmd] == 'right' or \
                    self.CmdName[icmd] == 'middle':  # 指数由self统一输入了
                print('找', self.pic_name, self.CmdName[icmd], self.CmdDetail)
                if self.redoFAMPic() == 0:  # 如果没找到
                    Mylog('没找到，查看超时行为')
                    return self.OtAct()  # 调用并return超时行为，跳出了work
                time.sleep(0.5)  # 为什么会还没移动到就点击了？
                pyautogui.click(clicks=int(self.CmdDetail[icmd]), interval=self.interval.value,
                                duration=self.move_duration.value,
                                button=self.CmdName[icmd])  # 这咋会读到输入啊
            elif self.CmdName[icmd] == '滚动':
                pyautogui.scroll(int(self.CmdDetail[icmd]))
            elif self.CmdName[icmd] == '偏移':
                # offseted = True  # ？不能理解它的必要性
                Split = re.split('/', self.CmdDetail[icmd])
                Mylog('鼠标相对移动', Split)
                pyautogui.moveRel(xOffset=int(Split[0]), yOffset=int(Split[1]))
                # tween = pyautogui.linear  里面的参数，不知道什么用
            elif self.CmdName[icmd] == '相对拖拽':
                Split = re.split('/', self.CmdDetail[icmd])
                Mylog('相对拖拽', Split)
                pyautogui.dragRel(xOffset=int(Split[0]), yOffset=int(Split[1]), duration=0.11,
                                  button='left', mouseDownUp=True)  # mouseDownUp是个啥
            elif self.CmdName[icmd] == '绝对拖拽':
                Split = re.split('/', self.CmdDetail[icmd])
                Mylog('绝对拖拽', Split)
                pyautogui.dragTo(x=int(Split[0]), y=int(Split[1]),
                                 duration=3, button='left')  # 这里duration不能过快
            elif self.CmdName[icmd] == '热键':  # 大羽版本方案
                newinput = self.CmdDetail[icmd].split('+')  # +分割热键读入
                pyautogui.hotkey(*tuple(newinput))  # 调用热键
            elif self.CmdName[icmd] == '按键':
                pyautogui.press(str(self.CmdDetail[icmd]))
            elif self.CmdName[icmd] == '按下':
                if self.CmdDetail[icmd] == '左键':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                elif self.CmdDetail[icmd] == '右键':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                else:
                    pyautogui.keyDown(str(self.CmdDetail[icmd]))
            elif self.CmdName[icmd] == '松开':
                if self.CmdDetail[icmd] == '左键':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                elif self.CmdDetail[icmd] == '右键':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                else:
                    pyautogui.keyUp(str(self.CmdDetail[icmd]))
            elif self.CmdName[icmd] == '输入':  # 不知道为啥复制粘贴好像有问题
                pyautogui.typewrite(self.CmdDetail[icmd], interval=0.25)
            elif self.CmdName[icmd] == '输入中文':  # 但是中文键入很麻烦，如果能粘贴进去就粘贴吧
                strtemp = pyperclip.paste()
                # mylog("上次剪切板内容：", strtemp)
                pyperclip.copy(str(self.CmdDetail[icmd]))
                time.sleep(0.6)
                pyautogui.hotkey('ctrl', 'v')
                Mylog('输入', str(self.CmdDetail[icmd]))
                time.sleep(0.5)
                # mylog("恢复上次剪切板内容")
                pyperclip.copy(strtemp)
                # pyautogui.typewrite(str(NowRowValue[local]), interval=0.1)
            elif self.CmdName[icmd] == '命令行':
                wincmd = self.CmdDetail[icmd]
                os.system(wincmd)
                Mylog('运行系统指令', wincmd)
            elif self.CmdName[icmd] == '等待' or '延时':
                time.sleep(float(self.CmdDetail[icmd]))
            elif self.CmdName[icmd] == '截屏':  # 用以在某些反应过快无法截屏时使用
                if os.path.exists('Screenshot') is not True:
                    os.mkdir('Screenshot')
                ShotImgpath = 'Screenshot/Shot_' + f'{time.strftime("%m%d%H%M%S ")}.png'
                pyautogui.screenshot().save(ShotImgpath)  # 保存为ShotImgpath
            elif self.CmdName[icmd] == '弹窗' or '提示':
                pyautogui.alert(text=self.CmdName[icmd], title='tiresonne Rpa')
            elif self.CmdName[icmd] == '跳往':
                return self.CmdDetail[icmd]
            elif self.CmdName[icmd] == '结束':
                return '退出'
            elif self.CmdName[icmd] == '':
                pass
            Mylog('完成', self.CmdName[icmd])
            icmd += 1
            time.sleep(0.2)  # 每执行一行的命令暂停0.2s

    def FAMPic(self):  # find and move to picture找图片，找到则移动到位置并返回1
        ImgPath = (WorkPath + '\\' + self.pic_name.value)
        if self.pic_name.value != '' and os.path.exists(ImgPath) is True:
            Mylog(ImgPath, '图片有效')
            location = pyautogui.locateCenterOnScreen(ImgPath,
                                                      confidence=self.confidence.value)  # 找到图片坐标
            if location is not None:
                pyautogui.moveTo(location, duration=self.move_duration.value)
                Mylog('找到图片', self.pic_name.value, '并完成移动')
                return 1
            return 0  # if不成立，即没找到，返回0
        else:
            Mylog(ImgPath, '图片无效')

    def redoFAMPic(self):  # 如果在限次内未找到，返回0.
        if self.out_time.value != -1 and run == 1:
            Outtype = []
            sourceStr = str(self.out_time.value)
            replaceStr = sourceStr
            SplitStr = re.split('=', replaceStr)
            Outtype.append(SplitStr[0])
            Outvalue = int(SplitStr[1])
            Mylog('开始找图', self.pic_name.value)
            if Outtype[0] == '次数':  # Outtype 是个列表，就算只有一个元素也是列表
                i = 0
                while i < Outvalue and run == 1:
                    key1 = self.FAMPic()  # 运行函数同时key取一个返回值
                    if key1 == 1:  # 如果返回值为1，func成功，则跳出
                        print('调用的函数成功')
                        return 1
                    i += 1
                    print('找图', i, '/', Outvalue, '次')
                    time.sleep(0.2)
                return 0
            if Outtype[0] == '时间':
                BeginTime = time.time()
                while run:
                    if self.FAMPic() == 1:
                        Mylog('调用的函数成功')
                        return 1
                    else:
                        if time.time() - BeginTime > Outvalue:
                            Mylog('已超时，调用的函数未完成')
                            return 0
        else:  # 不限尝试直到成功
            i = 1
            while run:
                key1 = self.FAMPic()  # 运行函数同时key取一个返回值
                if key1 == 1:  # 如果返回值为1，func成功，则跳出
                    Mylog('调用的函数成功')
                    return 1  # 如果要开启跳出，func需要有返回值，当成功时返回1.此外一直重复没有其他返回
                i += 1
                Mylog('运行', i, '次')
                time.sleep(0.2)

    def OtAct(self):  # out time act  查看并返回超时行为
        ActName = []
        sourceStr = str(self.out_time_act.value)
        SplitStr = re.split('=', sourceStr)
        ActName = SplitStr[0]  # ActName用append的话 也是列表不是字符串，不加序号这里判断全部失效
        if ActName == '跳往':
            Gotoline = int(SplitStr[1])
            Mylog('跳到', Gotoline, '行')
            return Gotoline  # Gotoline 是个列表,刚开始的时候
        else:
            if ActName == '弹窗':
                pyautogui.alert(text=self.pic_name.value + '查找超时', title='tiresonneRpa')
            if ActName == '退出':
                Mylog('中途退出')
                return '退出'
        Mylog('跳过，执行下一行')
        return -1  # 剩下是个跳过，弹窗和跳过都返回-1，在下一个判断都是跳过


def Readcmd():
    global sheet1, nrows, WorkPath
    filename = glob2.glob(WorkPath + '\\*.xls')[0]
    wb = xlrd.open_workbook(filename)  # 读文件
    sheet1 = wb.sheet_by_index(0)  # 读文件中的第一表
    nrows = sheet1.nrows
    commands = []  # 把每一行cmd存在commands中,列表
    NowRow = 1  # 从第二行开始读，第一行是表头读进去会报错  # 之前怎么没报错？
    while NowRow < sheet1.nrows and run == 1:  # 为了保证读全，nrows+1  # 不用加，nrows比nowrow大1，正好等于的时候停止了
        # print('读入了', NowRow, '行')  # 报错检查
        cmd = Cmd(sheet1.row(NowRow)[0], sheet1.row(NowRow)[1], sheet1.row(NowRow)[2],
                  sheet1.row(NowRow)[3], sheet1.row(NowRow)[4], sheet1.row(NowRow)[5],
                  sheet1.row(NowRow)[6], sheet1.row(NowRow)[7], sheet1.row(NowRow)[8],
                  sheet1.row(NowRow)[9], NowRow)
        commands.append(cmd)
        NowRow += 1
    return commands


def AllRowRun(commands):  # 遍历一次表  # n个commands，n+1行表格，nrows=n，i从0开始，共n个，最后是n-1
    global run
    i = 0
    while i < nrows-1:  # 其实这里commands[0]是表格里的第二行，commands行数比表格计数少1，nrows=sheet1.nrows比表格计行多一
        if commands[i].redo.value < 0:  # 先看是否执行redo 如果为-1则无限执行
            while run:
                key = commands[i].work()
                if key is not None:  # int 没有ctype属性
                    if type(key) == str and key == '退出':  # ctype  空0 字符串1 数字2 日期3 布尔4 error5
                        Mylog('已退出')
                        i = nrows
                        break
                    elif type(key) == int and key != -1:  # 只会是数字了
                        i = key - 3  # 输入是表格第n行，计算机第n-1行，执行完本次循环i会+1，所以再-1，还有个1不知道在哪\\找到了
                        break
        else:  # redo大于0，
            j = 0
            while j < commands[i].redo.value and run == 1:  # 执行redo次，默认为1
                key = commands[i].work()
                if key is not None:  # int 没有ctype属性  #为啥会没执行啊
                    if type(key) == str and key == '退出':  # ctype  空0 字符串1 数字2 日期3 布尔4 error5
                        Mylog('已退出')
                        i = nrows
                        break
                    elif type(key) == int and key != -1:  # 只会是数字了
                        i = key - 3
                        break
                j += 1
        i += 1
    if run == 1:
        Mylog('完全执行一次表')


def allwork(redotimes, workpath, logMtd):  # log_method
    global LogMethod, WorkPath, run
    run = 1
    WorkPath = workpath
    LogMethod =logMtd
    week = time.localtime().tm_wday + 1
    Mylog('今天是周', week)  # 以后可以加个周功能

    commands = Readcmd()  # 读取表中的命令
    # 检查数据
    check = 1
    if nrows < 1:
        print('没数据')
        check = 0
    key = check  # 加个判断,key正好定义和赋值一起了
    for i, cmd in enumerate(commands):  # 遍历command中的cmd，检查命令
        check = cmd.rowdataCheck()  # dataCheck发现错误返回0
        if check == 0:
            key = check  # 如果返回了0，key就为0，下一步不执行。返回1继续执行
        if run == 0:  # 增加停止功能
            break
    if key != 0:  # 检查通过，选择功能
        while run:
            # if type(redotimes) == str and redotimes == 'c':
            #     os.system('cls')
            if int(redotimes) == 1:
                AllRowRun(commands)
                time.sleep(0.1)
                break
            elif int(redotimes) < 1:  # 无限循环
                count = 0
                while True:
                    count += 1
                    Mylog("列表正在执行第", count, "次")
                    AllRowRun(commands)
                    time.sleep(0.1)
                    Mylog("已经完成第", count, "次", "列表")
            else:
                count = 0
                while count < redotimes:
                    count += 1
                    Mylog("列表正在执行第", count, "次")
                    AllRowRun(commands)
                    Mylog("已经完成第", count, "次", "列表")
                    time.sleep(0.1)
                    Mylog("等待0.1秒")
                Mylog('操作已全部完成。')
                break
