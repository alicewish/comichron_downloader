import codecs
import csv
import datetime
import os
import random
import re
import threading
import time
from pathlib import Path

import requests
import wx
from dateutil.relativedelta import relativedelta

p_data_body = re.compile(r'<!-- CC TOP COMICS Module -->(.*?)<!-- END CC TOP COMICS Module -->', re.I | re.DOTALL)

p_table = re.compile(r'<table .*?>(.*?)</table>')
p_a = re.compile(r'<a .*?>(.*?)</a>')

p_attri = re.compile(r'</?(strong|thead|tbody)>')

p_td = re.compile(r'</td> *<td>|</th> *<th>')

p_redun = re.compile(r'</?(td|th|tr)>')

comichron_url = 'www.comichron.com/monthlycomicssales'


# ================创建目录================
def make_dir(file_path):
    if not os.path.exists(file_path):
        try:
            os.mkdir(file_path)
        except:
            pass


# ================保存 CSV================
def save_csv(csv_file_path, headers, rows):
    if rows:
        with codecs.open(csv_file_path, 'w', 'utf_8_sig') as f:
            f_csv = csv.writer(f)
            if headers:
                f_csv.writerow(headers)
            f_csv.writerows(rows)


# ================运行时间计时================
def run_time(start_time):
    '''
    :param start_time:
    :return: 运行时间
    '''
    run_time = time.time() - start_time
    if run_time < 60:  # 两位小数的秒
        show_run_time = '{:.2f}秒'.format(run_time)
    elif run_time < 3600:  # 分秒取整
        show_run_time = '{:.0f}分{:.0f}秒'.format(run_time // 60, run_time % 60)
    else:  # 时分秒取整
        show_run_time = '{:.0f}时{:.0f}分{:.0f}秒'.format(run_time // 3600, run_time % 3600 // 60, run_time % 60)
    return show_run_time


# ================确定可用日期================
def get_dates(start_day):
    result = []

    today = datetime.date.today()
    end_day = today - relativedelta(months=1)

    current = start_day

    while current <= end_day:
        result.append(current)
        current += relativedelta(months=1)

    return result


class HelloFrame(wx.Frame):

    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(HelloFrame, self).__init__(*args, **kw)

        # self.SetBackgroundColour(wx.Colour(224, 224, 224))

        mm = wx.DisplaySize()
        self.SetSize(0.5 * mm[0], 0.5 * mm[1])

        self.Center()

        self.SetToolTip(wx.ToolTip('这是一个框架！'))
        # self.SetCursor(wx.StockCursor(wx.CURSOR_MAGNIFIER))  # 改变鼠标样式

        self.createUI()

    def createUI(self):
        self.pnl = wx.Panel(self)

        # ================框架================
        self.button1 = wx.Button(self.pnl, wx.ID_ANY, '执行任务')

        self.st1 = wx.StaticText(self.pnl, label='下载自：')
        self.tc1 = wx.TextCtrl(self.pnl, wx.ID_ANY, value=comichron_url, style=wx.TE_READONLY)

        self.st2 = wx.StaticText(self.pnl, label='保存到文件夹：')
        self.tc2 = wx.TextCtrl(self.pnl, wx.ID_ANY, value=str(info_dir))

        self.st3 = wx.StaticText(self.pnl, label='调试信息：')
        self.tc3 = wx.TextCtrl(self.pnl, style=wx.TE_READONLY)

        self.st4 = wx.StaticText(self.pnl, label='调试日志：')
        self.tc4 = wx.TextCtrl(self.pnl, wx.ID_ANY, style=wx.TE_MULTILINE)

        # ================尺寸器================
        # sBox = wx.BoxSizer()  # 水平尺寸器，不带参数则为默认的水平尺寸器
        self.vBox = wx.BoxSizer(wx.VERTICAL)  # 垂直尺寸器

        # 给尺寸器添加组件，从左往右，从上到下
        # sBox.Add(pathText, proportion=3, flag=wx.EXPAND | wx.ALL, border=5)
        # sBox.Add(but1, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        # sBox.Add(but2, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        self.vBox.Add(self.button1, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        self.vBox.Add(self.st1, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        self.vBox.Add(self.tc1, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        self.vBox.Add(self.st2, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        self.vBox.Add(self.tc2, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        self.vBox.Add(self.st3, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        self.vBox.Add(self.tc3, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)

        self.vBox.Add(self.st4, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        self.vBox.Add(self.tc4, proportion=4, flag=wx.EXPAND | wx.ALL, border=5)

        # 设置主尺寸
        self.pnl.SetSizer(self.vBox)  # 因为sBox被嵌套在vBox上，所以以vBox为主尺寸

        # ================绑定================
        self.button1.Bind(wx.EVT_BUTTON, self.onButton)

        # ================状态栏================
        self.CreateStatusBar()
        self.SetStatusText('准备就绪')

        # ================菜单栏================
        fileMenu = wx.Menu()  # 文件菜单

        helloItem = fileMenu.Append(-1, '你好\tCtrl-H', '程序帮助')
        fileMenu.AppendSeparator()
        exitItem = fileMenu.Append(wx.ID_EXIT, '退出\tCtrl-Q', '退出程序')

        helpMenu = wx.Menu()  # 帮助菜单

        aboutItem = helpMenu.Append(wx.ID_ABOUT, '关于\tCtrl-G', '关于程序')

        menuBar = wx.MenuBar()  # 菜单栏
        menuBar.Append(fileMenu, '文件')
        menuBar.Append(helpMenu, '其他')

        self.SetMenuBar(menuBar)

        # ================绑定================
        self.Bind(wx.EVT_MENU, self.OnHello, helloItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

        # # ================工具栏================
        # self.toolbar = self.CreateToolBar()
        #
        # save_ico = wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE, wx.ART_TOOLBAR, (16, 16))
        # saveTool = self.toolbar.AddSimpleTool(wx.ID_ANY, save_ico, '保存', '开始保存')
        #
        # self.toolbar.AddSeparator()
        #
        # print_ico = wx.ArtProvider.GetBitmap(wx.ART_PRINT, wx.ART_TOOLBAR, (16, 16))
        # self.printTool = self.toolbar.AddSimpleTool(wx.ID_ANY, print_ico, '打印', '开始打印')
        #
        # delete_ico = wx.ArtProvider.GetBitmap(wx.ART_DELETE, wx.ART_TOOLBAR, (16, 16))
        # self.deleteTool = self.toolbar.AddSimpleTool(wx.ID_ANY, delete_ico, '删除', '开始删除')
        #
        # self.toolbar.AddSeparator()
        #
        # undo_ico = wx.ArtProvider.GetBitmap(wx.ART_UNDO, wx.ART_TOOLBAR, (16, 16))
        # self.undoTool = self.toolbar.AddSimpleTool(wx.ID_UNDO, undo_ico, '撤销', '开始撤销')
        # self.toolbar.EnableTool(wx.ID_UNDO, False)
        #
        # redo_ico = wx.ArtProvider.GetBitmap(wx.ART_REDO, wx.ART_TOOLBAR, (16, 16))
        # self.redoTool = self.toolbar.AddSimpleTool(wx.ID_REDO, redo_ico, '重做', '开始重做')
        # self.toolbar.EnableTool(wx.ID_REDO, False)
        #
        # # ================绑定================
        # self.Bind(wx.EVT_MENU, self.onSave, saveTool)
        # self.Bind(wx.EVT_MENU, self.onPrint, self.printTool)
        # self.Bind(wx.EVT_MENU, self.onDelete, self.deleteTool)
        # self.Bind(wx.EVT_TOOL, self.onUndo, self.undoTool)
        # self.Bind(wx.EVT_TOOL, self.onRedo, self.redoTool)
        #
        # self.toolbar.Realize()  # 准备显示

    # ================获取网页信息================
    def process_page(self, page_url, csv_file):
        start_time = time.time()  # 初始时间戳

        page = requests.get(url=page_url, headers=mac_header)
        page_text = page.text
        ulist = []

        m_data_body = re.search(p_data_body, page_text)
        if m_data_body:
            data_body = m_data_body.group(1)
            data_body = data_body.replace('\r', '').replace('\n', '')
            data_body = re.sub(p_table, r'\1', data_body)
            data_body = re.sub(p_a, r'\1', data_body)
            data_body = re.sub(p_attri, r'', data_body)

            entries = data_body.split('</tr><tr>')

            for entry in entries:
                rows = re.split(p_td, entry)
                rows = [re.sub(p_redun, '', row) for row in rows]
                ulist.append(rows)

            headers = ulist[0]
            rows = ulist[1:]
            if rows:
                save_csv(csv_file, headers, rows)
                # ================运行时间计时================
                show_run_time = run_time(start_time)
                label_str = '保存成功：' + str(csv_file) + ' ' + show_run_time
                self.show_label_str(label_str)

    def download_comichron(self, dates, info_dir):
        start_time = time.time()  # 初始时间戳

        entry_str = self.tc2.GetValue()
        new_info_dir = Path(entry_str)
        if new_info_dir.exists():
            info_dir = new_info_dir

        make_dir(info_dir)

        if info_dir.exists():
            for i in range(len(dates)):
                date = dates[i]
                year = date.year
                month = date.month

                page_url = prefix + str(year) + '/' + str(year) + '-' + str(month).zfill(2) + '.html'

                csv_file = info_dir / (Path(page_url).stem + '.csv')

                if csv_file.exists():
                    label_str = '已存在：' + str(csv_file)
                    self.show_label_str(label_str)
                else:
                    label_str = '正在下载：' + page_url
                    self.show_label_str(label_str)
                    try:
                        self.process_page(page_url, csv_file)
                        time.sleep(random.randint(300, 500) / 1000)
                    except:
                        pass
                # time.sleep(random.randint(50, 100) / 1000)

        # ================运行时间计时================
        show_run_time = run_time(start_time)
        label_str = '程序结束！' + show_run_time
        self.show_label_str(label_str)

    def onButton(self, event):
        self.thread_it(self.download_comichron, dates, info_dir)
        event.GetEventObject().Disable()
        # event.GetEventObject().Enable()

    def OnExit(self, event):
        self.Close(True)

    def OnHello(self, event):
        wx.MessageBox('来自 wxPython', '你好')

    # def onSave(self, event):
    #     wx.MessageBox('保存自 wxPython')
    #
    # def onPrint(self, event):
    #     wx.MessageBox('打印自 wxPython')
    #
    # def onDelete(self, event):
    #     wx.MessageBox('删除自 wxPython')
    #
    # def onUndo(self, event):
    #     wx.MessageBox('撤销自 wxPython')
    #
    # def onRedo(self, event):
    #     wx.MessageBox('重做自 wxPython')

    def OnAbout(self, event):
        wx.MessageBox(message=about_me,
                      caption='关于' + app_name,
                      style=wx.OK | wx.ICON_INFORMATION)

    def show_label_str(self, label_str):
        # print(label_str)

        wx.CallAfter(lambda: self.tc3.Clear())
        wx.CallAfter(lambda: self.tc3.AppendText(label_str))

        if not label_str.endswith('\n\r'):
            label_str += '\n'

        wx.CallAfter(lambda: self.tc4.AppendText(label_str))

    @staticmethod
    def thread_it(func, *args):
        t = threading.Thread(target=func, args=args)
        t.setDaemon(True)  # 守护--就算主界面关闭，线程也会留守后台运行（不对!）
        t.start()  # 启动
        # t.join()          # 阻塞--会卡死界面！


if __name__ == '__main__':
    mac_header = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
    }

    prefix = 'http://www.comichron.com/monthlycomicssales/'

    current_dir = os.path.dirname(os.path.abspath(__file__))
    current_dir = Path(current_dir)

    info_dir = current_dir / 'comichron'

    start_day = datetime.date(1996, 9, 1)
    dates = get_dates(start_day)

    app_name = 'comichron历史销量下载器 v1.0 by 墨问非名'
    about_me = '这是下载comichron历史销量至csv表格的软件。'

    app = wx.App()

    frm = HelloFrame(None, title=app_name)
    frm.Show()

    app.MainLoop()
