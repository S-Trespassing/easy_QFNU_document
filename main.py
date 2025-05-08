#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：main.py 
@File    ：main.py
@IDE     ：PyCharm
@Author  ：Trespassing
@Date    ：2025/5/4 21:08
"""
import pythoncom
from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
                            QMetaObject, QObject, QPoint, QRect,
                            QSize, QTime, QUrl, Qt, QThread, Signal, QThreadPool, QRunnable)
from PySide6.QtGui import (QAction, QBrush, QColor, QConicalGradient,
    QCursor, QFont, QFontDatabase, QGradient,
    QIcon, QImage, QKeySequence, QLinearGradient,
    QPainter, QPalette, QPixmap, QRadialGradient,
    QTransform)
from PySide6.QtWidgets import (QApplication, QGroupBox, QHBoxLayout, QLabel,
                               QLineEdit, QMainWindow, QMenu, QMenuBar,
                               QPushButton, QSizePolicy, QStatusBar, QWidget, QFileDialog)
from win32com import client
import re
from pathlib import Path
import os
import pyautogui
import webbrowser
first_level = [
    '一、', '二、', '三、', '四、', '五、',
    '六、', '七、', '八、', '九、', '十、',
    '十一、', '十二、', '十三、', '十四、', '十五、',
    '十六、', '十七、', '十八、', '十九、', '二十、'
]
second_level=[
    '（一）', '（二）', '（三）', '（四）', '（五）',
    '（六）', '（七）', '（八）', '（九）', '（十）',
    '（十一）', '（十二）', '（十三）', '（十四）', '（十五）',
    '（十六）', '（十七）', '（十八）', '（十九）', '（二十）'
]
# 处理信号
class WorkerSignals(QObject):
    finished = Signal(str)
class  Worker(QRunnable):
    # finished_signal = Signal(str)
    def __init__(self,path):
        super().__init__()
        self.finished_signal=WorkerSignals()
        self.path = path
        # print(f"获得路径{path}")
    def run(self):
        try:
            pythoncom.CoInitialize()
            app=client.Dispatch('Word.Application')
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Documents.Open(self.path)
            # 设置纸张
            for section in doc.Sections:
                section.PageSetup.PageWidth = 595.3  # 210mm ≈ 595.3 磅
                section.PageSetup.PageHeight = 841.9  # 297mm ≈ 841.9 磅
                section.PageSetup.TopMargin = 3.7 * 28.35  # 上边距 3.7cm
                section.PageSetup.BottomMargin = 3.2 * 28.35  # 下边距 3.2cm
                section.PageSetup.LeftMargin = 2.7 * 28.35  # 左边距 2.7cm
                section.PageSetup.RightMargin = 2.7 * 28.35  # 右边距 2.7cm
            # 单独处理标题
            doc.Paragraphs(1).Range.Font.Name = "方正小宋简体"
            doc.Paragraphs(1).Range.Font.Size = "35"
            doc.Paragraphs(1).Format.Alignment = 1
            doc.Paragraphs(1).LineSpacingRule = 4
            doc.Paragraphs(1).Format.LineSpacing = 35
            cnt = 0
            for paragraph in doc.Paragraphs:
                if int(cnt) == 0:
                    cnt = 1
                    continue
                if re.match(r'^\s*[一二三四五六七八九十]\s*、.*', paragraph.Range.Text):
                    paragraph.Range.Font.Name = '黑体'
                elif re.match(r'^\s*\([一二三四五六七八九十]\).*', paragraph.Range.Text):
                    paragraph.Range.Font.Name = '楷体'
                elif paragraph.Range.ListFormat.ListType == 3:
                    if paragraph.Range.ListFormat.ListString in first_level:
                        paragraph.Range.Font.Name = '黑体'
                        # print(paragraph.Range.ListFormat.ListString)
                        # print(paragraph.Range.Text)
                    elif paragraph.Range.ListFormat.ListString in second_level:
                        paragraph.Range.Font.Name = '楷体'
                        # print("这是二级标题")
                    else:
                        paragraph.Range.Font.Name = "仿宋"
                else:
                    paragraph.Range.Font.Name = "仿宋"
                paragraph.Range.Font.Size = "16"
                paragraph.Format.Alignment = 3
                paragraph.LineSpacingRule = 4
                paragraph.Format.LineSpacing = 29
                paragraph.Format.CharacterUnitFirstLineIndent = 2
            path = Path(self.path)
            dir_name = 'output_' + str(path.parent.name)
            os.makedirs(dir_name, exist_ok=True)
            file_name = str(path.name)
            opt_path = Path(os.getcwd()) / Path(dir_name) / Path(file_name)
            doc.SaveAs(FileName=str(opt_path))
            doc.Close()
            app.Quit()
            # 释放COM资源
            pythoncom.CoUninitialize()
        except:
            pyautogui.alert(text="出错了",title="异常")
        #子线程每完成一个任务便发送一次完成信号
        self.finished_signal.finished.emit("done")
        return

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(640, 400)
        MainWindow.setMinimumSize(QSize(640, 400))
        MainWindow.setMaximumSize(QSize(640, 400))
        icon = QIcon()
        icon.addFile(u"../../../OIP-C.png", QSize(), QIcon.Mode.Normal, QIcon.State.Off)
        MainWindow.setWindowIcon(icon)
        self.action1 = QAction(MainWindow)
        self.action1.setObjectName(u"action1")
        self.action2 = QAction(MainWindow)
        self.action2.setObjectName(u"action2")
        self.action3 = QAction(MainWindow)
        self.action3.setObjectName(u"action3")
        self.action4 = QAction(MainWindow)
        self.action4.setObjectName(u"action4")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setObjectName(u"groupBox")
        self.groupBox.setGeometry(QRect(40, 20, 551, 301))
        self.bt_dl = QPushButton(self.groupBox)
        self.bt_dl.setObjectName(u"bt_dl")
        self.bt_dl.setGeometry(QRect(230, 220, 91, 31))
        self.label = QLabel(self.groupBox)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(40, 100, 101, 16))
        self.layoutWidget = QWidget(self.groupBox)
        self.layoutWidget.setObjectName(u"layoutWidget")
        self.layoutWidget.setGeometry(QRect(41, 133, 491, 26))
        self.horizontalLayout = QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.notice = QLineEdit(self.layoutWidget)
        self.notice.setObjectName(u"notice")
        self.horizontalLayout.addWidget(self.notice)
        self.bt1 = QPushButton(self.layoutWidget)
        self.bt1.setObjectName(u"bt1")
        self.horizontalLayout.addWidget(self.bt1)
        self.bt2 = QPushButton(self.layoutWidget)
        self.bt2.setObjectName(u"bt2")
        self.horizontalLayout.addWidget(self.bt2)
        self.process = QLabel(self.groupBox)
        self.process.setObjectName(u"process")
        self.process.setGeometry(QRect(40, 190, 481, 20))
        self.process.setAlignment(Qt.AlignmentFlag.AlignCenter)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menuBar = QMenuBar(MainWindow)
        self.menuBar.setObjectName(u"menuBar")
        self.menuBar.setGeometry(QRect(0, 0, 640, 33))
        self.menuAbout = QMenu(self.menuBar)
        self.menuAbout.setObjectName(u"menuAbout")
        MainWindow.setMenuBar(self.menuBar)
        self.menuBar.addAction(self.menuAbout.menuAction())
        self.menuAbout.addAction(self.action1)
        self.menuAbout.addAction(self.action2)
        self.menuAbout.addAction(self.action3)
        self.menuAbout.addAction(self.action4)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", "曲Star公文处理demo", None))
        self.action1.setText(QCoreApplication.translate("MainWindow", "遇到问题?", None))
        self.action2.setText(QCoreApplication.translate("MainWindow", "关于本项目", None))
        self.action3.setText(QCoreApplication.translate("MainWindow", "项目直达", None))
        self.action4.setText(QCoreApplication.translate("MainWindow", "联系作者", None))
        self.groupBox.setTitle(QCoreApplication.translate("MainWindow", "曲Star公文处理demo", None))
        self.bt_dl.setText(QCoreApplication.translate("MainWindow", "开始处理", None))
        self.label.setText(QCoreApplication.translate("MainWindow", "文件路径:", None))
        self.bt1.setText(QCoreApplication.translate("MainWindow", "选择文件", None))
        self.bt2.setText(QCoreApplication.translate("MainWindow", "选择文件夹", None))
        self.process.setText("")
        self.menuAbout.setTitle(QCoreApplication.translate("MainWindow", "更多", None))
    # retranslateUi
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.len_file_list = None
        self.remaining_tasks = None
        self.ui=Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.bt1.clicked.connect(self.select_file)
        self.ui.bt2.clicked.connect(self.select_folder)
        self.ui.bt_dl.clicked.connect(self.start_to_del)
        self.path=None
        self.ui.action1.triggered.connect(lambda:pyautogui.alert(text="您可以通过qq:2669807502联系作者或者到github提交issue\n项目地址:https://github.com/S-Trespassing/easy_QFNU_document\n期待您的宝贵意见~~~",title="遇到问题?"))
        self.ui.action2.triggered.connect(lambda:pyautogui.alert(text="本项目的初衷是一键处理繁琐的公文格式,您可以对项目对项目进行各种二创~\n项目地址:https://github.com/S-Trespassing/easy_QFNU_document\n",title="关于项目"))
        self.ui.action3.triggered.connect(lambda:webbrowser.open_new_tab("https://github.com/S-Trespassing/easy_QFNU_document"))
        self.ui.action4.triggered.connect(lambda:webbrowser.open_new_tab("https://qm.qq.com/cgi-bin/qm/qr?k=qIaENQvu9-jWTr5x2NUC2s-jTupUokdk"))
        # 创建线程池
        self.thread_pool=QThreadPool()
        #最大线程数
        self.thread_pool.setMaxThreadCount(8)
    def show_process(self,i,all):
        i=int(i)
        all=int(all)
        con= f"{('*'*int(i*30/all)).ljust(30,' ')} {i}/{all}"
        self.ui.process.setText(con)
    def start_to_del(self):
        file_list=self.get_file(self.path)
        if  not file_list:
            pyautogui.alert(text="未找到有效的doc/docx文件...",title="提示")
            return
        #禁用按钮
        self.ui.bt_dl.setEnabled(False)
        self.ui.process.setText("正在处理,请稍等...")
        self.remaining_tasks=len(file_list)
        self.len_file_list=self.remaining_tasks
        for file in file_list:
            worker=Worker(file)
            worker.finished_signal.finished.connect(self.on_task_finished)
            self.thread_pool.start(worker)
        # self.thread_pool.waitForDone()
        # print("走到下一环节")


    def on_task_finished(self, status):
        self.remaining_tasks-=1
        if self.remaining_tasks==0:
            # 恢复按钮
            self.ui.bt_dl.setEnabled(True)
            path = Path(self.path)
            if self.len_file_list > 1:
                dir_name = 'output_' + str(path.name)
            else:
                dir_name = 'output_' + str(path.parent.name)
            opt_path = Path(os.getcwd()) / Path(dir_name)
            self.ui.process.setText("处理完毕(✪ω✪)~")
            pyautogui.alert(text=f"处理完毕,保存在{opt_path}", title="提示")
            os.startfile(opt_path.resolve())
    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="选择文件",
            dir=""
        )
        if path:
            self.ui.notice.setText(path)
            self.path=path
    def select_folder(self):
        path = QFileDialog.getExistingDirectory(
            parent=self,
            caption="选择文件夹",
            dir=""
        )
        if path:
            self.ui.notice.setText(path)
            self.path=path

    def get_file(self,path):
        path = Path(path)
        file_list = []
        if not Path.exists(path):
            # print("文件不存在")
            return
        # 防止路径不是绝对路径
        path = path.resolve()
        if Path.is_file(path):
            if path.suffix in ['.doc','.docx']:
                file_list.append(str(path))
                # print("文件类型正确")
                return [str(path)]
            else:
                # print("文件类型错误")
                return
        elif Path.is_dir(path):
            file_list.extend([str(file) for file in path.glob('*.doc')])
            file_list.extend([str(file) for file in path.glob('*.docx')])
            return file_list

if __name__=='__main__':
    try:
        _=client.Dispatch('Word.Application')
        _.Quit()
    except:
        pyautogui.alert(title="提示", text="您的电脑貌似并没有安装microsoft word...")
        exit(0)
    appui = QApplication([])
    mainw = MainWindow()
    mainw.show()
    appui.exec()