#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
@Project ：main.py 
@File    ：main.py
@IDE     ：PyCharm 
@Author  ：Trespassing
@Date    ：2025/5/4 21:08 
"""
from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
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
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"\u66f2Star\u516c\u6587\u5904\u7406demo", None))
        self.action1.setText(QCoreApplication.translate("MainWindow", u"\u9047\u5230\u95ee\u9898?", None))
        self.action2.setText(QCoreApplication.translate("MainWindow", u"\u5173\u4e8e\u672c\u9879\u76ee", None))
        self.action3.setText(QCoreApplication.translate("MainWindow", u"\u9879\u76ee\u76f4\u8fbe", None))
        self.action4.setText(QCoreApplication.translate("MainWindow", u"\u8054\u7cfb\u4f5c\u8005", None))
        self.groupBox.setTitle(QCoreApplication.translate("MainWindow", u"\u66f2Star\u516c\u6587\u5904\u7406demo", None))
        self.bt_dl.setText(QCoreApplication.translate("MainWindow", u"\u5f00\u59cb\u5904\u7406", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"\u8f93\u5165\u6587\u4ef6\u8def\u5f84:", None))
        self.bt1.setText(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u6587\u4ef6", None))
        self.bt2.setText(QCoreApplication.translate("MainWindow", u"\u9009\u62e9\u6587\u4ef6\u5939", None))
        self.process.setText("")
        self.menuAbout.setTitle(QCoreApplication.translate("MainWindow", u"\u66f4\u591a", None))
    # retranslateUi
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui=Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.bt1.clicked.connect(self.select_file)
        self.ui.bt2.clicked.connect(self.select_folder)
        self.ui.bt_dl.clicked.connect(self.start_to_del)
        self.path=""
        self.ui.action1.triggered.connect(lambda:pyautogui.alert(text="您可以通过qq:2669807502联系作者或者到github提交issue\n项目地址:https://github.com/S-Trespassing/easy_QFNU_document\n期待您的宝贵意见~~~",title="遇到问题?"))
        self.ui.action2.triggered.connect(lambda:pyautogui.alert(text="本项目的初衷是一键处理繁琐的公文格式,您可以对项目对项目进行各种二创~\n项目地址:https://github.com/S-Trespassing/easy_QFNU_document\n",title="关于项目"))
        self.ui.action3.triggered.connect(lambda:webbrowser.open_new_tab("https://github.com/S-Trespassing/easy_QFNU_document"))
        self.ui.action4.triggered.connect(lambda:webbrowser.open_new_tab("https://qm.qq.com/cgi-bin/qm/qr?k=qIaENQvu9-jWTr5x2NUC2s-jTupUokdk"))
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
        cnt=0
        for file in file_list:
            self.change_format(file)
            self.show_process(cnt,len(file_list))
            cnt+=1
        self.show_process(cnt,len(file_list))
        path = Path(self.path)
        if len(file_list)>1:
            dir_name = 'output_' + str(path.name)
        else:
            dir_name = 'output_' + str(path.parent.name)
        opt_path = Path(os.getcwd()) / Path(dir_name)
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
            print("文件不存在")
            return
        # 防止路径不是绝对路径
        path = path.resolve()
        if Path.is_file(path):
            if path.suffix in ['.doc','.docx']:
                file_list.append(str(path))
                print("文件类型正确")
                return [str(path)]
            else:
                print("文件类型错误")
                return
        elif Path.is_dir(path):
            file_list.extend([str(file) for file in path.glob('*.doc')])
            file_list.extend([str(file) for file in path.glob('*.docx')])
            return file_list
    def change_format(self,path):
        try:
            app.Visible = False
            app.DisplayAlerts = False
            doc = app.Documents.Open(path)
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
            path = Path(path)
            dir_name = 'output_' + str(path.parent.name)
            os.makedirs(dir_name, exist_ok=True)
            file_name = str(path.name)
            opt_path = Path(os.getcwd()) / Path(dir_name) / Path(file_name)
            doc.SaveAs(FileName=str(opt_path))
            doc.Close()
            return
        except:
            pyautogui.alert(text="很遗憾您的文档未被正确处理,出现错误的原因大概率是某个word文件已被占用",title="提示")
if __name__=='__main__':
    try:
        app=client.Dispatch('Word.Application')
    except:
        pyautogui.alert(title="提示", text="您的电脑貌似并没有安装microsoft word...")
        exit(0)
    appui = QApplication([])
    mainw = MainWindow()
    mainw.show()
    appui.exec()
    app.Quit()
