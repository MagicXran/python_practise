import socket
import sys
import time

import xlwings as xw
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QMainWindow, QApplication

'''常量'''
COL_NAME = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
            'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']

# HOSTNAME = "TrLoveXran"
HOSTNAME = "Karma"


# 1.重写一个类
class Threads(QThread):
    def __init__(self, *argv, **kwargs):
        super().__init__(*argv, **kwargs)

    # 4.创建信号
    pic_thread = pyqtSignal()

    # 2.设置休眠时间
    def run(self):
        # time.sleep(2)
        self.pic_thread.emit()  # 5.接受信号，并发送信号


class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()
        self.source_path = ''  # 选择文件的路径
        self.deal_action()

    def setupUi(self):
        self.resize(890, 532)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(8)
        font.setBold(False)
        font.setWeight(50)
        font.setStrikeOut(False)
        self.setFont(font)
        self.setLocale(QtCore.QLocale(QtCore.QLocale.Chinese, QtCore.QLocale.China))
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.log_text = QtWidgets.QTextEdit(self.centralwidget)
        self.log_text.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.log_text.setLocale(QtCore.QLocale(QtCore.QLocale.Chinese, QtCore.QLocale.China))
        self.log_text.setReadOnly(True)
        self.log_text.setObjectName("log_text")
        self.gridLayout_3.addWidget(self.log_text, 0, 0, 1, 2)
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox.sizePolicy().hasHeightForWidth())
        self.groupBox.setSizePolicy(sizePolicy)
        self.groupBox.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.groupBox.setObjectName("groupBox")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_2.sizePolicy().hasHeightForWidth())
        self.label_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.gridLayout_2.addWidget(self.label_2, 0, 0, 1, 1)
        self.classes_num = QtWidgets.QLineEdit(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.classes_num.sizePolicy().hasHeightForWidth())
        self.classes_num.setSizePolicy(sizePolicy)
        self.classes_num.setObjectName("classes_num")
        self.gridLayout_2.addWidget(self.classes_num, 1, 0, 1, 1)
        self.label_6 = QtWidgets.QLabel(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_6.sizePolicy().hasHeightForWidth())
        self.label_6.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_6.setFont(font)
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.gridLayout_2.addWidget(self.label_6, 4, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.gridLayout_2.addWidget(self.label_3, 2, 0, 1, 1)
        self.day_num = QtWidgets.QLineEdit(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.day_num.sizePolicy().hasHeightForWidth())
        self.day_num.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(8)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.day_num.setFont(font)
        self.day_num.setObjectName("day_num")
        self.gridLayout_2.addWidget(self.day_num, 3, 0, 1, 1)
        self.lesson_num = QtWidgets.QLineEdit(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lesson_num.sizePolicy().hasHeightForWidth())
        self.lesson_num.setSizePolicy(sizePolicy)
        self.lesson_num.setObjectName("lesson_num")
        self.gridLayout_2.addWidget(self.lesson_num, 5, 0, 1, 1)
        self.gridLayout_3.addWidget(self.groupBox, 1, 0, 1, 1)
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_2.sizePolicy().hasHeightForWidth())
        self.groupBox_2.setSizePolicy(sizePolicy)
        self.groupBox_2.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.groupBox_2.setObjectName("groupBox_2")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox_2)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("幼圆")
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(75)
        font.setStrikeOut(False)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.sel_path_btn = QtWidgets.QPushButton(self.groupBox_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sel_path_btn.sizePolicy().hasHeightForWidth())
        self.sel_path_btn.setSizePolicy(sizePolicy)
        self.sel_path_btn.setObjectName("sel_path_btn")
        self.gridLayout.addWidget(self.sel_path_btn, 1, 0, 1, 1)
        self.export_btn = QtWidgets.QPushButton(self.groupBox_2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.export_btn.sizePolicy().hasHeightForWidth())
        self.export_btn.setSizePolicy(sizePolicy)
        self.export_btn.setObjectName("export_btn")
        self.gridLayout.addWidget(self.export_btn, 2, 0, 1, 1)
        self.gridLayout_3.addWidget(self.groupBox_2, 1, 1, 1, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 890, 17))
        self.menubar.setObjectName("menubar")
        self.about = QtWidgets.QMenu(self.menubar)
        self.about.setObjectName("about")
        self.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        self.menubar.addAction(self.about.menuAction())

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "课表导出软件 with my darling 小鱼"))
        self.log_text.setHtml(_translate("MainWindow",
                                         "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                         "<html><head><meta name=\"qrichtext\" content=\"1\" /><style cope_type=\"text/css\">\n"
                                         "p, li { white-space: pre-wrap; }\n"
                                         "</style></head><body style=\" font-family:\'幼圆\',\'幼圆\'; font-size:8pt; font-weight:100; font-style:normal;\">\n"
                                         "<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'SimSun\'; font-size:14pt; font-weight:600;\">报警日志:</span></p>\n"
                                         "<p align=\"left\" style=\"-qt-paragraph-cope_type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'SimSun\'; font-size:8pt; font-weight:600;\"><br /></p></body></html>"))
        self.log_text.setPlaceholderText(_translate("MainWindow", "0"))
        self.groupBox.setTitle(_translate("MainWindow", "默认输入"))
        self.label_2.setText(_translate("MainWindow", "最大班级数"))
        self.classes_num.setPlaceholderText(_translate("MainWindow", "默认做多8个班级"))
        self.label_6.setText(_translate("MainWindow", "每天最多课时数"))
        self.label_3.setText(_translate("MainWindow", "每周工作几天"))
        self.day_num.setPlaceholderText(_translate("MainWindow", "默认每周工作5天"))
        self.lesson_num.setPlaceholderText(_translate("MainWindow", "默认每天8节课"))
        self.groupBox_2.setTitle(_translate("MainWindow", "处理"))
        self.label.setText(_translate("MainWindow", "选择要处理的表的路径"))
        self.sel_path_btn.setText(_translate("MainWindow", "选择表路径"))
        self.export_btn.setText(_translate("MainWindow", "拆分表结构"))
        self.about.setTitle(_translate("MainWindow", "关于"))

    def deal_action(self):
        self.sel_path_btn.clicked.connect(self.select_source_file)
        self.export_btn.clicked.connect(self.export_action)

    def show_message(self):
        QtWidgets.QMessageBox.warning(self, "严重错误", "当前计算机并非小鱼所有, 此软件无法正常运行~",
                                      QtWidgets.QMessageBox.Cancel)

    def LOG(self, text):
        """Write console output to text widget."""
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text + '\n')
        self.log_text.setTextCursor(cursor)
        self.log_text.ensureCursorVisible()

    def export_action(self):
        hostname = socket.gethostname()
        if hostname != HOSTNAME:
            self.show_message()
            time.sleep(3)
            sys.exit(app.exec_())

        if self.source_path:
            sybus = Syllabus(self, self.source_path)
            sybus.done()
        else:
            self.LOG("ERROR: 请先选择待处理文件.............\n")

    def select_source_file(self):
        hostname = socket.gethostname()
        if hostname == HOSTNAME:
            self.LOG("请选择文件:")
            directory = QtWidgets.QFileDialog.getOpenFileName(None, "选取文件", "./",
                                                              "All Files (*);;Text Files (*.txt)")
            self.log_text.moveCursor(QTextCursor.End)
            self.LOG(directory[0])
            self.source_path = directory[0]
            self.LOG("请稍等约3分钟 ")
            self.LOG("等待完成提示即可")
        else:
            self.LOG("当前计算机并非小鱼所有, 此软件无法正常运行~")
            self.show_message()
            time.sleep(3)
            sys.exit(app.exec_())


class Syllabus(object):
    def __init__(self, ui: Ui_MainWindow, source_path_=r"C:\\init.xlsx"):
        # if ui.classes_num.text():
        #     MAX_CLASSES_NUM = int(ui.classes_num.text())
        #
        # if ui.day_num.text():
        #     MAX_DAY_NUM = int(ui.day_num.text())
        #
        # if ui.lesson_num.text():
        #     MAX_LESSON_NUM = int(ui.lesson_num.text())

        self.col_nums = None
        self.row_nums = None
        self.app = xw.App(visible=False, add_book=False)
        # self.wb: xw.Book = self.app.books.open(ui.source_path)
        self.wb: xw.Book = self.app.books.open(source_path_)
        self.sheet: xw.Sheet = self.wb.sheets[0]
        self.weeks = dict()
        self.teacher_names = list()

    def __del__(self):
        self.wb.close()
        self.app.quit()
        print("执行析构")

    def sel_teacher_table(self):
        """提前教师课程表"""

        MAX_DAY_NUM = 5  # 一周上几天学
        MAX_LESSON_NUM = 8  # 每天最大课时数
        WEEK_LESSION_BEGIN = "A3"
        WEEK_LESSION_END = "AP50"
        WEEK_LESSION = "{}:{}".format(WEEK_LESSION_BEGIN, WEEK_LESSION_END)
        print('范围:{}\n'.format(WEEK_LESSION))

        scope = self.sheet.range(WEEK_LESSION)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))

        row_datas_vec = []
        '''每个教师一周上课班级的集合'''

        for i in range(int(WEEK_LESSION_BEGIN[1:]), int(WEEK_LESSION_END[2:]) + 1):
            row_datas = list()
            # 所有列
            for j in range(2, self.col_nums):
                if str(self.sheet.range(COL_NAME[j] + str(i)).value).strip():
                    row_datas.append(str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() + '班')
                else:
                    row_datas.append(str(self.sheet.range(COL_NAME[j] + str(i)).value).strip())  # print(row_datas)

            self.teacher_names.append(str(self.sheet.range(COL_NAME[0] + str(i)).value).strip())
            row_datas_vec.append(row_datas)

        #
        print(row_datas_vec)
        #
        print(self.teacher_names)

        # 教师人数

        print("教师总人数:{}".format(len(self.teacher_names)))

        teas_course = list()

        cycle = MAX_LESSON_NUM
        # 遍历每位老师
        for tea_index in range(len(self.teacher_names)):
            tea_course = list()
            """[[星期一第一节,12,13,..,18],...,[星期五第一节,52,...,58]]"""
            #  每8节课拆分为一天
            count = 0
            start = 0
            end = MAX_LESSON_NUM
            # 一周五天
            while count < 5:
                # 每天课程所在班级
                every_day = row_datas_vec[tea_index][start + count * cycle:end + count * cycle]
                count = count + 1
                # print(
                #     "老师{},星期{}, len:{}, {}".format(self.teacher_names[tea_index], count, len(every_day), every_day))

                tea_course.append(every_day)  # 此种顺序与teacher_names中顺序一一对应,即tea_course[0] 是 teacher_names[0] 该位老师的上课表
            # print(tea_course)
            teas_course.append(tea_course)

        print("teas_course len:{}".format(len(teas_course)))

        """创建教师课程表"""

        export_path2 = r"C:\\教师课程表.xlsx"
        wb2 = self.app.books.add()
        sheet_arr2 = list()

        for i in range(len(self.teacher_names)):
            sheet_arr2.append(wb2.sheets.add("{}课程表({})".format(self.teacher_names[i], i)))

        print("sheet_arr2 len:{}".format(len(sheet_arr2)))

        for sht in sheet_arr2:
            # 画框
            sht.range("A1:I7").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I7").api.Borders(9).Weight = 3
            sht.range("A1:I7").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I7").api.Borders(10).Weight = 3
            sht.range("A1:I7").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I7").api.Borders(11).Weight = 2
            sht.range("A1:I7").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I7").api.Borders(12).Weight = 2

            sht.autofit()
            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("A1:I7").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '星期/节'
            # 按行写
            sht.range('B2').value = [1, 2, 3, 4, 5, 6, 7, 8]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['一', '二', '三', '四', '五']

        for cn_index in range(len(self.teacher_names)):
            sheet_arr2[cn_index].range('A1').value = '教师{}课程表'.format(self.teacher_names[cn_index])
            for day in range(0, MAX_DAY_NUM):
                sheet_arr2[cn_index].range('B' + str(day + 3)).value = teas_course[cn_index][day]

        wb2.save(export_path2)

    def sel_classes_table(self):

        """
        提取年级课程表
        """

        MAX_CLASSES_NUM = 5
        MAX_DAY_NUM = 5  # 一周上几天学
        WEEK_LESSION_BEGIN = "A3"
        WEEK_LESSION_END = "AP50"
        WEEK_LESSION = "{}:{}".format(WEEK_LESSION_BEGIN, WEEK_LESSION_END)

        print('范围:{}\n'.format(WEEK_LESSION))

        scope = self.sheet.range(WEEK_LESSION)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))

        col_datas_vec = []
        '''所有列集合'''

        # 遍历列: 列索引从0开始
        for i in range(0, self.col_nums):
            col_datas = []
            """每节课班级集合"""
            for j in range(int(WEEK_LESSION_BEGIN[1:]), int(WEEK_LESSION_END[2:]) + 1):
                # 遍历scope中每列数据
                # print(COL_NAME[i] + str(j))
                if self.sheet.range(COL_NAME[i] + str(j)).value is not None:
                    col_datas.append(self.sheet.range(COL_NAME[i] + str(j)).value)  # print(COL_NAME[i] + str(j))
            # print(col_datas)
            col_datas_vec.append(col_datas)
        print("采取:{}".format(col_datas_vec))

        cycle = 8
        # 一共MAX_CLASSES_NUM个班
        for cl_index in range(1, MAX_CLASSES_NUM + 1):
            days = list()
            count = 0
            start = 2
            end = 10
            # 一周五天
            while count < 5:
                every_day = col_datas_vec[start + count * cycle:end + count * cycle]
                count = count + 1
                # print("周{},len={}: {}".format(count, len(every_week), every_week))
                subjects = col_datas_vec[1]
                lesson_nums = len(every_day)  # 每天几节课
                lession_subject_list = list(dict())
                """一天上课集合 [{第一节:课程},{第二节:课程},]"""
                # 周一的第一节~第八节课
                for lesson_index in range(lesson_nums):
                    # 寻找一班上的
                    for classes_index in range(0, self.row_nums):
                        if str(every_day[lesson_index][classes_index]).strip() == str(cl_index):
                            # print("一班周{}第{}节课:{}".format(count, lesson_index + 1, subjects[classes_index]))
                            # lession_subject = {lesson_index + 1: subjects[classes_index]}
                            lession_subject = subjects[classes_index]
                            """节:课程"""
                            lession_subject_list.append(lession_subject)
                # print("{}班周{}的课程表:{}".format(cl_index, count, lession_subject_list))
                # print(lession_subject_list)
                days.append(lession_subject_list)
            #
            # print("{}班一周的课程表:{}".format(cl_index, days))
            # self.weeks.append(self.days)
            self.weeks[cl_index] = days
        #
        print("所有班级一周的课程表[{}]:{}".format(len(self.weeks), self.weeks))

        """创建班级课程表的excel"""
        export_path = r"C:\\班级课程表.xlsx"
        wb1 = self.app.books.add()
        sheet_arr = list()

        for i in range(MAX_CLASSES_NUM):
            sheet_arr.append(wb1.sheets.add("{}班课程表".format(i + 1)))

        for sht in sheet_arr:
            # 画框
            sht.range("A1:I7").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I7").api.Borders(9).Weight = 3
            sht.range("A1:I7").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I7").api.Borders(10).Weight = 3
            sht.range("A1:I7").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I7").api.Borders(11).Weight = 2
            sht.range("A1:I7").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I7").api.Borders(12).Weight = 2

            sht.autofit()
            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("A1:I7").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '星期/节'
            # 按行写
            sht.range('B2').value = [1, 2, 3, 4, 5, 6, 7, 8]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['一', '二', '三', '四', '五']

        """将班级课程表写入数据库"""
        for cn_index in range(MAX_CLASSES_NUM):
            # for sht in sheet_arr:
            sheet_arr[cn_index].range('A1').value = '{}班课程表'.format(cn_index + 1)
            for day in range(0, MAX_DAY_NUM):
                sheet_arr[cn_index].range('B' + str(day + 3)).value = self.weeks.get(cn_index + 1, '')[day]

        wb1.save(export_path)
        wb1.close()

    def done_syllabus(self):
        self.sel_classes_table()
        self.sel_teacher_table()

    def cope_asc(self):
        """补课课程表"""
        WEEK_LESSION_BEGIN = "A3"
        WEEK_LESSION_END = "AX33"
        WEEK_LESSION = "A3:AX33"
        MAX_DAY_NUM = 6
        MAX_LESSON_NUM = 8
        DATA_AREA_START = 2  # 数据区起始列名: 1:B

        teacher_names_vec = list()

        self.sheet: xw.Sheet = self.wb.sheets[1]
        scope = self.sheet.range(WEEK_LESSION)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])

        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))
        row_datas_vec = []
        '''每个教师一周上课班级的集合'''

        # 遍历每行,即每位教师 row3 ~ row37 最多支持35位老师
        for i in range(int(WEEK_LESSION_BEGIN[1:]), int(WEEK_LESSION_END[2:]) + 1):
            row_datas = list()
            # 所有列
            for j in range(DATA_AREA_START, self.col_nums):
                row_datas.append('' if (self.sheet.range(COL_NAME[j] + str(i)).value is None or self.sheet.range(
                    COL_NAME[j] + str(i)).value == '') else (self.sheet.range(COL_NAME[j] + str(i)).value + '班'))
            #
            if self.sheet.range(COL_NAME[0] + str(i)).value is not None:
                teacher_names_vec.append(str(self.sheet.range(COL_NAME[0] + str(i)).value).strip())
            row_datas_vec.append(row_datas)

        print(row_datas_vec)
        #
        #

        # 教师人数
        print("教师总人数:{}".format(len(teacher_names_vec)))

        teas_course = list()

        cycle = MAX_LESSON_NUM
        # 遍历每位老师
        for tea_index in range(0, len(teacher_names_vec)):
            tea_course = list()
            #  每8节课拆分为一天
            count = 0
            start = 0
            end = MAX_LESSON_NUM
            # 一周五天
            while count < MAX_DAY_NUM:
                # 每天课程所在班级
                every_day = row_datas_vec[tea_index][start + count * cycle:end + count * cycle]
                count = count + 1
                tea_course.append(every_day)  # 此种顺序与teacher_names中顺序一一对应,即tea_course[0] 是 teacher_names[0] 该位老师的上课表
            # print(tea_course)
            #
            teas_course.append(tea_course)
            print("teas_course len:{}, {}".format(len(teas_course), tea_course))

        export_path3 = r"C:\\教师补课表.xlsx"
        wb3 = self.app.books.add()
        sheet_arr3 = list()

        for i in range(len(teacher_names_vec)):
            sheet_arr3.append(wb3.sheets.add("{}补课表({})".format(teacher_names_vec[i], i)))

        print("sheet_arr3 len:{}".format(len(sheet_arr3)))

        for sht in sheet_arr3:
            # 画框
            sht.range("A1:I8").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I8").api.Borders(9).Weight = 3
            sht.range("A1:I8").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I8").api.Borders(10).Weight = 3
            sht.range("A1:I8").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I8").api.Borders(11).Weight = 2
            sht.range("A1:I8").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I8").api.Borders(12).Weight = 2

            sht.autofit()
            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("A1:I8").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '组/节'
            # 按行写
            sht.range('B2').value = [x for x in range(1, MAX_LESSON_NUM + 1)]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

        for cn_index in range(len(teacher_names_vec)):
            sheet_arr3[cn_index].range('A1').value = '教师{}补课表'.format(teacher_names_vec[cn_index])
            for day in range(0, MAX_DAY_NUM):
                sheet_arr3[cn_index].range('B' + str(day + 3)).value = teas_course[cn_index][day]

        wb3.save(export_path3)
        wb3.close()

    def classes_asc(self):
        """班级补课表"""

        WEEK_LESSION_BEGIN = "A3"
        WEEK_LESSION_END = "AX33"
        WEEK_LESSION = "A3:AX33"
        MAX_CLASSES_NUM = 6
        MAX_DAY_NUM = 6
        MAX_LESSON_NUM = 8
        DATA_AREA_START = 2  # 数据区起始列名: 2:C
        weeks = dict()

        self.sheet: xw.Sheet = self.wb.sheets[1]
        scope = self.sheet.range(WEEK_LESSION)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))

        col_datas_vec = []
        '''所有列集合'''

        # 遍历列: 列索引从0开始
        for i in range(0, self.col_nums):
            col_datas = []
            """每节课班级集合"""
            for j in range(int(WEEK_LESSION_BEGIN[1:]), int(WEEK_LESSION_END[2:]) + 1):
                col_datas.append(self.sheet.range(COL_NAME[i] + str(j)).value)
                #
            col_datas_vec.append(col_datas)
        print("采集所有列数据:{}".format(col_datas_vec))

        cycle = MAX_LESSON_NUM
        # 一共MAX_CLASSES_NUM个班
        for cl_index in range(1, MAX_CLASSES_NUM + 1):
            days = list()
            count = 0
            start = DATA_AREA_START
            end = 10
            # 一周五天
            while count < MAX_DAY_NUM:
                every_day = col_datas_vec[start + count * cycle:end + count * cycle]
                count = count + 1
                # print("周{},len={}: {}".format(count, len(every_day), every_day))
                #
                subjects = col_datas_vec[1]
                lesson_nums = len(every_day)  # 每天几节课
                lession_subject_list = list(dict())
                """一天上课集合 [{第一节:课程},{第二节:课程},]"""
                # 周一的第一节~第八节课
                for lesson_index in range(lesson_nums):
                    # 寻找一班上的
                    for classes_index in range(0, self.row_nums):
                        if str(every_day[lesson_index][classes_index]).strip() == str(cl_index):
                            # print("一班周{}第{}节课:{}".format(count, lesson_index + 1, subjects[classes_index]))
                            #
                            # lession_subject = {lesson_index + 1: subjects[classes_index]}
                            lession_subject = subjects[classes_index]
                            """节:课程"""
                            lession_subject_list.append(lession_subject)

                # print("{}班周{}的班级补课表:{}".format(cl_index, count, lession_subject_list))
                #
                # print(lession_subject_list)
                days.append(lession_subject_list)
            # print("{}班一周的班级补课表:{}".format(cl_index, days))
            #
            # self.weeks.append(self.days)
            weeks[cl_index] = days

        print("所有班级一周的班级补课表[{}]:{}".format(len(weeks), weeks))

        """.:cvar......................................创建表........................................."""
        export_path4 = r"C:\\班级补课表.xlsx"
        wb4 = self.app.books.add()
        sheet_arr4 = list()

        for i in range(MAX_CLASSES_NUM):
            sheet_arr4.append(wb4.sheets.add("{}班补课表".format(i + 1)))

        print("sheet_arr4 len:{}".format(len(sheet_arr4)))
        for sht in sheet_arr4:
            # 画框
            sht.range("A1:I8").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I8").api.Borders(9).Weight = 3
            sht.range("A1:I8").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I8").api.Borders(10).Weight = 3
            sht.range("A1:I8").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I8").api.Borders(11).Weight = 2
            sht.range("A1:I8").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I8").api.Borders(12).Weight = 2

            sht.autofit()
            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("A1:I8").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '周/节'
            # 按行写
            sht.range('B2').value = [x for x in range(1, MAX_LESSON_NUM + 1)]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

        """将班级课程表写入数据库"""
        for cn_index in range(MAX_CLASSES_NUM):
            sheet_arr4[cn_index].range('A1').value = '{}班补课表'.format(cn_index + 1)
            for day in range(0, MAX_DAY_NUM):
                sheet_arr4[cn_index].range('B' + str(day + 3)).value = weeks.get(cn_index + 1, '')[day]

        wb4.save(export_path4)
        wb4.close()

    def done(self):
        self.done_syllabus()
        self.cope_asc()
        self.classes_asc()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = Ui_MainWindow()
    ui.show()
    sys.exit(app.exec_())
