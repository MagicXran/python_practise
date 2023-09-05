import math

import xlwings as xw

'''常量'''
COL_NAME = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
            'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']

# HOSTNAME = "TrLoveXran"
HOSTNAME = "Karma"

ASC_LESSION_BEGIN = "A3"
'''补课表数据起始'''
ASC_LESSION_END = "AX23"
'''补课表数据终止'''
ASC_LESSION_SCOPE = ASC_LESSION_BEGIN + ":" + ASC_LESSION_END
ASC_MAX_CLASSES_NUM = 6
'''补课表最大班级数'''
ASC_MAX_DAY_NUM = 6  # 3周(每周有周六周日)
'''补课表最大天数'''
ASC_MAX_LESSON_NUM = 8
'''补课表每天最多节数'''
ASC_DATA_AREA_START = 2  # 数据区起始列名: 2:C
'''数据起始'''

# 课程表参数
MAX_DAY_NUM = 5
'''一周上几天学'''
MAX_CLASSES_NUM = 5  # 最多班级数
MAX_LESSON_NUM = 8  # 每天最大课时数
WEEK_LESSION_BEGIN = "A9"
WEEK_LESSION_END = "AP64"


class Syllabus(object):
    def __init__(self, grade, source_path_):
        self.col_nums = None
        self.row_nums = None
        self.app = xw.App(visible=False, add_book=False)
        # self.wb: xw.Book = self.app.books.open(ui.source_path)
        self.wb: xw.Book = self.app.books.open(source_path_)
        self.sheet: xw.Sheet = self.wb.sheets[0]
        self.weeks = dict()
        self.teacher_names = list()
        self.grade = grade

    def __del__(self):
        self.wb.close()
        self.app.quit()
        print("执行析构")

    def sel_teacher_table(self):
        """提前教师课程表"""

        print("处理教师课程表******************************************************************")
        WEEK_LESSION = "{}:{}".format(WEEK_LESSION_BEGIN, WEEK_LESSION_END)
        print('范围:{}'.format(WEEK_LESSION))

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

                if str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() == '集备' or str(
                        self.sheet.range(COL_NAME[j] + str(i)).value).strip() == '集':
                    row_datas.append(str(self.sheet.range(COL_NAME[j] + str(i)).value).strip())
                elif str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() != 'None' and str(
                        self.sheet.range(COL_NAME[j] + str(i)).value).strip():
                    row_datas.append(str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() + '班')
                elif str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() == 'None':
                    row_datas.append('')
                else:
                    row_datas.append(str(self.sheet.range(COL_NAME[j] + str(i)).value).strip())  # print(row_datas)

            self.teacher_names.append(str(self.sheet.range(COL_NAME[0] + str(i)).value).strip())
            row_datas_vec.append(row_datas)

        print('打印课程:{}'.format(row_datas_vec))
        print('打印教师:{}'.format(self.teacher_names))
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
                if row_datas_vec[tea_index][start + count * cycle:end + count * cycle] == 'None':
                    every_day = ''
                else:
                    every_day = row_datas_vec[tea_index][start + count * cycle:end + count * cycle]
                count = count + 1
                # print(
                #     "老师{},星期{}, len:{}, {}".format(self.teacher_names[tea_index], count, len(every_day), every_day))

                tea_course.append(every_day)  # 此种顺序与teacher_names中顺序一一对应,即tea_course[0] 是 teacher_names[0] 该位老师的上课表
            # print(tea_course)
            teas_course.append(tea_course)

        print("teas_course len:{},{}".format(len(teas_course), tea_course))

        """创建教师课程表"""

        export_path2 = r"C:\\" + self.grade + "教师课程表.xlsx"
        wb2 = self.app.books.add()
        sheet_arr2 = list()

        # 每个sheet 装 14 个教师课程表
        sheet_nums = math.ceil(len(self.teacher_names) / 12)
        print('math.ceil((len(self.teacher_names)) / 12) = {}'.format(sheet_nums))
        for i in range(sheet_nums):
            sheet_arr2.append(wb2.sheets.add("{}课程表({})".format(self.grade, i)))

        print("sheet_arr2 len:{}".format(len(sheet_arr2)))

        for sht in sheet_arr2:
            # 画框
            # sht.autofit() # 自动适应行高
            sht.range("A1:T55").api.RowHeight = 15
            sht.range("A1:T55").api.ColumnWidth = 3

            sht.range("A1:T55").api.Font.Size = 9  # 字体大小
            sht.range("A1:T55").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range("A1").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L1").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L1").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A9").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A9").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L9").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L9").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A17").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A17").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L17").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L17").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A25").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A25").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L25").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L25").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A33").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A33").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L33").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L33").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A41").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A41").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L41").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L41").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A49").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A49").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L49").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L49").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A1:T55").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

        # 遍历每个教师
        cyc_tea = 0
        # 每个sheet
        for sht_index in range(sheet_nums):
            print("遍历")

            count = 0
            rigth_start_row = 0  # 右边每个表左上角起始位置
            left_start_row = 0
            for cn_index in range(cyc_tea, len(self.teacher_names)):
                # 14个教师一组
                count = count + 1
                # 偶数 => 右边
                if (count & 1) == 0:
                    rigth_start_row = rigth_start_row + 1
                    # 合并单元格
                    sheet_arr2[sht_index].range('L' + str(((rigth_start_row - 1) * 8 + 1)) + ':' + 'T' + str(
                        ((rigth_start_row - 1) * 8 + 1))).merge()

                    sheet_arr2[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 8 + 1))).value = '教师:{}课程表'.format(
                        self.teacher_names[cn_index])

                    # 画框
                    # a_n=1+(n-1)*8
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(8).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(8).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(9).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(9).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(7).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(7).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(10).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(10).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(11).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(11).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(12).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("L" + str((rigth_start_row - 1) * 8 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 8 + 7)).api.Borders(12).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range('L' + str(((rigth_start_row - 1) * 8 + 2))).value = '星期/节'
                    sheet_arr2[sht_index].range('M' + str(((rigth_start_row - 1) * 8 + 2))).value = [1, 2, 3, 4, 5, 6,
                                                                                                     7, 8]
                    sheet_arr2[sht_index].range('L' + str(((rigth_start_row - 1) * 8 + 3))).options(
                        transpose=True).value = ['一', '二', '三', '四', '五']

                    sheet_arr2[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 8 + 2))).api.Font.Size = 5.5  # 设置单元格字体大小
                    sheet_arr2[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 8 + 2))).api.Font.Bold = True  # 设置单元格字体是否加粗

                    for day in range(0, MAX_DAY_NUM):
                        sheet_arr2[sht_index].range('M' + str(day + ((rigth_start_row - 1) * 8 + 3))).value = \
                            teas_course[cn_index][day]

                else:
                    left_start_row = left_start_row + 1
                    sheet_arr2[sht_index].range('A' + str(((left_start_row - 1) * 8 + 1)) + ':' + 'I' + str(
                        ((left_start_row - 1) * 8 + 1))).merge()
                    sheet_arr2[sht_index].range(
                        'A' + str(((left_start_row - 1) * 8 + 1))).value = '教师:{}课程表'.format(
                        self.teacher_names[cn_index])

                    sheet_arr2[sht_index].range('A' + str(((left_start_row - 1) * 8 + 2))).value = '星期/节'
                    sheet_arr2[sht_index].range('B' + str(((left_start_row - 1) * 8 + 2))).value = [1, 2, 3, 4, 5, 6,
                                                                                                    7, 8]
                    sheet_arr2[sht_index].range('A' + str(((left_start_row - 1) * 8 + 3))).options(
                        transpose=True).value = ['一', '二', '三', '四', '五']
                    # 画框
                    # a_n=1+(n-1)*8
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(8).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(8).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(9).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(9).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(7).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(7).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(10).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(10).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(11).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(11).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(12).LineStyle = 1  # 划顶部边框
                    sheet_arr2[sht_index].range("A" + str((left_start_row - 1) * 8 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 8 + 7)).api.Borders(12).Weight = 3  # 划顶部边框

                    sheet_arr2[sht_index].range(
                        'A' + str(((left_start_row - 1) * 8 + 2))).api.Font.Size = 5.5  # 设置单元格字体大小
                    sheet_arr2[sht_index].range(
                        'A' + str(((left_start_row - 1) * 8 + 2))).api.Font.Bold = True  # 设置单元格字体是否加粗

                    for day in range(0, MAX_DAY_NUM):
                        sheet_arr2[sht_index].range('B' + str(day + ((left_start_row - 1) * 8 + 3))).value = \
                            teas_course[cn_index][day]

                # 则终止(或跳出)与break最贴近的那层循环
                if count == 12:
                    cyc_tea = cyc_tea + 12
                    print('跳出内层循环!')
                    break

        wb2.save(export_path2)

    def sel_classes_table(self):
        """
        提取年级课程表
        """
        print("处理班级课程表******************************************************************")
        WEEK_LESSION = "{}:{}".format(WEEK_LESSION_BEGIN, WEEK_LESSION_END)

        print('年级课程表范围:{}'.format(WEEK_LESSION))

        scope = self.sheet.range(WEEK_LESSION)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("课程表行数:{} 列数:{}".format(self.row_nums, self.col_nums))

        col_datas_vec = []
        '''所有列集合'''

        # 遍历列: 列索引从0开始
        for i in range(0, self.col_nums):
            col_datas = []
            """每节课班级集合"""
            for j in range(int(WEEK_LESSION_BEGIN[1:]), int(WEEK_LESSION_END[2:]) + 1):
                # 遍历scope中每列数据
                # print(COL_NAME[i] + str(j))
                if self.sheet.range(COL_NAME[i] + str(j)).value is None:
                    col_datas.append('')  # print(COL_NAME[i] + str(j))
                else:
                    col_datas.append(
                        str(self.sheet.range(COL_NAME[i] + str(j)).value).strip(' ').replace('.', '').replace('0',
                                                                                                              ''))  # print(COL_NAME[i] + str(j))
            # print(col_datas)
            col_datas_vec.append(col_datas)
        print("总数:{},列集合:{}".format(len(col_datas_vec), col_datas_vec))
        # print('15:{}', col_datas_vec[15])

        # return 0
        cycle = 8
        # 一共MAX_CLASSES_NUM个班
        # for cl_index in range(6, 6 + 1):
        for cl_index in range(1, MAX_CLASSES_NUM + 1):
            '''遍历每个班级'''
            days = list()
            count = 0
            start = 2
            end = 10
            # 一周五天
            while count < MAX_DAY_NUM:
                every_day = col_datas_vec[start + count * cycle:end + count * cycle]
                # print(every_day)
                count = count + 1
                # print("周{},len={}: {}".format(count, len(every_week), every_week))
                subjects = col_datas_vec[1]
                lesson_nums = len(every_day)  # 每天几节课
                # print("每天{}节课{}:".format(len(every_day), every_day))

                lession_subject_list = list(dict())
                class_oneday_subject = list()
                """一天上课集合 [{第一节:课程},{第二节:课程},]"""

                # 周一的第一节~第八节课
                for lesson_index in range(lesson_nums):
                    # 寻找一班上的
                    # for classes_index in range(0, self.row_nums):
                    if str(cl_index) in every_day[lesson_index]:
                        index_class = every_day[lesson_index].index(str(cl_index))
                        class_oneday_subject.append(subjects[index_class])
                    else:
                        class_oneday_subject.append('')
                        """节:课程"""
                print("{}班周{}的课程表:{}".format(cl_index, count, class_oneday_subject))
                days.append(class_oneday_subject)
            self.weeks[cl_index] = days
        print("所有班级一周的课程表[{}]:{}".format(len(self.weeks), self.weeks))

        """创建班级课程表的excel"""
        export_path = r"C:\\" + self.grade + "班级课程表.xlsx"
        wb1 = self.app.books.add()
        sheet_arr = list()

        for i in range(MAX_CLASSES_NUM):
            sheet_arr.append(wb1.sheets.add("{}班课程表".format(i + 1)))

        for sht in sheet_arr:
            # 画框
            sht.range("A1:I7").api.Borders(8).LineStyle = 1  # 划顶部边框
            sht.range("A1:I7").api.Borders(8).Weight = 3

            sht.range("A1:I7").api.Borders(7).LineStyle = 1  # 划左部边框
            sht.range("A1:I7").api.Borders(7).Weight = 3

            sht.range("A1:I7").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I7").api.Borders(9).Weight = 3

            sht.range("A1:I7").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I7").api.Borders(10).Weight = 3

            sht.range("A1:I7").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I7").api.Borders(11).Weight = 2

            sht.range("A1:I7").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I7").api.Borders(12).Weight = 2

            # sht.autofit() # 自动适应行高
            sht.range("A1:I7").api.RowHeight = 105
            sht.range("A1:I7").api.ColumnWidth = 8.22

            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A1:I7").api.Font.Size = 28  # 字体大小
            sht.range("A1:I7").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '星期/节'
            sht.range('A2').api.Font.Size = 14  # 字体大小
            # 按行写
            sht.range('B2').value = [1, 2, 3, 4, 5, 6, 7, 8]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['一', '二', '三', '四', '五']

        """将班级课程表写入数据库"""
        for cn_index in range(MAX_CLASSES_NUM):
            # for sht in sheet_arr:
            sheet_arr[cn_index].range('A1').value = '{}:{}班课程表'.format(self.grade, cn_index + 1)
            for day in range(0, MAX_DAY_NUM):
                sheet_arr[cn_index].range('B' + str(day + 3)).value = self.weeks.get(cn_index + 1, '')[day]

                # 设置 单元格中 两个汉字的字号  22
                for lesson_index in range(MAX_LESSON_NUM):
                    # 找到每班每天每节课名称
                    if len(self.weeks.get(cn_index + 1, '')[day][lesson_index]) > 1:
                        sheet_arr[cn_index].range(COL_NAME[lesson_index + 1] + str(day + 3)).api.Font.Size = 22  # 字体大小

        wb1.save(export_path)
        wb1.close()

    def done_syllabus(self):
        self.sel_classes_table()
        self.sel_teacher_table()

    def teacher_asc(self):
        """教师补课表"""

        print("处理教师补课表******************************************************************")

        teacher_names_vec = list()

        self.sheet: xw.Sheet = self.wb.sheets[1]
        scope = self.sheet.range(ASC_LESSION_SCOPE)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])

        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))
        row_datas_vec = []
        '''每个教师一周上课班级的集合'''

        # 遍历每行,即每位教师 row3 ~ row37 最多支持35位老师
        for i in range(int(ASC_LESSION_BEGIN[1:]), int(ASC_LESSION_END[2:]) + 1):
            row_datas = list()
            # 所有列
            for j in range(ASC_DATA_AREA_START, self.col_nums):
                row_datas.append('' if (self.sheet.range(COL_NAME[j] + str(i)).value is None or self.sheet.range(
                    COL_NAME[j] + str(i)).value == '') else str(
                    int(self.sheet.range(COL_NAME[j] + str(i)).value)) + '班')
            #
            if self.sheet.range(COL_NAME[0] + str(i)).value is not None:
                teacher_names_vec.append(str(self.sheet.range(COL_NAME[0] + str(i)).value).strip())
            row_datas_vec.append(row_datas)

        print(row_datas_vec)

        # 教师人数
        print("教师总人数:{}".format(len(teacher_names_vec)))

        teas_course = list()

        cycle = ASC_MAX_LESSON_NUM
        # 遍历每位老师
        for tea_index in range(0, len(teacher_names_vec)):
            tea_course = list()
            #  每8节课拆分为一天
            count = 0
            start = 0
            end = ASC_MAX_LESSON_NUM
            # 一周五天
            while count < ASC_MAX_DAY_NUM:
                # 每天课程所在班级
                every_day = row_datas_vec[tea_index][start + count * cycle:end + count * cycle]
                count = count + 1
                tea_course.append(every_day)  # 此种顺序与teacher_names中顺序一一对应,即tea_course[0] 是 teacher_names[0] 该位老师的上课表
            # print(tea_course)
            #
            teas_course.append(tea_course)
            print("teas_course len:{}, {}".format(len(teas_course), teas_course))

        export_path3 = r"C:\\" + self.grade + "教师补课表.xlsx"
        wb3 = self.app.books.add()
        sheet_arr3 = list()

        # 每个sheet 装 14 个教师课程表
        sheet_nums = math.ceil(len(teacher_names_vec) / 12)
        print('math.ceil((len(self.teacher_names)) / 12) = {}'.format(sheet_nums))
        for i in range(sheet_nums):
            sheet_arr3.append(wb3.sheets.add("{}补课表({})".format(self.grade, i)))

        print("sheet_arr3 len:{}".format(len(sheet_arr3)))

        for sht in sheet_arr3:
            # 画框
            # sht.autofit() # 自动适应行高
            sht.range("A1:T55").api.RowHeight = 15
            sht.range("A1:T55").api.ColumnWidth = 3

            sht.range("A1:T55").api.Font.Size = 9  # 字体大小
            sht.range("A1:T55").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range("A1").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L1").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L1").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A10").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A10").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L10").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L10").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A19").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A19").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L19").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L19").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A28").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A28").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L28").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L28").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A37").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A37").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L37").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L37").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A46").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("A46").api.Font.Bold = True  # 设置单元格字体是否加粗
            sht.range("L46").api.Font.Size = 10  # 设置单元格字体大小
            sht.range("L46").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A1:T55").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

        # 遍历每个教师
        cyc_tea = 0
        # 每个sheet
        for sht_index in range(sheet_nums):
            print("遍历")

            count = 0
            rigth_start_row = 0  # 右边每个表左上角起始位置
            left_start_row = 0
            for cn_index in range(cyc_tea, len(teacher_names_vec)):
                # 14个教师一组
                count = count + 1
                # 偶数 => 右边
                if (count & 1) == 0:
                    rigth_start_row = rigth_start_row + 1
                    # 合并单元格
                    sheet_arr3[sht_index].range('L' + str(((rigth_start_row - 1) * 9 + 1)) + ':' + 'T' + str(
                        ((rigth_start_row - 1) * 9 + 1))).merge()

                    sheet_arr3[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 9 + 1))).value = '教师:{}课程表'.format(
                        teacher_names_vec[cn_index])

                    # 画框
                    # a_n=1+(n-1)*9
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(8).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(8).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(9).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(9).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(7).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(7).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(10).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(10).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(11).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(11).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(12).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("L" + str((rigth_start_row - 1) * 9 + 1) + ':' + 'T' + str(
                        (rigth_start_row - 1) * 9 + 8)).api.Borders(12).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range('L' + str(((rigth_start_row - 1) * 9 + 2))).value = '组/节'
                    sheet_arr3[sht_index].range('M' + str(((rigth_start_row - 1) * 9 + 2))).value = [1, 2, 3, 4, 5, 6,
                                                                                                     7, 8]
                    sheet_arr3[sht_index].range('L' + str(((rigth_start_row - 1) * 9 + 3))).options(
                        transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

                    sheet_arr3[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 9 + 2))).api.Font.Size = 5.5  # 设置单元格字体大小
                    sheet_arr3[sht_index].range(
                        'L' + str(((rigth_start_row - 1) * 9 + 2))).api.Font.Bold = True  # 设置单元格字体是否加粗

                    for day in range(0, ASC_MAX_DAY_NUM):
                        sheet_arr3[sht_index].range('M' + str(day + ((rigth_start_row - 1) * 9 + 3))).value = \
                            teas_course[cn_index][day]

                else:
                    left_start_row = left_start_row + 1
                    sheet_arr3[sht_index].range('A' + str(((left_start_row - 1) * 9 + 1)) + ':' + 'I' + str(
                        ((left_start_row - 1) * 9 + 1))).merge()
                    sheet_arr3[sht_index].range(
                        'A' + str(((left_start_row - 1) * 9 + 1))).value = '教师:{}课程表'.format(
                        teacher_names_vec[cn_index])

                    sheet_arr3[sht_index].range('A' + str(((left_start_row - 1) * 9 + 2))).value = '组/节'
                    sheet_arr3[sht_index].range('B' + str(((left_start_row - 1) * 9 + 2))).value = [1, 2, 3, 4, 5, 6,
                                                                                                    7, 9]
                    sheet_arr3[sht_index].range('A' + str(((left_start_row - 1) * 9 + 3))).options(
                        transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']
                    # 画框
                    # a_n=1+(n-1)*9
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(8).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(8).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(9).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(9).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(7).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(7).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(10).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(10).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(11).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(11).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(12).LineStyle = 1  # 划顶部边框
                    sheet_arr3[sht_index].range("A" + str((left_start_row - 1) * 9 + 1) + ':' + 'I' + str(
                        (left_start_row - 1) * 9 + 8)).api.Borders(12).Weight = 3  # 划顶部边框

                    sheet_arr3[sht_index].range(
                        'A' + str(((left_start_row - 1) * 9 + 2))).api.Font.Size = 5.5  # 设置单元格字体大小
                    sheet_arr3[sht_index].range(
                        'A' + str(((left_start_row - 1) * 9 + 2))).api.Font.Bold = True  # 设置单元格字体是否加粗

                    for day in range(0, ASC_MAX_DAY_NUM):
                        sheet_arr3[sht_index].range('B' + str(day + ((left_start_row - 1) * 9 + 3))).value = \
                            teas_course[cn_index][day]

                # 则终止(或跳出)与break最贴近的那层循环
                if count == 12:
                    cyc_tea = cyc_tea + 12
                    print('跳出内层循环!')
                    break

        wb3.save(export_path3)
        wb3.close()

    def classes_asc(self):
        """班级补课表"""
        print("处理班级补课表******************************************************************")
        weeks = dict()
        '''一周每班课程表: {1班:[[[周一第一天第一节课程,...],第二天,....,第八天],[周二第一天,第二天,...,第八天],[]],2:[],...,6:[]} '''
        self.sheet: xw.Sheet = self.wb.sheets[1]
        scope = self.sheet.range(ASC_LESSION_SCOPE)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("采集范围:{},行数:{} 列数:{}".format(ASC_LESSION_SCOPE, self.row_nums, self.col_nums))

        col_datas_vec = []
        '''所有列集合'''

        # 遍历列: 列索引从0开始
        for i in range(0, self.col_nums):
            col_datas = []
            """每节课班级集合"""
            for j in range(int(ASC_LESSION_BEGIN[1:]), int(ASC_LESSION_END[2:]) + 1):
                col_datas.append(self.sheet.range(COL_NAME[i] + str(j)).value)  #
            col_datas_vec.append(col_datas)
        print("采集所有列,len:{},数据:{}".format(len(col_datas_vec), col_datas_vec))

        subjects = col_datas_vec[1]  # 课程

        cycle = ASC_MAX_LESSON_NUM  # 每天最大节数

        # 遍历每个班
        for cl_index in range(1, ASC_MAX_CLASSES_NUM + 1):
            """遍历每班每天每节上什么课"""
            days = list()
            '''一周内每天课程: [[1,2,3,..,8],[周二],[],...,[周五]]'''
            count = 0
            start = ASC_DATA_AREA_START
            end = 10
            # 一周五天
            while count < ASC_MAX_DAY_NUM:
                #  遍历每天每节课
                every_day = col_datas_vec[start + count * cycle:end + count * cycle]
                '''每天的每节课全部科目'''
                count = count + 1
                # print("周{},共{}节: {}".format(count, len(every_day), every_day))
                class_oneday_subject = list()
                """指定班级的一天上课集合['','语','数',...,'第八节课']"""
                # 遍历每天每节
                for lesson_index in range(len(every_day)):
                    # 遍历第一节上课的班级
                    # print('every_day[lesson_index]={}'.format(every_day[lesson_index]))
                    if str(cl_index) in every_day[lesson_index]:
                        index_class = every_day[lesson_index].index(str(cl_index))
                        class_oneday_subject.append(subjects[index_class])
                        # print('{}班,A{},第{}节,{}'.format(cl_index, count, lesson_index + 1, subjects[index_class]))
                    else:
                        class_oneday_subject.append('')
                # print(class_oneday_subject)
                days.append(class_oneday_subject)
                # print('{}班,A1~A6课程{}'.format(cl_index, days))
            weeks[cl_index] = days
        print('每班A1~A6课程{}'.format(weeks))
        """.:cvar......................................创建表........................................."""
        export_path4 = r"C:\\" + self.grade + "班级补课表.xlsx"
        wb4 = self.app.books.add()
        sheet_arr4 = list()

        for i in range(ASC_MAX_CLASSES_NUM):
            sheet_arr4.append(wb4.sheets.add("{}班补课表".format(i + 1)))

        print("sheet_arr4 len:{}".format(len(sheet_arr4)))
        for sht in sheet_arr4:
            # 页边距
            # sht.api.PageSetup.LeftMargin = 1.9

            sht.range("A1:I8").api.Borders(8).LineStyle = 1  # 划顶部边框
            sht.range("A1:I8").api.Borders(8).Weight = 3

            sht.range("A1:I8").api.Borders(7).LineStyle = 1  # 划左部边框
            sht.range("A1:I8").api.Borders(7).Weight = 3

            sht.range("A1:I8").api.Borders(9).LineStyle = 1  # 划底部边框
            sht.range("A1:I8").api.Borders(9).Weight = 3

            sht.range("A1:I8").api.Borders(10).LineStyle = 1  # 划右部边框
            sht.range("A1:I8").api.Borders(10).Weight = 3

            sht.range("A1:I8").api.Borders(11).LineStyle = 1  # 划内部竖线
            sht.range("A1:I8").api.Borders(11).Weight = 2

            sht.range("A1:I8").api.Borders(12).LineStyle = 1  # 划内部横线
            sht.range("A1:I8").api.Borders(12).Weight = 2

            # sht.autofit() # 自动适应行高
            sht.range("A1:I8").api.RowHeight = 105
            sht.range("A1:I8").api.ColumnWidth = 8.22

            # 标题
            sht.range('A1:I1').merge()
            sht.range("A1").api.Font.Size = 15  # 设置单元格字体大小
            sht.range("A1").api.Font.Bold = True  # 设置单元格字体是否加粗

            sht.range("A1:I8").api.Font.Size = 28  # 字体大小
            sht.range("A1:I8").api.HorizontalAlignment = -4108
            # -4108 水平居中。 -4131 靠左，-4152 靠右

            sht.range('A2').value = '周/节'
            sht.range('A2').api.Font.Size = 14
            # 按行写
            sht.range('B2').value = [x for x in range(1, ASC_MAX_LESSON_NUM + 1)]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

        """将班级课程表写入数据库"""
        for cn_index in range(ASC_MAX_CLASSES_NUM):
            sheet_arr4[cn_index].range('A1').value = '{}:{}班补课表'.format(self.grade, cn_index + 1)
            for day in range(0, ASC_MAX_DAY_NUM):
                sheet_arr4[cn_index].range('B' + str(day + 3)).value = weeks.get(cn_index + 1, '')[day]

                # 设置 单元格中 两个汉字的字号  22
                for lesson_index in range(ASC_MAX_LESSON_NUM):
                    # 找到每班每天每节课名称
                    if len(weeks.get(cn_index + 1, '')[day][lesson_index]) > 1:
                        sheet_arr4[cn_index].range(COL_NAME[lesson_index + 1] + str(day + 3)).api.Font.Size = 22  # 字体大小

        wb4.save(export_path4)
        wb4.close()

    def done(self):
        self.done_syllabus()
        self.teacher_asc()
        self.classes_asc()


if __name__ == '__main__':
    # doesss = Syllabus('高三', r"C:\Projects\老婆\课表\高三课程表总表.xlsx")
    # doesss = Syllabus('高一', r"C:\Projects\老婆\课表\副本高一课程表.xlsx")
    doesss = Syllabus('高一', r"C:\Projects\老婆\课表\高一课程表+补课表  总表.xlsx")
    # doesss.done()
    doesss.sel_classes_table()
    # doesss.sel_teacher_table()
    # doesss.teacher_asc()
