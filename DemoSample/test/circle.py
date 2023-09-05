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
ASC_LESSION_SCOPE = "A3:AX23"
ASC_MAX_CLASSES_NUM = 6
'''补课表最大班级数'''
ASC_MAX_DAY_NUM = 6  # 3周(每周有周六周日)
'''补课表最大天数'''
ASC_MAX_LESSON_NUM = 8
'''补课表每天最多节数'''
ASC_DATA_AREA_START = 2  # 数据区起始列名: 2:C
'''数据起始'''

# 课程表参数
MAX_DAY_NUM = 5  # 一周上几天学
MAX_CLASSES_NUM = 6  # 最多班级数
MAX_LESSON_NUM = 8  # 每天最大课时数
WEEK_LESSION_BEGIN = "A3"
WEEK_LESSION_END = "AP32"


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

                if str(self.sheet.range(COL_NAME[j] + str(i)).value).strip() == '集备':
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
                    col_datas.append(self.sheet.range(COL_NAME[i] + str(j)).value)  # print(COL_NAME[i] + str(j))
            # print(col_datas)
            col_datas_vec.append(col_datas)
        print("总数:{},列集合:{}".format(len(col_datas_vec), col_datas_vec))

        cycle = 8
        # 一共MAX_CLASSES_NUM个班
        for cl_index in range(1, MAX_CLASSES_NUM + 1):
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
                """一天上课集合 [{第一节:课程},{第二节:课程},]"""

                # 周一的第一节~第八节课
                for lesson_index in range(lesson_nums):
                    # 寻找一班上的
                    for classes_index in range(0, self.row_nums):
                        if str(every_day[lesson_index][classes_index]).strip() == str(cl_index):
                            # lession_subject = {lesson_index + 1: subjects[classes_index]}
                            lession_subject = subjects[classes_index]
                            # print("{}班周{}第{}节课:{}".format(cl_index, count, lesson_index + 1,
                            #                                    subjects[classes_index]))
                            """节:课程"""
                            lession_subject_list.append(lession_subject)
                print("{}班周{}的课程表:{}".format(cl_index, count, lession_subject_list))
                # print(lession_subject_list)
                days.append(lession_subject_list)
            #
            # print("{}班一周的课程表:{}".format(cl_index, days))
            # self.weeks.append(self.days)
            self.weeks[cl_index] = days
        #
        print("所有班级一周的课程表[{}]:{}".format(len(self.weeks), self.weeks))

        """创建班级课程表的excel"""
        export_path = r"C:\\" + self.grade + "班级课程表.xlsx"
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
        #
        #

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
            print("teas_course len:{}, {}".format(len(teas_course), tea_course))

        export_path3 = r"C:\\" + self.grade + "教师补课表.xlsx"
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
            sht.range('B2').value = [x for x in range(1, ASC_MAX_LESSON_NUM + 1)]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

        for cn_index in range(len(teacher_names_vec)):
            sheet_arr3[cn_index].range('A1').value = '教师{}补课表'.format(teacher_names_vec[cn_index])
            for day in range(0, ASC_MAX_DAY_NUM):
                sheet_arr3[cn_index].range('B' + str(day + 3)).value = teas_course[cn_index][day]

        wb3.save(export_path3)
        wb3.close()

    def classes_asc(self):
        """班级补课表"""
        print("处理班级补课表******************************************************************")
        # ASC_LESSION_BEGIN = "A3"
        # ASC_LESSION_END = "AX33"
        # ASC_LESSION_SCOPE = "A3:AX33"
        # ASC_MAX_CLASSES_NUM = 5  # 五个班
        # ASC_MAX_DAY_NUM = 6  # 3周(每周有周六周日)
        # ASC_MAX_LESSON_NUM = 8
        # ASC_DATA_AREA_START = 2  # 数据区起始列名: 2:C
        weeks = dict()

        self.sheet: xw.Sheet = self.wb.sheets[1]
        scope = self.sheet.range(ASC_LESSION_SCOPE)
        self.row_nums = len(scope.value)
        self.col_nums = len(scope.value[0])
        print("行数:{} 列数:{}".format(self.row_nums, self.col_nums))

        col_datas_vec = []
        '''所有列集合'''

        # 遍历列: 列索引从0开始
        for i in range(0, self.col_nums):
            col_datas = []
            """每节课班级集合"""
            for j in range(int(ASC_LESSION_BEGIN[1:]), int(ASC_LESSION_END[2:]) + 1):
                col_datas.append(self.sheet.range(COL_NAME[i] + str(j)).value)  #
            col_datas_vec.append(col_datas)
        print("采集所有列数据:{}".format(col_datas_vec))

        cycle = ASC_MAX_LESSON_NUM  # 每天最大节数

        # 一共ASC_MAX_CLASSES_NUM个班
        for cl_index in range(1, ASC_MAX_CLASSES_NUM + 1):
            """遍历每班每天每节上什么课"""
            days = list()
            count = 0
            start = ASC_DATA_AREA_START
            end = 10
            # 一周五天
            while count < ASC_MAX_DAY_NUM:
                #  遍历每天每节课
                every_day = col_datas_vec[start + count * cycle:end + count * cycle]
                count = count + 1
                # print("周{},len={}: {}".format(count, len(every_day), every_day))

                subjects = col_datas_vec[1]  # 课程
                lesson_nums = len(every_day)  # 每天几节课
                print('第{}天,{}节课:{}'.format(count, lesson_nums, every_day))
                lession_subject_list = list(dict())

                """一天上课集合 [[],[],[]]"""
                # 周一的第一节~第八节课
                for lesson_index in range(lesson_nums):
                    lession_subject = ''
                    # 寻找一班上的
                    for classes_index in range(0, self.row_nums):
                        # 考虑去除 None '1.0' (字符串情况)
                        if str(every_day[lesson_index][classes_index]) == str(cl_index):
                            print("{}班,第{}个休息日,第{}节课:{}".format(cl_index, count, lesson_index + 1,
                                                                         subjects[classes_index]))
                            #
                            # lession_subject = {lesson_index + 1: subjects[classes_index]}
                            lession_subject = subjects[classes_index]
                        else:
                            lession_subject = ''

                        lession_subject_list.append(lession_subject)
                        """每天课程"""

                # print("{}班周{}的班级补课表:{}".format(cl_index, count, lession_subject_list))
                # print(lession_subject_list)
                # print("{}班一周的班级补课表:{}".format(cl_index, days))
                days.append(lession_subject_list)
            # self.weeks.append(self.days)
            weeks[cl_index] = days

        print("所有班级一周的班级补课表[{}]:{}".format(len(weeks), weeks))

        """.:cvar......................................创建表........................................."""
        export_path4 = r"C:\\" + self.grade + "班级补课表.xlsx"
        wb4 = self.app.books.add()
        sheet_arr4 = list()

        for i in range(ASC_MAX_CLASSES_NUM):
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
            sht.range('B2').value = [x for x in range(1, ASC_MAX_LESSON_NUM + 1)]
            # 按列写
            sht.range('A3').options(transpose=True).value = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']

        """将班级课程表写入数据库"""
        for cn_index in range(ASC_MAX_CLASSES_NUM):
            sheet_arr4[cn_index].range('A1').value = '{}班补课表'.format(cn_index + 1)
            for day in range(0, ASC_MAX_DAY_NUM):
                sheet_arr4[cn_index].range('B' + str(day + 3)).value = weeks.get(cn_index + 1, '')[day]

        wb4.save(export_path4)
        wb4.close()

    def done(self):
        self.done_syllabus()
        self.cope_asc()
        self.classes_asc()


if __name__ == '__main__':
    doesss = Syllabus('高二', r"C:\Projects\老婆\课表\高二课程表.xlsx")
    # doesss.done()
    doesss.classes_asc()
