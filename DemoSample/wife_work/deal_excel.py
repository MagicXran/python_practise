import xlwings as xw

SUBJECT_QUANTITY = 9
'''最多科目：9科'''

CLASS_QUANTITY = 7
'''总班级数'''

GRADE_NO = 2
'''要处理的年级：高一：1'''

DATA_AREA_BEGIN = 'A2'
'''data begin position in sheet[0] of the excel'''
DATA_AREA_END = 'L285'
'''data end position in sheet[0] of the excel'''
CLASSES_AREA_BEGIN = 'C2'

SUBJECT_AREA_BEGIN = 'D2'
'''subject begin position'''

COL_NAME = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
            'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']
'''常量'''

"""扣人表参数"""
#############################################################
EX_SENIOR1_CLASSES_CELL = 'A3:A19'  # 高一 班级单元格
EX_SENIOR1_NAME_CELL = 'B3:B19'  # 高一姓名单元格

EX_SENIOR2_CLASSES_CELL = 'D3:D17'  # 高2 班级单元格
EX_SENIOR2_NAME_CELL = 'E3:E17'  # 高2姓名单元格

EX_SENIOR3_CLASSES_CELL = 'G3:G24'  # 高2 班级单元格
EX_SENIOR3_NAME_CELL = 'H3:H24'  # 高2姓名单元格

##################################### 总成绩表参数 #####################################

SENIOR1_TOTAL_SCORE = 150 * 3 + 6 * 100
'''高一全科满分'''

SENIOR2_TOTAL_SCORE = 150 * 3 + 3 * 100
'''高2全科满分'''

SENIOR3_TOTAL_SCORE = 150 * 3 + 3 * 100
'''高3全科满分'''

PASS_RATE100 = 60
PASS_RATE150 = 90
'''及格分数线'''
FAIL_RATE100 = 40
FAIL_RATE150 = 60
'''不及格分数线'''

ROW_BEGIN = '2'
'''数据开始行号'''
ROW_END = '200'

SENIOR1_MAX_CLASSES_NUM = 5
SENIOR1_CLASSES_CELL = 'C'  # 高一 班级单元格
SENIOR1_NAME_CELL = 'B'  # 高一姓名单元格
SENIOR1_TOTAL_SCORES = 'D'  # 高一总分
SENIOR1_CHINESE_SCORES = 'J'  # 语文
SENIOR1_MATH_SCORES = 'P'  # 数学
SENIOR1_ENGLISH_SCORES = 'M'  # 外语
SENIOR1_PHYSICS_SCORES = 'S'  # 物理
SENIOR1_CHEMISTRY_SCORES = 'Y'  # 化学
SENIOR1_BIOLOGY_SCORES = 'AE'  # 生物
SENIOR1_POLITICS_SCORES = 'AK'  # 政治
SENIOR1_HISTORY_SCORES = 'V'  # 历史
SENIOR1_GEOGRAPHY_SCORES = 'AQ'  # 地理

SENIOR2_MAX_CLASSES_NUM = 5
SENIOR2_CLASSES_CELL = 'D'  # 高2 班级单元格
SENIOR2_COMBINE_CELL = 'C'  # 高2 类别
SENIOR2_NAME_CELL = 'B'  # 高2姓名单元格
SENIOR2_TOTAL_SCORES = 'E'  # 高2总分
SENIOR2_CHINESE_SCORES = 'K'  # 语文
SENIOR2_MATH_SCORES = 'Q'  # 数学
SENIOR2_ENGLISH_SCORES = 'N'  # 外语
SENIOR2_PHYSICS_SCORES = 'T'  # 物理
SENIOR2_CHEMISTRY_SCORES = 'Z'  # 化学
SENIOR2_BIOLOGY_SCORES = 'AF'  # 生物
SENIOR2_POLITICS_SCORES = 'AL'  # 政治
SENIOR2_HISTORY_SCORES = 'W'  # 历史
SENIOR2_GEOGRAPHY_SCORES = 'AR'  # 地理

SENIOR3_MAX_CLASSES_NUM = 6
SENIOR3_CLASSES_CELL = 'D'  # 高3 班级单元格
SENIOR3_COMBINE_CELL = 'C'  # 高3 类别
SENIOR3_NAME_CELL = 'B'  # 高3姓名单元格
SENIOR3_TOTAL_SCORES = 'E'  # 高3总分
SENIOR3_CHINESE_SCORES = 'K'  # 语文
SENIOR3_MATH_SCORES = 'Q'  # 数学
SENIOR3_ENGLISH_SCORES = 'N'  # 外语
SENIOR3_PHYSICS_SCORES = 'T'  # 物理
SENIOR3_CHEMISTRY_SCORES = 'Z'  # 化学
SENIOR3_BIOLOGY_SCORES = 'AF'  # 生物
SENIOR3_POLITICS_SCORES = 'AL'  # 政治
SENIOR3_HISTORY_SCORES = 'W'  # 历史
SENIOR3_GEOGRAPHY_SCORES = 'AR'  # 地理

SENIOR2_COMBINE_CATEGORY = ['物化政', '物化地', '物生政', '物政地', '物生地', '历政地']
SENIOR3_COMBINE_CATEGORY = ['物生政', '物生地', '物政地', '历政地', '物化生', '物化政']


class Grade:
    def __init__(self, source_path=r"C:\Projects\VPF\成绩\成绩\高中成绩名单.xls"):

        self.app = xw.App(visible=False, add_book=False)
        self.app.screen_updating = False  # 加速脚本运行速度

        self.wb: xw.Book = self.app.books.open(source_path)
        self.score_sheet: xw.Sheet = self.wb.sheets[0]
        '''总分名单'''
        self.score_sheet.autofit()
        self.tea_sheet: xw.Sheet = self.wb.sheets[1]
        '''学科教师名单'''
        self.tea_sheet.autofit()

        self.headcount = list()
        '''总人数'''

        ###################  考试成绩统计表 #############################################

        self.statistical_table_examination_results = r'C:\Projects\Python_codes\DemoSample\wife_work\考试成绩统计表（模板）.xlsx'

        self.ex_path = r'C:\Projects\Python_codes\DemoSample\wife_work\缺考.xlsx'
        '''扣人，缺考excel全路径'''

        self.scores_path = r'C:\Projects\Python_codes\DemoSample\wife_work\一中_高一总成绩.xls'
        '''高x总成绩excel全路径'''

        self.classes_results = list()
        self.classes_scores_info = list()  #############################################################################

    def ready(self):
        scores_list = list()
        '''各科分数总表[[],[],[],...[]] :语数外，物历化，生政地'''

        '''遍历所有班级[A1:K?]'''
        class_list = self.score_sheet.range(CLASSES_AREA_BEGIN + ":" + CLASSES_AREA_BEGIN[0] + DATA_AREA_END[1:]).value
        '''班级列表'''

        print('班级列表：{}'.format(class_list))
        stu_info = self.score_sheet.range(DATA_AREA_BEGIN + ':' + DATA_AREA_BEGIN[0] + DATA_AREA_END[1:]).value
        '''学生姓名表'''

        sub_info = self.tea_sheet.range('A2:A125').value
        '''课程表'''
        print("课程表：{}".format(sub_info))
        cls_info = self.tea_sheet.range('B2:B125').value
        print("班级表：{}".format(cls_info))
        tea_info = self.tea_sheet.range('C2:C125').value
        print("教师表：{}".format(tea_info))

        for index in [x + 3 for x in range(SUBJECT_QUANTITY)]:
            """遍历各科成绩"""
            scores_list.append(self.score_sheet.range(
                COL_NAME[index] + DATA_AREA_BEGIN[1] + ':' + COL_NAME[index] + DATA_AREA_END[1:]).value)
        print("所有课成绩：{}".format(scores_list))

        grade_stu_score_info = dict()
        '''所有班级学生各科成绩信息表'''

        for class_index in range(1, CLASS_QUANTITY + 1):
            '''遍历所有班级'''
            # print(class_index)
            the_classes_stu_info = list()
            hc = 0
            for index in range(len(class_list)):
                '''只保留班级号'''
                temp_cls_no = str(class_list[index]).replace('0', '').replace('班', '') if class_list[
                                                                                               index] is not None else ''
                # print(temp_cls_no)
                if int(temp_cls_no if (temp_cls_no != '') else 0) == class_index:
                    '''每个班级'''
                    class_no = (str(GRADE_NO) + '.' + str(class_index))
                    print("班级：{}: {}, 语文：{} ".format(str(GRADE_NO) + '.' + str(class_index), stu_info[index],
                                                         scores_list[0][index]))

                    hc = hc + 1  # print('rang：{}，{}'.format(self.tea_sheet.range('B2:B125').value[index],  #                           float(str(GRADE_NO) + '.' + str(class_index))))  # the_classes_stu_info.append(class_no)  #     the_classes_stu_info.append(stu_info[index])  # for i in range(SUBJECT_QUANTITY):  #     the_classes_stu_info.append((scores_list[i][index] if scores_list[i][index] is not None else 0))
            self.headcount.append(hc)
            # print(the_classes_stu_info)
            pass
        print(self.headcount)

    def __del__(self):
        self.wb.save()
        self.wb.close()
        self.app.screen_updating = True
        self.app.quit()
        print("执行析构")

    def sort(self):
        shang_name_row_begin = '2'
        shang_name_row_end = '256'

        total_name_row_begin = '237'
        total_name_row_end = '257'

        wb: xw.Book = self.app.books.open(r"C:\Projects\Python_codes\DemoSample\wife_work\高三总排名.xlsx")
        s1: xw.Sheet = wb.sheets[0]
        s1_names = s1.range('A' + total_name_row_begin + ':A' + total_name_row_end).value

        wb_shang: xw.Book = self.app.books.open(r"C:\Projects\Python_codes\DemoSample\wife_work\高三上次.xlsx")
        s2: xw.Sheet = wb_shang.sheets[0]  # 上一次表
        s2_names = s2.range('A' + shang_name_row_begin + ':A' + shang_name_row_end).value  # 人名

        print('s1 len:{},{}'.format(len(s1_names), s1_names))
        print('s2 len:{},{}'.format(len(s2_names), s2_names))

        s2_no = list()
        for i in range(len(s1_names)):
            for j in range(len(s2_names)):
                if s1_names[i].strip() == s2_names[j].strip():
                    # print(j+1)
                    s2_no.append(int(s2.range('D' + str(j + 2)).value))  # print(s2.range('B'+str(j)).value)

        print('len:{}: {}'.format(len(s2_no), s2_no))
        # s2_no.sort()
        print('len:{}: {}'.format(len(s2_no), s2_no))
        s1.range('P' + total_name_row_begin + ':P' + total_name_row_end).options(transpose=True).value = s2_no

        wb.save()
        wb_shang.close()
        wb.close()

    def cope_senior1(self, scores_tb_path, grade_no: str = '1', classes_no: str = '1', ex_list_=None):
        """
        处理高一班级统计
        :param ex_list_:
        :param grade_no:
        :param scores_tb_path:
        :param classes_no:
        :return: return_param
        """

        #############################################################
        SENIOR_CLASSES_CELL = SENIOR1_CLASSES_CELL
        SENIOR_NAME_CELL = SENIOR1_NAME_CELL
        SENIOR_TOTAL_SCORES = SENIOR1_TOTAL_SCORES
        SENIOR_CHINESE_SCORES = SENIOR1_CHINESE_SCORES
        SENIOR_MATH_SCORES = SENIOR1_MATH_SCORES
        SENIOR_ENGLISH_SCORES = SENIOR1_ENGLISH_SCORES
        SENIOR_PHYSICS_SCORES = SENIOR1_PHYSICS_SCORES
        SENIOR_CHEMISTRY_SCORES = SENIOR1_CHEMISTRY_SCORES
        SENIOR_BIOLOGY_SCORES = SENIOR1_BIOLOGY_SCORES
        SENIOR_POLITICS_SCORES = SENIOR1_POLITICS_SCORES
        SENIOR_HISTORY_SCORES = SENIOR1_HISTORY_SCORES
        SENIOR_GEOGRAPHY_SCORES = SENIOR1_GEOGRAPHY_SCORES
        SENIOR_TOTAL_SCORE = SENIOR1_TOTAL_SCORE

        if int(classes_no) > SENIOR1_MAX_CLASSES_NUM:
            print("指定班级数：{}超过最大班级数：{}！！！".format(classes_no, SENIOR1_MAX_CLASSES_NUM))
            return

        ex_list = ex_list_

        wb_score: xw.Book = self.app.books.open(scores_tb_path)
        sht_score: xw.Sheet = wb_score.sheets[0]

        print(
            'cope_senior1#############################{}.{}班开始计算：########################################'.format(
                grade_no, classes_no))
        names_list = sht_score.range(SENIOR_NAME_CELL + ROW_BEGIN + ':' + SENIOR_NAME_CELL + ROW_END).value
        print("高{}总名单 len:{},{}".format(grade_no, len(names_list), names_list))

        actual_stu_nums = len(names_list) - len(ex_list)
        print("{}年级实际计算人数（扣除例外学生）:{}".format(grade_no, actual_stu_nums))

        actual_no_nums = 0
        '''实际班总人数'''

        #######################  60%率 #########################################
        chinese_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        chinese_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        math_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        math_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        english_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        english_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        physics_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        physics_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        chemistry_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        chemistry_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        biology_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        biology_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        politics_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        politics_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        history_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        history_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        geography_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        geography_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        total_subject_pass_count = 0
        '''学生全科总成绩 >=  全科满成绩的60% 的 人数'''
        total_subject_fail_count = 0
        '''学生全科总成绩 < 全科满成绩的60% 的 人数'''
        #######################  各科总成绩 #########################################
        chinese_total_score = 0
        math_total_score = 0
        english_total_score = 0
        physics_total_score = 0
        chemistry_total_score = 0
        biology_total_score = 0
        politics_total_score = 0
        history_total_score = 0
        geography_total_score = 0

        classes_stu_names = list()
        for i in range(len(names_list)):
            classes_index = str(sht_score.range(SENIOR_CLASSES_CELL + str(i + 2)).value).strip()
            if names_list[i].strip() not in ex_list:
                if classes_index == grade_no + '.' + classes_no:
                    classes_stu_names.append(sht_score.range(SENIOR_NAME_CELL + str(i + 2)).value)
                    actual_no_nums = actual_no_nums + 1
                    #######################  60%,40%率 #########################################
                    if sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value >= PASS_RATE150:
                        chinese_sixty_score_count = chinese_sixty_score_count + 1
                    elif sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value < FAIL_RATE150:
                        chinese_forty_score_count = chinese_forty_score_count + 1

                    if sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value >= PASS_RATE150:
                        math_sixty_score_count = math_sixty_score_count + 1
                    elif sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value < FAIL_RATE150:
                        math_forty_score_count = math_forty_score_count + 1  # print('{}'.format(sht_score.range(SENIOR_NAME_CELL + str(i + 2)).value))

                    if sht_score.range(SENIOR_ENGLISH_SCORES + str(i + 2)).value >= PASS_RATE150:
                        english_sixty_score_count = english_sixty_score_count + 1
                    elif sht_score.range(SENIOR_ENGLISH_SCORES + str(i + 2)).value < FAIL_RATE150:
                        english_forty_score_count = english_forty_score_count + 1

                    if sht_score.range(SENIOR_PHYSICS_SCORES + str(i + 2)).value >= PASS_RATE100:
                        physics_sixty_score_count = physics_sixty_score_count + 1
                    elif sht_score.range(SENIOR_PHYSICS_SCORES + str(i + 2)).value < FAIL_RATE100:
                        physics_forty_score_count = physics_forty_score_count + 1

                    if sht_score.range(SENIOR_CHEMISTRY_SCORES + str(i + 2)).value >= PASS_RATE100:
                        chemistry_sixty_score_count = chemistry_sixty_score_count + 1
                    elif sht_score.range(SENIOR_CHEMISTRY_SCORES + str(i + 2)).value < FAIL_RATE100:
                        chemistry_forty_score_count = chemistry_forty_score_count + 1

                    if sht_score.range(SENIOR_BIOLOGY_SCORES + str(i + 2)).value >= PASS_RATE100:
                        biology_sixty_score_count = biology_sixty_score_count + 1
                    elif sht_score.range(SENIOR_BIOLOGY_SCORES + str(i + 2)).value < FAIL_RATE100:
                        biology_forty_score_count = biology_forty_score_count + 1

                    if sht_score.range(SENIOR_POLITICS_SCORES + str(i + 2)).value >= PASS_RATE100:
                        politics_sixty_score_count = politics_sixty_score_count + 1
                    elif sht_score.range(SENIOR_POLITICS_SCORES + str(i + 2)).value < FAIL_RATE100:
                        politics_forty_score_count = politics_forty_score_count + 1

                    if sht_score.range(SENIOR_HISTORY_SCORES + str(i + 2)).value >= PASS_RATE100:
                        history_sixty_score_count = history_sixty_score_count + 1
                    elif sht_score.range(SENIOR_HISTORY_SCORES + str(i + 2)).value < FAIL_RATE100:
                        history_forty_score_count = history_forty_score_count + 1

                    if sht_score.range(SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value >= PASS_RATE100:
                        geography_sixty_score_count = geography_sixty_score_count + 1
                    elif sht_score.range(SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value < FAIL_RATE100:
                        geography_forty_score_count = geography_forty_score_count + 1

                    #######################  各科总成绩 #########################################
                    # print(sht_score.range(SENIOR_CHINESE_SCORES[0] + str(i + 2)).value)
                    chinese_total_score = chinese_total_score + sht_score.range(
                        SENIOR_CHINESE_SCORES + str(i + 2)).value  # 语文分数
                    math_total_score = math_total_score + sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value  # 分数
                    english_total_score = english_total_score + sht_score.range(
                        SENIOR_ENGLISH_SCORES + str(i + 2)).value  # 分数
                    physics_total_score = physics_total_score + sht_score.range(
                        SENIOR_PHYSICS_SCORES + str(i + 2)).value  # 分数
                    chemistry_total_score = chemistry_total_score + sht_score.range(
                        SENIOR_CHEMISTRY_SCORES + str(i + 2)).value  # 分数
                    biology_total_score = biology_total_score + sht_score.range(
                        SENIOR_BIOLOGY_SCORES + str(i + 2)).value  # 分数
                    politics_total_score = politics_total_score + sht_score.range(
                        SENIOR_POLITICS_SCORES + str(i + 2)).value  # 分数
                    history_total_score = history_total_score + sht_score.range(
                        SENIOR_HISTORY_SCORES + str(i + 2)).value  # 分数
                    geography_total_score = geography_total_score + sht_score.range(
                        SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value  # 分数
                    #################################################班级及格率#############################

                    the_stu_total = (sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_MATH_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_ENGLISH_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_PHYSICS_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_CHEMISTRY_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_BIOLOGY_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_POLITICS_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_HISTORY_SCORES + str(i + 2)).value + sht_score.range(
                        SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value)

                    if the_stu_total >= (SENIOR_TOTAL_SCORE * 0.6):
                        total_subject_pass_count = total_subject_pass_count + 1
                    elif the_stu_total < (SENIOR_TOTAL_SCORE * 0.4):
                        total_subject_fail_count = total_subject_fail_count + 1

        wb_score.close()

        print('{}.{}班班级总人数:{},{}'.format(grade_no, classes_no, actual_no_nums, classes_stu_names))
        print('60%率=语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            (chinese_sixty_score_count / actual_no_nums) * 100, (math_sixty_score_count / actual_no_nums) * 100,
            (english_sixty_score_count / actual_no_nums) * 100, (physics_sixty_score_count / actual_no_nums) * 100,
            (chemistry_sixty_score_count / actual_no_nums) * 100, (biology_sixty_score_count / actual_no_nums) * 100,
            (politics_sixty_score_count / actual_no_nums) * 100, (history_sixty_score_count / actual_no_nums) * 100,
            (geography_sixty_score_count / actual_no_nums) * 100))

        print('40%率=语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            (chinese_forty_score_count / actual_no_nums) * 100, (math_forty_score_count / actual_no_nums) * 100,
            (english_forty_score_count / actual_no_nums) * 100, (physics_forty_score_count / actual_no_nums) * 100,
            (chemistry_forty_score_count / actual_no_nums) * 100, (biology_forty_score_count / actual_no_nums) * 100,
            (politics_forty_score_count / actual_no_nums) * 100, (history_forty_score_count / actual_no_nums) * 100,
            (geography_forty_score_count / actual_no_nums) * 100))

        print('平均分：语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            chinese_total_score / actual_no_nums, math_total_score / actual_no_nums,
            english_total_score / actual_no_nums, physics_total_score / actual_no_nums,
            chemistry_total_score / actual_no_nums, biology_total_score / actual_no_nums,
            politics_total_score / actual_no_nums, history_total_score / actual_no_nums,
            geography_total_score / actual_no_nums))

        #######################  返回参数 #########################################

        chinese_score_dict = dict()
        math_score_dict = dict()
        english_score_dict = dict()
        physics_score_dict = dict()
        chemistry_score_dict = dict()
        biology_score_dict = dict()
        politics_score_dict = dict()
        history_score_dict = dict()
        geography_score_dict = dict()
        total_subject_score = dict()

        total_subject_score['total_subject_pass_rate'] = (total_subject_pass_count / actual_no_nums) * 100
        total_subject_score['total_subject_fail_rate'] = (total_subject_fail_count / actual_no_nums) * 100

        print('全科达到几个标准人数：{}, 不及格人数：{}'.format(total_subject_pass_count, total_subject_fail_count))

        chinese_score_dict['total'] = chinese_total_score
        chinese_score_dict['stu_nums'] = actual_no_nums
        chinese_score_dict['avg'] = chinese_total_score / actual_no_nums
        chinese_score_dict['pass_rate'] = (chinese_sixty_score_count / actual_no_nums) * 100
        chinese_score_dict['fail_rate'] = (chinese_forty_score_count / actual_no_nums) * 100

        math_score_dict['total'] = math_total_score
        math_score_dict['stu_nums'] = actual_no_nums
        math_score_dict['avg'] = math_total_score / actual_no_nums
        math_score_dict['pass_rate'] = (math_sixty_score_count / actual_no_nums) * 100
        math_score_dict['fail_rate'] = (math_forty_score_count / actual_no_nums) * 100

        english_score_dict['total'] = english_total_score
        english_score_dict['stu_nums'] = actual_no_nums
        english_score_dict['avg'] = english_total_score / actual_no_nums
        english_score_dict['pass_rate'] = (english_sixty_score_count / actual_no_nums) * 100
        english_score_dict['fail_rate'] = (english_forty_score_count / actual_no_nums) * 100

        physics_score_dict['total'] = physics_total_score
        physics_score_dict['stu_nums'] = actual_no_nums
        physics_score_dict['avg'] = physics_total_score / actual_no_nums
        physics_score_dict['pass_rate'] = (physics_sixty_score_count / actual_no_nums) * 100
        physics_score_dict['fail_rate'] = (physics_forty_score_count / actual_no_nums) * 100

        chemistry_score_dict['total'] = chemistry_total_score
        chemistry_score_dict['stu_nums'] = actual_no_nums
        chemistry_score_dict['avg'] = chemistry_total_score / actual_no_nums
        chemistry_score_dict['pass_rate'] = (chemistry_sixty_score_count / actual_no_nums) * 100
        chemistry_score_dict['fail_rate'] = (chemistry_forty_score_count / actual_no_nums) * 100

        biology_score_dict['total'] = biology_total_score
        biology_score_dict['stu_nums'] = actual_no_nums
        biology_score_dict['avg'] = biology_total_score / actual_no_nums
        biology_score_dict['pass_rate'] = (biology_sixty_score_count / actual_no_nums) * 100
        biology_score_dict['fail_rate'] = (biology_forty_score_count / actual_no_nums) * 100

        politics_score_dict['total'] = politics_total_score
        politics_score_dict['stu_nums'] = actual_no_nums
        politics_score_dict['avg'] = politics_total_score / actual_no_nums
        politics_score_dict['pass_rate'] = (politics_sixty_score_count / actual_no_nums) * 100
        politics_score_dict['fail_rate'] = (politics_forty_score_count / actual_no_nums) * 100

        history_score_dict['total'] = history_total_score
        history_score_dict['stu_nums'] = actual_no_nums
        history_score_dict['avg'] = history_total_score / actual_no_nums
        history_score_dict['pass_rate'] = (history_sixty_score_count / actual_no_nums) * 100
        history_score_dict['fail_rate'] = (history_forty_score_count / actual_no_nums) * 100

        geography_score_dict['total'] = geography_total_score
        geography_score_dict['stu_nums'] = actual_no_nums
        geography_score_dict['avg'] = geography_total_score / actual_no_nums
        geography_score_dict['pass_rate'] = (geography_sixty_score_count / actual_no_nums) * 100
        geography_score_dict['fail_rate'] = (geography_forty_score_count / actual_no_nums) * 100

        return_param = dict()
        return_param['chinese'] = chinese_score_dict
        return_param['math'] = math_score_dict
        return_param['english'] = english_score_dict
        return_param['physics'] = physics_score_dict
        return_param['chemistry'] = chemistry_score_dict
        return_param['biology'] = biology_score_dict
        return_param['politics'] = politics_score_dict
        return_param['history'] = history_score_dict
        return_param['geography'] = geography_score_dict
        return_param['total_subject_rate'] = total_subject_score

        print('return_param:{}'.format(return_param))
        print('######################################################################')
        return return_param

    def cope_senior23(self, path: str, grade_no: str, classes_no: str, ex_list_=None):
        """

        :param ex_list_: 
        :param grade_no:年级
        :param classes_no: 班级号：1~8
        :param path:
        :return: return_param, category_score_dict
        """

        """处理高二，三的成绩表"""
        if grade_no.strip() == '2':
            senior_combine_category = SENIOR2_COMBINE_CATEGORY
            SENIOR_COMBINE_CELL = SENIOR2_COMBINE_CELL
            SENIOR_CLASSES_CELL = SENIOR2_CLASSES_CELL
            SENIOR_NAME_CELL = SENIOR2_NAME_CELL
            SENIOR_TOTAL_SCORES = SENIOR2_TOTAL_SCORES
            SENIOR_CHINESE_SCORES = SENIOR2_CHINESE_SCORES
            SENIOR_MATH_SCORES = SENIOR2_MATH_SCORES
            SENIOR_ENGLISH_SCORES = SENIOR2_ENGLISH_SCORES
            SENIOR_PHYSICS_SCORES = SENIOR2_PHYSICS_SCORES
            SENIOR_CHEMISTRY_SCORES = SENIOR2_CHEMISTRY_SCORES
            SENIOR_BIOLOGY_SCORES = SENIOR2_BIOLOGY_SCORES
            SENIOR_POLITICS_SCORES = SENIOR2_POLITICS_SCORES
            SENIOR_HISTORY_SCORES = SENIOR2_HISTORY_SCORES
            SENIOR_GEOGRAPHY_SCORES = SENIOR2_GEOGRAPHY_SCORES
            MAX_CLASSES_NUM = SENIOR2_MAX_CLASSES_NUM
            SENIOR_TOTAL_SCORE = SENIOR2_TOTAL_SCORE
        else:
            senior_combine_category = SENIOR3_COMBINE_CATEGORY
            SENIOR_COMBINE_CELL = SENIOR3_COMBINE_CELL
            SENIOR_CLASSES_CELL = SENIOR3_CLASSES_CELL
            SENIOR_NAME_CELL = SENIOR3_NAME_CELL
            SENIOR_TOTAL_SCORES = SENIOR3_TOTAL_SCORES
            SENIOR_CHINESE_SCORES = SENIOR3_CHINESE_SCORES
            SENIOR_MATH_SCORES = SENIOR3_MATH_SCORES
            SENIOR_ENGLISH_SCORES = SENIOR3_ENGLISH_SCORES
            SENIOR_PHYSICS_SCORES = SENIOR3_PHYSICS_SCORES
            SENIOR_CHEMISTRY_SCORES = SENIOR3_CHEMISTRY_SCORES
            SENIOR_BIOLOGY_SCORES = SENIOR3_BIOLOGY_SCORES
            SENIOR_POLITICS_SCORES = SENIOR3_POLITICS_SCORES
            SENIOR_HISTORY_SCORES = SENIOR3_HISTORY_SCORES
            SENIOR_GEOGRAPHY_SCORES = SENIOR3_GEOGRAPHY_SCORES
            MAX_CLASSES_NUM = SENIOR3_MAX_CLASSES_NUM
            SENIOR_TOTAL_SCORE = SENIOR3_TOTAL_SCORE

        if int(classes_no) > MAX_CLASSES_NUM:
            print("指定班级数：{}超过最大班级数：{}！！！".format(classes_no, MAX_CLASSES_NUM))
            return

        ex_list = ex_list_

        wb: xw.Book = self.app.books.open(path)
        sht_score = wb.sheets[int(classes_no) - 1]

        print(
            'cope_senior23#############################{}.{}班开始计算：########################################'.format(
                grade_no, classes_no))

        names_list = list(
            filter(None, sht_score.range(SENIOR_NAME_CELL + ROW_BEGIN + ':' + SENIOR_NAME_CELL + ROW_END).value))
        print("{}.{}班名单 len:{},{}".format(grade_no, classes_no, len(names_list), names_list))

        #######################  60%率 #########################################
        chinese_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        chinese_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        math_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        math_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        english_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        english_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        physics_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        physics_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        chemistry_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        chemistry_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        biology_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        biology_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        politics_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        politics_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        history_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        history_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        geography_sixty_score_count = 0
        '''每班成绩>= 60的学生个数'''
        geography_forty_score_count = 0
        '''每班成绩< 40的学生个数'''

        total_subject_pass_count = 0
        '''学生全科总成绩 >=  全科满成绩的60% 的 人数'''
        total_subject_fail_count = 0
        '''学生全科总成绩 < 全科满成绩的60% 的 人数'''
        #######################  各科总成绩 #########################################
        chinese_total_score = 0
        math_total_score = 0
        english_total_score = 0
        physics_total_score = 0
        chemistry_total_score = 0
        biology_total_score = 0
        politics_total_score = 0
        history_total_score = 0
        geography_total_score = 0

        actual_no_nums = 0
        '''实际班总人数'''

        category_score_list = list()
        category_score_dict = dict()
        '''组合-各个成绩集合'''
        classes_stu_names = list()

        for i in range(len(names_list)):
            classes_index = str(sht_score.range(SENIOR_CLASSES_CELL + str(i + 2)).value).strip()
            if names_list[i].strip() not in ex_list:
                if classes_index == grade_no + '.' + classes_no:
                    classes_stu_names.append(sht_score.range(SENIOR_NAME_CELL + str(i + 2)).value)
                    actual_no_nums = actual_no_nums + 1

                    #######################  组合-各科成绩#########################################

                    for category in senior_combine_category:
                        if sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value == category:
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_CHINESE_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_CHINESE_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_MATH_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_MATH_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_ENGLISH_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_ENGLISH_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_PHYSICS_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_PHYSICS_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_CHEMISTRY_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_CHEMISTRY_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_BIOLOGY_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_BIOLOGY_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_POLITICS_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_POLITICS_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_HISTORY_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_HISTORY_SCORES + str(i + 2)).value)
                            category_score_list.append(0 if sht_score.range(
                                SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value is None else sht_score.range(
                                SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value)

                            category_score_dict[category] = category_score_list
                    #######################  60%,40%率 #########################################

                    if sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value >= PASS_RATE150:
                        chinese_sixty_score_count = chinese_sixty_score_count + 1
                    elif sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value < FAIL_RATE150:
                        chinese_forty_score_count = chinese_forty_score_count + 1

                    if sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value >= PASS_RATE150:
                        math_sixty_score_count = math_sixty_score_count + 1
                    elif sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value < FAIL_RATE150:
                        math_forty_score_count = math_forty_score_count + 1  # print('{}'.format(sht_score.range(SENIOR_NAME_CELL + str(i + 2)).value))

                    if sht_score.range(SENIOR_ENGLISH_SCORES + str(i + 2)).value >= PASS_RATE150:
                        english_sixty_score_count = english_sixty_score_count + 1
                    elif sht_score.range(SENIOR_ENGLISH_SCORES + str(i + 2)).value < FAIL_RATE150:
                        english_forty_score_count = english_forty_score_count + 1

                    if '物' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_PHYSICS_SCORES + str(i + 2)).value >= PASS_RATE100:
                            physics_sixty_score_count = physics_sixty_score_count + 1
                        elif sht_score.range(SENIOR_PHYSICS_SCORES + str(i + 2)).value < FAIL_RATE100:
                            physics_forty_score_count = physics_forty_score_count + 1

                    if '化' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_CHEMISTRY_SCORES + str(i + 2)).value >= PASS_RATE100:
                            chemistry_sixty_score_count = chemistry_sixty_score_count + 1
                        elif sht_score.range(SENIOR_CHEMISTRY_SCORES + str(i + 2)).value < FAIL_RATE100:
                            chemistry_forty_score_count = chemistry_forty_score_count + 1

                    if '生' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_BIOLOGY_SCORES + str(i + 2)).value >= PASS_RATE100:
                            biology_sixty_score_count = biology_sixty_score_count + 1
                        elif sht_score.range(SENIOR_BIOLOGY_SCORES + str(i + 2)).value < FAIL_RATE100:
                            biology_forty_score_count = biology_forty_score_count + 1

                    if '政' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_POLITICS_SCORES + str(i + 2)).value >= PASS_RATE100:
                            politics_sixty_score_count = politics_sixty_score_count + 1
                        elif sht_score.range(SENIOR_POLITICS_SCORES + str(i + 2)).value < FAIL_RATE100:
                            politics_forty_score_count = politics_forty_score_count + 1

                    if '史' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_HISTORY_SCORES + str(i + 2)).value >= PASS_RATE100:
                            history_sixty_score_count = history_sixty_score_count + 1
                        elif sht_score.range(SENIOR_HISTORY_SCORES + str(i + 2)).value < FAIL_RATE100:
                            history_forty_score_count = history_forty_score_count + 1

                    if '地' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        if sht_score.range(SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value >= PASS_RATE100:
                            geography_sixty_score_count = geography_sixty_score_count + 1
                        elif sht_score.range(SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value < FAIL_RATE100:
                            geography_forty_score_count = geography_forty_score_count + 1

                    #######################  各科总成绩 #########################################
                    # print(sht_score.range(SENIOR_CHINESE_SCORES[0] + str(i + 2)).value)
                    chinese_total_score = chinese_total_score + sht_score.range(
                        SENIOR_CHINESE_SCORES + str(i + 2)).value  # 语文分数
                    math_total_score = math_total_score + sht_score.range(SENIOR_MATH_SCORES + str(i + 2)).value  # 分数
                    english_total_score = english_total_score + sht_score.range(
                        SENIOR_ENGLISH_SCORES + str(i + 2)).value  # 分数

                    if '物' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        physics_total_score = physics_total_score + sht_score.range(
                            SENIOR_PHYSICS_SCORES + str(i + 2)).value  # 分数

                    if '化' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        chemistry_total_score = chemistry_total_score + sht_score.range(
                            SENIOR_CHEMISTRY_SCORES + str(i + 2)).value  # 分数

                    if '生' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        biology_total_score = biology_total_score + sht_score.range(
                            SENIOR_BIOLOGY_SCORES + str(i + 2)).value  # 分数

                    if '政' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        politics_total_score = politics_total_score + sht_score.range(
                            SENIOR_POLITICS_SCORES + str(i + 2)).value  # 分数

                    if '史' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        history_total_score = history_total_score + sht_score.range(
                            SENIOR_HISTORY_SCORES + str(i + 2)).value  # 分数

                    if '地' in sht_score.range(SENIOR_COMBINE_CELL + str(i + 2)).value:
                        geography_total_score = geography_total_score + sht_score.range(
                            SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value  # 分数

                        #################################################班级及格率#############################

                        # the_stu_total = (sht_score.range(SENIOR_CHINESE_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_MATH_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_ENGLISH_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_PHYSICS_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_CHEMISTRY_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_BIOLOGY_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_POLITICS_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_HISTORY_SCORES + str(i + 2)).value + sht_score.range(  #     SENIOR_GEOGRAPHY_SCORES + str(i + 2)).value)  #  # if the_stu_total >= (SENIOR_TOTAL_SCORE * 0.6):  #     total_subject_pass_count = total_subject_pass_count + 1  # elif the_stu_total < (SENIOR_TOTAL_SCORE * 0.4):  #     total_subject_fail_count = total_subject_fail_count + 1

        wb.close()

        print('{}.{}班级实际人数(扣去去除):{},{}'.format(grade_no, classes_no, actual_no_nums, classes_stu_names))
        print('60%率=语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            (chinese_sixty_score_count / actual_no_nums) * 100, (math_sixty_score_count / actual_no_nums) * 100,
            (english_sixty_score_count / actual_no_nums) * 100, (physics_sixty_score_count / actual_no_nums) * 100,
            (chemistry_sixty_score_count / actual_no_nums) * 100, (biology_sixty_score_count / actual_no_nums) * 100,
            (politics_sixty_score_count / actual_no_nums) * 100, (history_sixty_score_count / actual_no_nums) * 100,
            (geography_sixty_score_count / actual_no_nums) * 100))

        print('40%率=语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            (chinese_forty_score_count / actual_no_nums) * 100, (math_forty_score_count / actual_no_nums) * 100,
            (english_forty_score_count / actual_no_nums) * 100, (physics_forty_score_count / actual_no_nums) * 100,
            (chemistry_forty_score_count / actual_no_nums) * 100, (biology_forty_score_count / actual_no_nums) * 100,
            (politics_forty_score_count / actual_no_nums) * 100, (history_forty_score_count / actual_no_nums) * 100,
            (geography_forty_score_count / actual_no_nums) * 100))

        print('平均分：语文:{},数：{},外:{},物:{},化:{},生:{},政:{},历:{},地:{}'.format(
            chinese_total_score / actual_no_nums, math_total_score / actual_no_nums,
            english_total_score / actual_no_nums, physics_total_score / actual_no_nums,
            chemistry_total_score / actual_no_nums, biology_total_score / actual_no_nums,
            politics_total_score / actual_no_nums, history_total_score / actual_no_nums,
            geography_total_score / actual_no_nums))

        print('######################################################################')
        print('组合-各科成绩：{}'.format(category_score_dict))
        #######################  返回参数 #########################################

        chinese_score_dict = dict()
        math_score_dict = dict()
        english_score_dict = dict()
        physics_score_dict = dict()
        chemistry_score_dict = dict()
        biology_score_dict = dict()
        politics_score_dict = dict()
        history_score_dict = dict()
        geography_score_dict = dict()

        chinese_score_dict['total'] = chinese_total_score
        chinese_score_dict['stu_nums'] = actual_no_nums
        chinese_score_dict['avg'] = chinese_total_score / actual_no_nums
        chinese_score_dict['pass_rate'] = (chinese_sixty_score_count / actual_no_nums) * 100
        chinese_score_dict['fail_rate'] = (chinese_forty_score_count / actual_no_nums) * 100

        math_score_dict['total'] = math_total_score
        math_score_dict['stu_nums'] = actual_no_nums
        math_score_dict['avg'] = math_total_score / actual_no_nums
        math_score_dict['pass_rate'] = (math_sixty_score_count / actual_no_nums) * 100
        math_score_dict['fail_rate'] = (math_forty_score_count / actual_no_nums) * 100

        english_score_dict['total'] = english_total_score
        english_score_dict['stu_nums'] = actual_no_nums
        english_score_dict['avg'] = english_total_score / actual_no_nums
        english_score_dict['pass_rate'] = (english_sixty_score_count / actual_no_nums) * 100
        english_score_dict['fail_rate'] = (english_forty_score_count / actual_no_nums) * 100

        physics_score_dict['total'] = physics_total_score
        physics_score_dict['stu_nums'] = actual_no_nums
        physics_score_dict['avg'] = physics_total_score / actual_no_nums
        physics_score_dict['pass_rate'] = (physics_sixty_score_count / actual_no_nums) * 100
        physics_score_dict['fail_rate'] = (physics_forty_score_count / actual_no_nums) * 100

        chemistry_score_dict['total'] = chemistry_total_score
        chemistry_score_dict['stu_nums'] = actual_no_nums
        chemistry_score_dict['avg'] = chemistry_total_score / actual_no_nums
        chemistry_score_dict['pass_rate'] = (chemistry_sixty_score_count / actual_no_nums) * 100
        chemistry_score_dict['fail_rate'] = (chemistry_forty_score_count / actual_no_nums) * 100

        biology_score_dict['total'] = biology_total_score
        biology_score_dict['stu_nums'] = actual_no_nums
        biology_score_dict['avg'] = biology_total_score / actual_no_nums
        biology_score_dict['pass_rate'] = (biology_sixty_score_count / actual_no_nums) * 100
        biology_score_dict['fail_rate'] = (biology_forty_score_count / actual_no_nums) * 100

        politics_score_dict['total'] = politics_total_score
        politics_score_dict['stu_nums'] = actual_no_nums
        politics_score_dict['avg'] = politics_total_score / actual_no_nums
        politics_score_dict['pass_rate'] = (politics_sixty_score_count / actual_no_nums) * 100
        politics_score_dict['fail_rate'] = (politics_forty_score_count / actual_no_nums) * 100

        history_score_dict['total'] = history_total_score
        history_score_dict['stu_nums'] = actual_no_nums
        history_score_dict['avg'] = history_total_score / actual_no_nums
        history_score_dict['pass_rate'] = (history_sixty_score_count / actual_no_nums) * 100
        history_score_dict['fail_rate'] = (history_forty_score_count / actual_no_nums) * 100

        geography_score_dict['total'] = geography_total_score
        geography_score_dict['stu_nums'] = actual_no_nums
        geography_score_dict['avg'] = geography_total_score / actual_no_nums
        geography_score_dict['pass_rate'] = (geography_sixty_score_count / actual_no_nums) * 100
        geography_score_dict['fail_rate'] = (geography_forty_score_count / actual_no_nums) * 100

        return_param = dict()
        return_param['chinese'] = chinese_score_dict
        return_param['math'] = math_score_dict
        return_param['english'] = english_score_dict
        return_param['physics'] = physics_score_dict
        return_param['chemistry'] = chemistry_score_dict
        return_param['biology'] = biology_score_dict
        return_param['politics'] = politics_score_dict
        return_param['history'] = history_score_dict
        return_param['geography'] = geography_score_dict

        return return_param, category_score_dict

    def cope_exclude_table(self, exclude_tb_path, grade_no: str = '1', cope_type: int = 0):
        """
        :param exclude_tb_path: 扣人表路径
        :param grade_no: 指定年级 1~3
        :param cope_type: 类型：获取扣人的姓名(0)，班级(1),全部(3)
        :return: type=0,1: list() ;  type=3: ; list(list(班级），list(name))
        """

        #############################################################

        wb_ex: xw.Book = self.app.books.open(exclude_tb_path)
        sht_ex: xw.Sheet = wb_ex.sheets[0]

        if grade_no == '1':
            if cope_type == 0:
                return_list = list(sht_ex.range(EX_SENIOR1_NAME_CELL).value)  # 去重
                '''高一扣人名单'''
                print("高一扣人名单 len:{},{}".format(len(return_list), return_list))
            elif cope_type == 1:
                return_list = list(sht_ex.range(EX_SENIOR1_CLASSES_CELL).value)  # 去重
                '''高一扣人名单'''
                print("高一扣人名单 len:{},{}".format(len(return_list), return_list))
            else:
                return_list = list(sht_ex.range(EX_SENIOR1_CLASSES_CELL).value)  # 去重
                return_list.append(list(sht_ex.range(EX_SENIOR1_NAME_CELL).value))
                '''高一扣人名单'''
                print("高一扣人名单 len:{},{}".format(len(return_list), return_list))

        elif grade_no == '2':
            if cope_type == 0:
                return_list = list(sht_ex.range(EX_SENIOR2_NAME_CELL).value)  # 去重
                '''高2扣人名单'''
                print("高2扣人名单 len:{},{}".format(len(return_list), return_list))
            elif cope_type == 1:
                return_list = list(sht_ex.range(EX_SENIOR2_CLASSES_CELL).value)  # 去重
                '''高2扣人名单'''
                print("高2扣人名单 len:{},{}".format(len(return_list), return_list))
            else:
                return_list = list(sht_ex.range(EX_SENIOR2_CLASSES_CELL).value)  # 去重
                return_list.append(list(sht_ex.range(EX_SENIOR2_NAME_CELL).value))
                '''高2扣人名单'''
                print("高2扣人名单 len:{},{}".format(len(return_list), return_list))

        else:
            if cope_type == 0:
                return_list = list(sht_ex.range(EX_SENIOR3_NAME_CELL).value)  # 去重
                '''高3扣人名单'''
                print("高3扣人名单 len:{},{}".format(len(return_list), return_list))
            elif cope_type == 1:
                return_list = list(sht_ex.range(EX_SENIOR3_CLASSES_CELL).value)  # 去重
                '''高3扣人名单'''
                print("高3扣人名单 len:{},{}".format(len(return_list), return_list))
            else:
                return_list = list(sht_ex.range(EX_SENIOR3_CLASSES_CELL).value)  # 去重
                return_list.append(list(sht_ex.range(EX_SENIOR3_NAME_CELL).value))
                '''高3扣人名单'''
                print("高3扣人名单 len:{},{}".format(len(return_list), return_list))

        wb_ex.close()
        return return_list

    def calcu_senior_1(self, ex_list):

        wb: xw.Book = self.app.books.open(self.statistical_table_examination_results)
        sht: xw.Sheet = wb.sheets[0]
        sht.autofit()

        s1_class_score_info = list()
        '''高x总成绩excel全路径'''

        for index in range(1, SENIOR1_MAX_CLASSES_NUM + 1):
            s1_class_score_info.append(self.cope_senior1(self.scores_path, '1', str(index), ex_list))

        # 写入excel
        for i in range(SENIOR1_MAX_CLASSES_NUM):
            sht.range(SENIOR_CHINESE_BEGIN_COL_NO + str(i + 5)).value = [s1_class_score_info[i]['chinese']['avg'], '',
                                                                         s1_class_score_info[i]['chinese']['pass_rate'],
                                                                         s1_class_score_info[i]['chinese']['fail_rate']]

            sht.range('I' + str(i + 5)).value = [s1_class_score_info[i]['math']['avg'], '',
                                                 s1_class_score_info[i]['math']['pass_rate'],
                                                 s1_class_score_info[i]['math']['fail_rate']]

            sht.range('O' + str(i + 5)).value = [s1_class_score_info[i]['english']['avg'], '',
                                                 s1_class_score_info[i]['english']['pass_rate'],
                                                 s1_class_score_info[i]['english']['fail_rate']]

            sht.range('C' + str(i + 13)).value = [s1_class_score_info[i]['physics']['avg'], '',
                                                  s1_class_score_info[i]['physics']['pass_rate'],
                                                  s1_class_score_info[i]['physics']['fail_rate']]

            sht.range('I' + str(i + 13)).value = [s1_class_score_info[i]['history']['avg'], '',
                                                  s1_class_score_info[i]['history']['pass_rate'],
                                                  s1_class_score_info[i]['history']['fail_rate']]

            sht.range('O' + str(i + 13)).value = [s1_class_score_info[i]['chemistry']['avg'], '',
                                                  s1_class_score_info[i]['chemistry']['pass_rate'],
                                                  s1_class_score_info[i]['chemistry']['fail_rate']]

            sht.range('C' + str(i + 21)).value = [s1_class_score_info[i]['biology']['avg'], '',
                                                  s1_class_score_info[i]['biology']['pass_rate'],
                                                  s1_class_score_info[i]['biology']['fail_rate']]

            sht.range('I' + str(i + 21)).value = [s1_class_score_info[i]['politics']['avg'], '',
                                                  s1_class_score_info[i]['politics']['pass_rate'],
                                                  s1_class_score_info[i]['politics']['fail_rate']]

            sht.range('O' + str(i + 21)).value = [s1_class_score_info[i]['geography']['avg'], '',
                                                  s1_class_score_info[i]['geography']['pass_rate'],
                                                  s1_class_score_info[i]['geography']['fail_rate']]

            # 总分平均分 =  班级每人总分 / 班级总人数
            avg_rate = (s1_class_score_info[i]['chinese']['total'] + s1_class_score_info[i]['math']['total'] +
                        s1_class_score_info[i]['english']['total'] + s1_class_score_info[i]['physics']['total'] +
                        s1_class_score_info[i]['history']['total'] + s1_class_score_info[i]['chemistry']['total'] +
                        s1_class_score_info[i]['biology']['total'] + s1_class_score_info[i]['politics']['total'] +
                        s1_class_score_info[i]['geography']['total']) / s1_class_score_info[i]['chinese']['stu_nums']

            # 班级及格率 = 学生个数（学生总分 > 全科满分*0.6） / 总人数
            # 班级总成绩统计
            sht.range('C' + str(i + 29)).value = [avg_rate, '', s1_class_score_info[i]['total_subject_rate'][
                'total_subject_pass_rate'], s1_class_score_info[i]['total_subject_rate']['total_subject_fail_rate']]

        wb.save()
        wb.close()
        pass

    def calcu_senior_23(self, grade_no: str, ex_list):

        s2_class_score_info = list()
        s2_category_score_info = list()

        if grade_no == '2':
            SENIOR_MAX_CLASSES_NUM = SENIOR2_MAX_CLASSES_NUM
        else:
            SENIOR_MAX_CLASSES_NUM = SENIOR3_MAX_CLASSES_NUM

        for index in range(1, SENIOR_MAX_CLASSES_NUM + 1):
            return_param, category_score_dict = self.cope_senior23(self.scores_path, grade_no, str(index), ex_list)
            s2_class_score_info.append(return_param)
            s2_category_score_info.append(category_score_dict)

            print('{}.{}班各科成绩,len:{},{}：'.format(grade_no, index, len(return_param), return_param))
            print('{}.{}班组合-成绩,len:{},{}：'.format(grade_no, index, len(category_score_dict), category_score_dict))

        wb: xw.Book = self.app.books.open(self.statistical_table_examination_results)
        sht1: xw.Sheet = wb.sheets[int(grade_no) - 1]
        sht1.autofit()

        for i in range(SENIOR_MAX_CLASSES_NUM):
            sht1.range(SENIOR_CHINESE_BEGIN_COL_NO + str(i + 5)).value = [s2_class_score_info[i]['chinese']['avg'], '',
                                                                          s2_class_score_info[i]['chinese'][
                                                                              'pass_rate'],
                                                                          s2_class_score_info[i]['chinese'][
                                                                              'fail_rate']]

            sht1.range('I' + str(i + 5)).value = [s2_class_score_info[i]['math']['avg'], '',
                                                  s2_class_score_info[i]['math']['pass_rate'],
                                                  s2_class_score_info[i]['math']['fail_rate']]

            sht1.range('O' + str(i + 5)).value = [s2_class_score_info[i]['english']['avg'], '',
                                                  s2_class_score_info[i]['english']['pass_rate'],
                                                  s2_class_score_info[i]['english']['fail_rate']]

            sht1.range('C' + str(i + 14)).value = [s2_class_score_info[i]['physics']['avg'], '',
                                                   s2_class_score_info[i]['physics']['pass_rate'],
                                                   s2_class_score_info[i]['physics']['fail_rate']]

            sht1.range('I' + str(i + 14)).value = [s2_class_score_info[i]['history']['avg'], '',
                                                   s2_class_score_info[i]['history']['pass_rate'],
                                                   s2_class_score_info[i]['history']['fail_rate']]

            sht1.range('O' + str(i + 14)).value = [s2_class_score_info[i]['chemistry']['avg'], '',
                                                   s2_class_score_info[i]['chemistry']['pass_rate'],
                                                   s2_class_score_info[i]['chemistry']['fail_rate']]

            sht1.range('C' + str(i + 23)).value = [s2_class_score_info[i]['biology']['avg'], '',
                                                   s2_class_score_info[i]['biology']['pass_rate'],
                                                   s2_class_score_info[i]['biology']['fail_rate']]

            sht1.range('I' + str(i + 23)).value = [s2_class_score_info[i]['politics']['avg'], '',
                                                   s2_class_score_info[i]['politics']['pass_rate'],
                                                   s2_class_score_info[i]['politics']['fail_rate']]

            sht1.range('O' + str(i + 23)).value = [s2_class_score_info[i]['geography']['avg'], '',
                                                   s2_class_score_info[i]['geography']['pass_rate'],
                                                   s2_class_score_info[i]['geography']['fail_rate']]

            # 班级总成绩统计

        wb.save()
        wb.close()


SENIOR_CHINESE_BEGIN_COL_NO = 'C'
'''语文开始列号'''

'''考试成绩统计表：高一各班语文平均分'''


def main():
    grade = Grade(r"C:\Projects\VPF\成绩\新.xls")
    # grade.statistical_table_examination_results = r'C:\Projects\Python_codes\DemoSample\wife_work\考试成绩统计表（模板）.xlsx'
    # grade.scores_path = r'C:\Projects\Python_codes\DemoSample\wife_work\一中_高一总成绩.xls'
    # grade.ex_path = r'C:\Projects\Python_codes\DemoSample\wife_work\缺考.xlsx'

    # grade.calcu_senior_1(ex_list)

    # ex_list = grade.cope_exclude_table(grade.ex_path, '2')
    # grade.scores_path = r'C:\Projects\Python_codes\DemoSample\wife_work\一中_高二各班成绩.xls'
    # grade.calcu_senior_23('2', ex_list)
    # ex_list = grade.cope_exclude_table(grade.ex_path, '3')
    # grade.scores_path = r'C:\Projects\Python_codes\DemoSample\wife_work\一中_高三各班成绩.xls'
    # grade.calcu_senior_23('3', ex_list)
    # grade.grade_statistic('3', 6)

    # classes_info = grade.grade_statistic('1', 5)
    # classes_avg = list()
    # for i in range(len(classes_info)):
    #     classes_avg.append((classes_info[i]['chinese']['total'] + classes_info[i]['math']['total'] +
    #                         classes_info[i]['english']['total'] + classes_info[i]['physics']['total'] +
    #                         classes_info[i]['chemistry']['total'] + classes_info[i]['biology']['total'] +
    #                         classes_info[i]['politics']['total'] + classes_info[i]['history']['total'] +
    #                         classes_info[i]['geography']['total']) / classes_info[i]['chinese']['stu_nums'])
    #     print('{}班,人数：{},总分{}：'.format(i + 1, classes_info[i]['chinese']['stu_nums'], classes_avg[i]))
    #

    grade.sort()
    pass


if __name__ == '__main__':
    main()
