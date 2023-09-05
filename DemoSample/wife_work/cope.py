import xlwings as xw

# 分科标志
GROUP_INDICATION = True

# 高2分科集合
SENIOR2_COMBINE_CATEGORY = ['物化政', '物化地', '物生政', '物政地', '物生地', '历政地']
# 高三分科集合
SENIOR3_COMBINE_CATEGORY = ['物生政', '物生地', '物政地', '历政地', '物化生', '物化政']
# 姓名	班级	语文	数学	英语	物理	历史	化学	生物	政治	地理	总分	班名	校名次(若分科则为这个科内排名)
SENIOR3_COMBINE_POSITION = ['B', 'C', 'M', 'Q', 'U', 'AC', 'AG', 'AK', 'AS', 'BA', 'BI', 'E', 'H', 'G']


class CopeExcel:
    def __init__(self, source_path=r"C:\Projects\老婆\普兰店一中-3各班级成绩.xls", grade=3):
        self.app = xw.App(visible=False, add_book=False)
        self.app.screen_updating = False  # 加速脚本运行速度

        self.wb: xw.Book = self.app.books.open(source_path)
        # self.score_sheet: xw.Sheet = self.wb.sheets[0]
        self.sheet_count = len(self.wb.sheets)
        print("一共{}个sheet".format(self.sheet_count))

        if grade == 3:
            self.category_ls = SENIOR3_COMBINE_CATEGORY
        else:
            self.category_ls = SENIOR2_COMBINE_CATEGORY

        # 每个班的所有信息
        self.grade_infos_ls: list = []

    def getBranchInfo(self):
        """
        获取分科的信息
        :return:
        """
        if GROUP_INDICATION:
            # 分科,遍历分科
            # for index in range(len(self.category_ls)):

            all_grade_info = []
            '''所有班级的信息'''

            # 遍历所有班级
            # for grade in range(5, 6):
            for grade in range(self.sheet_count):
                # self.wb.sheets[grade].autofit()
                # self.grade_infos_ls = self.wb.sheets[grade].range('B2:BP44').value
                # print(self.grade_infos_ls)
                # 获取有数据的所有行数和列数
                row_col_num_tuple: tuple = self.wb.sheets[grade].used_range.shape
                print('(总行数,列数)={}'.format(row_col_num_tuple))

                every_grade_info = []
                '''每个班的所有信息 '''
                # 遍历行
                for row in range(2, row_col_num_tuple[0] + 1):
                    row_data = []
                    for subject in SENIOR3_COMBINE_POSITION:

                        if not str(self.wb.sheets[grade].range('A' + str(row)).value).isnumeric():
                            break

                        row_data.append(self.wb.sheets[grade].range(subject + str(row)).value)  # print(row_data)

                    if len(row_data) < 2:
                        continue

                    every_grade_info.append(
                        row_data)  # self.grade_infos_ls.append(self.wb.sheets[grade].range(subject + str(row)))
                print('班级{}总信息{},人数{}'.format(grade + 1, every_grade_info, len(every_grade_info)))
            all_grade_info.append(every_grade_info)
            pass
        else:
            pass


if __name__ == '__main__':
    test = CopeExcel()
    test.getBranchInfo()
