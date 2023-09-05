import xlwings as xw

Flag: bool = 0
'''是否区分文理科，Flag=0，标识高一不区分文理，Flag=1 表示高二三，区分文理 '''

'''常量'''
COL_NAME = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
            'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
            'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']

# 理科类科目
Science_class = ['姓名', '班级', '类别', '总分', '等级总分', '语文分数', '英语分数', '数学分数', '物理分数', '化学分数',
                 '生物分数', '政治分数', '地理分数']

# 文科类
Liberal_arts = []


# 班级数
# ClassNum = 7


class DealScores:
    def __init__(self, file_path, sheet_name):
        self.classNum = 0
        self.col_index_list = None
        self.file_path = file_path
        self.target_list = [];

        self.app = xw.App(visible=False, add_book=False)
        self.app.screen_updating = False  # 加速脚本运行速度
        self.wb: xw.Book = self.app.books.open(file_path)
        self.score_sheet: xw.Sheet = self.wb.sheets[sheet_name]
        '''总分名单'''
        self.score_sheet.autofit()

        # self.max_row_num, self.max_col_num = self.get_max_row_column()

    def __del__(self):
        self.wb.close()
        self.app.quit()

    def get_matched_data(self, target_text: list):
        """
        获取所有行数，包含指定列名的属性值集合:
        :param target_text:
        :return:[[姓名1，班级，数学分数，...,],[姓名2,...]]
        """

        self.target_list = target_text

        # 连接到 Excel 应用程序
        worksheet = self.score_sheet

        # 获取指定列的所有数据
        self.col_index_list = []
        column_names = worksheet.range('A1').expand('right').value

        for name_ in target_text:
            for i in range(len(column_names)):
                if column_names[i] == name_:
                    self.col_index_list.append(i)

        # print(self.col_index_list)
        rows = worksheet.used_range.value
        row_datas = []
        for row in rows:
            row_data = []
            for col_index in self.col_index_list:
                row_data.append(row[col_index])
            row_datas.append(row_data)
        row_datas.pop(0)
        # print(row_datas)
        # print(len(row_datas))

        return row_datas

    def get_max_row_column(self):
        """
        获取当前工作表指定sheet中存在数据最大范围:
        注意：默认以A1横纵扩展，计数连续的行列数，若存在跳格，则此后不计入数，所以保证A1所在行列，必须是包含所有数据的区域
        :return:(row_num,column_num)
        """
        # 打开工作簿并选择工作表
        sheet = self.score_sheet

        # 选择起始单元格并获取最大行数
        start_cell = sheet.range('A1')
        max_row = start_cell.end('down').row

        # 选择起始单元格并获取最大列数
        max_column = start_cell.end('right').column

        return max_row, max_column

    def sorted_data(self, data, order_index):
        """
        按照指定元素降序排列，返回排列后的数据
        :param data:
        :param order_index:排序字段：根据target_text list 中 指定元素进行逆序排序，要索引从0开始，要求排序根据必须是数值型
        :return:
        """
        items = []
        # x[4]表示使用子列表中的第4个元素（等级总分）进行排序。
        sorted_data = sorted(data, key=lambda x: x[order_index], reverse=True)
        for item in sorted_data:
            items.append(item)

        return items

    def Split_class(self, data):
        """
        保存到总表中
        :param data:排序后的数据
        :return: 返回每个班级的，根据指定分数逆序的字典
        """

        # 创建一个字典，用于存储不同班级的列表
        class_dict = {}

        # 遍历 class_datas
        for data in data:
            class_name = data[1].strip()  # 班级名称在索引 1 处

            if class_name is None:
                print("出错，班级存在None")
                return

            # 检查班级是否已存在于字典中，如果不存在，则创建一个空列表
            if class_name not in class_dict:
                class_dict[class_name] = []

            # 将数据添加到对应班级的列表中
            class_dict[class_name].append(data)

        self.classNum = len(class_dict)

        # return class_dict
        # 打印每个班级的列表
        for class_name, data_list in class_dict.items():
            # print(f"{class_name}: {data_list}")
            self.saveClassInfo(data_list)

    def saveClassInfo(self, class_datas):
        """
        为王待续
        :param class_datas:
        :return:
        """
        print(class_datas)
        app = xw.App(visible=False)
        workbook = app.books.add()
        for i in range(0, self.classNum):
            '''多少个班级存储为多少个sheet'''
            sheet = workbook.sheets[i]

            # 写入数据
            for j, row in enumerate(class_datas, start=1):
                for col_ in range(len(self.target_list)):
                    print(f"{COL_NAME[col_] + str(col_+1)}")
                    print(row[col_])
                    sheet.range(f"{COL_NAME[col_] + str(col_+1)}").value = row[col_]  # 班级名称

                # 保存文件
            workbook.save(path=f"C:\\{i + 1}班总成绩(未扣人).xlsx")
            workbook.close()
        app.quit()


def main():
    temp_list = DealScores(r"C:\Projects\老婆\成绩\普1中_物理类总成绩.xls", '总分')
    data = temp_list.get_matched_data(Science_class);
    # print(data)
    data = temp_list.sorted_data(data, (Science_class.index('总分') + 1))
    # print(data)
    # Split_class
    temp_list.Split_class(data)


if __name__ == '__main__':
    main()
