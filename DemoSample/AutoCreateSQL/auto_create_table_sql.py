"""
根据excel配置字段,自动生成sql语句

"""

import os

import xlwings as xw


def filter_func(n):
    """
    过滤函数
    :param n:
    :return:
    """
    if n is None or n == ' ' or n == '':
        return False
    else:
        return True


class GeneralOracleDDL(object):
    def __init__(self, source_path_, scope_row: str, scope_col: str, encoding='utf-8'):

        self.encode = encoding
        '''字符编码'''
        self.default = []
        '''默认值'''
        self.isPrimaryKey = []
        '''主键标识'''
        self.comment = []
        '''注释'''
        self.data_type = []
        '''数据类型'''
        self.col_name = []
        '''字段名'''
        self.primarys = []
        '''主键集合'''

        self.file_path = source_path_
        self.app = xw.App(visible=False, add_book=False)
        self.wb: xw.Book = self.app.books.open(source_path_)
        self.scope = scope_row + ':' + scope_col
        self.col_nums = int(scope_row[1:])
        self.row_nums = int(scope_col[1:])
        self.table_names = []
        print('scope: {}, rows:{}, cols:{}'.format(self.scope, self.row_nums, self.col_nums))
        for i in range(self.wb.sheets.count):
            self.table_names.append(self.wb.sheets[i].name)
            print('sheet count: {}, sheet name: {}'.format(self.wb.sheets.count, self.table_names))

    def __del__(self):
        self.wb.close()
        self.app.quit()
        print("执行析构")

    def fetch_col_value(self):
        # 每个sheet
        for i in range(len(self.table_names)):
            # 获取初始列数据清洗
            column_names = list(filter(filter_func, self.wb.sheets[i].range('A2:A' + str(self.row_nums)).value))
            data_type = list(filter(filter_func, self.wb.sheets[i].range('B2:B' + str(self.row_nums)).value))
            # default = list(filter(filter_func, self.wb.sheets[i].range('C2:C' + str(self.row_nums)).value))
            default = self.wb.sheets[i].range('C2:C' + str(self.row_nums)).value
            comment = self.wb.sheets[i].range('D2:D' + str(self.row_nums)).value
            isPrimaryKey = self.wb.sheets[i].range('E2:E' + str(self.row_nums)).value
            pri_key = []

            if len(column_names) != len(data_type):
                print('此三列:column_name,	data_type,	default 数量不一致~,退出')
                return

            # 修饰组数据
            for j in range(len(column_names)):
                column_names[j] = str(column_names[j]).strip()
                data_type[j] = str(data_type[j]).strip()
                default[j] = str(default[j]).strip().replace('None', '')
                comment[j] = str(comment[j]).strip().replace('\n', '').replace('None', '')
                isPrimaryKey[j] = str(isPrimaryKey[j]).strip().replace('None', '')

            for k in range(len(column_names)):
                if isPrimaryKey[k] != '':
                    pri_key.append(column_names[k])

            self.primarys.append(pri_key)
            self.col_name.append(column_names)
            self.default.append(default)
            self.data_type.append(data_type)
            self.comment.append(comment)
            self.isPrimaryKey.append(isPrimaryKey)
            print(
                '第{}张表: column_names字段:{}\n,data_type:{}\n,default:{}\n,comment:{}\n,isPrimaryKey:{}\n'.format(
                    i + 1,
                    column_names,
                    data_type,
                    default,
                    comment,
                    isPrimaryKey))

    def generalize_ddl(self):
        self.fetch_col_value()
        # print('***************{}'.format(self.primarys))
        for table_index in range(len(self.table_names)):
            # 判断是否为无效表 有sheet但是内容为空
            if len(self.col_name[table_index]) < 1:
                continue

            sql = 'create table {} ( '.format(self.table_names[table_index])
            comment_sql = ''
            primary_key_sql = 'alter table {}  add constraint pk_{}  primary key  ({});'.format(
                self.table_names[table_index], self.table_names[table_index],
                str(self.primarys[table_index]).replace('[', '').replace(']', '').replace("'", ''))

            for col_index in range(len(self.col_name[table_index])):
                sql += str(self.col_name[table_index][col_index])
                sql += ' '
                sql += str(self.data_type[table_index][col_index])
                sql += ' '
                sql += str(self.default[table_index][col_index])
                if col_index != (len(self.col_name[table_index]) - 1):
                    sql += ','
                    sql += '\n'

                comment_sql += 'comment on column '
                comment_sql += str(self.table_names[table_index]) + '.' + str(self.col_name[table_index][col_index])
                comment_sql += ' is ' + "'" + str(self.comment[table_index][col_index]) + "'"
                comment_sql += ';'
                comment_sql += '\n'

            sql += ');'
            total_sql = sql + '\n' + comment_sql + '\n' + primary_key_sql + '\n' + 'commit;'
            print(sql)
            print(comment_sql)
            print(primary_key_sql)

            self.general_file(str(self.table_names[table_index]), total_sql)

    def general_file(self, table_name: str, ddl: str):
        dir = os.path.dirname(self.file_path)
        sql_file_path = dir + os.sep + table_name + '.sql'
        print('生成文件位置:{}'.format(sql_file_path))
        with open(sql_file_path, 'w', encoding=self.encode) as f:
            f.write(ddl)
            f.flush()


if __name__ == '__main__':
    # autoSQL = GeneralOracleDDL(r"C:\Projects\Python_codes\DemoSample\AutoCreateSQL\table_fields.xlsx", 'A2', 'E94')
    autoSQL = GeneralOracleDDL(r"table_fields.xlsx", 'A2', 'E94')
    autoSQL.generalize_ddl()
    # autoSQL.fetch_col_value()
