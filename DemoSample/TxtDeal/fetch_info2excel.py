import os
import re

import xlwings as xw


def read_files(directory, items):
    """
    读取指定目录下所有文件内容
    :param directory:
    :return:
    """
    try:
        datas = {}
        for item in items:
            datas[item] = []

        # 检查目录是否存在
        if not os.path.isdir(directory):
            raise ValueError("指定的目录不存在")

        # 获取目录下的所有文件
        files = os.listdir(directory)

        print(files)
        # 遍历每个文件
        for file in files:
            # 拼接文件路径
            file_path = os.path.join(directory, file)

            # 检查文件是否存在且是文件而不是文件夹
            if os.path.isfile(file_path):
                try:
                    # 打开文件并读取内容
                    with open(file_path, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                        for line in lines:
                            # 处理文件内容，这里只打印内容
                            for item in items:
                                if item in line:
                                    val = extract_field_value(line.strip(','), item)
                                    if val is not None:
                                        datas[item].append(val)
                except OSError as e:
                    print(f"无法读取文件 '{file}': {str(e)}")

        for key in datas.keys():
            print('摘取字段:{0}， 个数:{1}'.format(key, len(datas[key])))
        return datas

    except ValueError as e:
        print("读取文件失败:" + str(e))


def extract_field_value(text, field):
    # pattern = r"\b" + re.escape(field) + r"\S*=\s*([\d.\S]+)"
    pattern = r"[^_]" + re.escape(field) + r"\s*=[\s]*([\d.\S]*)"
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    else:
        return None


def readFile(filename, items):
    datas = {}
    for item in items:
        datas[item] = []
    # 读取文件并提取匹配的字段
    # with open(filename, 'r', encoding='utf-8') as file:
    with open(filename, 'r', encoding='ansi') as file:
        lines = file.readlines()
        for line in lines:
            for item in items:
                if item in line:
                    val = extract_field_value(line.strip(','), item)
                    # if val is not None:
                    datas[item].append(val)
    # print(datas)  # for item in items:

    for key in datas.keys():
        print('摘取字段:{0}， 个数:{1}'.format(key, len(datas[key])))
    return datas


def write_dict_to_excel(data_dict, filename):
    # 打开 Excel 应用程序
    app = xw.App(visible=False, add_book=False)
    # 创建一个新的工作簿
    wb = app.books.add()

    # 选择第一个工作表
    sheet = wb.sheets[0]

    # 写入标题
    keys = list(data_dict.keys())
    for i, key in enumerate(keys):
        sheet.range((1, i + 1)).value = key

    # 写入数据
    max_len = max(len(data_dict[key]) for key in keys)
    for i in range(max_len):
        for j, key in enumerate(keys):
            values = data_dict[key]
            if i < len(values):
                sheet.range((i + 2, j + 1)).value = values[i]

    # 保存工作簿
    wb.save(filename)

    # 关闭工作簿和 Excel 应用程序
    wb.close()
    app.quit()


if __name__ == '__main__':
    # 获取命令行参数
    # args = sys.argv

    # 第一个元素是脚本的名称，忽略它
    # script_name = args[0]

    # src_path = args[1]
    src_path = r'C:\Repository\Nercar\ShaGang\Log分析'
    items = ['Decode_steelGrade', 'thkCold', 'widCold', 'widCle', 'yieldClr','tenPrRf','tenClrRf','gradeFam']
    # items = []
    # 后续元素是传递给脚本的参数
    # 参数索引从 1 开始
    # for i in range(2, len(args)):
    #     arg = args[i]
    #     items.append(arg)
    #     print('Argument {}: {}'.format(i, arg))

    # print('源文件路径:{}'.format(src_path))
    # print('参数路径:{}'.format(items))

    values = read_files(src_path, items)
    write_dict_to_excel(values, r'C:\摘取数据.xlsx')
    print('导出成功!')
    # read_files(r'C:/Repository/Nercar/ShaGang')
