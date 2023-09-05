import xlwings as xw


class WorkBookOpr(object):
    def __init__(self, path=None):
        self.app = xw.Book(path).save()
        # 新建工作簿 (如果不接下一条代码的话，Excel只会一闪而过，卖个萌就走了）
        self.wb = self.app.books.add()


def main():
    app = xw.App(visible=False, add_book=False)
    wb1 = app.books.open('mes.xlsx')
    wb2 = app.books.open('sccxlsx.xlsx')
    shtmes = wb1.sheets(1)
    shtscc = wb2.sheets(1)
    list_mes = shtmes.range('B2:B2994').value
    list_scc = shtscc.range('B2:B3049').value
    # list_mesthk = shtmes.range('M2:M2994').value
    list_mesthk = shtmes.range('I2:I2994').value
    list_sccthk = shtscc.range('C2:C3049').value
    count = 0
    unequilty = 0

    # print(list_scc)

    for i in range(0, len(list_mes)):
        for j in range(0, len(list_scc)):
            if list_mes[i] == list_scc[j]:
                count = count + 1
                if list_mesthk[i] == list_sccthk[j]:
                    pass
                else:
                    unequilty = unequilty + 1
                    print("{0},{1},{2}".format(list_mes[i], list_mesthk[i], list_sccthk[j]))
                    # list_mesthk[i] = list_sccthk[j]

    print("{0},{1}".format(count, unequilty))
    # shtmes.range('M2:M5').options(transpose=True).value = [1, 2, 3, 4, 5]
    # shtmes.range('I2:I2996').options(transpose=True).value = list_mesthk
    # shtmes.range('M2:M2994').options(transpose=True).value = list_mesthk
    # print(list_mesthk)
    wb1.save()
    print("This is ok")
    app.quit()


if __name__ == '__main__':
    main()

# Apps = xw.apps
# count = xw.apps.count
# keys = xw.apps.keys()
# print("{0},{1},{2}".format(Apps, count, keys))
#
# # app = xw.App()
# # pid = app.pid
# # # app = xw.apps[1668]
# # # app = xw.apps.active
# #
# # print("{0}, {1}, {2}".format(app, pid, 00))
