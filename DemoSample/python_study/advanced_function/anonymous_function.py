# 
#   @FileName       :anonymous_function.py
#   @author         :徐潇然
#   @description    :匿名函数 : 关键字lambda表示匿名函数，冒号前面的x表示函数参数。匿名函数有个限制，就是只能有一个表达式，不用写return，返回值就是该表达式的结果。
#   @create-time    :20:36
#   @version        :1.0
#   


def map_anonymous():
    """
    lambda x: x * x
    实际上就是:
    def f(x):
        return x * x

    this sample is map + anonymous function  :
        1. map将传入的函数依次作用到序列的每个元素
    """
    temp_list = list(map(lambda x: x * x, [1, 2, 3, 4, 5, 6, 7, 8, 9]))
    print(temp_list)


def return_a_anontmous(x, y):
    """
    返回一个匿名函数 !!! 而非数值.
    """
    return lambda: x + y


if __name__ == '__main__':
    # map_anonymous()
    # f = lambda x: x + 1
    # print(f(3))
    print(return_a_anontmous(1, 2)())
