"""

"""


####################################################
# 
#   @author:${name}
#   @description:   函数式编程: 返回函数
#   @create-time:${time}
#   @version 1.0
#   
#####################################################

def lazy_sum(*args):
    def sum_():
        ax = 0
        for n in args:
            ax = ax + n
        return ax

    return sum_


def inc():
    """
    内部函数使用外部函数参数: nonlocal
    """
    x = 0

    def fn():
        """
        原因是x作为局部变量并没有初始化，直接计算x+1是不行的。
        但我们其实是想引用inc()函数内部的x，所以需要在fn()函数内部加一个nonlocal x的声明。
        加上这个声明后，解释器把fn()的x看作外层函数的局部变量，它已经被初始化了，可以正确计算x+1。
        """
        nonlocal x
        x = x + 1
        return x

    return fn


if __name__ == '__main__':
    f = inc()
    # print(f())  # 1
    # print(f())  # 2

    f = lazy_sum(1, 3, 5, 7, 9)
    print(f())
