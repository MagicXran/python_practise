import re

if __name__ == '__main__':
    line = "bobby123"
    regex_str = r"^b.*3$"  # 匹配模式: 以`b`开头,其后接任意个数任意字符,以3结尾.

    line1 = "booooobaaaaooobdhdhffbby123"
    line2 = "study in 尼玛假大学 "
    # 匹配两个b...b之间的字符串
    regex_str_greed = ".*(b.*b).*"  # 默认贪婪模式,从右向左扫描.符合的直接输出:  bdhdhffb
    regex_str1_nonGreed = ".*?(b.*?b).*"  # 非贪婪模式:从左向右匹配首次遇到的b...b的字符串: booooob

    match_obj = re.match(r'.*?([\u4E00-\u9FA5]+大学)', line2)
    # if match_obj:
    #     print(match_obj.group(1))

    import sys

    print(sys.getdefaultencoding())
