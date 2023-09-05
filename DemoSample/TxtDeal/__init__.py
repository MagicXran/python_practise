"""
用于文本处理
"""
import  re
import  logging

logging.basicConfig(level=logging.CRITICAL,format=' %(asctime)s -%(levelname)s - %(message)s')

def main():
    with open('text.txt', 'r', encoding='utf-8') as f:
        datas = f.readlines()

    with open('text.txt', 'r', encoding='utf-8') as fw:
        for line in datas:
            # fw.write(line.replace('', ''))
            print(line)


if __name__ == '__main__':
    # main()
    text = '今天是：11/28/2018'
    datepat = re.compile(r'(\d+)(/)(\d+)/(\d+)')
    # print(datepat.sub(r'\3-\1-\2', text))
    print(datepat.sub(r'\4_\1-\3', text))
    print(text)