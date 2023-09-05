# 自动刷新网页
import random  # 用于生成一个随机数
import time  # 用于控制访问间隔

import requests  # 访问网页所必须用到的头文件

i = 1  # 记录下当前是第几轮（在刷新多个网页时可以看到）
count = 1  # 记录下当前总共刷新了多少次
while True:
    file = open('web.txt', 'r', encoding='utf-8', errors='ignore')
    while True:
        url = file.readline().rstrip()
        header = {"user-agent": "Mozilla/5.0"}
        try:
            data = requests.get(url=url, headers=header)
        except ValueError:
            break
        else:
            print(data.status_code, end='')
            if data.status_code == 200:
                print(f"访问{url}成功")
            else:
                print(f"访问{url}失败")
            k = random.randint(5, 10)  # 生成一个5-10s的随机数   可以自己调整
            time.sleep(k)
            count += 1
            print(f"随机数为{k}，现在是第{count}次刷新")
    file.close()
    print(f"txt文件第{i}轮刷新完毕")
    time.sleep(5)  # 防止被网页认出你是恶意刷新，当然可以修改
    i += 1
