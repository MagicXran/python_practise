import time

import itchat
import requests
from bs4 import BeautifulSoup


def test_send():
    itchat.auto_login(hotReload=True)
    friends = itchat.search_friends(name='老婆')
    print(friends[0])
    print(friends[0].get('UserName'))
    count = 1
    while True:
        itchat.send("老婆先你了，不生气了好么~{}".format('宝贝'), toUserName=friends[0].get('UserName'))
        print("发送成功!")
        time.sleep(1)
        count = count + 1  # print(friends)

    # itchat.auto_login()
    # itchat.send('Hello, filehelper', toUserName='🐠')
    # friends = itchat.get_friends()  # 好友列表 返回一个list
    # # groups = itchat.get_chatrooms()
    # # print(friends)
    # # count = 1
    # # for i in friends:
    # #     if i['RemarkName'] == '老婆':
    #         while True:
    #             itchat.send("{}只羊".format(count), toUserName=i['UserName'])
    #             itchat.msg_register()
    #             print(i['NickName'])
    #             print("发送成功!")
    #             time.sleep(0.5)
    #             count = count + 1  # print(friends)


# 返回多条语录
def findLoveWord():
    url = "https://www.1juzi.com/new/150542.html"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE",
    }

    content = requests.get(url, headers=headers, verify=False).content.decode("gb2312", errors="ignore")
    soup = BeautifulSoup(content, 'html.parser')
    contentDocument = soup.find(class_="content").find_all("p")[:50]
    loveList = []
    for dom in contentDocument:
        domString = dom.string
        domString = domString[domString.index("、") + 1:]
        loveList.append(domString)

    return loveList


if __name__ == '__main__':
    # test_send()

    findLoveWord()
