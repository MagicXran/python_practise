"""
#   @FileName       :common.py
#   @author         :徐潇然
#   @create-time    :2021/11/15
#   @version        :1.0
#   @description    : 将url固定长度的md5
#
#
"""
import hashlib


def get_md5(url):
    if isinstance(url, str):
        url = url.encode('utf-8')
    m = hashlib.md5()
    m.update(url)

    return m.hexdigest()


if __name__ == '__main__':
    url = "https://www.baidu.com"
    url1 = "https://www."
    print(get_md5(url))
    print(get_md5(url1))
