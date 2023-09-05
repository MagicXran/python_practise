import base64

from Crypto.Hash import SHA256
from Crypto.PublicKey import RSA
from Crypto.Signature import PKCS1_v1_5


def encode():
    # 导入公钥和私钥
    with open('sign') as f:
        private = f.read()

    private_key = RSA.import_key(private)

    # 对需要签名的数据生成摘要
    digest = SHA256.new()
    message = '2022-03-13'  # 这里可以修改截止时间
    digest.update(message.encode())

    # 数据进行签名
    signer = PKCS1_v1_5.new(private_key)
    code = signer.sign(digest)
    signature = base64.b64encode(code)
    print(signature.decode())
    # 软件第一次验证通过以后，就可以把这个过期时间的字符串和签名字符串一起用文件的形式存到硬盘上，每次启动软件的时候都检查一遍。
    # 发现合法并且没有过期就正常运行。发现过期了或者不合法就就重新弹出输入注册码的对话框
    return signature.decode()


def decode(pub):
    with open('sign.pub') as fp:
        public = fp.read()

    public_key = RSA.import_key(public)

    message = '2022-03-13'
    signature = pub
    digest = SHA256.new()
    digest.update(message.encode())
    reader = PKCS1_v1_5.new(public_key)
    print(reader.verify(digest, base64.b64decode(signature.encode())))


def main():
    # pub = encode()
    decode(
        'DTNtMHRCMUKT2g+vTUGg5z+zJHvmlnlq2pP5GPrS3EYyWkPe+8vWCKe4VbkUxJfKIew+p/ZY/LWqqxyE9ONf+bpPFLQYwu+uJ9HVv9h0Ik9XSrWdTn4r94dxHSvoCco98eqdjszwFroF6epnWFCmN5SaJ8KBVFmH55RdTkCgSyfdDZ9pD+CFkhWbUubU8BLCNHnW13InkQXHwasIkJgafRuKxLQqstpbG31EkP/VPeITSnqk5uzxCn7dLzmWeT3Wu7u3FkI1gNKQzNouGROYvMgOWaENqFR3EwcO04s/qLQwP/4o0A1V+fYpilaxjZdw2EX2wSkeKj+jE3RVOLon37v26l5nttWWK0DW3cqal1anUChFhfJ2apZWMo1wZ8h6ZVQ70IXZ7nZAgW0gtXmgEuyBtSwFUwWICLgHwsaa/rCsBSBSvnLHhJ3/dZarYJ+7nAQdmcGklAXSmmaGFY96rapw/oPXuiAX2Typem1Hjm5knwyyhHrXHBGCzA1IMZBS')
    pass


if __name__ == '__main__':
    main()
