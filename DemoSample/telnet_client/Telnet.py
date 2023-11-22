import telnetlib
import time

import winsound  # 蜂鸣器

Delay_time = 1
''' 延迟'''

filter_len = 5
''' 过滤长度 '''

# 定义亮晶晶的音节
notes = [262, 294, 330, 349, 392, 440, 494, 523]
# 定义音符频率
C4 = 262
D4 = 294
E4 = 330
F4 = 349
G4 = 392
A4 = 440
B4 = 494

# 定义音符时长
# 设置音节的持续时间和暂停时间
duration = 100  # 持续时间（毫秒）
pause = 50  # 暂停时间（毫秒）


# 播放音符
def play_note(note, duration):
    winsound.Beep(note, duration)


# 播放乐曲
def play_music():
    # 祝你生日快乐
    play_note(G4, duration)
    # play_note(G4, duration)
    # play_note(G4, duration)
    # play_note(A4, duration)
    # play_note(A4, duration)
    # play_note(B4, h)
    # play_note(G4, q)
    # play_note(G4, q)
    # play_note(A4, q)
    # play_note(A4, q)
    # play_note(C4, h)
    # play_note(G4, q)
    # play_note(G4, q)
    # play_note(G4, en)
    # play_note(F4, en)
    # play_note(F4, en)
    # play_note(E4, en)
    # play_note(E4, en)
    # play_note(D4, en)
    # play_note(D4, en)
    # play_note(C4, en)
    # play_note(C4, en)


def beep():
    # 调用电脑蜂鸣器
    winsound.Beep(1899, 1000)
    # 循环播放音节
    # for note in notes:
    # winsound.Beep(note, duration)
    # 播放生日快乐曲目
    # play_music()
    # time.sleep(pause / 1000)  # 将毫秒转换为秒


# Demo how to use argparse
from argparse import ArgumentParser, SUPPRESS


def build_argparser():
    """
    命令行帮助
    Returns
    -------

    """
    parser = ArgumentParser(add_help=False)

    args = parser.add_argument_group("Options")
    args.add_argument('-h', '--help', action='help', default=SUPPRESS, help='显示如何使用')

    args.add_argument("-i", "--ip", help="指定监控的ip,字符串类型",
                      required=True, type=str)

    args.add_argument("-p", "--port", help="指定监控的ip的端口号，数值类型， 默认为9110",
                      default=9110, type=int)

    args.add_argument("-u", "--user", help="远程主机的用户名，,字符串类型，默认为 administrator",
                      default='administrator', type=str)

    args.add_argument("-w", "--pwd", help="远程主机的密码,字符串类型,默认为 Nercar505",
                      default='Nercar505', type=str)

    args.add_argument("-a", "--arg", help="输入要监控的文本字样 可以是多个,用空格分割,默认是 '[ERR]'",
                      default='[ERR]', type=str, nargs='+')

    args.add_argument("-s", "--show", help="是否显示匹配的文本, 取值： True or False ,默认为False",
                      default=False, type=bool)

    return parser


isShow = False


def analyse_info(args: list, txt):
    """
    读取数据,分析报警
    Parameters
    ----------
    args
    txt

    Returns
    -------

    """
    # least_line = self.tn.read_very_eager().decode('gbk')
    try:
        all_exist = 0
        for ele in args:
            if ele in txt:
                all_exist = all_exist + 1

        if all_exist == len(args):
            beep()
            if isShow:
                print(txt)
    except Exception as e:
        print(f'发生异常：{e}')


pass


class TelnetClient:
    """
    连接远程telnet 类
    """

    def __init__(self):
        self.tn = telnetlib.Telnet()  # 此函数实现telnet登录主机

    def login_host(self, host_ip, username, password, port=9110):
        """
        根据ip:port 以及远程主机账户密码,来连接对方telnet server, 连接至指定端口
        Parameters
        ----------
        host_ip
        username
        password
        port

        Returns
        -------

        """
        try:
            # self.tn = telnetlib.Telnet(host_ip,port=23)
            self.tn.open(host_ip, port)
            # 等待login出现后输入用户名，最多等待10秒
            # self.tn.read_until(b'login: ', timeout=10)
            self.tn.write(username.encode('GBK') + b'\n')
            # 等待Password出现后输入用户名，最多等待10秒
            # self.tn.read_until(b'Password: ', timeout=10)
            self.tn.write(password.encode('GBK') + b'\n')
            # 延时两秒再收取返回结果，给服务端足够响应时间
            time.sleep(Delay_time)
            # 获取登录结果
            # read_very_eager()获取到的是的是上次获取之后本次获取之前的所有输出
            command_result = self.tn.read_very_eager().decode('GBK')
            if 'Login incorrect' not in command_result:
                print('%s 登录成功' % host_ip)
                return True
            else:
                print('%s 登录失败，用户名或密码错误' % host_ip)
                self.logout_host()
                return False
        except BaseException as e:
            print('%s 网络连接失败' % host_ip)
            self.logout_host()
            return False

    def execute_some_command(self, args: list):
        """
        获取该端口数据流, 默认GBK编码
        :return:
        """

        print("正在监听...")
        try:
            while True:
                # 获取命令结果
                command_result = self.tn.read_very_eager().decode('gbk')
                if len(command_result) < filter_len:
                    time.sleep(Delay_time)  # return command_result
                else:
                    analyse_info(args, command_result)

        except BaseException as e:
            print("Unexpected Error in execute_some_command() : {}".format(e))
            raise

    def logout_host(self):
        """
        断开连接
        :return:
        """
        self.tn.write(b"exit\n")
        self.tn.close()
        print('断开连接...')

    def listening(self, ip, user, pwd, port, args: list):
        """
        监听指定端口
        :param ip:
        :param user:
        :param pwd:
        :return:
        """
        try:
            if self.login_host(ip, user, pwd, port):
                self.execute_some_command(args)
                # telnet_client.logout_host()
            else:
                print('连接远程主机失败!')
                self.logout_host()
        except BaseException as e:
            print("Unexpected Error in listening() : {}".format(e))
            self.logout_host()


def listen(ip, user='administrator', pwd='Nercar505', port=9110, args=None):
    if args is None:
        args = ['ERR']
    telnet_client = TelnetClient()
    telnet_client.listening(ip, user, pwd, port, args)


if __name__ == '__main__':
    # isShow = True
    # listen('192.168.2.34', 'administrator', 'Sh0ugang', 9110, ['[ERR]', 'OpcPro'])

    args = build_argparser().parse_args()
    print(f"Type of args:{type(args)}")
    print(f"args.ip: {args.ip}")
    print(f"args.port: {args.port}")
    print(f"args.user: {args.user}")
    print(f"args.pwd: {args.pwd}")
    print(f"args.arg: {args.arg}, {type(args.arg)}")
    print(f"args.show: {args.show}")
    isShow = args.show
    listen(args.ip, args.user, args.pwd, args.port, args.arg)
