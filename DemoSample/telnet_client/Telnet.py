import telnetlib
import time

Delay_time = 2


class TelnetClient:
    """
    连接远程telnet 类
    """

    def __init__(self):
        self.tn = telnetlib.Telnet()  # 此函数实现telnet登录主机

    def login_host(self, host_ip, username, password, port=9110):
        """
        根据ip:port 以及远程主机账户密码,来连接对方telnet server, 连接至指定端口
        :param host_ip:
        :param username:
        :param password:
        :return:
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
                print('%s登录失败，用户名或密码错误' % host_ip)
                self.logout_host()
                return False
        except:
            print('%s网络连接失败' % host_ip)
            self.logout_host()
            return False

    def execute_some_command(self):
        """
        获取该端口数据流, 默认GBK编码
        :return:
        """

        try:
            while True:
                # self.tn.write(command.encode('gbk') + b'\n')
                # 获取命令结果
                command_result = self.tn.read_very_eager().decode('gbk')
                print(command_result)
                if len(command_result) < 10:
                    time.sleep(Delay_time)  # return command_result

        except BaseException as e:
            print("Unexpected Error in execute_some_command() : {}".format(e))
            raise e

    def logout_host(self):
        """
        断开连接
        :return:
        """
        self.tn.write(b"exit\n")
        print('断开连接...')

    def read_eager(self):
        """
        读取数据
        :return:
        """
        least_line = self.tn.read_very_eager().decode('gbk')
        return least_line

    def listening(self, ip, user, pwd, port):
        """
        监听指定端口
        :param ip:
        :param user:
        :param pwd:
        :return:
        """
        try:
            if self.login_host(ip, user, pwd, port):
                self.execute_some_command()
                # telnet_client.logout_host()
            else:
                print('连接远程主机失败!')
                self.logout_host()
        except BaseException as e:
            print("Unexpected Error in listening() : {}".format(e))
        finally:
            self.logout_host()


def listen(ip, user='administrator', pwd='Nercar505', port=9110):
    telnet_client = TelnetClient()
    telnet_client.listening(ip, user, pwd, port)


if __name__ == '__main__':
    listen('10.43.10.100', 'administrator', 'Nercar505', 9110)
