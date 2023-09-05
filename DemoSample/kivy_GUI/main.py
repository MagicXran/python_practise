import os

# os.environ['KIVY_NO_FILELOG'] = '1'  # 若设定了此环境变量，日志将不再输出到文件内。也可以在 "C:\Users\徐潇然\.kivy\config.ini" 中修改.
os.environ["KCFG_KIVY_LOG_LEVEL"] = "info"
import kivy
from kivy.clock import Clock

# during import it will map it to:
# Config.set("kivy", "log_level", "warning")

kivy.require('2.0.0')  # 注意要把这个版本号改变成你现有的Kivy版本号!

from kivy.app import App  # 译者注：这里就是从kivy.app包里面导入App类
from kivy.uix.label import Label  # 译者注：这里是从kivy.uix.label包中导入Label控件，这里都注意开头字母要大写
from kivy.uix.gridlayout import GridLayout  # 导入了一种名为Gridlayout的布局,作为根控件
from kivy.uix.textinput import TextInput  # 文本控件
from kivy.uix.button import Button


class LoginScreen(GridLayout):

    def __init__(self, **kwargs):
        """
        重新定义了初始化方法init()，这样来增加一些控件，并且定义了这些控件的行为：
        :param kwargs:
        """
        # super().__init__(**kwargs)
        super(LoginScreen, self).__init__(**kwargs)
        self.cols = 2
        # self.rows = 2
        self.add_widget(Label(text='User Name'))
        self.username = TextInput(multiline=False)
        self.add_widget(self.username)
        self.add_widget(Label(text='password'))
        self.password = TextInput(password=True, multiline=False)  # 作为密码输入框
        self.add_widget(self.password)
        self.btn = Button(text='Hello world', font_size=14)
        self.add_widget(self.btn)


class MyApp(App):

    def build(self):  # 译者注：这里是实现build()方法
        # 这个Label就是咱们这个应用的根控件了。
        timer_clock = Clock
        # timer_clock.max_iteration = 2
        # 间隔回调
        # timer_clock.schedule_interval(my_callback, 0)
        timer_clock.schedule_once(my_callback, 0)

        return LoginScreen()


def my_callback(dt):
    print('My callback is called !')
    # Clock.schedule_once(my_callback, 1)


if __name__ == '__main__':
    MyApp().run()  # 译者注：这里就是运行了。
'''
译者注：这一段的额外添加的备注是给萌新的.
就是要告诉萌新们，一定要每一句每一个函数每一个变量甚至每一个符号，都要读得懂！！！
如果是半懂不懂的状态，一定得学透了，要不然以后早晚得补课.
这时候又让我想起了结构化学。
总之更详细的内容后面会有，大家加油。
'''
