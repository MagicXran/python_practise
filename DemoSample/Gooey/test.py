"""
##########################################################
#   @FileName       :test.py
#   @author         :徐潇然
#   @create-time    :2021/11/13
#   @version        :1.0
#   @description    : 练习gooey
#
#
##########################################################  
"""

from gooey import Gooey, GooeyParser


@Gooey(program_name="示例")
def test1():
    parser = GooeyParser(description="第一个示例!")
    parser.add_argument('文件路径', widget="FileChooser")  # 文件选择框
    parser.add_argument('日期', widget="DateChooser")  # 日期选择框
    args = parser.parse_args()  # 接收界面传递的参数
    print(args)


# @Gooey(navigation='TABBED',
#        show_sidebar=True
#        )
def tabbed_layout():
    parser = GooeyParser(description='第一个示例')
    parser.add_argument('日期', widget='TextField')
    args = parser.parse_args()
    print(args, flush=True)
    pass


# @Gooey(
#     richtext_controls=True,  # 打开终端对颜色支持
#     program_name="MQTT连接订阅小工具",  # 程序名称
#     encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
#     progress_regex=r"^progress: (\d+)%$"  # 正则，用于模式化运行时进度信息  ,
#
# )
def MQITT():
    settings_msg = 'MQTT device activation information subscription'
    parser = GooeyParser(description=settings_msg)

    subs = parser.add_subparsers(help='commands', dest='command')

    my_wechat = subs.add_parser('微信轰炸')
    my_wechat.add_argument('boom', metavar='微信轰炸', )

    my_cool_parser = subs.add_parser('MQTT消息订阅')
    my_cool_parser.add_argument("connect", metavar='运行环境', help="请选择开发环境",
                                choices=['dev环境', 'staging环境'],
                                default='dev环境')
    my_cool_parser.add_argument("device_type", metavar='设备类型', help="请选择设备类型", choices=['H1', 'H3'],
                                default='H1')
    my_cool_parser.add_argument("serialNumber", metavar='设备SN号', default='LKVC19060047',
                                help='多个请用逗号或空格隔开')

    siege_parser = subs.add_parser('进度条控制')
    siege_parser.add_argument('num', help='请输入数字', default=100, metavar='数字')

    args = parser.parse_args()
    print(args, flush=True)  # 坑点：flush=True在打包的时候会用到
    print(type(args))
    # datas:argparse.Namespace =
    print(args.serialNumber)
    # print(args.)


@Gooey(
    richtext_controls=True,  # 打开终端对颜色支持
    program_name="WeChar Boom!",  # 程序名称
    encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
    progress_regex=r"^progress: (\d+)%$"  # 正则，用于模式化运行时进度信息  ,

)
def wechat_boom():
    parser = GooeyParser(description='settings_msg')
    parser.add_argument('textInput1', metavar='请输入是否轰炸', default=2, widget="TextField")

    args = parser.parse_args()
    print(args, flush=True)
    if args.textInput1 == 'y':
        from wechatTools import revoke_message
        revoke_message.test_send()


if __name__ == '__main__':
    # main()
    # tabbed_layout()
    # MQITT()
    # wechat_boom()
    test1()
